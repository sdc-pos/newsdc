VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PLN00501 
   Caption         =   "[商品化計画システム]商品化予定照会画面"
   ClientHeight    =   10992
   ClientLeft      =   2028
   ClientTop       =   -4476
   ClientWidth     =   15468
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
   ScaleHeight     =   10992
   ScaleWidth      =   15468
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   372
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表　示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
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
      ToolTipText     =   "商品化構成を読み込みます（Ｆ5）"
      Top             =   0
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7695
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   15135
      _ExtentX        =   26691
      _ExtentY        =   13568
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "品番／商品化予定日"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   720
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3493"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3387"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1566"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1461"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(27)  =   ":id=68,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(28)  =   ":id=68,.fontname=ＭＳ ゴシック"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=102,.parent=67,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=68"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=69"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=71"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=98,.parent=67,.alignment=1"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=71"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      ToolTipText     =   "商品化構成を保存します"
      Top             =   0
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid2 
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   8880
      Width           =   15135
      _ExtentX        =   26691
      _ExtentY        =   2773
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   720
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3493"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3387"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8194"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1566"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1461"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(27)  =   ":id=68,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(28)  =   ":id=68,.fontname=ＭＳ ゴシック"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=102,.parent=67,.alignment=1,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=68"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=69"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=71"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=98,.parent=67,.alignment=1"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=71"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      Caption         =   "から2週間"
      Height          =   252
      Index           =   2
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "商品化予定日"
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   1212
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "読込"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "登録"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   2
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "PLN00501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const ptxStart_Day% = 0


Dim PLN_S_YOTEI         As New XArrayDB
Dim PLN_S_YOTEI_G       As New XArrayDB



Private Const Min_Row% = 1              '最小行数
Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Max_Col As Long                 '最大列数

Private Const colHIN_GAI% = 0           '担当者名
Private Const colYOTEI_QTY% = 1         '商品化予定数

Private List_Week   As Long             '表示するn週間
Private KADOU_RITU  As Double           '稼働率
Private Day_KOUSU   As Double           '１日当たりの時間数(H)


Private Type S_YOTEI_Tbl_tag
    HIN_GAI         As String * 20
    YOTEI_QTY()     As Long
End Type
Private S_YOTEI_Tbl_flg As Boolean




Private Const LAST_UPDATE_DAY$ = "[PLN0050] 2011.11.10 09:30"

Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '読込み



            '取込みﾃﾞｰﾀ表示
            If List_Disp_Proc() Then
                Unload Me
            End If



        Case 1          '終了

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
        "[商品化計画システム]]商品化予定照会画面", Me.hwnd, 0)
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
    
                                '１日当たりの稼働時間(H)取り込み
    If GetIni(App.EXEName, "Day_KOUSU", App.EXEName, c) Then
        Day_KOUSU = 7#
    Else
        If Not IsNumeric(Trim(c)) Then
            Day_KOUSU = 7#
        Else
            If Val(Trim(c)) < 1 Then
                List_Week = 1
            Else
                Day_KOUSU = Val(Trim(c))
            End If
        End If
    End If
    
    
    
    
    Call List_Make_Proc


                                '稼働率取り込み
    If GetIni(App.EXEName, "KADOU_RITU", App.EXEName, c) Then
        KADOU_RITU = 100
    Else
        If Not IsNumeric(Trim(c)) Then
            KADOU_RITU = 100
        Else
            If Val(Trim(c)) < 0 Then
                KADOU_RITU = 100
            Else
                KADOU_RITU = Val(Trim(c))
            End If
        End If
    End If








    PLN00501.Caption = PLN00501.Caption & " " & LAST_UPDATE_DAY
    
    Label1(2).Caption = "から" & List_Week & "週間"
    Text1(ptxStart_Day).Text = Format(Now, "YYYY/MM/DD")


                                
                                
                                
                                '担当者別勤務時間データＯＰＥＮ
    If PLN_O_HOURS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化予定ファイルＯＰＥＮ
    If PLN_S_YOTEI_Open(BtOpenNomal) Then
        Unload Me
    End If





End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), K0_PLN_O_HOURS, Len(K0_PLN_O_HOURS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者別勤務時間データ")
        End If
    End If
    
    sts = BTRV(BtOpClose, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化予定ファイル")
        End If
    End If
    
    
    
    
    sts = BTRV(BtOpReset, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), K0_PLN_O_HOURS, Len(K0_PLN_O_HOURS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    
    Set PLN00501 = Nothing



    End

End Sub

Private Sub SHORI_Click(Index As Integer)

    Select Case Index
    
        Case 0
            Command1(0).Value = True
        Case 1
            Command1(1).Value = True
    End Select



End Sub
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「商品化予定データ」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer


Dim Row             As Long
Dim List_Day        As Long
Dim i               As Long
Dim j               As Long
Dim wkday           As Long


Dim Ing_Date        As String
Dim Start_Date      As String
Dim End_Date        As String

Dim Yobi            As Integer
Dim Yobi_NAME       As String

Dim S_YOTEI_Tbl()   As S_YOTEI_Tbl_tag

Dim S_JIKAN()       As Double
Dim S_KOUSU()       As Double
Dim O_Time()        As Double
Dim O_KOUSU()       As Double

    List_Disp_Proc = True


    If Not IsDate(Text1(ptxStart_Day).Text) Then
        MsgBox "入力した項目はエラーです。(商品化予定日)"
        Text1(ptxStart_Day).SetFocus
        List_Disp_Proc = False
        Exit Function
    End If







    Call Input_Lock



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定照会　[検索]処理開始！！", Me.hwnd, 0)


    
    List_Day = List_Week * 7
    Start_Date = Text1(ptxStart_Day).Text
    End_Date = DateAdd("d", List_Day - 1, Start_Date)


    For i = 0 To List_Day - 1
        
        
        Ing_Date = DateAdd("d", i, Start_Date)
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
                
                
        
        TDBGrid1.Columns(i + 1).Caption = Right(Format(Ing_Date, "YYYY/MM/DD"), 5) & "   " & Yobi_NAME
        TDBGrid2.Columns(i + 1).Caption = Right(Format(Ing_Date, "YYYY/MM/DD"), 5) & "   " & Yobi_NAME

    Next i

    
    Call UniCode_Conv(K2_PLN_S_YOTEI.YOTEI_DT, Format(Start_Date, "YYYYMMDD"))
    Call UniCode_Conv(K2_PLN_S_YOTEI.JGYOBU, "")
    Call UniCode_Conv(K2_PLN_S_YOTEI.NAIGAI, "")
    Call UniCode_Conv(K2_PLN_S_YOTEI.HIN_GAI, "")
    
    
    i = -1
    com = BtOpGetGreater
    
    S_YOTEI_Tbl_flg = False
    
    Do
        
        DoEvents
        
        sts = BTRV(com, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K2_PLN_S_YOTEI, Len(K2_PLN_S_YOTEI), 2)
        Select Case sts
            Case BtNoErr
            
                If StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) > Format(End_Date, "YYYYMMDD") Then
                    Exit Do
                End If
            
                wkday = DateDiff("d", Start_Date, Mid(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode), 7, 2))
                            
                            
                            
                If i = -1 Then
                    S_YOTEI_Tbl_flg = True
                    i = i + 1
                    ReDim S_YOTEI_Tbl(0 To 0)
                    ReDim S_YOTEI_Tbl(i).YOTEI_QTY(0 To List_Day - 1)
                    
                    For j = 0 To UBound(S_YOTEI_Tbl(i).YOTEI_QTY)
                        S_YOTEI_Tbl(i).YOTEI_QTY(j) = 0
                    Next j
                    
                    
                    S_YOTEI_Tbl(i).HIN_GAI = StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)
                    
                    
                    
                    
                Else
                
                    For i = 0 To UBound(S_YOTEI_Tbl)
                        If S_YOTEI_Tbl(i).HIN_GAI = StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode) Then
                            Exit For
                        End If
                    Next i
                
                    If i > UBound(S_YOTEI_Tbl) Then
                        
                        ReDim Preserve S_YOTEI_Tbl(0 To i)
                        ReDim S_YOTEI_Tbl(i).YOTEI_QTY(0 To List_Day - 1)
                        
                        For j = 0 To UBound(S_YOTEI_Tbl(i).YOTEI_QTY)
                            S_YOTEI_Tbl(i).YOTEI_QTY(j) = 0
                        Next j
                        
                        S_YOTEI_Tbl(i).HIN_GAI = StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)
                
                
                
                    End If
                
                
                
                    
                
                End If
            
            
                S_YOTEI_Tbl(i).YOTEI_QTY(wkday) = S_YOTEI_Tbl(i).YOTEI_QTY(wkday) + CLng(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode))
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "商品化予定データ")
                Call Input_UnLock
                Exit Function
        End Select
        
        com = BtOpGetNext
        
    Loop




    Set PLN_S_YOTEI = Nothing
    Row = Min_Row - 1
    
    
        
    If S_YOTEI_Tbl_flg Then
        
        
        For i = 0 To UBound(S_YOTEI_Tbl)
        
            Row = Row + 1
            PLN_S_YOTEI.ReDim Min_Row, Row, Min_Col, Max_Col
        
                    
            PLN_S_YOTEI(Row, colHIN_GAI) = S_YOTEI_Tbl(i).HIN_GAI
            
        
            For j = 0 To UBound(S_YOTEI_Tbl(i).YOTEI_QTY)
                
                PLN_S_YOTEI(Row, colYOTEI_QTY + j) = Format(S_YOTEI_Tbl(i).YOTEI_QTY(j), "#,##0")
            
            Next j
        
        Next i


        Set TDBGrid1.Array = PLN_S_YOTEI
        TDBGrid1.ReBind
        
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If
    
    
    
    
    
    '-------------------------------------------------------------  合計値集計
    ReDim S_JIKAN(0 To List_Day - 1)
    ReDim S_KOUSU(0 To List_Day - 1)
    ReDim O_Time(0 To List_Day - 1)
    ReDim O_KOUSU(0 To List_Day - 1)
    
    For i = 0 To UBound(S_JIKAN)
        S_JIKAN(i) = 0
        S_KOUSU(i) = 0
        O_Time(i) = 0
        O_KOUSU(i) = 0
    Next i

    Call UniCode_Conv(K2_PLN_S_YOTEI.YOTEI_DT, Format(Start_Date, "YYYYMMDD"))
    Call UniCode_Conv(K2_PLN_S_YOTEI.JGYOBU, "")
    Call UniCode_Conv(K2_PLN_S_YOTEI.NAIGAI, "")
    Call UniCode_Conv(K2_PLN_S_YOTEI.HIN_GAI, "")
    
    
    com = BtOpGetGreater
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K2_PLN_S_YOTEI, Len(K2_PLN_S_YOTEI), 2)
        Select Case sts
            Case BtNoErr
            
                If StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) > Format(End_Date, "YYYYMMDD") Then
                    Exit Do
                End If
            
                wkday = DateDiff("d", Start_Date, Mid(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode), 7, 2))
                            
                            
            
            
                S_JIKAN(wkday) = S_JIKAN(wkday) + CDbl(StrConv(PLN_S_YOTEI_R.S_JIKAN, vbUnicode))
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "商品化予定データ")
                Call Input_UnLock
                Exit Function
        End Select
        
        com = BtOpGetNext
        
    Loop



    '-----------------------    H-->M 2011.11.09
'    For i = 0 To UBound(S_JIKAN)
'
'        S_JIKAN(i) = Round(S_JIKAN(i) / 60, 2)
'
'    Next i
'
'
'    For i = 0 To UBound(S_JIKAN)
'
'        S_KOUSU(i) = Round(S_JIKAN(i) / Day_KOUSU, 2)
'
'    Next i
    
    
    For i = 0 To UBound(S_JIKAN)

        S_JIKAN(i) = Round(S_JIKAN(i), 2)

    Next i


    For i = 0 To UBound(S_JIKAN)

        S_KOUSU(i) = Round((S_JIKAN(i) / 60) / Day_KOUSU, 2)

    Next i
    
    
    '-----------------------    H-->M 2011.11.09




    Call UniCode_Conv(K1_PLN_O_HOURS.O_DATE, Format(Start_Date, "YYYYMMDD"))
    Call UniCode_Conv(K1_PLN_O_HOURS.TANTO_CODE, "")
    
    
    com = BtOpGetGreater
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), K1_PLN_O_HOURS, Len(K1_PLN_O_HOURS), 1)
        Select Case sts
            Case BtNoErr
            
                If StrConv(PLN_O_HOURS_REC.O_DATE, vbUnicode) > Format(End_Date, "YYYYMMDD") Then
                    Exit Do
                End If
            
                wkday = DateDiff("d", Start_Date, Mid(StrConv(PLN_O_HOURS_REC.O_DATE, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(PLN_O_HOURS_REC.O_DATE, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(PLN_O_HOURS_REC.O_DATE, vbUnicode), 7, 2))
                            
                            
            
                If IsNumeric(StrConv(PLN_O_HOURS_REC.O_Time, vbUnicode)) Then
                    O_Time(wkday) = O_Time(wkday) + CDbl(StrConv(PLN_O_HOURS_REC.O_Time, vbUnicode))
                Else
                    O_Time(wkday) = 0
                End If
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "担当者別勤務時間データ")
                Call Input_UnLock
                Exit Function
        End Select
        
        com = BtOpGetNext
        
    Loop

    '-----------------------    H-->M 2011.11.09
'    For i = 0 To UBound(O_Time)
'
'        O_Time(i) = Round(O_Time(i) * (KADOU_RITU / 100), 2)
'
'    Next i
'
'
'    For i = 0 To UBound(O_KOUSU)
'
'        O_KOUSU(i) = Round(Round((O_Time(i) / Day_KOUSU), 2) * (KADOU_RITU / 100), 2)
'
'    Next i

    For i = 0 To UBound(O_Time)

        O_Time(i) = Round((O_Time(i) * 60) * (KADOU_RITU / 100), 2)

    Next i


    For i = 0 To UBound(O_KOUSU)

        O_KOUSU(i) = Round(Round(((O_Time(i) / 60) / Day_KOUSU), 2) * (KADOU_RITU / 100), 2)

    Next i


    '-----------------------    H-->M 2011.11.09


    Set PLN_S_YOTEI_G = Nothing
    
    PLN_S_YOTEI_G.ReDim Min_Row, 1, Min_Col, Max_Col
    PLN_S_YOTEI_G(1, colHIN_GAI) = "工数(分)"           '2011.11.09 Ｈ>>分
    For i = 0 To UBound(S_JIKAN)
        PLN_S_YOTEI_G(1, colYOTEI_QTY + i) = Format(S_JIKAN(i), "#,##0.0")
    Next i
    
    PLN_S_YOTEI_G.ReDim Min_Row, 2, Min_Col, Max_Col
    PLN_S_YOTEI_G(2, colHIN_GAI) = "工数(人)"
    For i = 0 To UBound(S_KOUSU)
        PLN_S_YOTEI_G(2, colYOTEI_QTY + i) = Format(S_KOUSU(i), "#,##0.0")
    Next i
    
    PLN_S_YOTEI_G.ReDim Min_Row, 3, Min_Col, Max_Col
    PLN_S_YOTEI_G(3, colHIN_GAI) = "出勤予定(分)"       '2011.11.09 Ｈ>>分
    For i = 0 To UBound(O_Time)
        PLN_S_YOTEI_G(3, colYOTEI_QTY + i) = Format(O_Time(i), "#,##0.0")
    Next i
    
    PLN_S_YOTEI_G.ReDim Min_Row, 4, Min_Col, Max_Col
    PLN_S_YOTEI_G(4, colHIN_GAI) = "出勤予定(人)"
    For i = 0 To UBound(O_KOUSU)
        PLN_S_YOTEI_G(4, colYOTEI_QTY + i) = Format(O_KOUSU(i), "#,##0.0")
    Next i
    
    

    Set TDBGrid2.Array = PLN_S_YOTEI_G
    TDBGrid2.ReBind
    
    TDBGrid2.Update
    TDBGrid2.MoveFirst

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定照会　[検索]処理終了！！", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_Proc = False

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    PLN00501.MousePointer = vbHourglass

    Call Ctrl_Lock(PLN00501)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PLN00501)


    PLN00501.MousePointer = vbDefault

End Sub


Private Sub List_Make_Proc()
'----------------------------------------------------------------------------
'                   Grid行作成
'----------------------------------------------------------------------------

Dim i       As Long
Dim NISUU   As Long


Dim TColumn     As TrueDBGrid80.Column
 



    

    NISUU = List_Week * 7


    For i = 2 To NISUU
        
        Set TColumn = TDBGrid1.Columns.Add(i)
        With TColumn
            .Visible = True
            .Caption = ""
            .Width = TDBGrid1.Columns(1).Width
            .Font.NAME = "ＭＳ ゴシック"
            .Font.Size = 9
            .HeadFont.NAME = "ＭＳ ゴシック"
            .HeadFont.Size = 9
            .Alignment = dbgRight
        End With
    
    
        Set TColumn = TDBGrid2.Columns.Add(i)
        With TColumn
            .Visible = True
            .Caption = ""
            .Width = TDBGrid2.Columns(1).Width
            .Font.NAME = "ＭＳ ゴシック"
            .Font.Size = 9
            .HeadFont.NAME = "ＭＳ ゴシック"
            .HeadFont.Size = 9
            .Alignment = dbgRight
        End With
    
    Next i

    Max_Col = NISUU + 1

End Sub


