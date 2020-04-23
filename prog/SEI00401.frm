VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEI00401 
   Caption         =   "[請求システム]輸送箱請求書作成処理 ([SEI0040] 2016.10.26 13:15)"
   ClientHeight    =   11145
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   16020
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
   ScaleHeight     =   11145
   ScaleWidth      =   16020
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text2 
      Height          =   360
      Index           =   1
      Left            =   9240
      TabIndex        =   14
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Index           =   0
      Left            =   9240
      TabIndex        =   13
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "明  細"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4410
      TabIndex        =   4
      Top             =   120
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   10185
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3150
      TabIndex        =   1
      Top             =   960
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1470
      TabIndex        =   0
      Top             =   960
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表  紙"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2310
      TabIndex        =   3
      Top             =   120
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   7
      Top             =   10320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終  了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   6510
      TabIndex        =   5
      Top             =   120
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表  示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1905
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7935
      Left            =   315
      TabIndex        =   6
      Top             =   1800
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   13996
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "売上日付"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "品番"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "品名"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "才数"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "単価"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "出荷先"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "使用枚数"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "金額"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "才数"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2249"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2117"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1958"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1826"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=4207"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4075"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1535"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1402"
      Splits(0)._ColumnProps(16)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=2646"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2514"
      Splits(0)._ColumnProps(21)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=4154"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=4022"
      Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(27)=   "Column(6).Width=2461"
      Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2328"
      Splits(0)._ColumnProps(30)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(32)=   "Column(7).Width=3281"
      Splits(0)._ColumnProps(33)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(7)._WidthInPix=3149"
      Splits(0)._ColumnProps(35)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=2593"
      Splits(0)._ColumnProps(40)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(41)=   "Column(8).Order=9"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "請求明細"
      AllowArrows     =   0   'False
      MultipleLines   =   0
      EmptyRows       =   -1  'True
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ Ｐゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=20,.parent=87"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=17,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=18,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=19,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=87"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=106,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=103,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=104,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=105,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=110,.parent=87,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=107,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=108,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=109,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=114,.parent=87,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=111,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=112,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=113,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=24,.parent=87"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=118,.parent=87,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=115,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=116,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=117,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=16,.parent=87,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=13,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=14,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=15,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=28,.parent=87,.alignment=1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=91"
      _StyleDefs(72)  =   "Named:id=33:Normal"
      _StyleDefs(73)  =   ":id=33,.parent=0"
      _StyleDefs(74)  =   "Named:id=34:Heading"
      _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(76)  =   ":id=34,.wraptext=-1"
      _StyleDefs(77)  =   "Named:id=35:Footing"
      _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(79)  =   "Named:id=36:Selected"
      _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=37:Caption"
      _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(83)  =   "Named:id=38:HighlightRow"
      _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=39:EvenRow"
      _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(87)  =   "Named:id=40:OddRow"
      _StyleDefs(88)  =   ":id=40,.parent=33"
      _StyleDefs(89)  =   "Named:id=41:RecordSelector"
      _StyleDefs(90)  =   ":id=41,.parent=34"
      _StyleDefs(91)  =   "Named:id=42:FilterBar"
      _StyleDefs(92)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "～"
      Height          =   375
      Index           =   8
      Left            =   2835
      TabIndex        =   9
      Top             =   960
      Width           =   330
   End
   Begin VB.Label Label1 
      Caption         =   "日付範囲"
      Height          =   375
      Index           =   7
      Left            =   315
      TabIndex        =   8
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "表示"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL(表紙)"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEl(明細)"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   3
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   4
      End
   End
End
Attribute VB_Name = "SEI00401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Const ptxS_JITU_DATE% = 0       '日付範囲　開始
Private Const ptxE_JITU_DATE% = 1       '日付範囲　開始

Private Const ptxGK_MAISU% = 2          '枚数　合計
Private Const ptxGK_KINGAKU% = 3        '金額　合計
Private Const ptxGK_SAISU% = 4          '才数　合計




Dim SE_USOU_HAKO    As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 8              '最大列数

Private Const ColJITU_DATE% = 0         '売上日付
Private Const ColHIN_GAI% = 1           '品番
Private Const ColHIN_NAME% = 2          '品名
Private Const ColSAISU% = 3             '才数
Private Const ColG_ST_URITAN% = 4       '単価
Private Const ColMTS_CODE% = 5          '出荷先
Private Const ColMAISU% = 6             '使用枚数
Private Const ColKINGAKU% = 7           '金額
Private Const ColSAISU_G% = 8           '才数（使用枚数X才数）

Private Const EXCEL_OBJECT_NAME As String = "Excel.Application"

Private MUKE_TBL()  As String * 8       '対象向け先コード




'
Dim SE_USOU_HAKO_DET    As New XArrayDB


Dim SHIMEBI             As String

'--------------------------------------- EXCEL用定数    2016.10.24
Private Const xlCalculationManual% = -4135
Private Const xlLeft% = -4131
Private Const xlCenter% = -4108
Private Const xlBottom% = -4107
Private Const xlNone% = -4142
Private Const xlTop% = -4160
Private Const xlContinuous% = 1
Private Const xlThin% = 2
Private Const xlAutomatic% = -4105
Private Const xlRight% = -4152
Private Const xlDiagonalDown% = 5
Private Const xlDiagonalUp% = 6
Private Const xlEdgeLeft% = 7
Private Const xlEdgeTop% = 8
Private Const xlEdgeBottom% = 9
Private Const xlEdgeRight% = 10
Private Const xlInsideVertical% = 11
Private Const xlInsideHorizontal% = 12
Private Const xlThick% = 4
Private Const xlCalculationAutomatic% = -4105
Private Const xlPortrait% = 1
Private Const xlDot% = -4118
Private Const xlUnderlineStyleSingle% = 2

'--------------------------------------- EXCEL用定数




Private Sub Command1_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
        Case 0                              '再表示
            If List_Disp_Proc Then
                Unload Me
            End If
        
        Case 1                              'データ出力
        
            Beep
            ans = MsgBox("請求書(表紙)作成しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If ans = vbYes Then
                If COVER_Proc() Then
                    Unload Me
                End If
            End If
        
        
        Case 2                              'データ出力
        
            Beep
            ans = MsgBox("請求書(明細)作成しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If ans = vbYes Then
                If DETAIL_Proc() Then
                    Unload Me
                End If
            End If
        Case 3                             '終了
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
Dim E_DATE      As String
Dim S_YY        As String * 4
Dim S_MM        As String * 2
Dim S_DD        As String * 2
    

Dim lnghwnd     As Long     '2016.10.20
Dim retValue    As Long     '2016.10.20

Dim retHwnd As Long         '2016.10.20
    
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
    
    
    
    
    



    SEI00401.Show


    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]輸送箱請求書作成処理", Me.hwnd, 0)
    
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SETPARTS, 0, -1)



    Show
                                
                                


    Max_Row = 9999
                                
                                
    If GetIni(App.EXEName, "SHIMEBI", App.EXEName, c) Then
        SHIMEBI = ""
    Else
        SHIMEBI = Trim(c)
    End If
                                
                                

                                '倉庫マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '輸送箱実績ＯＰＥＮ
    If SE_USOU_HAKO_Open(BtOpenNomal) Then
        Unload Me
    End If



    ReDim MUKE_TBL(0 To 0)

    i = 1
    
    Do
        If GetIni("MUKE", "MUKE" & Format(i, "00"), "SEI_SYS", c) Then
            Exit Do
        Else
        
            If Trim(c) = "********" Then
                If i > 1 Then
                    ReDim Preserve MUKE_TBL(0 To i - 1)
                End If
                
                MUKE_TBL(i - 1) = Trim(c)
        
            Else
        
                Call UniCode_Conv(K0_MTS.MUKE_CODE, Trim(c))
                Call UniCode_Conv(K0_MTS.SS_CODE, "")
                sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                Select Case sts
                    Case BtNoErr
                        
                        
                        If i > 1 Then
                            ReDim Preserve MUKE_TBL(0 To i - 1)
                        End If
                        
                        MUKE_TBL(i - 1) = Trim(c)
                    
                    Case BtErrKeyNotFound
                        MsgBox "向け先コード（" & Trim(c) & "）登録させていません"
                        Unload Me
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
                        Unload Me
                End Select
            End If
        
        
        End If
        i = i + 1
    
    Loop



    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
            
        Case Else
            Unload Me
    End Select


    E_DATE = Format(Now, "YYYY/MM/DD")
    S_DATE = DateAdd("m", -1, Left(E_DATE, 8) & SHIMEBI)
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
    Text1(ptxE_JITU_DATE).Text = E_DATE

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '向け先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先マスタ")
        End If
    End If
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
                                            '輸送箱実績ＣＬＯＳＥ
    sts = BTRV(BtOpClose, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "輸送箱実績")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   明細表示
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim GK_MAISU    As Double
Dim GK_KINGAKU  As Double
Dim GK_SAISU    As Double
    
Dim Skip_Flg    As Boolean
    
Dim End_Date    As String
    
    
    List_Disp_Proc = True
                                    
    Call Input_Lock
    
    Set SE_USOU_HAKO = Nothing
    
    
    Row = Min_Row - 1
        
    GK_MAISU = 0
    GK_KINGAKU = 0
    GK_SAISU = 0
    
    
    If IsDate(Text1(ptxS_JITU_DATE).Text) Then
        Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Format((Text1(ptxS_JITU_DATE).Text), "YYYYMMDD"))
    Else
        Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Text1(ptxS_JITU_DATE).Text)
    End If
    
    If IsDate(Text1(ptxE_JITU_DATE).Text) Then
        End_Date = Format((Text1(ptxE_JITU_DATE).Text), "YYYYMMDD")
    Else
        End_Date = Text1(ptxE_JITU_DATE).Text
    End If
    
    
    Call UniCode_Conv(K0_SE_USOU_HAKO.JGYOBU, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.NAIGAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.HIN_GAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.MTS_CODE, "")
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        
        sts = BTRV(com, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    
        Select Case sts
            Case BtNoErr
        
                If StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode) > End_Date Then
                    Exit Do
                End If
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "輸送箱実績")
                Exit Function
        End Select
                
        Skip_Flg = False
        If StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode) = "********" Then
            Skip_Flg = True
        Else
            If CLng(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode)) = 0 Then
                Skip_Flg = True
            End If
        End If
                        
                
        If Not Skip_Flg Then
            Row = Row + 1
            If Row > Max_Row Then
                Beep
                MsgBox "最大表示行数を超えました。"
                Exit Do
            End If
                    
            If Grid_Set_Proc(Row, GK_MAISU, GK_KINGAKU, GK_SAISU) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
        
    Loop
    
                                
                                'DBテーブルリンク
    
    Set TDBGrid1.Array = SE_USOU_HAKO
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
    Text1(ptxGK_MAISU).Text = Format(GK_MAISU, "#,##0")
    Text1(ptxGK_KINGAKU).Text = Format(ToRoundUp(GK_KINGAKU, 0), "#,##0.00")
    Text1(ptxGK_SAISU).Text = Format(GK_SAISU, "#,##0.0")
    
    Call Input_UnLock
    
    
    Text1(ptxS_JITU_DATE).SetFocus
    
    List_Disp_Proc = False

    
End Function

Private Function DETAIL_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（明細）出力
'----------------------------------------------------------------------------

'Dim excelApplication    As excel.Application       '2016.10.24
'Dim excelWorkBook       As excel.Workbook          '2016.10.24
'Dim excelSheet          As excel.Worksheet         '2016.10.24

Dim excelApplication    As Object                   '2016.10.24
Dim excelWorkBook       As Object                   '2016.10.24
Dim excelSheet          As Object                   '2016.10.24



Dim i                   As Integer
Dim j                   As Integer
Dim Row                 As Integer
    
    
Dim sts                 As Integer
Dim com                 As Integer
    
Dim End_Date            As String

Dim svJGYOBU            As String * 1
Dim svNAIGAI            As String * 1
Dim svHIN_GAI           As String * 20

Dim GK_MAISU            As Long
Dim GK_KINGAKU          As Long
Dim GK_SAISU            As Long

Dim Fast_Flg            As Boolean

Dim c                   As String * 128             '2016.10.24
Dim MEISAI_TITLE        As String                   '2016.10.24
    
    
    DETAIL_Proc = True
    
    Call Input_Lock
    
    
    
    Set SE_USOU_HAKO_DET = Nothing
    
    
    Set excelApplication = CreateObject("Excel.Application")
'2009.02.24    excelApplication.Visible = True


    
    Set excelWorkBook = excelApplication.Workbooks.Add
    Set excelSheet = excelWorkBook.Worksheets(1)
    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "ＭＳ　ゴシック"

    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 4)).Select
    With excelSheet.Application.Selection.Font
                
        .Size = 16
    End With
    
'    excelSheet.Application.Cells(1, 1).Value = "輸送箱請求書"      '2016.10.24
    
    If GetIni(App.EXEName, "MEISAI_TITLE", App.EXEName, c) Then     '2016.10.24
        MEISAI_TITLE = ""                                           '2016.10.24
    Else                                                            '2016.10.24
        MEISAI_TITLE = Trim(c)                                      '2016.10.24
    End If                                                          '2016.10.24
    excelSheet.Application.Cells(1, 1).Value = MEISAI_TITLE         '2016.10.24
    
    
    
    excelSheet.Application.Cells(1, 4).Value = Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode)) & _
                                    "（" & Text1(ptxS_JITU_DATE).Text & "～" & _
                                    Text1(ptxE_JITU_DATE).Text & "）"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 4), excelSheet.Application.Cells(2, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 4), excelSheet.Application.Cells(2, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 4), excelSheet.Application.Cells(2, 6)).MergeCells = True
    excelSheet.Application.Cells(2, 4).Value = "合計"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(3, 1).Value = "箱番号"
    excelSheet.Application.Cells(3, 2).Value = "才数"
    excelSheet.Application.Cells(3, 3).Value = "単価"
    excelSheet.Application.Cells(3, 4).Value = "数量"
    excelSheet.Application.Cells(3, 5).Value = "金額"
    excelSheet.Application.Cells(3, 6).Value = "才数"

    '---------- 罫線
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 4), excelSheet.Application.Cells(2, 6)).Select
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
    
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 3)).Select
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
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 4), excelSheet.Application.Cells(3, 6)).Select
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
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    '---------- 罫線


    excelSheet.Application.Range(excelSheet.Application.Cells(2, 4), excelSheet.Application.Cells(2, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 4), excelSheet.Application.Cells(2, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 4), excelSheet.Application.Cells(2, 6)).MergeCells = True
    excelSheet.Application.Cells(2, 4).Value = "合計"
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 4), excelSheet.Application.Cells(3, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 4), excelSheet.Application.Cells(3, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(3, 4).Value = "数量"
    excelSheet.Application.Cells(3, 5).Value = "金額"
    excelSheet.Application.Cells(3, 6).Value = "才数"



    j = 5
    For i = 0 To UBound(MUKE_TBL)
        j = j + 2

        Call UniCode_Conv(K0_MTS.MUKE_CODE, MUKE_TBL(i))
        Call UniCode_Conv(K0_MTS.SS_CODE, "")
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)

        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(MTSREC.MUKE_NAME, "その他")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
                Exit Function
        End Select
        
'2008.06.04        excelSheet.Application.Range(excelSheet.Application.Cells(2, j), excelSheet.Application.Cells(2, j + 2)).HorizontalAlignment = xlCenter
'2008.06.04        excelSheet.Application.Range(excelSheet.Application.Cells(2, j), excelSheet.Application.Cells(2, j + 2)).VerticalAlignment = xlCenter
'2008.06.04        excelSheet.Application.Range(excelSheet.Application.Cells(2, j), excelSheet.Application.Cells(2, j + 2)).MergeCells = True
'2008.06.04        excelSheet.Application.Cells(2, j).Value = Trim(StrConv(MTSREC.MUKE_NAME, vbUnicode))
'2008.06.04        excelSheet.Application.Range(excelSheet.Application.Cells(3, j), excelSheet.Application.Cells(3, j + 2)).HorizontalAlignment = xlCenter
'2008.06.04        excelSheet.Application.Range(excelSheet.Application.Cells(3, j), excelSheet.Application.Cells(3, j + 2)).VerticalAlignment = xlCenter
        
        
        
        
        '2008.06.04
        excelSheet.Application.Range(excelSheet.Application.Cells(2, j), excelSheet.Application.Cells(2, j + 1)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(2, j), excelSheet.Application.Cells(2, j + 1)).VerticalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(2, j), excelSheet.Application.Cells(2, j + 1)).MergeCells = True
        excelSheet.Application.Cells(2, j).Value = Trim(StrConv(MTSREC.MUKE_NAME, vbUnicode))
        excelSheet.Application.Range(excelSheet.Application.Cells(3, j), excelSheet.Application.Cells(3, j + 1)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(3, j), excelSheet.Application.Cells(3, j + 1)).VerticalAlignment = xlCenter
        '2008.06.04
        
        
        
        
        
        
        excelSheet.Application.Cells(3, j).Value = "数量"
'2008.06.04        excelSheet.Application.Cells(3, j + 1).Value = "金額"
'2008.06.04        excelSheet.Application.Cells(3, j + 2).Value = "才数"
        excelSheet.Application.Cells(3, j + 1).Value = "才数"
        
        '-----　罫線
'2080.06.04        excelSheet.Application.Range(excelSheet.Application.Cells(2, j), excelSheet.Application.Cells(2, j + 2)).Select
        excelSheet.Application.Range(excelSheet.Application.Cells(2, j), excelSheet.Application.Cells(2, j + 1)).Select
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
        
        excelSheet.Application.Range(excelSheet.Application.Cells(3, j), excelSheet.Application.Cells(3, j + 1)).Select
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
        With excelSheet.Application.Selection.Borders(xlInsideVertical)
            .LineStyle = xlDot
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        '-----　罫線
        

    Next i




    If IsDate(Text1(ptxS_JITU_DATE).Text) Then
        Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Format((Text1(ptxS_JITU_DATE).Text), "YYYYMMDD"))
    Else
        Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Text1(ptxS_JITU_DATE).Text)
    End If
 
 
    Call UniCode_Conv(K0_SE_USOU_HAKO.JGYOBU, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.NAIGAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.HIN_GAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.MTS_CODE, "")
    
    If IsDate(Text1(ptxE_JITU_DATE).Text) Then
        End_Date = Format((Text1(ptxE_JITU_DATE).Text), "YYYYMMDD")
    Else
        End_Date = Text1(ptxE_JITU_DATE).Text
    End If


    Fast_Flg = True

    com = BtOpGetGreater
    Do
        DoEvents
    
        sts = BTRV(com, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    
        Select Case sts
            Case BtNoErr
        
                If StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode) > End_Date Then
                    Exit Do
                End If
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "輸送箱実績")
                Exit Function
        End Select
    
    
        If CLng(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode)) = 0 Then
        Else
        
            If Fast_Flg Then
                
                
                SE_USOU_HAKO_DET.ReDim Min_Row, 1, Min_Col, UBound(MUKE_TBL) + 2
                
                SE_USOU_HAKO_DET(1, 1) = Trim(StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
                                
                Fast_Flg = False
            
            End If
        
        
        
            For i = 1 To SE_USOU_HAKO_DET.UpperBound(1)
            
                If Trim(SE_USOU_HAKO_DET(i, 1)) = Trim(StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode)) Then
                    Exit For
                End If
            
            Next i
        
            If i > SE_USOU_HAKO_DET.UpperBound(1) Then
                SE_USOU_HAKO_DET.ReDim Min_Row, i, Min_Col, UBound(MUKE_TBL) + 2
                SE_USOU_HAKO_DET(i, 1) = Trim(StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
            End If
        
        
                
            For j = 0 To UBound(MUKE_TBL)
            
                If Trim(MUKE_TBL(j)) = Trim(StrConv(SE_USOU_HAKOREC.MTS_CODE, vbUnicode)) Then
                    Exit For
                End If
            Next j
        
            If j > UBound(MUKE_TBL) Then
            Else
                SE_USOU_HAKO_DET(i, j + 2) = CLng(SE_USOU_HAKO_DET(i, j + 2)) + CLng(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode))
            End If
    
        End If
        com = BtOpGetNext
    
    Loop

    If Fast_Flg Then                '2016.10.24
    Else                            '2016.10.24
       SE_USOU_HAKO_DET.QuickSort Min_Row, SE_USOU_HAKO_DET.UpperBound(1), ColHIN_GAI, 0, XTYPE_STRING

        If Excel_Set_Proc(excelApplication, excelWorkBook, excelSheet) Then
            Exit Function
        End If
    
    End If
    
    excelApplication.Visible = True


    Set excelSheet = Nothing
    Set excelWorkBook = Nothing
    Set excelApplication = Nothing


    
    Call Input_UnLock
    DETAIL_Proc = False
    

End Function
Private Function COVER_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（表紙）出力
'----------------------------------------------------------------------------

'Dim excelApplication    As excel.Application       '2016.10.24
'Dim excelWorkBook       As excel.Workbook          '2016.10.24
'Dim excelSheet          As excel.Worksheet         '2016.10.24

Dim excelApplication    As Object                   '2016.10.24
Dim excelWorkBook       As Object                   '2016.10.24
Dim excelSheet          As Object                   '2016.10.24


Dim i                   As Integer
Dim j                   As Integer
Dim Row                 As Integer
    
    
Dim sts                 As Integer
Dim com                 As Integer
    
Dim End_Date            As String


Dim GK_KINGAKU          As Double
Dim WK_TANKA            As Double
Dim ZEI_KIN             As Long


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
    
    excelApplication.StandardFont = "ＭＳ　ゴシック"

    
    'ページ設定
    With excelSheet.Application.ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    
    '列の幅
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
    
    '行の幅
    excelSheet.Application.Rows(1).Select
    excelSheet.Application.Selection.RowHeight = 24
    excelSheet.Application.Rows("3:4").Select
    excelSheet.Application.Selection.RowHeight = 14.25
    excelSheet.Application.Rows(12).Select
    excelSheet.Application.Selection.RowHeight = 27
    excelSheet.Application.Rows("14:31").Select
    excelSheet.Application.Selection.RowHeight = 27
    
    'セルの結合
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "太字"
        .Size = 20
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(1, 1).Value = "請 求 書"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "標準"
        .Size = 14
    End With
    excelSheet.Application.Cells(12, 1).Value = "合 計 金 額"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "標準"
        .Size = 14
    End With
    excelSheet.Application.Cells(12, 4).Value = ""
    
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "標準"
        .Size = 11
    End With
    excelSheet.Application.Cells(29, 1).Value = "税 抜 き 金 額"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "標準"
        .Size = 11
    End With
    excelSheet.Application.Cells(30, 1).Value = "消    費    税"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "標準"
        .Size = 11
    End With
    excelSheet.Application.Cells(31, 1).Value = "税 込 み 金 額"
    
    
    
    '罫線
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
    '固定項目（見出し）
    excelSheet.Application.Range(excelSheet.Application.Cells(14, 1), excelSheet.Application.Cells(14, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(14, 1), excelSheet.Application.Cells(14, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(14, 1).Value = "月/日"
    excelSheet.Application.Cells(14, 2).Value = "品     名"
    excelSheet.Application.Cells(14, 3).Value = "数 量"
    excelSheet.Application.Cells(14, 4).Value = "単  価"
    excelSheet.Application.Cells(14, 5).Value = "金　額"
    excelSheet.Application.Cells(14, 6).Value = "摘　要"
    '固定項目（INI）
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 6), excelSheet.Application.Cells(2, 6)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(2, 6).Value = Left(Format(Now, "YYYY年MM月DD日"), 8) & SHIMEBI & "日"
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 1)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "太字"
        .Size = 11
    End With
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 6), excelSheet.Application.Cells(9, 6)).HorizontalAlignment = xlRight
    
    excelSheet.Application.Cells(3, 1).Value = Name1
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 1)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "太字"
        .Size = 11
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(4, 1).Value = Name2
    
    excelSheet.Application.Range(excelSheet.Application.Cells(8, 2), excelSheet.Application.Cells(8, 2)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "太字"
        .Size = 11
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(8, 2).Value = ITEM
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 6), excelSheet.Application.Cells(7, 6)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "太字"
        .Size = 9
    End With
    excelSheet.Application.Cells(4, 6).Value = ADDR1
    excelSheet.Application.Cells(5, 6).Value = ADDR2
    excelSheet.Application.Cells(6, 6).Value = SYAMEI
    excelSheet.Application.Cells(7, 6).Value = BIKOU1
    excelSheet.Application.Cells(8, 6).Value = BIKOU2
    excelSheet.Application.Cells(9, 6).Value = BIKOU3
    
    
    If IsDate(Text1(ptxS_JITU_DATE).Text) Then
        Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Format(Text1(ptxS_JITU_DATE).Text, "YYYYMMDD"))
    Else
        Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Text1(ptxS_JITU_DATE).Text)
    End If
    Call UniCode_Conv(K0_SE_USOU_HAKO.JGYOBU, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.NAIGAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.HIN_GAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.MTS_CODE, "")


    If IsDate(Text1(ptxE_JITU_DATE).Text) Then
        End_Date = Format(Text1(ptxE_JITU_DATE).Text, "YYYYMMDD")
    Else
        End_Date = Text1(ptxE_JITU_DATE).Text
    End If

    com = BtOpGetGreater
    GK_KINGAKU = 0
    
    Do
    
        DoEvents
    
        sts = BTRV(com, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    
        Select Case sts
            Case BtNoErr
        
                If StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode) > End_Date Then
                    Exit Do
                End If
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "輸送箱実績")
                Exit Function
        End Select
    
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    WK_TANKA = CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))
                Else
                    WK_TANKA = 0
                End If
            Case BtErrKeyNotFound
                    WK_TANKA = 0
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
    
        GK_KINGAKU = GK_KINGAKU + Round(CDbl(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode)) * _
                                            WK_TANKA, 2)
    
        com = BtOpGetNext
    
    Loop
    
    
    '月／日
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 1), excelSheet.Application.Cells(15, 1)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 1), excelSheet.Application.Cells(15, 1)).Select
    excelSheet.Application.Selection.NumberFormatLocal = "@"
    excelSheet.Application.Selection.HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(15, 1).Value = Format(CInt(Mid(Format(Now, "YYYYMMDD"), 5, 2)), "#") & "/" & SHIMEBI
    '品名
    excelSheet.Application.Cells(15, 2).Value = HIN_NAME
    '数量
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 3), excelSheet.Application.Cells(15, 3)).Select
    excelSheet.Application.Selection.NumberFormatLocal = "#,##0"
    excelSheet.Application.Cells(15, 3).Value = 1
    '単価～金額
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 4), excelSheet.Application.Cells(15, 5)).Select
    excelSheet.Application.Selection.NumberFormatLocal = "#,##0"
    excelSheet.Application.Cells(15, 4).Value = ToRoundUp(GK_KINGAKU, 0)
    excelSheet.Application.Cells(15, 5).Value = ToRoundUp(GK_KINGAKU, 0)
    '摘要
    excelSheet.Application.Cells(15, 6).Value = Trim(TEKIYO)
    
    
    
    
    
    '税抜き金額
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 5), excelSheet.Application.Cells(31, 5)).Select
    excelSheet.Application.Selection.NumberFormatLocal = "#,##0;""▲ ""#,##0"


    excelSheet.Application.Cells(29, 5).Value = ToRoundUp(GK_KINGAKU, 0)
    '消費税
    ZEI_KIN = Fix((GK_KINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10)
    excelSheet.Application.Cells(30, 5).Value = ZEI_KIN
    '税込み金額
    excelSheet.Application.Cells(31, 5).Value = GK_KINGAKU + ZEI_KIN
    '合計金額
    excelSheet.Application.Cells(12, 4).Value = Format(GK_KINGAKU + ZEI_KIN, "\\#,##0")


    Set excelSheet = Nothing
    Set excelWorkBook = Nothing
    Set excelApplication = Nothing


    
    Call Input_UnLock
    COVER_Proc = False
    

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    SEI00401.MousePointer = vbHourglass


    Call Ctrl_Lock(SEI00401)

    TDBGrid1.Enabled = False


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEI00401)
    
    TDBGrid1.Enabled = True


    SEI00401.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long, GK_MAISU As Double, GK_KINGAKU As Double, GK_SAISU As Double) As Integer
'----------------------------------------------------------------------------
'                   輸送箱実績-->Ｇｒｉｄ
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim wkDec   As Long
    
    Grid_Set_Proc = True

    

    SE_USOU_HAKO.ReDim Min_Row, Row, Min_Col, Max_Col
    
    
    
    
    '売上日付
    SE_USOU_HAKO(Row, ColJITU_DATE) = Mid(StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode), 7, 2)
    '品番
    SE_USOU_HAKO(Row, ColHIN_GAI) = Trim(StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            SE_USOU_HAKO(Row, ColHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
            
            
            If Not IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "0")
            End If
        Case BtErrKeyNotFound
            SE_USOU_HAKO(Row, ColHIN_NAME) = ""
            Call UniCode_Conv(ITEMREC.G_ST_URITAN, "0")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    '才数
    If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
        SE_USOU_HAKO(Row, ColSAISU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#0.0")
    Else
        SE_USOU_HAKO(Row, ColSAISU) = "0.0"
    End If
    '単価
    If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
        SE_USOU_HAKO(Row, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
    Else
        SE_USOU_HAKO(Row, ColG_ST_URITAN) = "0.00"
    End If
    '出荷先
    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(SE_USOU_HAKOREC.MTS_CODE, vbUnicode))
    Call UniCode_Conv(K0_MTS.SS_CODE, "")
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
            SE_USOU_HAKO(Row, ColMTS_CODE) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & Trim(StrConv(MTSREC.MUKE_NAME, vbUnicode))
        Case BtErrKeyNotFound
            SE_USOU_HAKO(Row, ColMTS_CODE) = "その他"
        Case Else
            Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
            Exit Function
    End Select
    '使用枚数
    SE_USOU_HAKO(Row, ColMAISU) = Format(CLng(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode)), "#0")
    GK_MAISU = GK_MAISU + CLng(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode))
    '金額
    SE_USOU_HAKO(Row, ColKINGAKU) = Format(Round(CDbl(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode)) * _
                                        CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), 2), "#,##0.00")
    
    
    
    
    GK_KINGAKU = GK_KINGAKU + Round(CDbl(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode)) * _
                                        CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), 2)
    
    
    '才数
    SE_USOU_HAKO(Row, ColSAISU_G) = Format(CLng(SE_USOU_HAKO(Row, ColMAISU)) * CDbl(SE_USOU_HAKO(Row, ColSAISU)), "#0.0")
    GK_SAISU = GK_SAISU + CLng(SE_USOU_HAKO(Row, ColMAISU)) * CDbl(SE_USOU_HAKO(Row, ColSAISU))
    
    
    
    
    
    
    
    
    
    Grid_Set_Proc = False
End Function

Private Sub SHORI_Click(Index As Integer)
    Select Case Index
    
        
        
        
        Case 0 To 3
        
        
            Command1(Index).Value = True
        
        
        
        Case 4      '画面印刷
        
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
                    
    
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
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
    
        Case ptxS_JITU_DATE     '開始日付
            
        Case ptxS_JITU_DATE       '終了日付
            
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function

Private Function Excel_Set_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object) As Integer
'Private Function Excel_Set_Proc(excelApplication As excel.Application, excelWorkBook As excel.Workbook, excelSheet As excel.Worksheet) As Integer
'----------------------------------------------------------------------------
'                   輸送箱実績-->EXCEL
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer


Dim Row         As Integer
Dim sts         As Integer
    
Dim ParaM1      As String
Dim ParaM2      As String
Dim ParaM3      As String
    
    
    
    Excel_Set_Proc = True
    
    
    Row = 3
    
    For i = 1 To SE_USOU_HAKO_DET.UpperBound(1)
    
        
        Row = Row + 1
        '品番
        excelSheet.Application.Cells(Row, 1).Value = SE_USOU_HAKO_DET(i, 1)
    
        '才数
        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, SE_USOU_HAKO_DET(i, 1))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
           Case BtNoErr
           Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.SAI_SU, "0.0")
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "0.00")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        excelSheet.Cells(Row, 2).NumberFormatLocal = "0.0_ "
        If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
            excelSheet.Application.Cells(Row, 2).Value = StrConv(ITEMREC.SAI_SU, vbUnicode)
        Else
            excelSheet.Application.Cells(Row, 2).Value = 0
        End If
    
    
        excelSheet.Cells(Row, 3).NumberFormatLocal = "#,##0.00_ "
        If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
            excelSheet.Application.Cells(Row, 3).Value = StrConv(ITEMREC.G_ST_URITAN, vbUnicode)
        Else
            excelSheet.Application.Cells(Row, 3).Value = 0
        End If
    
    
        k = 5
        For j = 0 To UBound(MUKE_TBL)
        
        
            k = k + 2
            
            '数量
            excelSheet.Application.Cells(Row, k).NumberFormatLocal = "0_ "
            excelSheet.Application.Cells(Row, k).Value = CLng(SE_USOU_HAKO_DET(i, j + 2))
            '金額
'2008.06.04            excelSheet.Application.Cells(Row, k + 1).NumberFormatLocal = "#,##0.00_ "
'2008.06.04            If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
'2008.06.04                excelSheet.Application.Cells(Row, k + 1).Value = Round(CDbl(SE_USOU_HAKO_DET(i, j + 2)) * _
'2008.06.04                                            CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), 3)
'2008.06.04            Else
'2008.06.04                excelSheet.Application.Cells(Row, k + 1).Value = 0
'2008.06.04            End If
            '才数
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, k + 1), excelSheet.Application.Cells(Row, k + 1)).Select
            excelSheet.Application.Cells(Row, k + 1).NumberFormatLocal = "#,##0.0_ "
            excelSheet.Application.ActiveCell.FormulaR1C1 = "=RC[-1]*RC[" & -k + 1 & "]"
            
            
'            If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
'                excelSheet.Application.Cells(Row, k + 1).Value = Round(CLng(SE_USOU_HAKO_DET(i, j + 2)) * _
'                                                                            CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), 1)
'            Else
'                excelSheet.Application.Cells(Row, k + 1).Value = 0
'            End If
        
        
        
        Next j
    
    
        ParaM1 = ""
    
    
        k = 1
        For j = 0 To UBound(MUKE_TBL)
                
                
            k = k + 2
            If ParaM1 = "" Then
                ParaM1 = "=RC[" & k & "]"
            Else
                ParaM1 = ParaM1 & "+RC[" & k & "]"
            End If
                    
        Next j
            
        
        
        ParaM2 = "=round(RC[-1]*RC[-2],2)"
        
        
        ParaM3 = ""
        k = 0
        For j = 0 To UBound(MUKE_TBL)
                
                
            k = k + 2
            If ParaM3 = "" Then
                ParaM3 = "=RC[" & k & "]"
            Else
                ParaM3 = ParaM3 & "+RC[" & k & "]"
            End If
                    
        Next j
            
            
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 4)).Select
        excelSheet.Application.Cells(Row, 4).NumberFormatLocal = "0_ "
        excelSheet.Application.ActiveCell.FormulaR1C1 = ParaM1
        
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 5)).Select
        excelSheet.Application.Cells(Row, 5).NumberFormatLocal = "#,##0.00_ "
        excelSheet.Application.ActiveCell.FormulaR1C1 = ParaM2
        
        
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 6)).Select
'        excelSheet.Application.Cells(Row, 6).NumberFormatLocal = "#,##0.0_ "
'        excelSheet.Application.ActiveCell.FormulaR1C1 = ParaM3
        
        excelSheet.Application.Cells(Row, 6).NumberFormatLocal = "#,##0.0_ "
        excelSheet.Application.ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-4]"
        
        
    
    
    
        '---------- 罫線
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 3)).Select
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
        With excelSheet.Application.Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        
        
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 6)).Select
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
        
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 6)).Select
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
        With excelSheet.Application.Selection.Borders(xlInsideVertical)
            .LineStyle = xlDot
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        
        
        
        j = 5
        For k = 0 To UBound(MUKE_TBL)
            
            
            DoEvents
            
            
            j = j + 2
    
            
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, j), excelSheet.Application.Cells(Row, j + 1)).Select
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
            
            excelSheet.Application.Range(excelSheet.Application.Cells(Row, j), excelSheet.Application.Cells(Row, j + 1)).Select
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
            With excelSheet.Application.Selection.Borders(xlInsideVertical)
                .LineStyle = xlDot
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        Next k
    
    
    
    
    Next i
    



    '----------------   縦計
    Row = Row + 1
    excelSheet.Application.Cells(Row, 1).Value = "合計"
    
    
    
    '数量
    excelSheet.Application.Cells(Row, 4).NumberFormatLocal = "0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 4)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
        
        
     '金額
    excelSheet.Application.Cells(Row, 5).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 5)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C),0)"
    '才数
    excelSheet.Application.Cells(Row, 6).NumberFormatLocal = "#,##0.0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 6), excelSheet.Application.Cells(Row, 6)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    
    j = 5
    For i = 0 To UBound(MUKE_TBL)
        j = j + 2
    
        '数量
        excelSheet.Application.Cells(Row, j).NumberFormatLocal = "#,##0_ "
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, j), excelSheet.Application.Cells(Row, j)).Select
        excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
        '金額
'        excelSheet.Application.Cells(Row, j + 1).NumberFormatLocal = "#,##0_ "
'        excelSheet.Application.Range(excelSheet.Application.Cells(Row, j + 1), excelSheet.Application.Cells(Row, j + 1)).Select
'        excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
        '才数
        excelSheet.Application.Cells(Row, j + 1).NumberFormatLocal = "#,##0.0_ "
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, j + 2), excelSheet.Application.Cells(Row, j + 1)).Select
        excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    
    
    
    Next i
    
    
    '---------- 罫線
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 2)).Select
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
    
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 3), excelSheet.Application.Cells(Row, 5)).Select
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
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 3), excelSheet.Application.Cells(Row, 5)).Select
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
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    j = 3
    For i = 0 To UBound(MUKE_TBL) + 1
        j = j + 2

        
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, j), excelSheet.Application.Cells(Row, j + 1)).Select
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
        
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, j), excelSheet.Application.Cells(Row, j + 1)).Select
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
        With excelSheet.Application.Selection.Borders(xlInsideVertical)
            .LineStyle = xlDot
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next i



           





    excelSheet.Application.Columns("C").EntireColumn.AutoFit


    Excel_Set_Proc = False


End Function



' ------------------------------------------------------------------------
'       指定した精度の数値に切り上げします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り上げられた数値。
' ------------------------------------------------------------------------
Private Function ToRoundUp(ByVal dValue As Double, ByVal iDigits As Integer) As Double
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

