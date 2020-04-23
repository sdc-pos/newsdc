VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEI00101 
   Caption         =   "[請求システム]見積書作成処理"
   ClientHeight    =   12195
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   19035
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
   ScaleHeight     =   12195
   ScaleWidth      =   19035
   StartUpPosition =   2  '画面の中央
   Begin TrueDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   855
      Left            =   1785
      TabIndex        =   42
      Top             =   8280
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1508
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).ValueItems(0)._DefaultItem=   0
      Columns(0).ValueItems(0).Value=   "aaaa"
      Columns(0).ValueItems(0).Value.vt=   8
      Columns(0).ValueItems(0).DisplayValue=   "aaaa"
      Columns(0).ValueItems(0).DisplayValue.vt=   8
      Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems.Count=   1
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4366"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   ""
      DataField       =   "ub_grid2"
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1200,.italic=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1785
      TabIndex        =   0
      Top             =   960
      Width           =   750
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   14
      Left            =   15015
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9480
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   13
      Left            =   7875
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9480
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   12
      Left            =   15435
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4800
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   11
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5280
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   9
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   330
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   8
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2400
      Width           =   330
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   7
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
      Width           =   330
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   6
      Left            =   1785
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   330
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   3
      Left            =   4305
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2040
      Width           =   4320
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   615
      Index           =   10
      Left            =   12600
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   15
      Left            =   1575
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   10200
      Width           =   9570
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   14385
      TabIndex        =   6
      Top             =   2040
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   4
      Left            =   10290
      TabIndex        =   5
      Top             =   2040
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1785
      TabIndex        =   3
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1785
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1680
      Width           =   2220
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
      Left            =   5250
      TabIndex        =   25
      Top             =   240
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXCEL"
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
      Left            =   3570
      TabIndex        =   24
      Top             =   240
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   9240
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "削  除"
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
      Left            =   1890
      TabIndex        =   22
      Top             =   240
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更　新"
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
      TabIndex        =   21
      Top             =   240
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   1935
      Index           =   0
      Left            =   315
      TabIndex        =   12
      Top             =   3360
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   3413
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "個装資材"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "名　称"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "員数"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "数量"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "単価"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "金額"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1958"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1826"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3572"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3440"
      Splits(0)._ColumnProps(8)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=1429"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1296"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=1561"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1429"
      Splits(0)._ColumnProps(17)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=1667"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1535"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=1640"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=1508"
      Splits(0)._ColumnProps(26)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "個装資材"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=102,.parent=87"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=106,.parent=87,.bgcolor=&HC0C0C0&"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=110,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=114,.parent=87,.bgcolor=&HC0C0C0&"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=111,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=112,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=113,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=118,.parent=87"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=115,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=116,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=117,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=16,.parent=87,.bgcolor=&HC0C0C0&"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=91"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   1455
      Index           =   1
      Left            =   9240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3360
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   2566
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "外装資材"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "名　称"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "入数"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "数量"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "単価"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "金額"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1958"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1826"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3572"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3440"
      Splits(0)._ColumnProps(8)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=1429"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1296"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=1561"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1429"
      Splits(0)._ColumnProps(17)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=1667"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1535"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=1640"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=1508"
      Splits(0)._ColumnProps(26)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "外装資材"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=102,.parent=87"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=106,.parent=87,.bgcolor=&HC0C0C0&"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=110,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=114,.parent=87,.bgcolor=&HC0C0C0&"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=111,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=112,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=113,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=118,.parent=87"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=115,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=116,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=117,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=16,.parent=87,.bgcolor=&HC0C0C0&"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=91"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   3615
      Index           =   2
      Left            =   315
      TabIndex        =   15
      Top             =   5880
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   6376
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   1
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "種別"
      Columns(0).DataField=   "TDBDropDown1"
      Columns(0).DropDown=   "TDBDropDown1"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "構成・同梱"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "名　称"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "員数"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "数量"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "単価"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "金額"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1879"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
      Splits(0)._ColumnProps(4)=   "Column(0).Button=1"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2540"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2408"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=3572"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=3440"
      Splits(0)._ColumnProps(13)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=1429"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1296"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=1561"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1429"
      Splits(0)._ColumnProps(22)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=1667"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=1535"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=1640"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=1508"
      Splits(0)._ColumnProps(31)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "構成・同梱"
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
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=106,.parent=87,.bgcolor=&HC0C0C0&"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=103,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=104,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=105,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=110,.parent=87"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=107,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=108,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=109,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=114,.parent=87,.bgcolor=&HC0C0C0&"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=111,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=112,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=113,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=118,.parent=87"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=115,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=116,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=117,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=16,.parent=87,.bgcolor=&HC0C0C0&"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=13,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=14,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=15,.parent=91"
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   3615
      Index           =   3
      Left            =   9240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   6376
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "作業内容"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "工数"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "単価"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "金額"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "請求先"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=767"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=635"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=2"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4789"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4657"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=1429"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1296"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=2408"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=2487"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2355"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(24)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=4366"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=4233"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "作業内容"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=87,.alignment=1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=87"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=110,.parent=87,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=118,.parent=87,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=115,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=116,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=117,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=16,.parent=87,.alignment=1,.bgcolor=&HC0C0C0&"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=20,.parent=87"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=17,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=18,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=19,.parent=91"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      Caption         =   "担当者"
      Height          =   255
      Index           =   12
      Left            =   945
      TabIndex        =   41
      Top             =   1080
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "小計"
      Height          =   255
      Index           =   11
      Left            =   14490
      TabIndex        =   40
      Top             =   9600
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "小計"
      Height          =   255
      Index           =   10
      Left            =   7245
      TabIndex        =   39
      Top             =   9600
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "小計"
      Height          =   255
      Index           =   9
      Left            =   14805
      TabIndex        =   37
      Top             =   4920
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "小計"
      Height          =   255
      Index           =   8
      Left            =   5985
      TabIndex        =   36
      Top             =   5400
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "-"
      Height          =   255
      Index           =   7
      Left            =   3045
      TabIndex        =   35
      Top             =   2520
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "-"
      Height          =   255
      Index           =   6
      Left            =   2625
      TabIndex        =   34
      Top             =   2520
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   2100
      TabIndex        =   33
      Top             =   2520
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "標準棚番"
      Height          =   255
      Index           =   4
      Left            =   735
      TabIndex        =   32
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "見積り合計"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   21.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10290
      TabIndex        =   31
      Top             =   2520
      Width           =   2325
   End
   Begin VB.Label Label3 
      Caption         =   "備考"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   21.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   315
      TabIndex        =   30
      Top             =   10200
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "見積り№"
      Height          =   255
      Index           =   3
      Left            =   13230
      TabIndex        =   29
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "標準ロット数"
      Height          =   255
      Index           =   2
      Left            =   9030
      TabIndex        =   28
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "品目コード"
      Height          =   255
      Index           =   1
      Left            =   525
      TabIndex        =   27
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "仕向け先"
      Height          =   255
      Index           =   0
      Left            =   735
      TabIndex        =   26
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Width           =   180
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "更新"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "削除"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   3
      End
   End
End
Attribute VB_Name = "SEI00101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxTanto_Code% = 0        '担当者コード
Private Const ptxTanto_Name% = 1        '担当者名称
Private Const ptxHin_Gai% = 2           '品番
Private Const ptxHin_Name% = 3          '品名
Private Const ptxDEF_LOT% = 4           '標準ロット
Private Const ptxMITSUMORI_NO% = 5      '見積書№
Private Const ptxST_SOKO% = 6           '標準棚番　倉庫
Private Const ptxST_RETU% = 7           '標準棚番　列
Private Const ptxST_REN% = 8            '標準棚番　蓮
Private Const ptxST_DAN% = 9            '標準棚番　段

Private Const ptxTotal% = 10            '見積り合計


Private Const ptxK_Total% = 11          '個装合計
Private Const ptxG_Total% = 12          '外装合計
Private Const ptxD_Total% = 13          '構成同梱合計
Private Const ptxS_Total% = 14          '作業合計

Private Const ptxBIKOU% = 15            '備考


Private Const pcmbSHIMUKE% = 0          '仕向け先

'------------------------------------   '個装資材
Private Const pGrdKOSOU% = 0

Dim KOSOU As New XArrayDB

Private Const K_Min_Row% = 1            '最小行数

Dim K_Max_Row   As Integer              'グリッド最大表示件数

Private Const K_Min_Col% = 0            '最小列数
Private Const K_Max_Col% = 5            '最大列数

Private Const ColK_HIN_GAI% = 0         '個装資材コード(品番)
Private Const ColK_HIN_NAME% = 1        '名称
Private Const ColK_QTY% = 2             '員数
Private Const ColK_SHIJI_QTY% = 3       '数量
Private Const ColK_TANKA% = 4           '単価
Private Const ColK_KIN% = 5             '金額

'------------------------------------   '外装資材
Private Const pGrdGAISOU% = 1


Dim GAISOU As New XArrayDB

Private Const G_Min_Row% = 1            '最小行数

Dim G_Max_Row   As Integer              'グリッド最大表示件数

Private Const G_Min_Col% = 0            '最小列数
Private Const G_Max_Col% = 5            '最大列数

Private Const ColG_HIN_GAI% = 0         '外装資材コード(品番)
Private Const ColG_HIN_NAME% = 1        '名称
Private Const ColG_QTY% = 2             '員数
Private Const ColG_SHIJI_QTY% = 3       '数量
Private Const ColG_TANKA% = 4           '単価
Private Const ColG_KIN% = 5             '金額

'------------------------------------   '構成／同梱
Private Const pGrdDOUKON% = 2


Dim DOUKON As New XArrayDB

Private Const D_Min_Row% = 1            '最小行数

Dim D_Max_Row   As Integer              'グリッド最大表示件数

Private Const D_Min_Col% = 0            '最小列数
Private Const D_Max_Col% = 6            '最大列数

Private Const ColD_SYUBETU% = 0         '種別
Private Const ColD_HIN_GAI% = 1         '同梱(品番)
Private Const ColD_HIN_NAME% = 2        '名称
Private Const ColD_QTY% = 3             '員数
Private Const ColD_SHIJI_QTY% = 4       '数量
Private Const ColD_TANKA% = 5           '単価
Private Const ColD_KIN% = 6             '金額

'------------------------------------   '作業
Private Const pGrdSAGYO% = 3

Dim SAGYO As New XArrayDB

Private Const S_Min_Row% = 1            '最小行数

Dim S_Max_Row   As Integer              'グリッド最大表示件数

Private Const S_Min_Col% = 0            '最小列数
Private Const S_Max_Col% = 5            '最大列数

Private Const ColS_No% = 0              '№
Private Const ColS_NAME% = 1            '名称
Private Const ColS_KOUSU% = 2           '員数
Private Const ColS_TANKA% = 3           '単価
Private Const ColS_KIN% = 4             '金額
Private Const ColS_SEIKYU% = 5          '請求先



'-----------------------------------    ドロップダウン
Dim SYUBETU As New XArrayDB

'-----------------------------------    入力内容のキープ
Private Type Item_Key_tag
    JGYOBU  As String * 1
    NAIGAI  As String * 1
End Type

Private K_Item_Tbl() As Item_Key_tag   '個装資材品目情報
Private G_Item_Tbl() As Item_Key_tag   '外装資材品目情報
Private D_Item_Tbl() As Item_Key_tag    '同梱品目情報

Private Sub Command1_Click(Index As Integer)
    If COVER_Proc() Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer


    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If


    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]見積書作成処理", Me.hwnd, 0)
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
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If

    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ")
        Unload Me
    End Select


    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    Combo1(pcmbSHIMUKE).ListIndex = 0

    Call Init_Proc

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
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
    
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    SEI00101.MousePointer = vbHourglass

    Call Ctrl_Lock(SEI00101)

    For i = pGrdKOSOU To pGrdSAGYO
        TDBGrid1(i).Enabled = False
    Next


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(SEI00101)
    For i = pGrdKOSOU To pGrdSAGYO
        TDBGrid1(i).Enabled = True
    Next


    SEI00101.MousePointer = vbDefault

End Sub


Private Sub SHORI_Click(Index As Integer)
    Select Case Index
    
        
        Case 0      '更新
        
        
            Command1(Index).Value = True
        
        
        Case 1      '終了
        
        
            Command1(Index).Value = True
        
        
        Case 2      '画面印刷
        
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
                    
    
    End Select


End Sub






Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   画面初期化
'----------------------------------------------------------------------------
Dim i           As Integer

Dim Row         As Integer
Dim KOTEI_NO    As Integer

Dim c           As String * 128
                                
Dim wkKOTEI As Variant
                                
                                
                                
                                
                                
    Init_Proc = True
                                
                                
    If SUBETU_Set_Proc(P_KBN06_CD, 1) Then
        Exit Function
    End If
                                
                                
                                
                                '作業工程情報取り込み
    Set SAGYO = Nothing
    
    
    
    
    
    Text1(ptxDEF_LOT).Text = Format(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode), "#0")
    
    
    
    
    Row = 0
    KOTEI_NO = 0
    For i = 1 To 10
        
        If GetIni("KOUTEI", "BEF" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                Row = Row + 1
                SAGYO.ReDim S_Min_Row, Row, S_Min_Col, S_Max_Col
                
                KOTEI_NO = KOTEI_NO + 10
                SAGYO(Row, ColS_No) = Format(KOTEI_NO, "00")
                
                
                
                wkKOTEI = Split(Trim(c), ",", -1)
                
                SAGYO(Row, ColS_NAME) = Trim(wkKOTEI(0))
                SAGYO(Row, ColS_KOUSU) = Trim(wkKOTEI(1))
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
                    SAGYO(Row, ColS_TANKA) = Format(SAGYO(Row, ColS_KOUSU) * CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#,##0.00")
                Else
                    SAGYO(Row, ColS_TANKA) = 0
                End If
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
                    SAGYO(Row, ColS_KIN) = Format(SAGYO(Row, ColS_TANKA) * CLng(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#,##0")
                Else
                    SAGYO(Row, ColS_KIN) = 0
                End If
            End If
        End If
    
    Next i
    
    For i = 1 To 10
        
        If GetIni("KOUTEI", "MAIN" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                Row = Row + 1
                SAGYO.ReDim S_Min_Row, Row, S_Min_Col, S_Max_Col
                
                KOTEI_NO = KOTEI_NO + 10
                SAGYO(Row, ColS_No) = Format(KOTEI_NO, "00")
                
                wkKOTEI = Split(Trim(c), ",", -1)
                
                SAGYO(Row, ColS_NAME) = Trim(wkKOTEI(0))
                SAGYO(Row, ColS_KOUSU) = Trim(wkKOTEI(1))
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
                    SAGYO(Row, ColS_TANKA) = Format(SAGYO(Row, ColS_KOUSU) * CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#,##0.00")
                Else
                    SAGYO(Row, ColS_TANKA) = 0
                End If
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
                    SAGYO(Row, ColS_KIN) = Format(SAGYO(Row, ColS_TANKA) * CLng(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#,##0")
                Else
                    SAGYO(Row, ColS_KIN) = 0
                End If
            End If
        End If
    
    Next i
                                
    For i = 1 To 10
        
        If GetIni("KOUTEI", "AFT" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                Row = Row + 1
                SAGYO.ReDim S_Min_Row, Row, S_Min_Col, S_Max_Col
                
                KOTEI_NO = KOTEI_NO + 10
                SAGYO(Row, ColS_No) = Format(KOTEI_NO, "00")
                
                wkKOTEI = Split(Trim(c), ",", -1)
                
                SAGYO(Row, ColS_NAME) = Trim(wkKOTEI(0))
                SAGYO(Row, ColS_KOUSU) = Trim(wkKOTEI(1))
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
                    SAGYO(Row, ColS_TANKA) = Format(SAGYO(Row, ColS_KOUSU) * CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#,##0.00")
                Else
                    SAGYO(Row, ColS_TANKA) = 0
                End If
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
                    SAGYO(Row, ColS_KIN) = Format(SAGYO(Row, ColS_TANKA) * CLng(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#,##0")
                Else
                    SAGYO(Row, ColS_KIN) = 0
                End If
            End If
        End If
    
    Next i
                                
                                
                                
    Set TDBGrid1(pGrdSAGYO).Array = SAGYO
    
    
    TDBGrid1(pGrdSAGYO).Bookmark = Null
    
    TDBGrid1(pGrdSAGYO).ReBind
    TDBGrid1(pGrdSAGYO).Update
    TDBGrid1(pGrdSAGYO).ScrollBars = dbgAutomatic

    Init_Proc = True


End Function
Private Function SUBETU_Set_Proc(KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   コードマスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
Dim OPTION2     As Integer

Dim wkOption    As String



Dim i           As Integer
    
    SUBETU_Set_Proc = True
    
    Set SYUBETU = Nothing
    
    
    For i = 0 To UBound(P_KBN_TBL)
    
        If KBN = P_KBN_TBL(i).KBN_CD Then
            Key_Len = P_KBN_TBL(i).KBN_Len
            Exit For
        End If
    
    Next i
    
    If i > UBound(P_KBN_TBL) Then
        Exit Function
    End If
    
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    i = 0
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
                Call File_Error(sts, com, "コードマスタ")
                Exit Function
        
        End Select

        wkOption = ""
        If P_KBN_TBL(i).KBN_OP1 Then
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        i = i + 1
        SYUBETU.ReDim 1, i, 0, 0
        
        
        SYUBETU(i, 0) = StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Set TDBDropDown1.Array = SYUBETU
    TDBDropDown1.ReBind

    SUBETU_Set_Proc = False
    



End Function


Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Error_Check_Proc(Index, 0, 0) Then   'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動
End Sub
Private Function Error_Check_Proc(Mode As Integer, chk As Integer, flg) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts         As Integer
    
    
Dim Mi_Qty      As Long
Dim Sumi_Qty    As Long
    
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
    
Dim yn          As Integer
    
    Error_Check_Proc = True
    
    Select Case Mode
    
    
        Case ptxTanto_Code     '担当者
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTanto_Code).Text)

            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTanto_Name).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTanto_Name).Text = ""
            
                    MsgBox "入力した項目はエラーです。(担当者)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
                
            
            
            End Select
        Case ptxHin_Gai         '品番
    
            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)


            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    Text1(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                
                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                        Text1(ptxST_SOKO).Text = ""
                        Text1(ptxST_RETU).Text = ""
                        Text1(ptxST_REN).Text = ""
                        Text1(ptxST_DAN).Text = ""
                    Else
                        Text1(ptxST_SOKO).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                        Text1(ptxST_RETU).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
                        Text1(ptxST_REN).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
                        Text1(ptxST_DAN).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
                    End If
                
                
                Case BtErrKeyNotFound

                    Text1(ptxHin_Name).Text = ""

                    MsgBox "入力した項目はエラーです。(品番)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function

            End Select
    
    
        
        Case ptxDEF_LOT
    
            If IsNumeric(ptxDEF_LOT) Then
    
    
    
                sts = P_COMPO_Disp_Proc()
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
    '                            MsgBox "入力した項目はエラーです。"
    '                            Text1(Mode).SetFocus
    '                            Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                        Exit Function
                End Select
                Text1(Mode + 1).SetFocus        '2008.01.15
            Else
                MsgBox "入力した項目はエラーです。(標準ロット)"
                Text1(Mode).SetFocus
                Exit Function
            End If
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function


Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   コードマスタをコンボにセットする。
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
                Call File_Error(sts, com, "コードマスタ")
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


Private Function P_COMPO_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   構成マスタの読み込み＆表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim k           As Integer
Dim g           As Integer
Dim d           As Integer
    
Dim KOSOU_Row   As Integer
Dim GAISOU_Row  As Integer
Dim DOUKON_Row  As Integer
    
    
    
    P_COMPO_Disp_Proc = True
    Call Input_Lock             '2008.01.15
    
        
    
    
            
    '出力対象
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)

    For i = 0 To UBound(K_Item_Tbl)
        K_Item_Tbl(i).JGYOBU = ""
        K_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(G_Item_Tbl)
        G_Item_Tbl(i).JGYOBU = ""
        G_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(D_Item_Tbl)
        D_Item_Tbl(i).JGYOBU = ""
        D_Item_Tbl(i).NAIGAI = ""
    
    Next i
    

    Set KOSOU = Nothing
    Set GAISOU = Nothing
    Set DOUKON = Nothing

    
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
       
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
       
    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case Else
            '出力対象
            Erase K_Item_Tbl
            Erase G_Item_Tbl
            Erase D_Item_Tbl
            
            ReDim K_Item_Tbl(0 To 4)
            ReDim G_Item_Tbl(0 To 2)
            ReDim D_Item_Tbl(0 To 49)
        
            For i = 0 To UBound(K_Item_Tbl)
                K_Item_Tbl(i).JGYOBU = ""
                K_Item_Tbl(i).NAIGAI = ""
            Next i
        
            For i = 0 To UBound(G_Item_Tbl)
                G_Item_Tbl(i).JGYOBU = ""
                G_Item_Tbl(i).NAIGAI = ""
            Next i
        
            For i = 0 To UBound(D_Item_Tbl)
                D_Item_Tbl(i).JGYOBU = ""
                D_Item_Tbl(i).NAIGAI = ""
            
            Next i
            
            Set KOSOU = Nothing
            Set GAISOU = Nothing
            Set DOUKON = Nothing
            
            
            Call Input_UnLock           '2008.01.15
            P_COMPO_Disp_Proc = sts
            Exit Function
    End Select

    '--------------------------------   「子」情報
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)
    
    KOSOU_Row = 0
    GAISOU_Row = 0
    DOUKON_Row = 0
    
    
    
    Do
        
        
        sts = BTRV(BtOpGetNext, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            
                            
                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                
                    Exit Do
            
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock             '2008.01.15
                Call File_Error(sts, BtOpGetNext, "構成マスタ")
                Exit Function
        
        
        End Select
        
        Select Case StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode)
        
            Case P_KOSOU    '個装資材
            
                k = k + 1
                K_Item_Tbl(k).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                K_Item_Tbl(k).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, K_Item_Tbl(k).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        '品名
                    
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
                        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                    Case Else
                        Call Input_UnLock             '2008.01.15
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
            
            
                KOSOU_Row = KOSOU_Row + 1
                KOSOU.ReDim K_Min_Row, KOSOU_Row, K_Min_Col, K_Max_Col
              
              
              
              
                KOSOU(KOSOU_Row, ColK_HIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                KOSOU(KOSOU_Row, ColK_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                KOSOU(KOSOU_Row, ColK_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "0.00")
                KOSOU(KOSOU_Row, ColK_SHIJI_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) * _
                                                    CLng(Text1(ptxDEF_LOT).Text), "0.00")
                
                If Not IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                End If
                KOSOU(KOSOU_Row, ColK_TANKA) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                KOSOU(KOSOU_Row, ColK_KIN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) * _
                                               KOSOU(KOSOU_Row, ColK_TANKA), "#0.00")
            
            
            
            Case P_GAISOU   '外装資材
                g = g + 1
                G_Item_Tbl(g).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                G_Item_Tbl(g).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, G_Item_Tbl(g).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
                        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                    Case Else
                        Call Input_UnLock             '2008.01.15
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
            
            
            
                GAISOU_Row = GAISOU_Row + 1
                GAISOU.ReDim G_Min_Row, GAISOU_Row, G_Min_Col, G_Max_Col
              
              
              
              
              
              
                GAISOU(GAISOU_Row, ColG_HIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                GAISOU(GAISOU_Row, ColG_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                GAISOU(GAISOU_Row, ColG_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "0.00")
                GAISOU(GAISOU_Row, ColG_SHIJI_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) * _
                                                    CLng(Text1(ptxDEF_LOT).Text), "0.00")
                
                If Not IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                End If
                GAISOU(GAISOU_Row, ColG_TANKA) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                GAISOU(GAISOU_Row, ColG_KIN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) * _
                                               GAISOU(GAISOU_Row, ColK_TANKA), "#0.00")
            
            Case P_DOUKON   '同梱／構成
            
                d = d + 1
                D_Item_Tbl(d).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                D_Item_Tbl(d).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                            
                '種別
                Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(d).JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(d).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        
                        Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
                        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                                    
                    Case Else
                        Call Input_UnLock             '2008.01.15
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
                
        
        
                DOUKON_Row = DOUKON_Row + 1
                DOUKON.ReDim D_Min_Row, DOUKON_Row, D_Min_Col, D_Max_Col
              
              
              
                Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
                Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        
                        Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                                    
                    Case Else
                        Call Input_UnLock             '2008.01.15
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
              
              
              
              
                DOUKON(DOUKON_Row, ColD_SYUBETU) = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
              
                DOUKON(DOUKON_Row, ColD_HIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                DOUKON(DOUKON_Row, ColD_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                DOUKON(DOUKON_Row, ColD_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "0.00")
                DOUKON(DOUKON_Row, ColD_SHIJI_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) * _
                                                    CLng(Text1(ptxDEF_LOT).Text), "0.00")
                
                If Not IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                End If
                DOUKON(DOUKON_Row, ColD_TANKA) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                DOUKON(DOUKON_Row, ColD_KIN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) * _
                                               DOUKON(DOUKON_Row, ColD_TANKA), "#0.00")
        
        
        End Select
        
        
        
        
        com = BtOpGetNext
    
    Loop


    Set TDBGrid1(pGrdKOSOU).Array = KOSOU
    TDBGrid1(pGrdKOSOU).Bookmark = Null
    TDBGrid1(pGrdKOSOU).ReBind
    TDBGrid1(pGrdKOSOU).Update
    TDBGrid1(pGrdKOSOU).ScrollBars = dbgAutomatic

    Set TDBGrid1(pGrdGAISOU).Array = GAISOU
    TDBGrid1(pGrdGAISOU).Bookmark = Null
    TDBGrid1(pGrdGAISOU).ReBind
    TDBGrid1(pGrdGAISOU).Update
    TDBGrid1(pGrdGAISOU).ScrollBars = dbgAutomatic

    Set TDBGrid1(pGrdDOUKON).Array = DOUKON
    TDBGrid1(pGrdDOUKON).Bookmark = Null
    TDBGrid1(pGrdDOUKON).ReBind
    TDBGrid1(pGrdDOUKON).Update
    TDBGrid1(pGrdDOUKON).ScrollBars = dbgAutomatic









    Call Input_UnLock             '2008.01.15

    
    
    P_COMPO_Disp_Proc = False

End Function


Private Function COVER_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（御見積書）出力
'----------------------------------------------------------------------------
Dim excelApplication    As Excel.Application
Dim excelWorkBook       As Excel.Workbook
Dim excelSheet          As Excel.Worksheet

    

    COVER_Proc = True
    
    Call Input_Lock
    



    
    Set excelApplication = CreateObject("Excel.Application")
    excelApplication.Visible = True
    
    Set excelWorkBook = excelApplication.Workbooks.Add
    Set excelSheet = excelWorkBook.Worksheets(1)
    

    
    
    excelApplication.StandardFontSize = 13
    
    excelApplication.StandardFont = "ＭＳ Ｐゴシック"

    
    
    'ページ設定
    With excelSheet.Application.ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    
    '列の幅
    excelSheet.Application.Cells.Select
    excelSheet.Application.Selection.ColumnWidth = 6.25
    excelSheet.Application.Columns(11).Select
    excelSheet.Application.Selection.ColumnWidth = 7.13

    '１行目
    excelSheet.Application.Rows(1).Select
    excelSheet.Application.Selection.RowHeight = 28.5
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 9)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 9)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 9)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 9)).Select
    With excelSheet.Application.Selection.Font
        .Size = 24
    End With
    excelSheet.Application.Cells(1, 5).Value = "御 見 積 書"
    
    '２行目
    excelSheet.Application.Rows(2).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 11), excelSheet.Application.Cells(2, 14)).HorizontalAlignment = xlRight
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 11), excelSheet.Application.Cells(2, 14)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 11), excelSheet.Application.Cells(2, 14)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 11), excelSheet.Application.Cells(2, 14)).Select
    
    With excelSheet.Application.Selection.Font
        .Size = 11
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(2, 11).NumberFormatLocal = "yyyy""年""m""月""d""日"";@"
    excelSheet.Application.Cells(2, 11).Value = Date
    
    '３行目
    excelSheet.Application.Rows(3).Select
    excelSheet.Application.Selection.RowHeight = 17.25
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 6)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    excelSheet.Application.Cells(3, 1).Value = "パナソニックパーツサプライ株式会社"
    
    
    '４行目
    excelSheet.Application.Rows(4).Select
    excelSheet.Application.Selection.RowHeight = 17.25
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 6)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    excelSheet.Application.Cells(4, 1).Value = "伊賀センター倉庫"
    excelSheet.Application.Cells(4, 5).Value = "西谷ＧＭ様"
    
    '５行目
    excelSheet.Application.Rows(5).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    With excelSheet.Application.Selection.Font
        .Size = 9
    End With
    excelSheet.Application.Cells(5, 1).Value = "　下記のとおり御見積りいたしましたので、何卒ご用命"
    excelSheet.Application.Cells(5, 14).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(5, 14).Select
    With excelSheet.Application.Selection.Font
        .Size = 11
    End With
    excelSheet.Application.Cells(5, 14).Value = "株式会社エスディーシィー"
    '６行目
    excelSheet.Application.Rows(6).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    With excelSheet.Application.Selection.Font
        .Size = 9
    End With
    excelSheet.Application.Cells(6, 1).Value = "賜りますよう宜しくお願い申し上げます。"
    excelSheet.Application.Cells(6, 14).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(6, 14).Select
    With excelSheet.Application.Selection.Font
        .Size = 8
    End With
    excelSheet.Application.Cells(6, 14).Value = "〒540-0028 大阪市中央区常磐町１丁目３番８号"
    '７行目
    excelSheet.Application.Rows(7).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    excelSheet.Application.Cells(7, 14).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(7, 14).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ Ｐゴシック"
        .Size = 8
    End With
    excelSheet.Application.Cells(7, 14).Value = "TEL.06 6942 9113 FAX.06 6942 9114"
    '８行目
    excelSheet.Application.Rows(8).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    excelSheet.Application.Cells(8, 11).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(8, 11).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ Ｐゴシック"
        .Size = 11
    End With
    excelSheet.Application.Cells(8, 11).Value = "草津商品化センター"
    '９行目
    excelSheet.Application.Rows(9).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    excelSheet.Application.Cells(9, 11).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(9, 11).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ Ｐゴシック"
        .Size = 7
    End With
    excelSheet.Application.Cells(9, 11).Value = "〒525-0071滋賀県草津市南笠東4-6-8"
    '10行目
    excelSheet.Application.Rows(10).Select
    excelSheet.Application.Selection.RowHeight = 29.5
    excelSheet.Application.Cells(10, 11).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(10, 11).VerticalAlignment = xlTop
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ Ｐゴシック"
        .Size = 7
    End With
    excelSheet.Application.Cells(10, 11).Value = "TEL.077-562-0945  FAX.077-562-0982"
    '９～10行目
    excelSheet.Application.Rows(10).Select
    ActiveSheet.Shapes.AddShape(1, 525#, 117.75, 70.5, 13.5). _
        Select
    Selection.Characters.Text = "承認印"
    With Selection.Characters(Start:=1, Length:=3).Font
        .NAME = "ＭＳ Ｐゴシック"
        .FontStyle = "標準"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .Orientation = xlHorizontal
        .AutoSize = False
        .AddIndent = False
    End With



    

    Set excelSheet = Nothing
    Set excelWorkBook = Nothing
    Set excelApplication = Nothing

End Function
