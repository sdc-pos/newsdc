VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEI00601 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "[請求システム]ミニマム売上入力処理"
   ClientHeight    =   11145
   ClientLeft      =   2010
   ClientTop       =   2535
   ClientWidth     =   18810
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   18810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   11
      Left            =   11025
      TabIndex        =   17
      Top             =   3000
      Width           =   5265
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   10
      Left            =   8610
      TabIndex        =   16
      Top             =   3000
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   8
      Left            =   3675
      TabIndex        =   14
      Top             =   3000
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   7
      Left            =   1155
      TabIndex        =   13
      Top             =   3000
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "伝票削除"
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
      Left            =   4935
      TabIndex        =   25
      Top             =   120
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "伝票終了"
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
      Left            =   1785
      TabIndex        =   23
      Top             =   120
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   4
      Left            =   5250
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   9
      Left            =   6195
      TabIndex        =   15
      Top             =   3000
      Width           =   1380
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   5
      Left            =   10920
      TabIndex        =   12
      Top             =   2520
      Width           =   5370
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      Left            =   2835
      TabIndex        =   11
      Top             =   2520
      Width           =   5370
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   14595
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   11550
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   8610
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1890
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   1920
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   8610
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1440
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   3
      Left            =   3885
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   1155
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1920
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "検  索"
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
      Index           =   5
      Left            =   4515
      TabIndex        =   20
      Top             =   3960
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   13
      Left            =   3045
      TabIndex        =   19
      Top             =   3960
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   12
      Left            =   1365
      TabIndex        =   18
      Top             =   3960
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1365
      TabIndex        =   2
      Top             =   1440
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
      Height          =   375
      Index           =   1
      Left            =   1890
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   3165
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1155
      MaxLength       =   5
      TabIndex        =   0
      Top             =   960
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終　　了"
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
      Index           =   4
      Left            =   6510
      TabIndex        =   26
      Top             =   120
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   12705
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "行削除"
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
      Left            =   3360
      TabIndex        =   24
      Top             =   120
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更  新"
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
      TabIndex        =   22
      Top             =   120
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   6135
      Left            =   315
      TabIndex        =   21
      Top             =   4560
      Width           =   18285
      _ExtentX        =   32253
      _ExtentY        =   10821
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
      Columns(1).Caption=   "請求№"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "計上年月"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "売上先"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "請求区分"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "経営項目"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "部署"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "請求項目（提出用）"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "請求項目（ＳＤＣ用）"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "数量"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "単価"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "金額"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "消費税"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "摘要"
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
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2646"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2514"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2037"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1905"
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
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2805"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2672"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=3387"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3254"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2249"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2117"
      Splits(0)._ColumnProps(40)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(42)=   "Column(10).Width=1826"
      Splits(0)._ColumnProps(43)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(10)._WidthInPix=1693"
      Splits(0)._ColumnProps(45)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(46)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(47)=   "Column(11).Width=1826"
      Splits(0)._ColumnProps(48)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(11)._WidthInPix=1693"
      Splits(0)._ColumnProps(50)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(51)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(52)=   "Column(12).Width=2037"
      Splits(0)._ColumnProps(53)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(12)._WidthInPix=1905"
      Splits(0)._ColumnProps(55)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(56)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(57)=   "Column(13).Width=4366"
      Splits(0)._ColumnProps(58)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(13)._WidthInPix=4233"
      Splits(0)._ColumnProps(60)=   "Column(13).Order=14"
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
      BorderStyle     =   1  '実線
      Caption         =   "摘要"
      Height          =   375
      Index           =   17
      Left            =   10395
      TabIndex        =   45
      Top             =   3000
      Width           =   645
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "消費税"
      Height          =   375
      Index           =   16
      Left            =   7770
      TabIndex        =   44
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "単　価"
      Height          =   375
      Index           =   15
      Left            =   2730
      TabIndex        =   43
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "数　量"
      Height          =   375
      Index           =   14
      Left            =   210
      TabIndex        =   42
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "－"
      Height          =   255
      Index           =   6
      Left            =   4935
      TabIndex        =   41
      Top             =   1560
      Width           =   330
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "金　額"
      Height          =   375
      Index           =   13
      Left            =   5250
      TabIndex        =   40
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "請求項目（ＳＤＣ用）"
      Height          =   375
      Index           =   12
      Left            =   8295
      TabIndex        =   39
      Top             =   2520
      Width           =   2640
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "請求項目（提出用）"
      Height          =   375
      Index           =   11
      Left            =   210
      TabIndex        =   38
      Top             =   2520
      Width           =   2640
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "部　署"
      Height          =   375
      Index           =   10
      Left            =   13545
      TabIndex        =   37
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "経営項目"
      Height          =   375
      Index           =   9
      Left            =   10500
      TabIndex        =   36
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "請求区分"
      Height          =   375
      Index           =   5
      Left            =   7560
      TabIndex        =   35
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "売上先"
      Height          =   375
      Index           =   4
      Left            =   210
      TabIndex        =   34
      Top             =   1920
      Width           =   960
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "計上年月"
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   33
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "請求№"
      Height          =   375
      Index           =   3
      Left            =   2940
      TabIndex        =   32
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "～"
      Height          =   375
      Index           =   8
      Left            =   2730
      TabIndex        =   31
      Top             =   4080
      Width           =   330
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "日付範囲"
      Height          =   375
      Index           =   7
      Left            =   210
      TabIndex        =   30
      Top             =   3960
      Width           =   1170
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "売上日付"
      Height          =   375
      Index           =   1
      Left            =   210
      TabIndex        =   29
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "担当者"
      Height          =   375
      Index           =   0
      Left            =   210
      TabIndex        =   28
      Top             =   960
      Width           =   960
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "更新"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "削除"
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
Attribute VB_Name = "SEI00601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const ptxTanto_Code% = 0        '担当者コード
Private Const ptxTanto_NAME% = 1        '担当者名称
Private Const ptxJITU_DATE% = 2         '売上日付
Private Const ptxDEN_NO% = 3            '伝票№
Private Const ptxGYO_NO% = 4            '伝票№

Private Const ptxKEIJYO_YM% = 5         '計上年月
Private Const ptxUKEHARAI_CODE% = 6     '売上先

Private Const ptxSURYO% = 7             '数量
Private Const ptxTANKA% = 8             '単価
Private Const ptxURI_KIN% = 9           '売上金額
Private Const ptxZEI_KIN% = 10          '消費税額
Private Const ptxTEKIYO% = 11           '摘要

Private Const ptxS_JITU_DATE% = 12      '日付範囲　開始
Private Const ptxE_JITU_DATE% = 13      '日付範囲　終了


Private Const pcmbUKEHARAI% = 0         '売上先
Private Const pcmbSE_KBN% = 1           '請求区分
Private Const pcmbMANA_KBN% = 2         '経営項目
Private Const pcmbPOST_CODE% = 3        '部署
Private Const pcmbSUB_ITEM% = 4         '請求項目（提出用）
Private Const pcmbSDC_ITEM% = 5         '請求項目（ＳＤＣ用）



Dim SE_MIN_URIAGE   As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 13             '最大列数

Private Const ColJITU_DATE% = 0         '注文区分
Private Const ColDEN_NO% = 1            '伝票№
Private Const ColKEIJYO_YM% = 2         '計上年月
Private Const ColUKEHARAI% = 3          '売上先

Private Const ColSE_KBN% = 4            '請求区分
Private Const ColMANA_KBN% = 5          '経営項目
Private Const ColPOST_CODE% = 6         '部署
Private Const ColSUB_ITEM% = 7          '請求項目（提出用）
Private Const ColSDC_ITEM% = 8          '請求項目（SDC用）


Private Const ColSURYO% = 9             '数量
Private Const ColTANKA% = 10            '単価
Private Const ColURI_KIN% = 11          '売上金額
Private Const ColZEI_KIN% = 12          '消費税

Private Const ColTEKIYO% = 13           '摘要


'請求項目
Private Type SE_ITEM_Tag
    No          As Integer
    SUB_ITEM    As String
    SDC_ITEM    As String
End Type
Private SE_ITEM()   As SE_ITEM_Tag

'請求区分
Private Type SE_KBN_Tag
    No          As Integer
    SE_KBN      As String
End Type
Private SE_KBN()    As SE_KBN_Tag

Dim SHIMEBI         As String


'
Private svSURYO     As Double
Private svTANKA     As Double
Private svURI_KIN   As Long
Private svZEI_KIN   As Long

Private Sub Combo1_GotFocus(Index As Integer)


    Select Case Index
        Case ptxSURYO
            If IsNumeric(Text1(Index).Text) Then
                svSURYO = CDbl(Text1(Index).Text)
            End If
        Case ptxTANKA
            If IsNumeric(Text1(Index).Text) Then
                svTANKA = CDbl(Text1(Index).Text)
            End If
        Case ptxURI_KIN
            If IsNumeric(Text1(Index).Text) Then
                svURI_KIN = CLng(Text1(Index).Text)
            End If
        Case ptxZEI_KIN
            If IsNumeric(Text1(Index).Text) Then
                svZEI_KIN = CLng(Text1(Index).Text)
            End If
    End Select
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)


    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Call Tab_Ctrl(Shift)        '移動


End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
    
        Case pcmbUKEHARAI
            Text1(ptxUKEHARAI_CODE).Text = Trim(Right(Combo1(Index).Text, 5))
    
    End Select




End Sub

Private Sub Command1_Click(Index As Integer)

Dim yn  As Integer
Dim i   As Integer







    Select Case Index
    
        Case 0          '更新
        
        
            For i = ptxTanto_Code To ptxTEKIYO
            
                If Error_Check_Proc(Index) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
        
        
            yn = MsgBox("更新しますか？", vbYesNo, "確認入力")
                        
            If yn = vbYes Then
                        
            
                If Update_Proc() Then
                    Unload Me
                End If
            
            
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
                Call Init_Proc(ptxGYO_NO)
            
            
            
            
            End If
        
        
        
        
        Case 1          '伝票終了
        
        
        
            yn = MsgBox("伝票終了しますか？", vbYesNo, "確認入力")
                        
            If yn = vbYes Then
                        
            
                Call Init_Proc(ptxDEN_NO)
            
            
            
            
            End If
        
        
        
        
        
        
        
        Case 2          '行削除
        
        
        
            If Error_Check_Proc(ptxTanto_Code) Then     'エラーチェック
                Exit Sub
            End If
        
            yn = MsgBox("「行」削除しますか？", vbYesNo, "確認入力")
                        
            If yn = vbYes Then
                        
            
                If Delete_Proc() Then
                    Unload Me
                End If
                
            
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
                
                Call Init_Proc(ptxGYO_NO)
            
            
            End If
        
        Case 3          '伝票削除
        
        
        
            If Error_Check_Proc(ptxTanto_Code) Then     'エラーチェック
                Exit Sub
            End If
        
            yn = MsgBox("「伝票」削除しますか？", vbYesNo, "確認入力")
                        
            If yn = vbYes Then
                        
            
                If DEN_Delete_Proc() Then
                    Unload Me
                End If
                
            
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
                
                Call Init_Proc(ptxDEN_NO)
            
            
            End If
        
        
        Case 4          '終了
            Unload Me
    
        Case 5          '検索
    
            If List_Disp_Proc() Then
                Unload Me
            End If
            
            Call Init_Proc(ptxGYO_NO)
    
    
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
    
Dim wkITEM      As Variant
    
Dim S_DATE      As String
Dim E_DATE      As String
Dim S_YY        As String * 4
Dim S_MM        As String * 2
Dim S_DD        As String * 2
    
    
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If


    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]ミニマム売上入力処理", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SETPARTS, 0, -1)


    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)

    Max_Row = 9999
                                

                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
        
    sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "管理マスタ")
            Unload Me
    End Select
                                
                                
                                
                                
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '売上実績ＯＰＥＮ
    If SE_MIN_URIAGE_Open(BtOpenNomal) Then
        Unload Me
    End If



'請求項目取り込み
    i = 0
    Do
        i = i + 1
        If GetIni("ITEM", Format(i, "00"), "SEI_SYS", c) Then
            Exit Do
        End If
        wkITEM = Split(RTrim(c), ",", -1)
    
        ReDim Preserve SE_ITEM(0 To i - 1)
    
        SE_ITEM(i - 1).No = i
        SE_ITEM(i - 1).SUB_ITEM = wkITEM(0)
        SE_ITEM(i - 1).SDC_ITEM = wkITEM(1)
    
    
    Loop

    If GetIni(App.EXEName, "SHIMEBI", App.EXEName, c) Then
        SHIMEBI = ""
    Else
        SHIMEBI = Trim(c)
    End If


'コンボに設定
    Combo1(pcmbSUB_ITEM).Clear

    For i = 0 To UBound(SE_ITEM)
        Combo1(pcmbSUB_ITEM).AddItem Trim(SE_ITEM(i).SUB_ITEM)
    Next i

    Combo1(pcmbSDC_ITEM).Clear

    For i = 0 To UBound(SE_ITEM)
        Combo1(pcmbSDC_ITEM).AddItem Trim(SE_ITEM(i).SDC_ITEM)
    Next i



'請求区分取り込み
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

    Combo1(pcmbSE_KBN).Clear

    For i = 0 To UBound(SE_KBN)
        Combo1(pcmbSE_KBN).AddItem SE_KBN(i).SE_KBN & "                " & Format(SE_KBN(i).No, "00")
    Next i

    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc


    '経営項目のセット
    If Code_Set_Proc(pcmbMANA_KBN, P_KBN09_CD, 0) Then
        Unload Me
    End If
    
    '部署のセット
    If Code_Set_Proc(pcmbPOST_CODE, P_KBN10_CD, 0) Then
        Unload Me
    End If
    
    '受払先
    If Ukeharai_Set_Proc() Then
        Unload Me
    End If


    '初期表示
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
    
    If List_Disp_Proc() Then
        Unload Me
    End If

    Call Init_Proc(ptxDEN_NO)

    Text1(ptxTanto_Code).SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
    
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '売上実績ＣＬＯＳＥ
    sts = BTRV(BtOpClose, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "売上実績")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   指定範囲の売上データを表示する
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim E_DATE      As String
    
    
    List_Disp_Proc = True
                                    
    Call Input_Lock
                                    
                                    
                                    'テーブルリセット
    Set SE_MIN_URIAGE = Nothing
                                    '売上実績読み込み開始
    
    If IsDate(Text1(ptxS_JITU_DATE).Text) Then
        Call UniCode_Conv(K0_SE_MIN_URIAGE.JITU_DATE, Format(Text1(ptxS_JITU_DATE).Text, "YYYYMMDD"))
    Else
        Call UniCode_Conv(K0_SE_MIN_URIAGE.JITU_DATE, Text1(ptxS_JITU_DATE).Text)
    End If
    
    If IsDate(Text1(ptxE_JITU_DATE).Text) Then
        E_DATE = Format(Text1(ptxE_JITU_DATE).Text, "YYYYMMDD")
    Else
        E_DATE = Text1(ptxS_JITU_DATE).Text
    End If
    
    
    Call UniCode_Conv(K0_SE_MIN_URIAGE.DEN_NO, "")
    Call UniCode_Conv(K0_SE_MIN_URIAGE.GYO_NO, "")

        
    
    
    
    
    Row = Min_Row - 1
        
    
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        
        sts = BTRV(com, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
    
    
    
        Select Case sts
            Case BtNoErr
        
                If StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode) > E_DATE Then
                    Exit Do
                End If
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "売上実績")
                Exit Function
        End Select
            
        Row = Row + 1
                    
        If Grid_Set_Proc(Row) Then
            Exit Function
        End If
        
        com = BtOpGetNext
        
    Loop
    
                                
                                'DBテーブルリンク
    Set TDBGrid1.Array = SE_MIN_URIAGE
    
    TDBGrid1.style.Locked = True
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
    
    Call Input_UnLock
    
    
    List_Disp_Proc = False

    
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    SEI00601.MousePointer = vbHourglass

    TDBGrid1.Enabled = False


    Call Ctrl_Lock(SEI00601)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEI00601)

    TDBGrid1.Enabled = True

    SEI00601.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'                   売上データ---＞Grid
'----------------------------------------------------------------------------

Dim sts As Integer
Dim i   As Integer
    
    Grid_Set_Proc = True

    SE_MIN_URIAGE.ReDim Min_Row, Row, Min_Col, Max_Col


    SE_MIN_URIAGE(Row, ColJITU_DATE) = Mid(StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode), 7, 2)
    SE_MIN_URIAGE(Row, ColDEN_NO) = StrConv(SE_MIN_URIAGEREC.DEN_NO, vbUnicode) & "-" & StrConv(SE_MIN_URIAGEREC.GYO_NO, vbUnicode)
    
    SE_MIN_URIAGE(Row, ColKEIJYO_YM) = Mid(StrConv(SE_MIN_URIAGEREC.KEIJYO_YM, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(SE_MIN_URIAGEREC.KEIJYO_YM, vbUnicode), 5, 2)

    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(SE_MIN_URIAGEREC.UKEHARAI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    
    Select Case sts
        Case BtNoErr
    
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払マスタ")
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
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
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
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Function
    End Select
    
    
    SE_MIN_URIAGE(Row, ColPOST_CODE) = StrConv(SE_MIN_URIAGEREC.POST_CODE, vbUnicode) & " " & Trim(StrConv(P_CODEREC.C_NAME, vbUnicode))
    
    SE_MIN_URIAGE(Row, ColSUB_ITEM) = Trim(StrConv(SE_MIN_URIAGEREC.SUB_ITEM, vbUnicode))
    SE_MIN_URIAGE(Row, ColSDC_ITEM) = Trim(StrConv(SE_MIN_URIAGEREC.SDC_ITEM, vbUnicode))
    
    SE_MIN_URIAGE(Row, ColSURYO) = Format(CLng(StrConv(SE_MIN_URIAGEREC.SURYO, vbUnicode)), "#,##0.00")
    SE_MIN_URIAGE(Row, ColTANKA) = Format(CLng(StrConv(SE_MIN_URIAGEREC.TANKA, vbUnicode)), "#,##0.00")
    SE_MIN_URIAGE(Row, ColURI_KIN) = Format(CLng(StrConv(SE_MIN_URIAGEREC.URI_KIN, vbUnicode)), "#,##0")
    SE_MIN_URIAGE(Row, ColZEI_KIN) = Format(CLng(StrConv(SE_MIN_URIAGEREC.ZEI_KIN, vbUnicode)), "#,##0")
    
    SE_MIN_URIAGE(Row, ColTEKIYO) = Trim(StrConv(SE_MIN_URIAGEREC.TEKIYO, vbUnicode))
    

    SE_MIN_URIAGE.ReDim Min_Row, Row, Min_Col, Max_Col
    
    
    Grid_Set_Proc = False
End Function

Private Sub SHORI_Click(Index As Integer)
    Select Case Index
    
        
        Case 0      '更新
        
        
            Command1(Index).Value = True
        
        
        Case 1      '削除
        
        
            Command1(Index).Value = True
        
        Case 2      '終了
        
        
            Command1(Index).Value = True
        
        
        Case 3      '画面印刷
        
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
                    
    
    End Select


End Sub
Private Function Update_Proc() As Integer

'----------------------------------------------------------------------------
'                   更新処理
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
    
Dim DEN_NO  As String
Dim GYO_NO  As String



    Update_Proc = True
    '売上実績の読み込み
    If Trim(Text1(ptxDEN_NO).Text) = "" Or Trim(Text1(ptxGYO_NO).Text) = "" Then
        com = BtOpInsert
    Else
        
        Call UniCode_Conv(K0_SE_MIN_URIAGE.JITU_DATE, Format(Text1(ptxJITU_DATE), "YYYYMMDD"))
        Call UniCode_Conv(K0_SE_MIN_URIAGE.DEN_NO, Text1(ptxDEN_NO).Text)
        Call UniCode_Conv(K0_SE_MIN_URIAGE.GYO_NO, Text1(ptxGYO_NO).Text)
    
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    com = BtOpInsert
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            
                    ans = MsgBox("他端末でデータ使用中です。<SE_MIN_URIAGE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Update_Proc = False
                        Exit Function
                    End If
            
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "売上実績")
                    Exit Function
            End Select

        Loop
    End If
    
    Call UniCode_Conv(SE_MIN_URIAGEREC.JITU_DATE, Format(Text1(ptxJITU_DATE).Text, "YYYYMMDD"))
    If Trim(Text1(ptxDEN_NO).Text) = "" Then

        '管理ファイルより伝票番号の獲得
        Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
        
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Update_Proc = True
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "管理マスタ")
                    Exit Function
            
            End Select
        Loop
        
        '請求書№＋１
        If CLng(StrConv(P_KANRIREC.MIN_URIAGE_NO, vbUnicode)) = 99999999 Then
            Call UniCode_Conv(P_KANRIREC.SASHIZU_NO, "00000001")
        Else
            Call UniCode_Conv(P_KANRIREC.MIN_URIAGE_NO, Format(CLng(StrConv(P_KANRIREC.MIN_URIAGE_NO, vbUnicode)) + 1, "00000000"))
        End If
    
    
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "管理マスタ")
                        End If
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "管理マスタ")
                    Exit Function
            End Select
        Loop
        
        DEN_NO = StrConv(P_KANRIREC.MIN_URIAGE_NO, vbUnicode)
    
        GYO_NO = "001"
    
    
    Else
       If Trim(Text1(ptxGYO_NO).Text) = "" Then
    
            DEN_NO = Text1(ptxDEN_NO).Text
    
            GYO_NO = "001"
        
        
        
        Else
            DEN_NO = Text1(ptxDEN_NO).Text
    
            GYO_NO = Text1(ptxGYO_NO).Text
        
        End If
    
    End If
    
    
    Call UniCode_Conv(SE_MIN_URIAGEREC.DEN_NO, DEN_NO)
    Call UniCode_Conv(SE_MIN_URIAGEREC.GYO_NO, GYO_NO)
    
                                    
    Call UniCode_Conv(SE_MIN_URIAGEREC.KEIJYO_YM, Left(Format(Text1(ptxKEIJYO_YM).Text & "/01", "YYYYMMDD"), 6))
    
    Call UniCode_Conv(SE_MIN_URIAGEREC.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)
    Call UniCode_Conv(SE_MIN_URIAGEREC.SE_KBN, Right(Combo1(pcmbSE_KBN).Text, 2))
    Call UniCode_Conv(SE_MIN_URIAGEREC.MANA_KBN, Right(Combo1(pcmbMANA_KBN).Text, 2))
    Call UniCode_Conv(SE_MIN_URIAGEREC.POST_CODE, Right(Combo1(pcmbPOST_CODE).Text, 2))
                                    
    If IsNumeric(Right(Combo1(pcmbSUB_ITEM).Text, 2)) Then
        Call UniCode_Conv(SE_MIN_URIAGEREC.SUB_ITEM, Left(Combo1(pcmbSUB_ITEM).Text, Len(Combo1(pcmbSUB_ITEM).Text) - 2))
    Else
        Call UniCode_Conv(SE_MIN_URIAGEREC.SUB_ITEM, Combo1(pcmbSUB_ITEM).Text)
    End If
                                    
    If IsNumeric(Right(Combo1(pcmbSDC_ITEM).Text, 2)) Then
        Call UniCode_Conv(SE_MIN_URIAGEREC.SDC_ITEM, Left(Combo1(pcmbSDC_ITEM).Text, Len(Combo1(pcmbSUB_ITEM).Text) - 2))
    Else
        Call UniCode_Conv(SE_MIN_URIAGEREC.SDC_ITEM, Combo1(pcmbSDC_ITEM).Text)
    End If
                                    
    If CLng(Text1(ptxSURYO).Text) < 0 Then
        Call UniCode_Conv(SE_MIN_URIAGEREC.SURYO, Format(CLng(Text1(ptxSURYO).Text), "00000000.00"))
    Else
        Call UniCode_Conv(SE_MIN_URIAGEREC.SURYO, Format(CLng(Text1(ptxSURYO).Text), "000000000.00"))
    End If
                                    
    Call UniCode_Conv(SE_MIN_URIAGEREC.TANKA, Format(CLng(Text1(ptxTANKA).Text), "000000000.00"))
    
    If CLng(Text1(ptxURI_KIN).Text) < 0 Then
        Call UniCode_Conv(SE_MIN_URIAGEREC.URI_KIN, Format(CLng(Text1(ptxURI_KIN).Text), "0000000"))
    Else
        Call UniCode_Conv(SE_MIN_URIAGEREC.URI_KIN, Format(CLng(Text1(ptxURI_KIN).Text), "00000000"))
    End If
    
    If CLng(Text1(ptxZEI_KIN).Text) < 0 Then
        Call UniCode_Conv(SE_MIN_URIAGEREC.ZEI_KIN, Format(CLng(Text1(ptxZEI_KIN).Text), "0000000"))
    Else
        Call UniCode_Conv(SE_MIN_URIAGEREC.ZEI_KIN, Format(CLng(Text1(ptxZEI_KIN).Text), "00000000"))
    End If
    
    Call UniCode_Conv(SE_MIN_URIAGEREC.TEKIYO, Trim(Text1(ptxTEKIYO).Text))
    
    
    Call UniCode_Conv(SE_MIN_URIAGEREC.FILLER, "")
                                    
                                    
    Call UniCode_Conv(SE_MIN_URIAGEREC.UPD_TANTO, Text1(ptxTanto_Code).Text)
    Call UniCode_Conv(SE_MIN_URIAGEREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                                    
                                    
    '売上実績の書き込み
    Do
        sts = BTRV(com, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    ans = MsgBox("他端末でデータ使用中です。<SE_MIN_URIAGE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Update_Proc = False
                        Exit Function
                    End If
        
            Case BtErrDuplicates
            
                If com = BtOpInsert Then
                    GYO_NO = Format(CInt(GYO_NO) + 1, "000")
                    Call UniCode_Conv(SE_MIN_URIAGEREC.GYO_NO, GYO_NO)
                Else
                    Call File_Error(sts, BtOpUpdate, "売上実績")
                    Exit Function
                
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "売上実績")
                Exit Function
        End Select
    Loop
                                        
                                        
    Update_Proc = False
    


End Function

Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   削除処理
'----------------------------------------------------------------------------
Dim sts As Integer
Dim ans As Integer
    
    
    Delete_Proc = True
    '売上実績の読み込み
    If Trim(Text1(ptxDEN_NO).Text) = "" Or Trim(Text1(ptxGYO_NO).Text) = "" Then
        Delete_Proc = False
        Exit Function
    Else
        Call UniCode_Conv(K0_SE_MIN_URIAGE.JITU_DATE, Format(Text1(ptxJITU_DATE).Text, "YYYYMMDD"))
        Call UniCode_Conv(K0_SE_MIN_URIAGE.DEN_NO, Text1(ptxDEN_NO).Text)
        Call UniCode_Conv(K0_SE_MIN_URIAGE.GYO_NO, Text1(ptxGYO_NO).Text)
    
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    Delete_Proc = False
                    Exit Function
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            
                    ans = MsgBox("他端末でデータ使用中です。<SE_MIN_URIAGE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Delete_Proc = False
                        Exit Function
                    End If
            
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "売上実績")
                    Exit Function
            End Select

        Loop
    End If
    
    
    
                                    
                                    
    '売上実績の書き込み
    Do
        sts = BTRV(BtOpDelete, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    ans = MsgBox("他端末でデータ使用中です。<SE_MIN_URIAGE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Delete_Proc = False
                        Exit Function
                    End If
        
            Case Else
                Call File_Error(sts, BtOpDelete, "売上実績")
                Exit Function
        End Select
    Loop
                                        
                                        
    Delete_Proc = False
    


End Function


Private Function DEN_Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   伝票削除処理
'----------------------------------------------------------------------------
Dim sts As Integer
Dim ans As Integer
Dim com As Integer
    
    
    DEN_Delete_Proc = True
    '売上実績の読み込み
    If Trim(Text1(ptxDEN_NO).Text) = "" Or Trim(Text1(ptxGYO_NO).Text) = "" Then
        DEN_Delete_Proc = False
        Exit Function
    Else
        Call UniCode_Conv(K0_SE_MIN_URIAGE.JITU_DATE, Format(Text1(ptxJITU_DATE).Text, "YYYYMMDD"))
        Call UniCode_Conv(K0_SE_MIN_URIAGE.DEN_NO, Text1(ptxDEN_NO).Text)
        Call UniCode_Conv(K0_SE_MIN_URIAGE.GYO_NO, Text1(ptxGYO_NO).Text)
    
        com = BtOpGetGreaterEqual
        
        Do
        
            Do
                sts = BTRV(com + BtSNoWait, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrEOF
                        DEN_Delete_Proc = False
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                
                        ans = MsgBox("他端末でデータ使用中です。<SE_MIN_URIAGE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            DEN_Delete_Proc = False
                            Exit Function
                        End If
                
                
                    Case Else
                        Call File_Error(sts, com, "売上実績")
                        Exit Function
                End Select
    
    
            Loop
    
    
    
            If StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode) <> Format(Text1(ptxJITU_DATE).Text, "YYYYMMDD") Or _
                StrConv(SE_MIN_URIAGEREC.DEN_NO, vbUnicode) <> Text1(ptxDEN_NO).Text Or _
                StrConv(SE_MIN_URIAGEREC.GYO_NO, vbUnicode) <> Text1(ptxGYO_NO).Text Then

                Exit Do
    
            End If
    
    
            '売上実績の書き込み
            Do
                sts = BTRV(BtOpDelete, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            ans = MsgBox("他端末でデータ使用中です。<SE_MIN_URIAGE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                DEN_Delete_Proc = False
                                Exit Function
                            End If
                
                    Case Else
                        Call File_Error(sts, BtOpDelete, "売上実績")
                        Exit Function
                End Select
            Loop
    
    
            com = BtOpGetNext
        Loop
    
    
    
    End If
    
    
    
                                    
                                    
                                        
                                        
    DEN_Delete_Proc = False
    


End Function



Private Function Ukeharai_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   受払先マスタをコンボにセットする。
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
                Call File_Error(sts, com, "受払先マスタ")
                Exit Function
        
        End Select

        
        
        Combo1(pcmbUKEHARAI).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop

    Ukeharai_Set_Proc = False
    



End Function


Private Function Code_Set_Proc(Index As Integer, KBN As String, mode As Integer) As Integer
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
    
    If mode = 1 Then
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


Private Sub TDBGrid1_DblClick()

    Text1(ptxJITU_DATE).Text = SE_MIN_URIAGE(TDBGrid1.Bookmark, ColJITU_DATE)


    Text1(ptxDEN_NO).Text = Left(SE_MIN_URIAGE(TDBGrid1.Bookmark, ColDEN_NO), 8)
    Text1(ptxGYO_NO).Text = Right(SE_MIN_URIAGE(TDBGrid1.Bookmark, ColDEN_NO), 3)

    If Detail_Disp_Proc() Then
        Unload Me
    End If

    Text1(ptxKEIJYO_YM).SetFocus

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
Private Function Error_Check_Proc(mode As Integer) As Integer

'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts     As Integer
    
Dim i       As Integer
Dim ZEI_KIN As Long
    
    
    Error_Check_Proc = True
    
    Select Case mode
    
    
        Case ptxTanto_Code     '担当者ｺｰﾄﾞ
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTanto_Code).Text)
            
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTanto_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTanto_NAME).Text = ""
                    MsgBox "入力した項目はエラーです。(担当者)"
                    Text1(mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            End Select
            
        Case ptxJITU_DATE       '売上日付
            
            If Trim(Text1(ptxJITU_DATE).Text) = "" Then
                Text1(ptxJITU_DATE).Text = Format(Now, "YYYY/MM/DD")
            End If
            
            
            If Not IsDate(Text1(ptxJITU_DATE).Text) Then
                MsgBox "入力した項目はエラーです。(売上日付)"
                Text1(mode).SetFocus
                Exit Function
            End If
    
    
        Case ptxDEN_NO          '伝票№
    
        Case ptxKEIJYO_YM       '計上年月
            
            If Trim(Text1(ptxKEIJYO_YM).Text) = "" Then
                Text1(ptxKEIJYO_YM).Text = Mid(Format(Now, "YYYY/MM/DD"), 1, 7)
            End If
            
            If Not IsDate(Text1(ptxKEIJYO_YM).Text & "/" & "01") Then
                MsgBox "入力した項目はエラーです。(計上年月)"
                Text1(mode).SetFocus
                Exit Function
            End If
    
        Case ptxUKEHARAI_CODE   '売上先
            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)
            
            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            Select Case sts
                Case BtNoErr
                
                    For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
                    
                        If Trim(Text1(ptxUKEHARAI_CODE).Text) = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
                        
                            Combo1(pcmbUKEHARAI).ListIndex = i
                            Exit For
                        
                        End If
                    
                    Next i
                
                
                Case BtErrKeyNotFound
                    Combo1(pcmbUKEHARAI).ListIndex = -1
                    MsgBox "入力した項目はエラーです。(売上先)"
                    Text1(mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                    Exit Function
            End Select
    
    
        Case ptxSURYO       '数量
            If Not IsNumeric(Text1(ptxSURYO).Text) Then
                MsgBox "入力した項目はエラーです。(数量)"
                Text1(mode).SetFocus
                Exit Function
            Else
                Text1(ptxSURYO).Text = Format(CDbl(Text1(ptxSURYO).Text), "#,##0.00")
            
            
                If svSURYO <> CDbl(Text1(ptxSURYO).Text) Then
                    
                    If IsNumeric(Text1(ptxTANKA).Text) Then
            
                        Text1(ptxURI_KIN).Text = Format(CDbl(Text1(ptxSURYO).Text) * CDbl(Text1(ptxTANKA).Text), "#,##0")
                    
                    
                        ZEI_KIN = Fix((CLng(Text1(ptxURI_KIN).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                        CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10)
                    
                    
                        Text1(ptxZEI_KIN).Text = Format(ZEI_KIN, "#,##0")
                    
                    End If
                
                
                End If
            
            End If
    
    
        Case ptxTANKA       '単価
            If Not IsNumeric(Text1(ptxSURYO).Text) Then
                MsgBox "入力した項目はエラーです。(単価)"
                Text1(mode).SetFocus
                Exit Function
            Else
                Text1(ptxTANKA).Text = Format(CDbl(Text1(ptxTANKA).Text), "#,##0.00")
            
            
                If svTANKA <> CDbl(Text1(ptxTANKA).Text) Then
                    
                    If IsNumeric(Text1(ptxSURYO).Text) Then
            
                        Text1(ptxURI_KIN).Text = Format(CDbl(Text1(ptxSURYO).Text) * CDbl(Text1(ptxTANKA).Text), "#,##0")
                    
                    
                        ZEI_KIN = Fix((CLng(Text1(ptxURI_KIN).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                        CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10)
                    
                    
                        Text1(ptxZEI_KIN).Text = Format(ZEI_KIN, "#,##0")
                    
                    End If
                
                
                End If
            
            End If
    
    
    
    
    
        Case ptxURI_KIN     '売上金額
            If Not IsNumeric(Text1(ptxURI_KIN).Text) Then
                MsgBox "入力した項目はエラーです。(金額)"
                Text1(mode).SetFocus
                Exit Function
            Else
                Text1(ptxURI_KIN).Text = Format(CLng(Text1(ptxURI_KIN).Text), "#,##0")
            
                If svURI_KIN <> CLng(Text1(ptxURI_KIN).Text) Then
                    
                    
                
                    ZEI_KIN = Fix((CLng(Text1(ptxURI_KIN).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                    CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10)
                    
                    
                    Text1(ptxZEI_KIN).Text = Format(ZEI_KIN, "#,##0")
                    
                End If
                
                
            
            End If
    
    
        Case ptxZEI_KIN     '消費税
            If Not IsNumeric(Text1(ptxURI_KIN).Text) Then
                MsgBox "入力した項目はエラーです。(消費税)"
                Text1(mode).SetFocus
                Exit Function
            Else
                Text1(ptxZEI_KIN).Text = Format(CLng(Text1(ptxZEI_KIN).Text), "#,##0")
            End If
    
        Case ptxS_JITU_DATE '日付範囲　開始
            
            If Trim(Text1(ptxS_JITU_DATE).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxS_JITU_DATE).Text) Then
                    MsgBox "入力した項目はエラーです。(日付範囲　開始)"
                    Text1(mode).SetFocus
                    Exit Function
                End If
            End If
    
        Case ptxS_JITU_DATE '日付範囲　終了
            
            If Trim(Text1(ptxE_JITU_DATE).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxE_JITU_DATE).Text) Then
                    MsgBox "入力した項目はエラーです。(日付範囲　終了)"
                    Text1(mode).SetFocus
                    Exit Function
                End If
            End If
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function


Private Function Detail_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   売上実績明表示
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim yn      As Integer

Dim i        As Integer

    
    Detail_Disp_Proc = True

    Call UniCode_Conv(K0_SE_MIN_URIAGE.JITU_DATE, Format(Text1(ptxJITU_DATE).Text, "YYYYMMDD"))
    Call UniCode_Conv(K0_SE_MIN_URIAGE.DEN_NO, Text1(ptxDEN_NO).Text)
    Call UniCode_Conv(K0_SE_MIN_URIAGE.GYO_NO, Text1(ptxGYO_NO).Text)


    sts = BTRV(BtOpGetEqual, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            yn = MsgBox("データ内容が変更されています。最新表示を行います。", vbOK, "確認入力")
            If List_Disp_Proc() Then
                Exit Function
            End If
        
            TDBGrid1.SetFocus
            Detail_Disp_Proc = False
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "売上実績")
            Exit Function
    End Select

    Text1(ptxJITU_DATE).Text = Mid(StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode), 7, 2)


    Text1(ptxKEIJYO_YM).Text = Mid(StrConv(SE_MIN_URIAGEREC.KEIJYO_YM, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(SE_MIN_URIAGEREC.KEIJYO_YM, vbUnicode), 5, 2)



    Text1(ptxUKEHARAI_CODE).Text = StrConv(SE_MIN_URIAGEREC.UKEHARAI_CODE, vbUnicode)
    
    Combo1(pcmbUKEHARAI).ListIndex = -1
    For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
    
        If Text1(ptxUKEHARAI_CODE).Text = Right(Combo1(pcmbUKEHARAI).List(i), 5) Then
        
            Combo1(pcmbUKEHARAI).ListIndex = i
            Exit For
        
        End If
    
    Next i

    Combo1(pcmbSE_KBN).ListIndex = -1
    For i = 0 To Combo1(pcmbSE_KBN).ListCount - 1
    
        If StrConv(SE_MIN_URIAGEREC.SE_KBN, vbUnicode) = Right(Combo1(pcmbSE_KBN).List(i), 2) Then
            Combo1(pcmbSE_KBN).ListIndex = i
            Exit For
        End If
    
    Next i

    Combo1(pcmbMANA_KBN).ListIndex = -1
    For i = 0 To Combo1(pcmbMANA_KBN).ListCount - 1
    
        If StrConv(SE_MIN_URIAGEREC.MANA_KBN, vbUnicode) = Right(Combo1(pcmbMANA_KBN).List(i), 2) Then
            Combo1(pcmbMANA_KBN).ListIndex = i
            Exit For
        End If
    
    Next i

    Combo1(pcmbPOST_CODE).ListIndex = -1
    For i = 0 To Combo1(pcmbPOST_CODE).ListCount - 1
    
        If StrConv(SE_MIN_URIAGEREC.POST_CODE, vbUnicode) = Right(Combo1(pcmbPOST_CODE).List(i), 2) Then
            Combo1(pcmbPOST_CODE).ListIndex = i
            Exit For
        End If
    
    Next i


    Combo1(pcmbSUB_ITEM).Text = Trim(StrConv(SE_MIN_URIAGEREC.SUB_ITEM, vbUnicode))
    Combo1(pcmbSDC_ITEM).Text = Trim(StrConv(SE_MIN_URIAGEREC.SDC_ITEM, vbUnicode))


    Text1(ptxSURYO).Text = Format(CDbl(StrConv(SE_MIN_URIAGEREC.SURYO, vbUnicode)), "#,##0.00")
    Text1(ptxTANKA).Text = Format(CDbl(StrConv(SE_MIN_URIAGEREC.TANKA, vbUnicode)), "#,##0.00")


    Text1(ptxURI_KIN).Text = Format(CLng(StrConv(SE_MIN_URIAGEREC.URI_KIN, vbUnicode)), "#,##0")
    Text1(ptxZEI_KIN).Text = Format(CLng(StrConv(SE_MIN_URIAGEREC.ZEI_KIN, vbUnicode)), "#,##0")

    Text1(ptxTEKIYO).Text = Trim(StrConv(SE_MIN_URIAGEREC.TEKIYO, vbUnicode))


    Detail_Disp_Proc = False


End Function

Public Sub Init_Proc(Start_Index As Integer)
'----------------------------------------------------------------------------
'                   画面クリアー
'----------------------------------------------------------------------------
Dim i   As Integer


    For i = Start_Index To ptxTEKIYO
        Text1(i).Text = ""
    Next i

    For i = pcmbUKEHARAI To pcmbSDC_ITEM
        
        If Combo1(i).ListCount > 0 Then
            Combo1(i).ListIndex = 0
        End If
    Next i

    Text1(ptxKEIJYO_YM).SetFocus

End Sub
