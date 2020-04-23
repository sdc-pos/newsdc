VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEM00301 
   Caption         =   "[請求システム]商品化単価マスタメンテナンス"
   ClientHeight    =   11670
   ClientLeft      =   2025
   ClientTop       =   2550
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
   ScaleHeight     =   11670
   ScaleWidth      =   18810
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   10185
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   8190
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1200
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   5565
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1200
      Width           =   1410
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
      Index           =   2
      Left            =   3780
      TabIndex        =   11
      Top             =   360
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
      Height          =   375
      Index           =   1
      Left            =   1785
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2430
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1050
      MaxLength       =   5
      TabIndex        =   0
      Top             =   1200
      Width           =   750
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終　了"
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
      Left            =   2100
      TabIndex        =   7
      Top             =   360
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
      Left            =   420
      TabIndex        =   6
      Top             =   360
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   9135
      Left            =   210
      TabIndex        =   5
      Top             =   2040
      Width           =   18030
      _ExtentX        =   31803
      _ExtentY        =   16113
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "コード"
      Columns(0).DataField=   ""
      Columns(0).DataWidth=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "名　　称"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "個装形態"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "工料 原価"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "工料 売価"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "設定日"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "資材 原価"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "資材 売価"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "メモ"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "設定日"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "更新日時"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "担当者"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1879"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=5371"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5239"
      Splits(0)._ColumnProps(10)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=4366"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=4233"
      Splits(0)._ColumnProps(15)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2117"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1984"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2117"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1984"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2434"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2302"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=8196"
      Splits(0)._ColumnProps(31)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(33)=   "Column(6).Width=2117"
      Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=1984"
      Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(38)=   "Column(7).Width=2117"
      Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=1984"
      Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(43)=   "Column(8).Width=3096"
      Splits(0)._ColumnProps(44)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(8)._WidthInPix=2963"
      Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(47)=   "Column(9).Width=2434"
      Splits(0)._ColumnProps(48)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(9)._WidthInPix=2302"
      Splits(0)._ColumnProps(50)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(51)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(52)=   "Column(10).Width=3149"
      Splits(0)._ColumnProps(53)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(10)._WidthInPix=3016"
      Splits(0)._ColumnProps(55)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(56)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(57)=   "Column(11).Width=2831"
      Splits(0)._ColumnProps(58)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(11)._WidthInPix=2699"
      Splits(0)._ColumnProps(60)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(61)=   "Column(11).Order=12"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=1200"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.bold=0,.fontsize=1200,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=43,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=43,.fontname=ＭＳ ゴシック"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9,.bgcolor=&HFFFF80&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10,.bgcolor=&HFFFF00&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=58,.parent=43,.alignment=0,.bgcolor=&HC0C0C0&"
      _StyleDefs(39)  =   ":id=58,.bold=0,.fontsize=1200,.italic=0,.underline=0,.strikethrough=0"
      _StyleDefs(40)  =   ":id=58,.charset=128"
      _StyleDefs(41)  =   ":id=58,.fontname=ＭＳ ゴシック"
      _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=44"
      _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=45"
      _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=47"
      _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=16,.parent=43,.bgcolor=&HC0C0C0&"
      _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=44"
      _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=45"
      _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=47"
      _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=70,.parent=43,.bgcolor=&HC0C0C0&"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=44"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=45"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=47"
      _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=28,.parent=43,.alignment=1,.bold=0,.fontsize=1200"
      _StyleDefs(54)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(55)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(56)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=44"
      _StyleDefs(57)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=45"
      _StyleDefs(58)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=47"
      _StyleDefs(59)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.bold=0,.fontsize=1200"
      _StyleDefs(60)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(61)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(62)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(63)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(64)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(65)  =   "Splits(0).Columns(5).Style:id=62,.parent=43,.bgcolor=&HC0C0C0&,.locked=-1"
      _StyleDefs(66)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=44"
      _StyleDefs(67)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=45"
      _StyleDefs(68)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=47"
      _StyleDefs(69)  =   "Splits(0).Columns(6).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(70)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=44"
      _StyleDefs(71)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=45"
      _StyleDefs(72)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=47"
      _StyleDefs(73)  =   "Splits(0).Columns(7).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(74)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=44"
      _StyleDefs(75)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=45"
      _StyleDefs(76)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=47"
      _StyleDefs(77)  =   "Splits(0).Columns(8).Style:id=74,.parent=43"
      _StyleDefs(78)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=44"
      _StyleDefs(79)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=45"
      _StyleDefs(80)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=47"
      _StyleDefs(81)  =   "Splits(0).Columns(9).Style:id=32,.parent=43,.bgcolor=&HC0C0C0&"
      _StyleDefs(82)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=44"
      _StyleDefs(83)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=45"
      _StyleDefs(84)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=47"
      _StyleDefs(85)  =   "Splits(0).Columns(10).Style:id=78,.parent=43,.bgcolor=&HC0C0C0&"
      _StyleDefs(86)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=44"
      _StyleDefs(87)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=45"
      _StyleDefs(88)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=47"
      _StyleDefs(89)  =   "Splits(0).Columns(11).Style:id=86,.parent=43,.bgcolor=&HC0C0C0&"
      _StyleDefs(90)  =   "Splits(0).Columns(11).HeadingStyle:id=83,.parent=44"
      _StyleDefs(91)  =   "Splits(0).Columns(11).FooterStyle:id=84,.parent=45"
      _StyleDefs(92)  =   "Splits(0).Columns(11).EditorStyle:id=85,.parent=47"
      _StyleDefs(93)  =   "Named:id=33:Normal"
      _StyleDefs(94)  =   ":id=33,.parent=0"
      _StyleDefs(95)  =   "Named:id=34:Heading"
      _StyleDefs(96)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(97)  =   ":id=34,.wraptext=-1"
      _StyleDefs(98)  =   "Named:id=35:Footing"
      _StyleDefs(99)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(100) =   "Named:id=36:Selected"
      _StyleDefs(101) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(102) =   "Named:id=37:Caption"
      _StyleDefs(103) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(104) =   "Named:id=38:HighlightRow"
      _StyleDefs(105) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(106) =   "Named:id=39:EvenRow"
      _StyleDefs(107) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(108) =   "Named:id=40:OddRow"
      _StyleDefs(109) =   ":id=40,.parent=33"
      _StyleDefs(110) =   "Named:id=41:RecordSelector"
      _StyleDefs(111) =   ":id=41,.parent=34"
      _StyleDefs(112) =   "Named:id=42:FilterBar"
      _StyleDefs(113) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      Caption         =   "品番"
      Height          =   255
      Index           =   3
      Left            =   9660
      TabIndex        =   14
      Top             =   1320
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "個装形態"
      Height          =   255
      Index           =   2
      Left            =   7140
      TabIndex        =   13
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "事業部"
      Height          =   255
      Index           =   1
      Left            =   4725
      TabIndex        =   12
      Top             =   1320
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "担当者"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   10
      Top             =   1320
      Width           =   750
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
      Left            =   210
      TabIndex        =   8
      Top             =   9360
      Width           =   180
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "更新"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "検査"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   3
      End
   End
End
Attribute VB_Name = "SEM00301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxTanto_Code% = 0        '担当者ｺｰﾄﾞ
Private Const ptxTanto_Name% = 1        '担当者名称
Private Const ptxSELK_KEITAI% = 2       '選択　個装形態
Private Const ptxSELHIN_GAI% = 3        '選択　品番


Private Const pcmbJGYOBU% = 0           '事業部


Dim ITEM As New XArrayDB

Private Const Min_Row% = 1              '最小行数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 11             '最大列数


Private Const ColHIN_GAI% = 0           '品番
Private Const ColName% = 1              '名称

Private Const ColK_KEITAI% = 2          '名称


Private Const ColS_KOUSU_GENKA% = 3     '商品化　工数原価
Private Const ColS_KOUSU_BAIKA% = 4     '商品化　工数売価
Private Const ColS_KOUSU_SET_DATE% = 5  '商品化　単価設定日

Private Const ColS_SHIZAI_GENKA% = 6    '商品化　工数原価
Private Const ColS_SHIZAI_BAIKA% = 7    '商品化　工数売価

Private Const ColS_MEMO% = 8            '商品化　メモ


Private Const ColS_SHIZAI_SET_DATE% = 9 '商品化　単価設定日


Private Const ColUPD_DATETIME% = 10      '更新　日時
Private Const ColUPD_TANTO% = 11         '更新　担当者


Private INPUT_mode As Integer

Private Last_Naigai As String * 1       '指定　国内外


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Command1_Click(Index As Integer)

Dim yn  As Integer
Dim i   As Integer


    Select Case Index
    
        Case 0
    
    
            If Not INPUT_mode Then
                Exit Sub
            End If
            
            
            For i = ptxTanto_Code To ptxTanto_Name
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
    
            If Grid_Error_Check_Proc() Then
                Exit Sub
            End If
    
    
            yn = MsgBox("更新を行いますか？", vbYesNo, "確認入力")
    
            If yn = vbYes Then
        
                If Update_Proc() Then
                    Unload Me
                End If
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
            End If
        
        
        
        
        Case 1
    
            Unload Me
    
    
        Case 2
            
            If List_Disp_Proc() Then
                Unload Me
            End If
    
    End Select





End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer




    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If


    'コモンコントロールを初期化する
'    cc.dwSize = Len(cc)
'    cc.dwICC = ICC_BAR_CLASSES
    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]商品化単価マスタメンテナンス", Me.hwnd, 0)
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
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
        
                                
    Combo1(pcmbJGYOBU).Clear
    For i = 0 To UBound(JGYOBU_T)
        Combo1(pcmbJGYOBU).AddItem JGYOBU_T(i).NAME & "                 " & JGYOBU_T(i).CODE
    Next i
                                
                                
                                    
                                'デフォルト事業部取り込み
    If GetIni("JIGYOBU", "DEF_NO", "SYS", c) Then
    Else
    
    
        For i = 0 To Combo1(pcmbJGYOBU).ListCount - 1
            If Trim(c) = Right(Combo1(pcmbJGYOBU).List(i), 1) Then
                Combo1(pcmbJGYOBU).ListIndex = i
                Exit For
            End If
        Next i
    
    
    End If
                                
                                
                                
                                
    Last_Naigai = NAIGAI_NAI
                                
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                



    Text1(ptxTanto_Code).SetFocus


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ内容の表示
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
    
Dim Found_F     As Boolean
    
Dim S_TIME      As String
    
    
    List_Disp_Proc = True
                                    
    Call Input_Lock
                                    
'    Me.MousePointer = vbArrowHourglass
                                    
                        
S_TIME = Format(Now, "HHMMSS")
                        
                        'テーブルリセット
    Set ITEM = Nothing
    Row = Min_Row - 1
        
    Last_JGYOBU = Right(Combo1(pcmbJGYOBU).Text, 1)
        
                                    
                        '品目ﾏｽﾀ読み込み開始
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Last_Naigai)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxSELHIN_GAI).Text)
    
    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If Last_JGYOBU <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                    Last_Naigai <> StrConv(ITEMREC.NAIGAI, vbUnicode) Then
                    Exit Do
                End If
            
            
            
            
                Found_F = True
            
            
                If Trim(Text1(ptxSELHIN_GAI)) = "" Then
                Else
                    For i = 1 To Len(Trim(Text1(ptxSELHIN_GAI).Text))
                    
                        If Mid(Text1(ptxSELHIN_GAI).Text, i, 1) <> Mid(StrConv(ITEMREC.HIN_GAI, vbUnicode), i, 1) Then
                        
                            Exit Do
                        
                        End If
                    
                    Next i
                End If
            
            
                    
                If Trim(Text1(ptxSELK_KEITAI)) = "" Then
                Else
                    For i = 1 To Len(Trim(Text1(ptxSELK_KEITAI).Text))
                    
                        If Mid(Text1(ptxSELK_KEITAI).Text, i, 1) <> Mid(StrConv(ITEMREC.K_KEITAI, vbUnicode), i, 1) Then
                        
                            Found_F = False
                            Exit For
                        
                        End If
                    
                    Next i
                End If
            
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目ﾏｽﾀ")
                List_Disp_Proc = SYS_ERR
                Exit Function
        End Select
            
        
        If Found_F Then
        
        
            Row = Row + 1
                        
                        
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        
        End If
        
Call SendMessageStr(hStatusWnd, SB_SETTEXT, 0, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        com = BtOpGetNext
        
    Loop
                                'DBテーブルリンク
    Set TDBGrid1.Array = ITEM
    
    
    TDBGrid1.Bookmark = Null
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic
    
    If ITEM.Count(1) > 0 Then
        TDBGrid1.MoveFirst
    End If
    
Call SendMessageStr(hStatusWnd, SB_SETTEXT, 0, S_TIME & " " & Format(Now, "HHMMSS"))
    
    
    Call Input_UnLock
    
'    Me.MousePointer = vbDefault
    
    
    
    List_Disp_Proc = False

    
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    SEM00301.MousePointer = vbHourglass

    Call Ctrl_Lock(SEM00301)

    TDBGrid1.Enabled = False

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEM00301)

    TDBGrid1.Enabled = True

    SEM00301.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    ITEM.ReDim Min_Row, Row, Min_Col, Max_Col
    
    '品目ｺｰﾄﾞ
    ITEM(Row, ColHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
    '名称
    ITEM(Row, ColName) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    '個装形態
    ITEM(Row, ColK_KEITAI) = Trim(StrConv(ITEMREC.K_KEITAI, vbUnicode))
    
    '工数 原価
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_GENKA, vbUnicode)) Then
        ITEM(Row, ColS_KOUSU_GENKA) = Format(StrConv(ITEMREC.S_KOUSU_GENKA, vbUnicode), "#0.00")
    Else
        ITEM(Row, ColS_KOUSU_GENKA) = ""
    End If
    '工数 売価
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
        ITEM(Row, ColS_KOUSU_BAIKA) = Format(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode), "#0.00")
    Else
        ITEM(Row, ColS_KOUSU_BAIKA) = ""
    End If
    '単価設定日
    If Trim(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode)) <> "" Then
        ITEM(Row, ColS_KOUSU_SET_DATE) = Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 7, 2)
    Else
        ITEM(Row, ColS_KOUSU_SET_DATE) = ""
    End If
    
    '資材 原価
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_GENKA, vbUnicode)) Then
        ITEM(Row, ColS_SHIZAI_GENKA) = Format(StrConv(ITEMREC.S_SHIZAI_GENKA, vbUnicode), "#0.00")
    Else
        ITEM(Row, ColS_SHIZAI_GENKA) = ""
    End If
    '資材 売価
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
        ITEM(Row, ColS_SHIZAI_BAIKA) = Format(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode), "#0.00")
    Else
        ITEM(Row, ColS_SHIZAI_BAIKA) = ""
    End If
    
    'メモ
    ITEM(Row, ColS_MEMO) = Trim(StrConv(ITEMREC.SE_TANKA_MEMO, vbUnicode))
    
    '単価設定日
    If Trim(StrConv(ITEMREC.S_SHIZAI_SET_DATE, vbUnicode)) <> "" Then
        ITEM(Row, ColS_SHIZAI_SET_DATE) = Mid(StrConv(ITEMREC.S_SHIZAI_SET_DATE, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(ITEMREC.S_SHIZAI_SET_DATE, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(ITEMREC.S_SHIZAI_SET_DATE, vbUnicode), 7, 2)
    Else
        ITEM(Row, ColS_SHIZAI_SET_DATE) = ""
    End If
    
    
    
    '更新日時
    ITEM(Row, ColUPD_DATETIME) = Trim(StrConv(ITEMREC.UPD_DATETIME, vbUnicode))
    '更新担当者
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(ITEMREC.UPD_TANTO, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            ITEM(Row, ColUPD_TANTO) = StrConv(ITEMREC.UPD_TANTO, vbUnicode) & " " & Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
        Case BtErrKeyNotFound
            ITEM(Row, ColUPD_TANTO) = StrConv(ITEMREC.UPD_TANTO, vbUnicode)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    End Select
    
    
    
    Grid_Set_Proc = False
End Function



Private Sub SHORI_Click(Index As Integer)
    Select Case Index
    
        
        Case 0      '更新
        
        
            Command1(Index).Value = True
        
        
        Case 1      '終了
        
        
            Command1(Index).Value = True
        
        Case 2      '検索
        
        
            Command1(Index).Value = True
        
        
        Case 3      '画面印刷
        
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
                    
    
    End Select

End Sub




Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ更新
'----------------------------------------------------------------------------
Dim sts         As Integer
    
Dim i           As Integer
    
Dim com         As Integer
    
Dim CHANGE_Flg  As Boolean
    
    
    
    Update_Proc = True
                                     
    Set TDBGrid1.Array = ITEM
    TDBGrid1.Refresh
    
    TDBGrid1.Update
                                     
    If ITEM.Count(1) < 1 Then
        Update_Proc = False
        Exit Function
    End If
                                     
    Call Input_Lock
                                     
                                     'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
                                    
                                    
                                    
                                    
    For i = 1 To ITEM.Count(1)
                                    
        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, Last_Naigai)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, ITEM(i, ColHIN_GAI))
            
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
            Case BtErrKeyNotFound
                com = BtOpInsert
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        
        Select Case com
            Case BtOpUpdate
               '変更
                CHANGE_Flg = False
        
                If Trim(ITEM(i, ColS_KOUSU_GENKA)) = "" Then
                    ITEM(i, ColS_KOUSU_GENKA) = "00000000.00"
                End If
                
                If Not IsNumeric(StrConv(ITEMREC.S_KOUSU_GENKA, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "00000000.00")
                End If
                If CDbl(ITEM(i, ColS_KOUSU_GENKA)) <> CDbl(StrConv(ITEMREC.S_KOUSU_GENKA, vbUnicode)) Then
                    CHANGE_Flg = True
                    Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, Format(CDbl(ITEM(i, ColS_KOUSU_GENKA)), "00000000.00"))
                End If
                
                If Trim(ITEM(i, ColS_KOUSU_BAIKA)) = "" Then
                    ITEM(i, ColS_KOUSU_BAIKA) = "00000000.00"
                End If
                If Not IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                End If
                If CDbl(ITEM(i, ColS_KOUSU_BAIKA)) <> CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
                    CHANGE_Flg = True
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, Format(CDbl(ITEM(i, ColS_KOUSU_BAIKA)), "00000000.00"))
                    If CDbl(ITEM(i, ColS_KOUSU_BAIKA)) <> 0 Then
                        Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, Format(Now, "YYYYMMDD"))
                    Else
                        Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")
                    End If
                End If
        
        
                If Trim(ITEM(i, ColS_SHIZAI_GENKA)) = "" Then
                    ITEM(i, ColS_SHIZAI_GENKA) = "00000000.00"
                End If
                
                If Not IsNumeric(StrConv(ITEMREC.S_SHIZAI_GENKA, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "00000000.00")
                End If
                If CDbl(ITEM(i, ColS_SHIZAI_GENKA)) <> CDbl(StrConv(ITEMREC.S_SHIZAI_GENKA, vbUnicode)) Then
                    CHANGE_Flg = True
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, Format(CDbl(ITEM(i, ColS_SHIZAI_GENKA)), "00000000.00"))
                End If
                
                If Trim(ITEM(i, ColS_SHIZAI_BAIKA)) = "" Then
                    ITEM(i, ColS_SHIZAI_BAIKA) = "00000000.00"
                End If
                
                If Not IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                End If
                If CDbl(ITEM(i, ColS_SHIZAI_BAIKA)) <> CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
                    CHANGE_Flg = True
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, Format(CDbl(ITEM(i, ColS_SHIZAI_BAIKA)), "00000000.00"))
                    If CDbl(ITEM(i, ColS_SHIZAI_BAIKA)) <> 0 Then
                        Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, Format(Now, "YYYYMMDD"))
                    Else
                        Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")
                    End If
                End If
        
                Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, Trim(ITEM(i, ColS_MEMO)))
        
        
                If CHANGE_Flg Then
                    Call UniCode_Conv(ITEMREC.UPD_TANTO, Text1(ptxTanto_Code).Text)
                    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                End If
        
        
        
                sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        com = BtOpUpdate
                    Case Else
                        Call File_Error(sts, com, "品目マスタ")
                        Exit Function
                End Select
        
        
        
        End Select


    
    
    
    Next i
                                    
                                    
                                    
    Call Input_UnLock
                                        
                                        
End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function



Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    
    
'    Set TDBGrid1.Array = SE_LOC_TANKA_M
'
'    TDBGrid1.Refresh
'
'    TDBGrid1.Update
'
'
'    Select Case ColIndex
'
'        Case ColIN_KOUSU
'
'            If Trim(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) = "" Then
'            Else
'
'                If Not IsNumeric(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) Then
'                Else
'                    SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex) = Format(CDbl(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)), "#0.000")
'
'                    TDBGrid1.Bookmark = Null
'
'                    TDBGrid1.ReBind
'                    TDBGrid1.Update
'                    TDBGrid1.ScrollBars = dbgAutomatic
'                    TDBGrid1.SetFocus
'                End If
'            End If
'
'        Case ColIN_TANKA
'
'            If Trim(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) = "" Then
'            Else
'
'                If Not IsNumeric(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) Then
'                Else
'                    SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex) = Format(CDbl(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)), "#0.00")
'
'                    TDBGrid1.Bookmark = Null
'
'                    TDBGrid1.ReBind
'                    TDBGrid1.Update
'                    TDBGrid1.ScrollBars = dbgAutomatic
'                    TDBGrid1.SetFocus
'                End If
'            End If
'
'
'
'        Case ColOUT_KOUSU
'
'            If Trim(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) = "" Then
'            Else
'
'                If Not IsNumeric(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) Then
'                Else
'                    SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex) = Format(CDbl(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)), "#0.000")'
'
'                    TDBGrid1.Bookmark = Null
'
'                    TDBGrid1.ReBind
'                    TDBGrid1.Update
'                    TDBGrid1.ScrollBars = dbgAutomatic
'                    TDBGrid1.SetFocus
'                End If
'            End If
'
'        Case ColOUT_TANKA
'            If Trim(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) = "" Then
'            Else
'
'                If Not IsNumeric(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) Then
'                Else
'                    SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex) = Format(CDbl(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)), "#0.00")
'
'                    TDBGrid1.Bookmark = Null
'
'                    TDBGrid1.ReBind
'                    TDBGrid1.Update
'                    TDBGrid1.ScrollBars = dbgAutomatic
'                    TDBGrid1.SetFocus
'                End If
'            End If
'        Case ColS_IN_KOUSU
'            If Trim(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) = "" Then
'            Else
'
'                If Not IsNumeric(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) Then
'                Else
'                    SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex) = Format(CDbl(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)), "#0.000")'
'
'                    TDBGrid1.Bookmark = Null
'
'                    TDBGrid1.ReBind
'                    TDBGrid1.Update
'                    TDBGrid1.ScrollBars = dbgAutomatic
'                    TDBGrid1.SetFocus
'                End If
'            End If
'        Case ColS_OUT_KOUSU
'            If Trim(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) = "" Then
'            Else
'
'                If Not IsNumeric(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)) Then
'                Else
'                    SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex) = Format(CDbl(SE_LOC_TANKA_M(TDBGrid1.Bookmark, ColIndex)), "#0.000")
'
'                    TDBGrid1.Bookmark = Null
'
'                    TDBGrid1.ReBind
'                    TDBGrid1.Update
'                    TDBGrid1.ScrollBars = dbgAutomatic
'                    TDBGrid1.SetFocus
'                End If
'            End If
'
'    End Select



End Sub

Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)
    
    ITEM.ReDim Min_Row, ITEM.Count(1), Min_Col, Max_Col

End Sub

Private Sub TDBGrid1_Change()
    
    INPUT_mode = True

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
    
    
        Case ptxTanto_Code     '担当者ｺｰﾄﾞ
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTanto_Code).Text)
            
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTanto_Name).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTanto_Name).Text = ""
                    MsgBox "入力した項目はエラーです。(担当者)"
                    Text1(ptxTanto_Code).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            End Select
            
            
        Case ptxSELK_KEITAI     '選択　個装形態
        Case ptxSELHIN_GAI      '選択　品番
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function

Private Function Grid_Error_Check_Proc() As Integer
'----------------------------------------------------------------------------
'                   グリッド入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim i   As Integer
Dim sts As Integer
    
    Grid_Error_Check_Proc = True
    
    
    
    
    Set TDBGrid1.Array = ITEM
    
'    TDBGrid1.Refresh
    
    TDBGrid1.Update
    
    If ITEM.Count(1) < 1 Then
        Grid_Error_Check_Proc = False
        Exit Function
    End If
    
    
    
    For i = 1 To ITEM.Count(1)
        
        
        
        
        
        If Trim(ITEM(i, ColS_KOUSU_GENKA)) = "" Then
        Else
            If IsNumeric(ITEM(i, ColS_KOUSU_GENKA)) Then
                ITEM(i, ColS_KOUSU_GENKA) = Format(CDbl(ITEM(i, ColS_KOUSU_GENKA)), "#0.00")
            Else
                MsgBox "入力した項目は、エラーです。（工数 原価）"
                
                TDBGrid1.Bookmark = i
                TDBGrid1.Col = ColS_KOUSU_GENKA
                TDBGrid1.SetFocus
                
                
                Exit Function
            End If
        End If
    
        If Trim(ITEM(i, ColS_KOUSU_BAIKA)) = "" Then
        Else
            If IsNumeric(ITEM(i, ColS_KOUSU_BAIKA)) Then
                ITEM(i, ColS_KOUSU_BAIKA) = Format(CDbl(ITEM(i, ColS_KOUSU_BAIKA)), "#0.00")
            Else
                MsgBox "入力した項目は、エラーです。（工数　売価）"
                
                TDBGrid1.Bookmark = i
                TDBGrid1.Col = ColS_KOUSU_BAIKA
                TDBGrid1.SetFocus
                
                
                Exit Function
            End If
        End If
    
        If Trim(ITEM(i, ColS_SHIZAI_GENKA)) = "" Then
        Else
            If IsNumeric(ITEM(i, ColS_SHIZAI_GENKA)) Then
                ITEM(i, ColS_SHIZAI_GENKA) = Format(CDbl(ITEM(i, ColS_SHIZAI_GENKA)), "#0.00")
            Else
                MsgBox "入力した項目は、エラーです。（資材 原価）"
                
                
                TDBGrid1.Bookmark = i
                TDBGrid1.Col = ColS_SHIZAI_GENKA
                TDBGrid1.SetFocus
                
                Exit Function
            End If
        End If
    
        If Trim(ITEM(i, ColS_SHIZAI_BAIKA)) = "" Then
        Else
            If IsNumeric(ITEM(i, ColS_SHIZAI_BAIKA)) Then
                ITEM(i, ColS_SHIZAI_BAIKA) = Format(CDbl(ITEM(i, ColS_SHIZAI_BAIKA)), "#0.00")
            Else
                MsgBox "入力した項目は、エラーです。（資材 売価）"
                
                TDBGrid1.Bookmark = i
                TDBGrid1.Col = ColS_SHIZAI_BAIKA
                TDBGrid1.SetFocus
                
                
                Exit Function
            End If
        End If
    
        
        
        
        
    Next i


    Grid_Error_Check_Proc = False

End Function
