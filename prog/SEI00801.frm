VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEI00801 
   Caption         =   "[請求システム]ミニマム集計表作成処理"
   ClientHeight    =   11145
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   15510
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
   ScaleWidth      =   15510
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   1470
      TabIndex        =   9
      Top             =   840
      Width           =   1380
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2835
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   840
      Width           =   4845
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
      Index           =   3
      Left            =   4620
      TabIndex        =   7
      Top             =   1200
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   3150
      TabIndex        =   6
      Top             =   1320
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   1470
      TabIndex        =   4
      Top             =   1320
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   10320
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
      Left            =   1890
      TabIndex        =   1
      Top             =   120
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
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   120
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7335
      Index           =   0
      Left            =   315
      TabIndex        =   11
      Top             =   2040
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   12938
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "売上先"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "請求区分"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "経営項目"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "部署"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "請求項目（提出用）"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "請求項目（ＳＤＣ用）"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "9999/99"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   18
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=18"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4630"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4498"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1879"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1746"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1879"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1746"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1244"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1111"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=4128"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3995"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=4842"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=4710"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2249"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2117"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2249"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2117"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=2249"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2117"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2249"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2117"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=2249"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2117"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=2249"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2117"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=2249"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=2117"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(53)=   "Column(13).Width=2249"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=2117"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=2249"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=2117"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=2249"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=2117"
      Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(65)=   "Column(16).Width=2249"
      Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=2117"
      Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(69)=   "Column(17).Width=2249"
      Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=2117"
      Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
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
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=110,.parent=87"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=107,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=108,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=109,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=114,.parent=87"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=111,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=112,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=113,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=118,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=115,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=116,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=117,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=16,.parent=87"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=13,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=14,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=15,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=20,.parent=87"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=17,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=18,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=19,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=24,.parent=87"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=87"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=87"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=46,.parent=87"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=50,.parent=87"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=47,.parent=88"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=48,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=49,.parent=91"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=54,.parent=87"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=88"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=89"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=91"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=58,.parent=87"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=55,.parent=88"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=56,.parent=89"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=57,.parent=91"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=62,.parent=87"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=59,.parent=88"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=60,.parent=89"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=61,.parent=91"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=66,.parent=87"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=63,.parent=88"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=64,.parent=89"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=65,.parent=91"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=70,.parent=87"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=67,.parent=88"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=68,.parent=89"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=69,.parent=91"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=74,.parent=87"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=71,.parent=88"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=72,.parent=89"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=73,.parent=91"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=78,.parent=87"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=75,.parent=88"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=76,.parent=89"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=77,.parent=91"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=82,.parent=87"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=79,.parent=88"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=80,.parent=89"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=81,.parent=91"
      _StyleDefs(108) =   "Named:id=33:Normal"
      _StyleDefs(109) =   ":id=33,.parent=0"
      _StyleDefs(110) =   "Named:id=34:Heading"
      _StyleDefs(111) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(112) =   ":id=34,.wraptext=-1"
      _StyleDefs(113) =   "Named:id=35:Footing"
      _StyleDefs(114) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(115) =   "Named:id=36:Selected"
      _StyleDefs(116) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(117) =   "Named:id=37:Caption"
      _StyleDefs(118) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(119) =   "Named:id=38:HighlightRow"
      _StyleDefs(120) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(121) =   "Named:id=39:EvenRow"
      _StyleDefs(122) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(123) =   "Named:id=40:OddRow"
      _StyleDefs(124) =   ":id=40,.parent=33"
      _StyleDefs(125) =   "Named:id=41:RecordSelector"
      _StyleDefs(126) =   ":id=41,.parent=34"
      _StyleDefs(127) =   "Named:id=42:FilterBar"
      _StyleDefs(128) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "売上先"
      Height          =   375
      Index           =   4
      Left            =   315
      TabIndex        =   10
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "～"
      Height          =   375
      Index           =   8
      Left            =   2835
      TabIndex        =   5
      Top             =   1320
      Width           =   330
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "表示範囲"
      Height          =   375
      Index           =   7
      Left            =   315
      TabIndex        =   3
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "更新"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   2
      End
   End
End
Attribute VB_Name = "SEI00801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit







Dim SE_LOC_TANKA_M As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 17             '最大列数

Private Const ColCYU_KBN% = 0           '注文区分
Private Const ColMUKE_CODE% = 1         '出荷先

Private Const ColID_NO% = 2             'ID№
Private Const ColDEN_NO% = 3            '伝票№
Private Const ColSYUKO_SYUSI& = 4       '出庫収支
Private Const ColHIN_GAI% = 5           '品番（外部）
Private Const ColHIN_NAME% = 6         '品名
Private Const ColYOTEI_QTY% = 7         '出荷予定数
Private Const ColFIX_QTY% = 8           '出荷実績
Private Const ColKENPIN_MARK% = 9       '検品
Private Const ColDEN_DT% = 10            '伝票日付
Private Const ColSort_Mark% = 11         'ＳＯＲＴマーク
Private Const ColPrint% = 12            '出庫表印刷マーク
Private Const ColIns_Date% = 13         '取込み日時

Private Const ColKENPIN_Date% = 14      '検品日
Private Const ColKENPIN_TANTO% = 15     '検品担当者

Private Const ColLK_SEQ_NO% = 16        'ﾘﾝｸ№

Private Const ColJGYOBU% = 17           '事業部


Private Const Sort_MISYUKO$ = "0"       '未出庫
Private Const Sort_SYUKOSUMI$ = "1"     '出庫済
Private Const Sort_KENPIN$ = "2"        '検品済

Private Const KENPIN_ON$ = "○"         '検品済
Private Const KENPIN_OFF$ = "×"        '未検品


Private Inspe_F As Integer              '検品方法



Private Sub Command_Click(index As Integer)

Dim ans As Integer

    Select Case index
        Case 7                              '再表示
            Text(ptxMUKE_CODE).Text = Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8)
            If List_Disp_Proc Then
                Unload Me
            End If
        Case 8                              'データ出力
        
            Beep
            ans = MsgBox("「出荷予定」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If ans = vbYes Then
                If OUTPUT_Proc() Then
                    Unload Me
                End If
            End If
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub
Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
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
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    
Dim cc As tagINITCOMMONCONTROLSEX
Dim PanePos(2) As Long

    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If


    'コモンコントロールを初期化する
    cc.dwSize = Len(cc)
    cc.dwICC = ICC_BAR_CLASSES
    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]ミニマム集計表作成処理", Me.hwnd, 0)
    'ペイン複数作る
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    PanePos(0) = 200
    PanePos(1) = 300
    PanePos(2) = -1
    Call SendMessageAny(hStatusWnd, SB_SETPARTS, 3, PanePos(0))


    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                


    Max_Row = 9999
                                

                                '倉庫マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'ロケーション別単価設定マスタＯＰＥＮ
    If SE_SHOHIN_TANKA_M_Open(BtOpenNomal) Then
        Unload Me
    End If

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
    
                                            '商品化単価マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SE_SHOHIN_TANKA_M_POS, SE_SHOHIN_TANKA_M_REC, Len(SE_SHOHIN_TANKA_M_REC), K0_SE_SHOHIN_TANKA_M, Len(K0_SE_SHOHIN_TANKA_M), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ロケーション別単価設定マスタ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub SubMenu_Click(index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        subMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1030611.Caption = "出荷確認（" + RTrim(JGYOBU_T(index).NAME) + ")"
    subMenu(index).Checked = True
    If Last_JGYOBU <> JGYOBU_T(index).CODE Then
        Last_JGYOBU = JGYOBU_T(index).CODE
        LabJIGYO.Caption = RTrim(JGYOBU_T(index).NAME)
        LabJIGYO.ForeColor = QBColor(JGYOBU_T(index).COLOR)

    End If

End Sub

Private Function MTS_Set_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim Edit        As String

    MTS_Set_Proc = True
    
    Call Input_Lock
    
    
    Combo(pcmbMUKE_CODE).Clear
    
    Combo(pcmbMUKE_CODE).AddItem "全て　　　" & "   " & Space(16)
        
    
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "向け先マスタ")
                Exit Function
        End Select
        
        Edit = StrConv(MTSREC.MUKE_NAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        
        
        Combo(pcmbMUKE_CODE).AddItem Edit
    
        com = BtOpGetNext
    
    Loop

    If Combo(pcmbMUKE_CODE).ListCount <= 0 Then
    Else
        Combo(pcmbMUKE_CODE).ListIndex = 0
    End If

    Call Input_UnLock

    MTS_Set_Proc = False
End Function


Private Function List_Disp_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim DEN_MAISU   As Long
Dim KAN_MAISU   As Long
    
Dim Skip_Flg    As Boolean
    
    
    List_Disp_Proc = True
                                    
'    Call Input_Lock
                                    
    Me.MousePointer = vbArrowHourglass
                                    
                                    'テーブルリセット
    Set SYUKA = Nothing
                                    '出荷予定読み込み開始
    
    If Last_JGYOBU = "*" Then
        Call UniCode_Conv(K2_Y_SYU.JGYOBU, "") '事業部
    Else
        Call UniCode_Conv(K2_Y_SYU.JGYOBU, Last_JGYOBU) '事業部
    End If
                                                    '注文区分
    Call UniCode_Conv(K2_Y_SYU.KEY_CYU_KBN, "")
                                                    '向け先
    Call UniCode_Conv(K2_Y_SYU.KEY_MUKE_CODE, "")
    Call UniCode_Conv(K2_Y_SYU.KEY_SS_CODE, "")
    
    
    Row = Min_Row - 1
        
    DEN_MAISU = 0
    KAN_MAISU = 0
    
    
    
    com = BtOpGetGreaterEqual
    
''com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)
    
    
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定")
                List_Disp_Proc = SYS_ERR
                Exit Function
        End Select
                                '事業部 KEYﾌﾞﾚｰｸ
        
        If Last_JGYOBU = "*" Then
        Else
            If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Exit Do
            End If
        End If
                                
        Skip_Flg = False
                                
                                '注文区分 KEYﾌﾞﾚｰｸ
        If Len(Trim(Right(Combo(pcmbCYU_KBN).Text, 1))) <> 0 Then
            If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCYU_KBN).Text, 1) Then
                Skip_Flg = True
            End If
        End If
                            '向け先 KEYﾌﾞﾚｰｸ
    
    
        If Len(Trim(Right(Combo(pcmbMUKE_CODE).Text, 16))) <> 0 Then
            If StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) <> Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8) Or _
                StrConv(Y_SYUREC.SS_CODE, vbUnicode) <> Right(Combo(pcmbMUKE_CODE).Text, 8) Then
                Skip_Flg = True
            End If
        End If
        
        If Len(Trim((Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text))) = 0 Then
        Else
            If (Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text) <> StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then
                Skip_Flg = True
            End If
        End If
        'ﾃﾞｰﾀ区分
        If Trim(Text(ptxDATA_KBN).Text) = "" Then
        Else
            
            If Trim(Text(ptxDATA_KBN).Text) = "*" Then
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" Or StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "3" Then
                Else
                    Skip_Flg = True
                End If
            Else
                If Text(ptxDATA_KBN).Text <> StrConv(Y_SYUREC.DATA_KBN, vbUnicode) Then
                    Skip_Flg = True
                End If
            End If
        End If
        '販売区分
        If Trim(Text(ptxHAN_KBN).Text) = "" Then
        Else
            If Text(ptxHAN_KBN).Text <> StrConv(Y_SYUREC.HAN_KBN, vbUnicode) Then
                Skip_Flg = True
            End If
        End If
                
                
        If Not Skip_Flg Then
            DEN_MAISU = DEN_MAISU + 1
            
                                        '検品完了
            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                KAN_MAISU = KAN_MAISU + 1
            End If
            
            Row = Row + 1
            If Row > Max_Row Then
                Beep
                MsgBox "最大表示行数を超えました。"
                Exit Do
            End If
                    
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
        
    Loop
    
                                
                                'DBテーブルリンク
    If DEN_MAISU < 1 Then
    Else
        SYUKA.QuickSort Min_Row, (SYUKA.UpperBound(1)), ColSort_Mark, XORDER_ASCEND, XTYPE_STRING, _
                                                        ColDEN_NO, XORDER_ASCEND, XTYPE_STRING
    End If
    
    Set TDBGrid1.Array = SYUKA
    
    TDBGrid1.style.Locked = True
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
    Text(ptxDEN_MAISU_JI).Text = Format(KAN_MAISU, "#,##0")
                                
    Text(ptxDEN_MAISU_YO).Text = Format(DEN_MAISU, "#,##0")
    
'    Call Input_UnLock
    
    Me.MousePointer = vbDefault
    
    
    Combo(pcmbMUKE_CODE).SetFocus
    
    List_Disp_Proc = False

    
End Function

Private Function OUTPUT_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

    
Dim Ret         As Integer
    

Dim FileNo      As Integer
Dim fileName    As String
    
Dim Skip_Flg    As Boolean
    
    
    OUTPUT_Proc = True
                                    
'    Call Input_Lock

    FileNo = FreeFile
    
    fileName = SYUKA_DATA
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (SYUKA_DATA) For Output As FileNo

    Write #FileNo, "注文区分", "出荷先", "ＩＤ№", "伝票№", "品番（外部）", "品番（内部）", "品名", "出荷予定数", "済み数", "検品", "伝票日付", Format(Now, "yyyy/mm/dd HH:mm:ss") & " 現在"

                                    '出荷予定読み込み開始
    Call UniCode_Conv(K2_Y_SYU.JGYOBU, Last_JGYOBU) '事業部
    
                                                    '注文区分
    Call UniCode_Conv(K2_Y_SYU.KEY_CYU_KBN, "")
                                                    '向け先
    Call UniCode_Conv(K2_Y_SYU.KEY_MUKE_CODE, "")
    Call UniCode_Conv(K2_Y_SYU.KEY_SS_CODE, "")
    
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定")
                OUTPUT_Proc = SYS_ERR
                Exit Function
        End Select
                                '事業部 KEYﾌﾞﾚｰｸ
        If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
            Exit Do
        End If
        
        Skip_Flg = False
        If Len(Trim(Right(Combo(pcmbCYU_KBN).Text, 1))) <> 0 Then
            If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCYU_KBN).Text, 1) Then
                Skip_Flg = True
            End If
        End If
                            '向け先 KEYﾌﾞﾚｰｸ
    
    
        If Len(Trim(Right(Combo(pcmbMUKE_CODE).Text, 16))) <> 0 Then
            If StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) <> Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8) Or _
                StrConv(Y_SYUREC.SS_CODE, vbUnicode) <> Right(Combo(pcmbMUKE_CODE).Text, 8) Then
                Skip_Flg = True
            End If
        End If
        
        If Len(Trim((Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text))) = 0 Then
        Else
            If (Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text) <> StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then
                Skip_Flg = True
            End If
        End If
        
        
        If Not Skip_Flg Then
            Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                Case CYU_KBN_TUK
                    Write #FileNo, CYU_KBN_1,
                Case CYU_KBN_SPO
                    Write #FileNo, CYU_KBN_2,
                Case CYU_KBN_HJU
                    Write #FileNo, CYU_KBN_3,
                Case CYU_KBN_TOK
                    Write #FileNo, CYU_KBN_4,
                Case CYU_KBN_BOU
                    Write #FileNo, CYU_KBN_E,
                Case CYU_KBN_KIN
                    Write #FileNo, CYU_KBN_T,
            End Select
            
            
            Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
            Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
            Select Case sts
                Case BtNoErr
                    Write #FileNo, StrConv(MTSREC.MUKE_DNAME, vbUnicode),
                Case BtErrKeyNotFound
                    Write #FileNo, ,
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
                    Exit Function
            End Select
            
            
            
            Write #FileNo, StrConv(Y_SYUREC.ID_NO, vbUnicode),
            Write #FileNo, StrConv(Y_SYUREC.DEN_NO, vbUnicode),
            Write #FileNo, StrConv(Y_SYUREC.HIN_NO, vbUnicode),
    '2004        Write #FileNo, StrConv(Y_SYUREC.HIN_NAI, vbUnicode),
                                    '品目マスタ読込み
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Write #FileNo, StrConv(ITEMREC.HIN_NAI, vbUnicode),
                    Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
                Case BtErrKeyNotFound
                    Write #FileNo,
                    Write #FileNo,
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
                                                                        '出荷予定数
            Write #FileNo, Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#,##0"),
                                                                        '出荷実績数
            Write #FileNo, Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "#,##0"),
                                                                        '検品マーク
            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) = 0 Then
                                    '未検品
                Write #FileNo, KENPIN_OFF,
            Else
                                    '検品済
                Write #FileNo, KENPIN_ON,
            End If
                
            Write #FileNo, Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" _
                            & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" _
                            & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 7, 2)
        End If
        com = BtOpGetNext
        
        DoEvents
    Loop

    Close #FileNo
    
'    Call Input_UnLock         '画面項目ロック解除
    
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"

    Combo(pcmbMUKE_CODE).SetFocus
    
    OUTPUT_Proc = False
    
    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If


End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1030611.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030611)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030611)


    F1030611.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
    
    Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
        Case CYU_KBN_TUK
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_1
        Case CYU_KBN_SPO
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_2
        Case CYU_KBN_HJU
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_3
        Case CYU_KBN_TOK
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_4
        Case CYU_KBN_BOU
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_E
        Case CYU_KBN_KIN
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_T
    End Select
    
    
    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColMUKE_CODE) = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.MUKE_DNAME, vbUnicode)
        Case BtErrKeyNotFound
            SYUKA(Row, ColMUKE_CODE) = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
            Exit Function
    End Select
    
    
    
    SYUKA(Row, ColID_NO) = StrConv(Y_SYUREC.ID_NO, vbUnicode)       'ＩＤ№
    SYUKA(Row, ColDEN_NO) = StrConv(Y_SYUREC.DEN_NO, vbUnicode)     '伝票№
    SYUKA(Row, ColSYUKO_SYUSI) = StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)   '出庫収支

    SYUKA(Row, ColHIN_GAI) = StrConv(Y_SYUREC.HIN_NO, vbUnicode)        '品番（外部）
    SYUKA(Row, ColLK_SEQ_NO) = StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)   '上位ﾘﾝｸ用連番
                                                                    '品目マスタ読込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
                                                                    '出荷予定数
    SYUKA(Row, ColYOTEI_QTY) = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#,##0")
                                                                    '出荷実績数
    SYUKA(Row, ColFIX_QTY) = Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "#,##0")
                                                                    '検品マーク
    If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) = 0 Then
                                '未検品
        SYUKA(Row, ColKENPIN_MARK) = KENPIN_OFF
    Else
                                '検品済
        SYUKA(Row, ColKENPIN_MARK) = KENPIN_ON
    End If
            
    SYUKA(Row, ColDEN_DT) = Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" _
                            & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" _
                            & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 7, 2)
    
    If CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) > CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) Then
                                '未出庫　または　出庫中
        SYUKA(Row, ColSort_Mark) = Sort_MISYUKO
    Else
                                '出庫完了　で　未検品
        If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) = 0 Then
            SYUKA(Row, ColSort_Mark) = Sort_SYUKOSUMI
        Else
            SYUKA(Row, ColSort_Mark) = Sort_KENPIN
        End If
    End If
    
    If Len(Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))) = 0 Then
            SYUKA(Row, ColPrint) = ""
    Else
            SYUKA(Row, ColPrint) = "○"
    End If
    If Trim(StrConv(Y_SYUREC.INS_NOW, vbUnicode)) <> "" Then
        SYUKA(Row, ColIns_Date) = Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 1, 4) & "/" _
                                    & Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 5, 2) & "/" _
                                    & Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 7, 2) & " " _
                                    & Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 9, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 11, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 13, 2)

    Else
        SYUKA(Row, ColIns_Date) = ""
    End If
    
    
    If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) <> "" Then
        SYUKA(Row, ColKENPIN_Date) = Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 1, 4) & "/" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 5, 2) & "/" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 7, 2) & " " _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 1, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 3, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 5, 2)

    Else
        SYUKA(Row, ColKENPIN_Date) = ""
    End If
    
    
    If Trim(StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode)) = "POS" Then
        SYUKA(Row, ColKENPIN_TANTO) = "出荷確認画面"
    Else
        Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                Exit Function
        End Select
        
        
        SYUKA(Row, ColKENPIN_TANTO) = StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode) & " " & StrConv(TANTOREC.TANTO_NAME, vbUnicode)
    End If
    
    
    SYUKA(Row, ColJGYOBU) = StrConv(Y_SYUREC.JGYOBU, vbUnicode)
    
    Grid_Set_Proc = False
End Function

Private Sub TDBGrid1_DblClick(index As Integer)

    If TDBGrid1.Bookmark = -1 Then
    Else
    
        If KENPIN_Update_Proc() Then
            Unload Me
        End If
    End If
    '再表示
'    If List_Disp_Proc Then
'        Unload Me
'    End If


End Sub

Private Sub TDBGrid1_HeadClick(index As Integer, ByVal ColIndex As Integer)
    TDBGrid1.Bookmark = -1
End Sub

Private Sub Text_GotFocus(index As Integer)
    
    If Text(index).TabStop = True Then
        Text(index) = Trim(Text(index).Text)
        Text(index).SelStart = 0
        Text(index).SelLength = Len(Text(index).Text)
    End If


End Sub

Private Sub Text_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)

Dim sts As Integer
Dim i   As Integer

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Select Case index
        
        Case ptxSyuka_YY
            If Len(Trim(Text(ptxSyuka_YY).Text)) = 0 Then
            Else
            
                If Not IsNumeric(Text(ptxSyuka_YY).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
            End If
        Case ptxSyuka_MM
            If Len(Trim(Text(ptxSyuka_MM).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSyuka_MM).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
                Text(ptxSyuka_MM).Text = Format(CInt(Text(ptxSyuka_MM).Text), "00")
            End If
        Case ptxSyuka_DD
            If Len(Trim(Text(ptxSyuka_DD).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSyuka_DD).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
                Text(ptxSyuka_DD).Text = Format(CInt(Text(ptxSyuka_DD).Text), "00")
            End If
        
        
        Case ptxDATA_KBN
            If Trim(Text(index).Text) = "" Or Text(index).Text = "1" Or Text(index).Text = "3" Or Text(index).Text = "*" Then
            Else
                Beep
                MsgBox "入力した項目はエラーです。"
                Exit Sub
            End If
        
        Case ptxHAN_KBN
            If Trim(Text(index).Text) = "" Or Text(index).Text = "1" Or Text(index).Text = "2" Then
            Else
                Beep
                MsgBox "入力した項目はエラーです。"
                Exit Sub
            End If
        
        Case ptxMUKE_CODE
            Call UniCode_Conv(K2_MTS.MUKE_CODE, Text(index).Text)
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                        Beep
                        MsgBox "入力した項目はエラーです。(向け先コード)"
                        Exit Sub
                    End If
                                
                Case BtErrKeyNotFound
                                
                    Call UniCode_Conv(K3_MTS.SS_CODE, Text(index).Text)
                                                        
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                    Select Case sts
                        Case BtNoErr
                                        
                        Case BtErrKeyNotFound
                            Beep
                            MsgBox "入力した項目はエラーです。(向け先コード)"
                            Exit Sub
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                            Unload Me
                    End Select

                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                    Unload Me
            End Select


            For i = 0 To Combo(pcmbMUKE_CODE).ListCount - 1 '向け先
    
                If Right(Combo(pcmbMUKE_CODE).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                    Combo(pcmbMUKE_CODE).ListIndex = i
                    Exit For
                End If
            
    
            Next

            Combo(pcmbMUKE_CODE).SetFocus
    End Select
    
    For i = index + 1 To Text_Max
        If Text(i).Visible And Text(i).Enabled And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Function KENPIN_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   検品済更新
'----------------------------------------------------------------------------
Dim sts As Integer
Dim ans As Integer
    
    If TDBGrid1.Bookmark = -1 Then
        Exit Function
    End If
    
    
    KENPIN_Update_Proc = True
                                     'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
                                    '出荷予定の読み込み
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, SYUKA(TDBGrid1.Bookmark, ColJGYOBU))     '事業部
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, SYUKA(TDBGrid1.Bookmark, ColID_NO))   ' ID№
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                MsgBox "他端末で内容が変更されています。最新表示を行ってください。"
                KENPIN_Update_Proc = False
                GoTo Abort_Tran
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    KENPIN_Update_Proc = False
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定")
                GoTo Abort_Tran
        End Select

    Loop
    
    
    If Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode)) <> "" Then
        MsgBox "他端末で処理中です。当画面では処理できません。。"
        KENPIN_Update_Proc = False
        GoTo Abort_Tran
    End If
    
    
                                    
    If Inspe_F = 0 Then
        If StrConv(Y_SYUREC.SURYO, vbUnicode) <> StrConv(Y_SYUREC.JITU_SURYO, vbUnicode) Then
            MsgBox "出庫作業未完了です。検品処理を実行できません。"
            KENPIN_Update_Proc = False
            GoTo Abort_Tran
        
        End If
            
    
    Else
    
''        If Not IsNumeric(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) Then
''        Else
''            If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> 0 Then
''
''                If Not IsNumeric(StrConv(Y_SYUREC.KENPIN_SURYO, vbUnicode)) Then
''                Else
''                    If CLng(StrConv(Y_SYUREC.KENPIN_SURYO, vbUnicode)) <> 0 Then
''
''                        MsgBox "出庫作業中です。検品処理を実行できません。"
''                        KENPIN_Update_Proc = False
''                        GoTo Abort_Tran
''                    End If
''                End If
''            End If
''        End If
    End If
                                    
                                    
    If Trim(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) <> "" Then
        MsgBox "GLICS引渡し済です。当画面では処理できません。"
        KENPIN_Update_Proc = False
        GoTo Abort_Tran
    End If
                                    
    If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
    Else
        If StrConv(Y_SYUREC.G_KENPIN_F, vbUnicode) = "1" Then
        Else
    
            MsgBox "ｽｷｬﾅ検品処理済です。当画面では処理できません。"
            KENPIN_Update_Proc = False
            GoTo Abort_Tran
        End If
    End If
    
    
    
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
    
    If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
        '検品済にする
        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
        '2006.07.20 検品担当者出力追加
        Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "POS")
        Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, Format(Now, "HHMMSS"))
        Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "1")
        '予定数--＞実績数（ここには未出庫設定時しか来ない）
        If Inspe_F = 1 Then
            If Not IsNumeric(StrConv(Y_SYUREC.KENPIN_SURYO, vbUnicode)) Then
                If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) = 0 Then
                    Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(Y_SYUREC.SURYO, vbUnicode))
                End If
            
            End If
            
            
            
        
        End If
    Else
        '未検品する
        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
        '2006.07.20 検品担当者出力追加
        Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")
        Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")
        Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "")
    
        If Inspe_F = 1 Then
            
            If IsNumeric(StrConv(Y_SYUREC.KENPIN_SURYO, vbUnicode)) Then
                        
                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(Y_SYUREC.KENPIN_SURYO, vbUnicode))
            End If
            Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, "")
        End If
    
    End If
                                    
                                    '出荷予定書込み
    Do
        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    KENPIN_Update_Proc = False
                    GoTo Abort_Tran
                End If
        
            Case Else
                Call File_Error(sts, BtOpUpdate, "出荷予定")
                GoTo Abort_Tran
        End Select
    Loop
                                        
                                        
End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
    
        SYUKA(TDBGrid1.Bookmark, ColKENPIN_MARK) = KENPIN_OFF
        SYUKA(TDBGrid1.Bookmark, ColKENPIN_Date) = ""
        SYUKA(TDBGrid1.Bookmark, ColKENPIN_TANTO) = ""
                                        
    Else
                                        
        SYUKA(TDBGrid1.Bookmark, ColKENPIN_MARK) = KENPIN_ON
        
        SYUKA(TDBGrid1.Bookmark, ColKENPIN_Date) = Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 1, 4) & "/" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 5, 2) & "/" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 7, 2) & " " _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 1, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 3, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 5, 2)
        
        If Trim(StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode)) = "POS" Then
            SYUKA(TDBGrid1.Bookmark, ColKENPIN_TANTO) = "出荷確認画面"
        Else
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode))
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            End Select
            
            
            SYUKA(TDBGrid1.Bookmark, ColKENPIN_TANTO) = StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode) & " " & StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                                            
        End If
                                        
    End If
    
    SYUKA(TDBGrid1.Bookmark, ColFIX_QTY) = Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "#,##0")
    
        
    
    
    Set TDBGrid1.Array = SYUKA
    TDBGrid1.Refresh
    
    TDBGrid1.Update

    
    If IsNumeric(Text(ptxDEN_MAISU_JI).Text) Then
        If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
            Text(ptxDEN_MAISU_JI).Text = Format(CInt(Text(ptxDEN_MAISU_JI).Text) - 1, "#,##0")
        Else
            Text(ptxDEN_MAISU_JI).Text = Format(CInt(Text(ptxDEN_MAISU_JI).Text) + 1, "#,##0")
        End If
    End If
    
    KENPIN_Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function
