VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PI000201 
   Caption         =   "商品化完了登録 "
   ClientHeight    =   10965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17535
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
   ScaleHeight     =   10965
   ScaleWidth      =   17535
   StartUpPosition =   2  '画面の中央
   Begin VB.ListBox lstGensankoku 
      Height          =   300
      Left            =   7680
      Sorted          =   -1  'True
      TabIndex        =   111
      Top             =   6600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      Index           =   50
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   55
      Top             =   4680
      Width           =   4812
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   360
      Index           =   2
      Left            =   1935
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   56
      Top             =   5160
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   42
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   41
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   30
      Top             =   4920
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   40
      Left            =   13245
      MaxLength       =   4
      TabIndex        =   29
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   39
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   38
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   27
      Top             =   4560
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   37
      Left            =   13245
      MaxLength       =   4
      TabIndex        =   26
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   36
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   35
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   24
      Top             =   4200
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   34
      Left            =   13245
      MaxLength       =   4
      TabIndex        =   23
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   48
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   47
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   38
      Top             =   5640
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   46
      Left            =   13245
      MaxLength       =   3
      TabIndex        =   37
      Top             =   5640
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      Left            =   9375
      TabIndex        =   36
      Top             =   5640
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   45
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   44
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   34
      Top             =   5280
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   43
      Left            =   13245
      MaxLength       =   4
      TabIndex        =   33
      Top             =   5280
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   9375
      TabIndex        =   32
      Top             =   5280
      Width           =   3855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   57
      Top             =   7080
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   5318
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "発行日時"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "指図票№"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "仕向け先"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "手配先"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "品番"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "品名"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "予定"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "完了"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "未完了"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "収単/担当者"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3149"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=512"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2011"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1905"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2778"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2672"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2778"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2672"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=512"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=3493"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3387"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=512"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=1826"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1720"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=1826"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1720"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=1826"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1720"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=3810"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=3704"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=1200"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=43,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=43,.fontname=ＭＳ ゴシック"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=58,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(39)  =   ":id=58,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(40)  =   ":id=58,.fontname=ＭＳ ゴシック"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=44"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=45"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=47"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=62,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(45)  =   ":id=62,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(46)  =   ":id=62,.fontname=ＭＳ ゴシック"
      _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=44"
      _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=28,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(51)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(52)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=44"
      _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=45"
      _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=47"
      _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=66,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(57)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(58)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(59)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=44"
      _StyleDefs(60)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=45"
      _StyleDefs(61)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=47"
      _StyleDefs(62)  =   "Splits(0).Columns(4).Style:id=32,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(63)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(64)  =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(65)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=44"
      _StyleDefs(66)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=45"
      _StyleDefs(67)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=47"
      _StyleDefs(68)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(69)  =   ":id=70,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(70)  =   ":id=70,.fontname=ＭＳ ゴシック"
      _StyleDefs(71)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
      _StyleDefs(82)  =   "Splits(0).Columns(8).Style:id=82,.parent=43,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(9).Style:id=16,.parent=43"
      _StyleDefs(87)  =   "Splits(0).Columns(9).HeadingStyle:id=13,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(9).FooterStyle:id=14,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(9).EditorStyle:id=15,.parent=47"
      _StyleDefs(90)  =   "Named:id=33:Normal"
      _StyleDefs(91)  =   ":id=33,.parent=0"
      _StyleDefs(92)  =   "Named:id=34:Heading"
      _StyleDefs(93)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(94)  =   ":id=34,.wraptext=-1"
      _StyleDefs(95)  =   "Named:id=35:Footing"
      _StyleDefs(96)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(97)  =   "Named:id=36:Selected"
      _StyleDefs(98)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(99)  =   "Named:id=37:Caption"
      _StyleDefs(100) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(101) =   "Named:id=38:HighlightRow"
      _StyleDefs(102) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(103) =   "Named:id=39:EvenRow"
      _StyleDefs(104) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(105) =   "Named:id=40:OddRow"
      _StyleDefs(106) =   ":id=40,.parent=33"
      _StyleDefs(107) =   "Named:id=41:RecordSelector"
      _StyleDefs(108) =   ":id=41,.parent=34"
      _StyleDefs(109) =   "Named:id=42:FilterBar"
      _StyleDefs(110) =   ":id=42,.parent=33"
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   49
      Left            =   14730
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   33
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   32
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   21
      Top             =   3840
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   31
      Left            =   13245
      MaxLength       =   4
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   30
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   29
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   18
      Top             =   3480
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   28
      Left            =   13245
      MaxLength       =   4
      TabIndex        =   17
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   27
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   26
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   15
      Top             =   3120
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   25
      Left            =   13245
      MaxLength       =   4
      TabIndex        =   14
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   24
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   23
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   12
      Top             =   2760
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   22
      Left            =   13245
      MaxLength       =   4
      TabIndex        =   11
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   21
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   20
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2400
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   19
      Left            =   13245
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   18
      Left            =   14730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   17
      Left            =   13995
      MaxLength       =   6
      TabIndex        =   6
      Top             =   2040
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   16
      Left            =   13245
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   48
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      IMEMode         =   3  'ｵﾌ固定
      Index           =   15
      Left            =   12720
      MaxLength       =   8
      TabIndex        =   4
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   48
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   9720
      MaxLength       =   8
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   48
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   54
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   53
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   52
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   9
      Left            =   4440
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   360
      Index           =   1
      Left            =   2640
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   49
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   48
      Top             =   2760
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      Style           =   1  '標準ｺﾝﾎﾞ
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   6
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   0
      Top             =   240
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
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
      Index           =   11
      Left            =   10440
      TabIndex        =   68
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
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
      Index           =   10
      Left            =   9600
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
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
      Index           =   9
      Left            =   8760
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
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
      Index           =   8
      Left            =   7920
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
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
      Index           =   7
      Left            =   6600
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
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
      Index           =   6
      Left            =   5760
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "最 新"
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
      Left            =   4080
      TabIndex        =   61
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   60
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
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
      Left            =   1920
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
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
      Index           =   1
      Left            =   1080
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更 新"
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
      Index           =   0
      Left            =   240
      TabIndex        =   58
      Top             =   10320
      Width           =   855
   End
   Begin VB.Label lblCOMPO_Check1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   110
      Top             =   6720
      Width           =   5235
   End
   Begin VB.Label lblCOMPO_Check1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   109
      Top             =   6480
      Width           =   5235
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "構成チェック"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   360
      TabIndex        =   108
      Top             =   6480
      Width           =   1515
   End
   Begin VB.Label lblHinban_Check1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   1560
      TabIndex        =   107
      Top             =   6120
      Width           =   5232
   End
   Begin VB.Label lblHinban_Check1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   106
      Top             =   5760
      Width           =   5235
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "品番チェック"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   360
      TabIndex        =   105
      Top             =   5760
      Width           =   1515
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "加工作業（BU請求）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   9375
      TabIndex        =   98
      Top             =   4920
      Width           =   3915
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "後片付け"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   9375
      TabIndex        =   97
      Top             =   4560
      Width           =   3915
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "部品搬入"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   9375
      TabIndex        =   96
      Top             =   4200
      Width           =   3915
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "③"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   9375
      TabIndex        =   89
      Top             =   3840
      Width           =   3915
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "②個装作業(見本確認含む)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   9375
      TabIndex        =   88
      Top             =   3480
      Width           =   3915
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "①ラベル貼り"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   9375
      TabIndex        =   87
      Top             =   3120
      Width           =   3915
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "同梱部品準備"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   9375
      TabIndex        =   86
      Top             =   2760
      Width           =   3915
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "副資材準備"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   9375
      TabIndex        =   85
      Top             =   2400
      Width           =   3915
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "部品準備"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   9375
      TabIndex        =   84
      Top             =   2040
      Width           =   3915
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "加工"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   56
      Left            =   8640
      TabIndex        =   104
      Top             =   4920
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "作業実績"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   9375
      TabIndex        =   83
      Top             =   1680
      Width           =   3915
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "工程"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   55
      Left            =   8640
      TabIndex        =   103
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "収単/担当者"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   54
      Left            =   465
      TabIndex        =   102
      Top             =   5280
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "後"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   8640
      TabIndex        =   101
      Top             =   4200
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "作業"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   8640
      TabIndex        =   100
      Top             =   3120
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "前"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   8640
      TabIndex        =   99
      Top             =   2040
      Width           =   765
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "他責"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   8640
      TabIndex        =   95
      Top             =   5640
      Width           =   765
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "自責"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   8640
      TabIndex        =   94
      Top             =   5280
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   8640
      TabIndex        =   93
      Top             =   6000
      Width           =   6105
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   14730
      TabIndex        =   92
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "分"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   13995
      TabIndex        =   91
      Top             =   1680
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "人数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   13245
      TabIndex        =   90
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "未完了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   12720
      TabIndex        =   82
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "完了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   9720
      TabIndex        =   81
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "予定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   6720
      TabIndex        =   80
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "内職ｸﾗｽ"
      Height          =   255
      Index           =   10
      Left            =   840
      TabIndex        =   79
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "付加ｸﾗｽ"
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   78
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "商品化ｸﾗｽ"
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   77
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "品番"
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   76
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "手配先"
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   75
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "仕向け先"
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   74
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "承認者"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   73
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "担当者"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   72
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "発行日"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   71
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "完了日"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   70
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "指図票№"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   69
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "PI000201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



    
'テキスト用添字
Private Const ptxSHIJI_NO% = 0              '指図票№
Private Const ptxUKEIRE_DT% = 1             '受入日付
Private Const ptxHAKKO_DT% = 2              '発行日
Private Const ptxTANTO_CODE% = 3            '担当者ｺｰﾄﾞ
Private Const ptxTANTO_NAME% = 4            '担当者名称
Private Const ptxSHONIN_CODE% = 5           '承認者ｺｰﾄﾞ
Private Const ptxSHONIN_NAME% = 6           '承認者名称
Private Const ptxUKEHARAI_CODE% = 7         '承認者ｺｰﾄﾞ
Private Const ptxHIN_GAI% = 8               '品番
Private Const ptxHIN_NAME% = 9              '品名
Private Const ptxS_CLASS_CODE% = 10         '商品化ｸﾗｽ
Private Const ptxF_CLASS_CODE% = 11         '付加ｸﾗｽ
Private Const ptxN_CLASS_CODE% = 12         '内職ｸﾗｽ

Private Const ptxSHIJI_QTY% = 13            '予定
Private Const ptxUKEIRE_QTY% = 14           '受入
Private Const ptxMI_QTY% = 15               '未完

Private Const ptxNIN01% = 16                '準備1(指図票/ﾗﾍﾞﾙ発行)  人
Private Const ptxTIMES01% = 17              '準備1(指図票/ﾗﾍﾞﾙ発行)  分
Private Const ptxTOTAL01% = 18              '準備1(指図票/ﾗﾍﾞﾙ発行)  合計

Private Const ptxNIN02% = 19                '準備2(部品搬入)　人
Private Const ptxTIMES02% = 20              '準備2(部品搬入)　分
Private Const ptxTOTAL02% = 21              '準備2(部品搬入)　合計

Private Const ptxNIN03% = 22                '準備3(資材出庫)　人
Private Const ptxTIMES03% = 23              '準備3(資材出庫)　分
Private Const ptxTOTAL03% = 24              '準備3(資材出庫)　合計

Private Const ptxNIN04% = 25                '準備4(同梱取説出庫など)　人
Private Const ptxTIMES04% = 26              '準備4(同梱取説出庫など)　分
Private Const ptxTOTAL04% = 27              '準備4(同梱取説出庫など)　合計

Private Const ptxNIN05% = 28                '作業1(ﾗﾍﾞﾙ貼り)　人
Private Const ptxTIMES05% = 29              '作業1(ﾗﾍﾞﾙ貼り)　分
Private Const ptxTOTAL05% = 30              '作業1(ﾗﾍﾞﾙ貼り)　合計

Private Const ptxNIN06% = 31                '作業2　人
Private Const ptxTIMES06% = 32              '作業2　分
Private Const ptxTOTAL06% = 33              '作業2　合計

Private Const ptxNIN07% = 34                '作業3　人
Private Const ptxTIMES07% = 35              '作業3　分
Private Const ptxTOTAL07% = 36              '作業3　合計

Private Const ptxNIN08% = 37                '後片1(部品搬入) 人
Private Const ptxTIMES08% = 38              '後片1(部品搬入) 分
Private Const ptxTOTAL08% = 39              '後片1(部品搬入) 合計

Private Const ptxNIN09% = 40                '後片2(部品搬入) 人
Private Const ptxTIMES09% = 41              '後片2(指図票完了登録) 分
Private Const ptxTOTAL09% = 42              '後片2(指図票完了登録) 合計

Private Const ptxNIN10% = 43                '自責　　　人
Private Const ptxTIMES10% = 44              '自責　　　分
Private Const ptxTOTAL10% = 45              '自責　　合計

Private Const ptxNIN11% = 46                '他責　　　人
Private Const ptxTIMES11% = 47              '他責　　　分
Private Const ptxTOTAL11% = 48              '他責　　合計


Private Const ptxTOTAL% = 49                '合計


Private Const ptxBIKOU% = 50                '備考


'コンボ用添字
Private Const pcmbSHIMUKE% = 0              '仕向け先
Private Const pcmbUKEHARAI% = 1             '手配先
Private Const pcmbS_TANTO% = 2              '収単／担当者コード
Private Const pcmbJISEKI% = 3               '自責要因
Private Const pcmbTASEKI% = 4               '他責要因



'Glid用環境

Dim SSHIJI  As New XArrayDB

Private Const Min_Row% = 1                  '最小行数
Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 9                  '最大列数


Private Const colHAKKO_DT% = 0              '発行日時
Private Const colSHIJI_NO% = 1              '指図票№
Private Const colSHIMUKE% = 2               '仕向け先
Private Const colUKEHARAI% = 3              '手配先
Private Const colHIN_GAI% = 4               '品番
Private Const colHIN_NAME% = 5              '品名
Private Const colSHIJI_QTY% = 6             '予定
Private Const colUKEIRE_QTY% = 7            '受入
Private Const colMI_QTY% = 8                '未完

Private Const colS_TANTO% = 9               '収単／担当者

Private Sort_Tbl(colHAKKO_DT To colS_TANTO) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
                                            
Private Save_UKEIRE_QTY     As Long             '受入数のセーブ


Private Time_Input_F        As Boolean      '2008.08.19



Private LOAD_LIMIT          As Long         '2010.09.21
Private HINBAN_CHECK        As Integer      '2010.09.21

Private COMPO_CHECK         As Integer      '2011.05.07


Private HINBAN_COMPO_CHECK  As Integer      '2011.05.07


Private GAISOU_CHECK        As Integer      '2015.11.07

Private CHECK_OK            As Integer      '2016.03.09

Private DEF_INP_SURYO       As Integer      '2017.08.09


'Private Const LAST_UPDATE_DAY$ = "[PI00020] 2018.04.19 13:15"
Private Const LAST_UPDATE_DAY$ = "[PI00020] 2019.04.05 10:15"


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PI000201.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000201)


    PI000201.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts     As Integer
Dim i       As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        Case ptxSHIJI_NO    '指図票№
        
            If IsNumeric(Text1(ptxSHIJI_NO).Text) Then
                Text1(ptxSHIJI_NO).Text = Format(CLng(Text1(ptxSHIJI_NO).Text), "00000000")
            End If
            '指図票ﾃﾞｰﾀのﾁｪｯｸ
            
            If Text1(ptxSHIJI_NO).Text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode) Then
                '207.11.19 ↓
                If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                    'MsgBox "完了登録済みです。"            '2017.07.22
                    MsgBox "完了登録済みの指図票です"       '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            
                If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                    'MsgBox "キャンセル処理済みです。"          '2017.07.22
                    MsgBox "キャンセル処理済みの指図票です。"   '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                End If
                '207.11.19 ↑
                
                
                
                
                
                
                
                
                sts = BtNoErr
            Else
                sts = P_SSHIJI_Read_Proc()
            End If
            Select Case sts
                Case False, BtNoErr
                            
                    If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                        'MsgBox "完了登録済みです。"            '2017.07.22
                        MsgBox "完了登録済みの指図票です"       '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                
                    If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                        'MsgBox "キャンセル処理済みです。"          '2017.07.22
                        MsgBox "キャンセル処理済みの指図票です。"   '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                    '2010.09.08
                    If Trim(StrConv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, vbUnicode)) = "" Then
                    End If
                    '2010.09.08
                
                
                Case BtErrKeyNotFound
                    'MsgBox "入力した項目はエラーです。"                '2017.07.22
                    MsgBox "入力した指図票№は登録されていません。"     '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Exit Function
            End Select
        
        Case ptxHAKKO_DT    '完了日
            
        Case ptxUKEIRE_DT   '完了日
            
            '207.11.19 ↓
            If Text1(ptxSHIJI_NO).Text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode) Then
                If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                    'MsgBox "完了登録済みです。"            '2017.07.22
                    MsgBox "完了登録済みの指図票です"       '2017.07.22
                    Text1(ptxSHIJI_NO).SetFocus
                    Exit Function
                End If
            
                If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                    'MsgBox "キャンセル処理済みです。"          '2017.07.22
                    MsgBox "キャンセル処理済みの指図票です。"   '2017.07.22
                    Text1(ptxSHIJI_NO).SetFocus
                    Exit Function
                End If
            End If
            '207.11.19 ↑
            
            If Not IsDate(Text1(ptxUKEIRE_DT).Text) Then
                'MsgBox "入力した項目はエラーです。"                        '2017.07.22
                MsgBox "完了日を正しく入力して下さい。(YYYY/MM/DD)"         '2017.07.22
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxUKEIRE_DT).Text = Format(CDate(Text1(ptxUKEIRE_DT).Text), "YYYY/MM/DD")
            End If
        
        
        Case ptxUKEHARAI_CODE   '手配先
            
                        
            Text1(Mode).Text = StrConv(Text1(Mode).Text, vbUpperCase)           '2017.07.22
                        
            Combo1(pcmbUKEHARAI).ListIndex = -1
            For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
                If Text1(ptxUKEHARAI_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
                    Combo1(pcmbUKEHARAI).ListIndex = i
                    Exit For
                End If
            
            Next i
     
            If i > Combo1(pcmbUKEHARAI).ListCount - 1 Then
                'MsgBox "入力した項目はエラーです。"
                MsgBox "入力した手配先は登録されていません。"               '2017.07.22
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        
        
    
        Case ptxS_CLASS_CODE    '商品化ｸﾗｽ
            
''            If Text1(ptxS_CLASS_CODE).Text = UNSET_CLASS Then
''                MsgBox "入力した項目はエラーです。"
''                Text1(Mode).SetFocus
''                Exit Function
''            End If
'
'
'            Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
'            Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxS_CLASS_CODE).Text)
'
'            sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
'            Select Case sts
'                Case BtNoErr
'
'                Case BtErrKeyNotFound
'
'                    'MsgBox "入力した項目はエラーです。"                    '2017.07.22
'                    MsgBox "入力した商品化クラスは登録されていません。"     '2017.07.22
'                    Text1(Mode).SetFocus
'                    Exit Function
'                Case Else
'                    Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ")
'                    Exit Function
'
'            End Select
    
        Case ptxF_CLASS_CODE    '付加ｸﾗｽ
        
'            If Trim(Text1(ptxF_CLASS_CODE).Text) = "" Then
'            Else
'
''                If Text1(ptxF_CLASS_CODE).Text = UNSET_CLASS Then
''                    MsgBox "入力した項目はエラーです。"
''                    Text1(Mode).SetFocus
''                    Exit Function
''                End If
'
'
'                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
'                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxF_CLASS_CODE).Text)
'
'                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'
'                        'MsgBox "入力した項目はエラーです。"
'                        MsgBox "入力した付加クラスは登録されていません。"       '2017.07.22
'                        Text1(Mode).SetFocus
'                        Exit Function
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ")
'                        Exit Function
'
'                End Select
'            End If
    
        Case ptxN_CLASS_CODE    '内職ｸﾗｽ
        
'            If Trim(Text1(ptxN_CLASS_CODE).Text) = "" Then
'            Else
'
''                If Text1(ptxF_CLASS_CODE).Text = UNSET_CLASS Then
''                    MsgBox "入力した項目はエラーです。"
''                    Text1(Mode).SetFocus
''                    Exit Function
''                End If
'
'
'                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
'                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxN_CLASS_CODE).Text)
'
'                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
'                Select Case sts
'                    Case BtNoErr
'
'
''                        Text1(ptxN_CLASS_NAME).Text = Trim(StrConv(P_CLASSREC.CLASS_NAME, vbUnicode))
'
'
'                    Case BtErrKeyNotFound
'
'                        'MsgBox "入力した項目はエラーです。"                '2017.07.22
'                        MsgBox "入力した内職クラスは登録されていません。"   '2017.07.22
'                        Text1(Mode).SetFocus
'                        Exit Function
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ")
'                        Exit Function
'
'                End Select
'            End If
        
        
        
        Case ptxUKEIRE_QTY       '受入数
    
            If Not IsNumeric(Text1(ptxUKEIRE_QTY).Text) Then
                'MsgBox "入力した項目はエラーです。"                '2017.07.22
                MsgBox "受入数は数値を入力して下さい。"             '2017.07.22
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxUKEIRE_QTY).Text = Format(CLng(Text1(ptxUKEIRE_QTY).Text), "#0")
            
                If CLng(Text1(ptxUKEIRE_QTY).Text) > CLng(Text1(ptxSHIJI_QTY).Text) Then
                    'MsgBox "入力した項目はエラーです。"            '2017.07.22
                    MsgBox "受入数が予定数をオーバーしています。"   '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                End If
                If Save_UKEIRE_QTY = CLng(Text1(ptxUKEIRE_QTY).Text) Then
                Else
            
                    Text1(ptxMI_QTY).Text = Format(CLng(Text1(ptxSHIJI_QTY).Text) - CLng(Text1(ptxUKEIRE_QTY).Text), "#0")
                    Save_UKEIRE_QTY = CLng(Text1(ptxUKEIRE_QTY).Text)
                End If
                    
                    
                    
            End If
    
        Case ptxMI_QTY          '未完数
    
            If Not IsNumeric(Text1(ptxMI_QTY).Text) Then
                'MsgBox "入力した項目はエラーです。"                '2017.07.22
                MsgBox "未完数は数値を入力して下さい。"             '2017.07.22
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxMI_QTY).Text = Format(CLng(Text1(ptxMI_QTY).Text), "#0")
                
                If CLng(Val(Text1(ptxSHIJI_QTY).Text) - Val(Text1(ptxUKEIRE_QTY).Text)) = Val(Text1(ptxMI_QTY).Text) Or _
                    CLng(Text1(ptxMI_QTY).Text) = 0 Then
                Else
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "未完数の値を正しく入力して下さい。"      '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            
                    
            End If
            
        Case ptxNIN01, ptxNIN02, ptxNIN03, ptxNIN04, ptxNIN05, ptxNIN06, ptxNIN07, ptxNIN08, ptxNIN09 '人数
            If Text1(Mode).Text = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    'MsgBox "入力した項目はエラーです。"
                    MsgBox "人数は数値を入力して下さい。"       '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    
                    Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.0")
                    
                    If InStr(1, Trim(Text1(Mode).Text), ".") > 2 Then
                        'MsgBox "入力した項目はエラーです。(9.9)"   '2017.07.22
                        MsgBox "人数は9.9以下で入力して下さい。"    '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                    
                    
                    If IsNumeric(Text1(Mode + 1).Text) Then
                        Text1(Mode + 2).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 1).Text), "#0.00")
                    Else
                        Text1(Mode + 2).Text = "0.00"
                    End If
                
                    Text1(ptxTOTAL).Text = "0"
                    For i = ptxTOTAL01 To ptxTOTAL11 Step 3
                        
                        If IsNumeric(Text1(i).Text) Then
                           Text1(ptxTOTAL).Text = Format(CDbl(Text1(ptxTOTAL).Text) + CDbl(Text1(i).Text), "#0.00")
                        End If
                    
                    Next i
                End If
            End If
    
        Case ptxTIMES01, ptxTIMES02, ptxTIMES03, ptxTIMES04, ptxTIMES05, ptxTIMES06, ptxTIMES07, ptxTIMES08, ptxTIMES09 '時間
            If Text1(Mode).Text = "" Then
                If Text1(Mode - 1).Text = "" Then
                    Text1(Mode + 1).Text = ""
                End If
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
'                    MsgBox "入力した項目はエラーです。"            '2017.07.22
                    MsgBox "時間は数値で入力して下さい。"           '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
                    If IsNumeric(Text1(Mode - 1).Text) Then
                        Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 1).Text) * CDbl(Text1(Mode).Text), "#0.00")
                    Else
                        Text1(Mode + 1).Text = "0.00"
                    End If
                
                    Text1(ptxTOTAL).Text = "0.00"
                    For i = ptxTOTAL01 To ptxTOTAL11 Step 3
                        
                        If IsNumeric(Text1(i).Text) Then
                           Text1(ptxTOTAL).Text = Format(CDbl(Text1(ptxTOTAL).Text) + CDbl(Text1(i).Text), "#0.00")
                        End If
                    
                    Next i
                End If
            End If
    
    
        Case ptxNIN10, ptxNIN11             '自責　人数
            
            If Text1(Mode).Text = "" Then
            
            
            
            Else
                
                If Mode = ptxNIN10 Then
                    
                    If Combo1(pcmbJISEKI).Text = "" Then
                                
                        'MsgBox "入力した項目はエラーです。"            '2017.07.22
                        MsgBox "自責内容を選択後に入力して下さい。"     '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                
                    End If
                
                Else
                
                    If Combo1(pcmbTASEKI).Text = "" Then
                                
                        'MsgBox "入力した項目はエラーです。"            '2017.07.22
                        MsgBox "他責内容を選択後に入力して下さい。"     '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                
                    End If
                
                End If
                
                
                If Not IsNumeric(Text1(Mode).Text) Then
                    'MsgBox "入力した項目はエラーです。"                '2017.07.22
                    MsgBox "人数は数値を入力して下さい。"               '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    
                    Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.0")
                    
                    If InStr(1, Trim(Text1(Mode).Text), ".") > 2 Then
                        'MsgBox "入力した項目はエラーです。(9.9)"   '2017.07.22
                        MsgBox "人数は9.9以下で入力して下さい。"    '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                    
                    
                    If IsNumeric(Text1(Mode + 1).Text) Then
                        Text1(Mode + 2).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 1).Text), "#0.00")
                    Else
                        Text1(Mode + 2).Text = "0.00"
                    End If
                
                    Text1(ptxTOTAL).Text = "0.00"
                    For i = ptxTOTAL01 To ptxTOTAL11 Step 3
                        
                        If IsNumeric(Text1(i).Text) Then
                           Text1(ptxTOTAL).Text = Format(CDbl(Text1(ptxTOTAL).Text) + CDbl(Text1(i).Text), "#0.00")
                        End If
                    
                    Next i
                End If
            End If
    
        Case ptxTIMES10, ptxTIMES11  '時間
            
            
            If Text1(Mode).Text = "" Then
                If Text1(Mode - 1).Text = "" Then
                    Text1(Mode + 1).Text = ""
                End If
            Else
                
                If Mode = ptxTIMES10 Then
                
                    If Combo1(pcmbJISEKI).Text = "" Then
                                
                        'MsgBox "入力した項目はエラーです。"            '2017.07.22
                        MsgBox "自責内容を選択後に入力して下さい。"     '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                
                    End If
            
                Else
            
                    If Combo1(pcmbTASEKI).Text = "" Then
                            
                        'MsgBox "入力した項目はエラーです。"            '2017.07.22
                        MsgBox "他責内容を選択後に入力して下さい。"     '2017.07.22
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
            
                End If
            
            
            
                If Not IsNumeric(Text1(Mode).Text) Then
'                    MsgBox "入力した項目はエラーです。"            '2017.07.22
                    MsgBox "時間は数値で入力して下さい。"           '2017.07.22
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
                    If IsNumeric(Text1(Mode - 1).Text) Then
                        Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 1).Text) * CDbl(Text1(Mode).Text), "#0.00")
                    Else
                        Text1(Mode + 1).Text = "0.00"
                    End If
                
                    Text1(ptxTOTAL).Text = "0"
                    For i = ptxTOTAL01 To ptxTOTAL11 Step 3
                        
                        If IsNumeric(Text1(i).Text) Then
                           Text1(ptxTOTAL).Text = Format(CDbl(Text1(ptxTOTAL).Text) + CDbl(Text1(i).Text), "#0.00")
                        End If
                    
                    Next i
                End If
            End If
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer


Dim wkDateTime      As String
Dim wkTanto         As String

Dim wkCNT           As String
    
Dim com             As Integer      '2012.04.20

Dim wkTanto_Code    As String * 5   '2012.04.20
Dim wkYMDHS         As String * 12  '2012.04.20
Dim wkALL_Cnt       As String * 2   '2012.04.20
Dim wkSumi_Cnt      As String * 2   '2012.04.20
    
    
    Item_Disp_Proc = True
    
        
    
    Text1(ptxSHIJI_NO).Text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)           '指図票№
    Text1(ptxUKEIRE_DT).Text = Format(Now, "YYYY/MM/DD")                            '受入日(今日を初期表示)
    Text1(ptxHAKKO_DT).Text = Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2)  '発行日
    
    Text1(ptxTANTO_CODE).Text = StrConv(P_SSHIJI_O_REC.TANTO_CODE, vbUnicode)       '担当者ｺｰﾄﾞ／名称
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).Text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxTANTO_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxTANTO_NAME).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    
    End Select
    
    Text1(ptxSHONIN_CODE).Text = StrConv(P_SSHIJI_O_REC.SHONIN_CODE, vbUnicode)     '承認者ｺｰﾄﾞ／名称
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).Text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxSHONIN_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxSHONIN_NAME).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    
    End Select
    
    For i = 0 To Combo1(pcmbSHIMUKE).ListCount - 1                                  '仕向け先ｺｰﾄﾞ
    
        If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE).List(i), 4), 1, 2) Then
            Combo1(pcmbSHIMUKE).ListIndex = i
            Exit For
        End If
    
    Next i
    Text1(ptxUKEHARAI_CODE).Text = Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))   '手配先
    For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
    
        If Text1(ptxUKEHARAI_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
            Combo1(pcmbUKEHARAI).ListIndex = i
            Exit For
        End If
    
    Next i
    
    
    Text1(ptxHIN_GAI).Text = Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))       '品番／品名／標準棚番／未商品／商品化済
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxHIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxHIN_NAME).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select
                                                                                    
    Text1(ptxS_CLASS_CODE).Text = Trim(StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode)) '商品化ｸﾗｽ
'    If Text1(ptxS_CLASS_CODE).Text = UNSET_CLASS Then
'
'        Text1(ptxS_CLASS_CODE).BackColor = G_INPUT_OK
'        Text1(ptxS_CLASS_CODE).Locked = False
'        Text1(ptxS_CLASS_CODE).TabStop = True
'    Else
'        Text1(ptxS_CLASS_CODE).BackColor = G_INPUT_NG
'        Text1(ptxS_CLASS_CODE).Locked = True
'        Text1(ptxS_CLASS_CODE).TabStop = False
'
'    End If
    
    Text1(ptxF_CLASS_CODE).Text = Trim(StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode)) '付加ｸﾗｽ
'    If Text1(ptxF_CLASS_CODE).Text = UNSET_CLASS Then
'
'        Text1(ptxF_CLASS_CODE).BackColor = G_INPUT_OK
'        Text1(ptxF_CLASS_CODE).Locked = False
'        Text1(ptxF_CLASS_CODE).TabStop = True
'    Else
'        Text1(ptxF_CLASS_CODE).BackColor = G_INPUT_NG
'        Text1(ptxF_CLASS_CODE).Locked = True
'        Text1(ptxF_CLASS_CODE).TabStop = False
'
'    End If
    
    Text1(ptxN_CLASS_CODE).Text = Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode)) '内職ｸﾗｽ
'    If Text1(ptxN_CLASS_CODE).Text = UNSET_CLASS Then
'
'        Text1(ptxN_CLASS_CODE).BackColor = G_INPUT_OK
'        Text1(ptxN_CLASS_CODE).Locked = False
'        Text1(ptxN_CLASS_CODE).TabStop = True
'    Else
'        Text1(ptxN_CLASS_CODE).BackColor = G_INPUT_NG
'        Text1(ptxN_CLASS_CODE).Locked = True
'        Text1(ptxN_CLASS_CODE).TabStop = False
'
'    End If
                                                                                        
                                                                                        
                                                                                        
'2010.09.21
'
'    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
'    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxN_CLASS_CODE).Text)
'
'    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
'    Select Case sts
'        Case BtNoErr
'
'
'            Text1(ptxN_CLASS_NAME).Text = Trim(StrConv(P_CLASSREC.CLASS_NAME, vbUnicode))
'
'
'        Case BtErrKeyNotFound
'            Text1(ptxN_CLASS_NAME).Text = ""
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ")
'            Exit Function
'
'    End Select
'2010.09.21
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        '指示数量
    Text1(ptxSHIJI_QTY).Text = Format(CDbl(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)) - CDbl(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "#0")
    
    
    If DEF_INP_SURYO = 1 Then                                           '2017.08.09
        Text1(ptxUKEIRE_QTY).Text = Text1(ptxSHIJI_QTY).Text            '2017.08.09
        Text1(ptxMI_QTY).Text = "0"                                     '未完数 2017.09.27
    Else                                                                '2017.08.09
        Text1(ptxUKEIRE_QTY).Text = ""                                                  '受入数（初期値）
        Text1(ptxMI_QTY).Text = Text1(ptxSHIJI_QTY).Text                                    '未完数 2017.09.27
    
    End If                                                              '2017.08.09
    'Text1(ptxMI_QTY).Text = Text1(ptxSHIJI_QTY).Text                                    '未完数 2017.09.27
    
    
    If Combo1(pcmbS_TANTO).ListCount = 0 Then                                           '収単／担当者
    Else
        For i = 0 To Combo1(pcmbS_TANTO).ListCount - 1
            If StrConv(P_SSHIJI_O_REC.S_TANTO, vbUnicode) = Right(Combo1(pcmbS_TANTO).List(i), 2) Then
                Combo1(pcmbS_TANTO).ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    
    
'2010.9.21
    
    Text1(ptxBIKOU).Text = ""
    
    
    
    
'>>>>>>>>>>>>>>>>   完了表示    2016.03.09

    Select Case CHECK_OK
        Case 0
    
            If IsNumeric(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode)) Then
                wkDateTime = Mid(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode), 7, 2) & " " & _
                                Mid(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode), 9, 2) & ":" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode), 11, 2)
        
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, vbUnicode))
            
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                        Exit Function
                
                End Select
                
                
                wkTanto = StrConv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, vbUnicode) & " " & Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
                lblHinban_Check1(0).Caption = wkDateTime & " " & wkTanto
                lblHinban_Check1(0).ForeColor = vbBlack
            
                wkCNT = "ラベル(" & Format(Val(StrConv(P_SSHIJI_O_REC.HIN_CHECK_LABEL_CNT, vbUnicode)), "#0") & ") "
                
                If GAISOU_CHECK Then                                                                                                    '2016.01.06
                    wkCNT = wkCNT & "外　装(" & Format(Val(StrConv(P_SSHIJI_O_REC.HIN_CHECK_GAISOU_CNT, vbUnicode)), "#0") & ") "       '2016.01.06
                End If                                                                                                                  '2016.01.06
                
                wkCNT = wkCNT & "現品票(" & Format(Val(StrConv(P_SSHIJI_O_REC.HIN_CHECK_GENPIN_CNT, vbUnicode)), "#0") & ") "
                
                        
                
                lblHinban_Check1(1).Caption = wkCNT
            
                Command1(P_CMD_Upd).Enabled = True
            
            Else
                lblHinban_Check1(0).Caption = "未チェック"
                lblHinban_Check1(0).ForeColor = vbRed
                
                lblHinban_Check1(1).Caption = ""
        
        
                Select Case HINBAN_CHECK
                    Case 0
                        Command1(P_CMD_Upd).Enabled = True
                    Case 1
                        Command1(P_CMD_Upd).Enabled = True
                    Case 2
                        Command1(P_CMD_Upd).Enabled = False
                End Select
        
        
            End If



        Case 1
        



            If IsNumeric(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode)) Then
                wkDateTime = Mid(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode), 7, 2) & " " & _
                                Mid(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode), 9, 2) & ":" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, vbUnicode), 11, 2)
        
        
                If Val(StrConv(P_SSHIJI_O_REC.HIN_CHECK_LABEL_CNT, vbUnicode)) > 0 And Val(StrConv(P_SSHIJI_O_REC.HIN_CHECK_GENPIN_CNT, vbUnicode)) > 0 Then
        
        
                    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, vbUnicode))
                
                    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                            Exit Function
                    
                    End Select
                    
                    
                    wkTanto = StrConv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, vbUnicode) & " " & Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
                    lblHinban_Check1(0).Caption = wkDateTime & " " & wkTanto
                    lblHinban_Check1(0).ForeColor = vbBlack
                
                    wkCNT = "ラベル(" & Format(Val(StrConv(P_SSHIJI_O_REC.HIN_CHECK_LABEL_CNT, vbUnicode)), "#0") & ") "
                    
                    If GAISOU_CHECK Then                                                                                                    '2016.01.06
                        wkCNT = wkCNT & "外　装(" & Format(Val(StrConv(P_SSHIJI_O_REC.HIN_CHECK_GAISOU_CNT, vbUnicode)), "#0") & ") "       '2016.01.06
                    End If                                                                                                                  '2016.01.06
                    
                    wkCNT = wkCNT & "現品票(" & Format(Val(StrConv(P_SSHIJI_O_REC.HIN_CHECK_GENPIN_CNT, vbUnicode)), "#0") & ") "
                    
                            
                    
                    lblHinban_Check1(1).Caption = wkCNT
                
                    Command1(P_CMD_Upd).Enabled = True
            
            
                Else
            
                    lblHinban_Check1(0).Caption = "未チェック"
                    lblHinban_Check1(0).ForeColor = vbRed
                    
                    lblHinban_Check1(1).Caption = ""
            
            
                    Select Case HINBAN_CHECK
                        Case 0
                            Command1(P_CMD_Upd).Enabled = True
                        Case 1
                            Command1(P_CMD_Upd).Enabled = True
                        Case 2
                            Command1(P_CMD_Upd).Enabled = False
                    End Select
                End If
            
            Else
                lblHinban_Check1(0).Caption = "未チェック"
                lblHinban_Check1(0).ForeColor = vbRed
                
                lblHinban_Check1(1).Caption = ""
        
        
                Select Case HINBAN_CHECK
                    Case 0
                        Command1(P_CMD_Upd).Enabled = True
                    Case 1
                        Command1(P_CMD_Upd).Enabled = True
                    Case 2
                        Command1(P_CMD_Upd).Enabled = False
                End Select
        
        
            End If




'2010.9.21


    End Select
'>>>>>>>>>>>>>>>>   完了表示    2016.03.09


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.04.20
    lblCOMPO_Check1(0).Caption = ""
    lblCOMPO_Check1(0).ForeColor = vbBlack

    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode))
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")

    
    
    
    
    
    
    
    
    
    
    
    
    com = BtOpGetGreaterEqual
    
        
    wkTanto_Code = ""
    wkYMDHS = ""
    wkALL_Cnt = ""
    wkSumi_Cnt = ""
    
    
    
    
    Do
        sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode) <> StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) Then
                    Exit Do
                End If
            
            
            
                wkTanto_Code = StrConv(P_SSHIJI_K_REC.COMPO_TANTO, vbUnicode)
                wkYMDHS = StrConv(P_SSHIJI_K_REC.COMPO_YMDHS, vbUnicode)
                wkALL_Cnt = StrConv(P_SSHIJI_K_REC.COMPO_ALL_Cnt, vbUnicode)
                wkSumi_Cnt = StrConv(P_SSHIJI_K_REC.COMPO_Sumi_Cnt, vbUnicode)
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "商品化指図データ（子）")
                Exit Function
        
        End Select
    
        com = BtOpGetNext
           
    Loop



    If Trim(wkTanto_Code) <> "" Then
        wkDateTime = Mid(wkYMDHS, 1, 4) & "/" & _
                        Mid(wkYMDHS, 5, 2) & "/" & _
                        Mid(wkYMDHS, 7, 2) & " " & _
                        Mid(wkYMDHS, 9, 2) & ":" & _
                        Mid(wkYMDHS, 11, 2)



        Call UniCode_Conv(K0_TANTO.TANTO_CODE, wkTanto_Code)
        sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                Exit Function
        
        End Select
        
        wkCNT = Format(Val(wkSumi_Cnt), "#0") & "/" & Format(Val(wkALL_Cnt), "#0")
        lblCOMPO_Check1(0).Caption = wkDateTime & " " & wkTanto_Code & " " & RTrim(StrConv(TANTOREC.TANTO_NAME, vbUnicode)) & " " & wkCNT
        lblCOMPO_Check1(0).ForeColor = vbBlack
        lblCOMPO_Check1(1).Caption = ""
    
        If wkSumi_Cnt <> wkALL_Cnt Then
            lblCOMPO_Check1(0).ForeColor = vbRed
        End If
    
    
    Else
        lblCOMPO_Check1(0).Caption = "未チェック"
        lblCOMPO_Check1(0).ForeColor = vbRed
        
        lblCOMPO_Check1(1).Caption = ""


        Select Case HINBAN_CHECK
            Case 0
                Command1(P_CMD_Upd).Enabled = True
            Case 1
                Command1(P_CMD_Upd).Enabled = True
            Case 2
                Command1(P_CMD_Upd).Enabled = False
        End Select




    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.04.20
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.05.07
        
    Select Case HINBAN_COMPO_CHECK
        Case 0
        Case 1
            
            
            If lblHinban_Check1(0).ForeColor = vbRed And lblHinban_Check1(0).ForeColor = vbRed Then
            
                If HINBAN_CHECK = 2 And COMPO_CHECK = 2 Then
                    Command1(P_CMD_Upd).Enabled = False
                Else
                    Command1(P_CMD_Upd).Enabled = True
                End If
            End If
        
        Case 2
    
            If lblHinban_Check1(0).ForeColor = vbRed Or lblHinban_Check1(0).ForeColor = vbRed Then
            
                If HINBAN_CHECK = 2 And COMPO_CHECK = 2 Then
                    Command1(P_CMD_Upd).Enabled = False
                Else
                    Command1(P_CMD_Upd).Enabled = True
                End If
            End If
    
    
    
    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.05.07

    
    
    Time_Input_F = False    '2008.08.19
        
    
    Item_Disp_Proc = False

End Function

Private Function Cancel_Proc() As Integer
'----------------------------------------------------------------------------
'                  商品化指図票ﾃﾞｰﾀｷｬﾝｾﾙ更新
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim SEQNO       As Integer



Dim i           As Integer


    Cancel_Proc = True
                                        
    Call Input_Lock
    DoEvents
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    
    
    
    '---------------------------------------------------    '指図票データ更新
    
    '指図票データ(ﾍｯﾀﾞｰ)処理
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Text1(ptxSHIJI_NO).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            
        Select Case sts
            Case BtNoErr
            
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "商品化指図票ﾃﾞｰﾀ(親)")
                GoTo Abort_Tran
        End Select

    Loop
    
    
    Call UniCode_Conv(P_SSHIJI_O_REC.CANCEL_F, P_CANCEL_ON)                     'ｷｬﾝｾﾙﾌﾗｸﾞ
                                                                                'ｷｬﾝｾﾙ日時
    Call UniCode_Conv(P_SSHIJI_O_REC.CANCEL_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                
                                                                                '更新日時
    Call UniCode_Conv(P_SSHIJI_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "商品化指図ﾃﾞｰﾀ(親)")
                    End If
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "商品化指図ﾃﾞｰﾀ(親)")
                GoTo Abort_Tran
        End Select
    
    Loop
    
    '指図票データ(子)処理
    
    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Text1(ptxSHIJI_NO).Text)
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
    
    com = BtOpGetGreater
    
    Do
        DoEvents        '2011.04.17
        Do
            sts = BTRV(com + BtSNoWait, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                
            Select Case sts
                Case BtNoErr
                
                    If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Text1(ptxSHIJI_NO).Text Then
                        
                        sts = BTRV(BtOpUnlock, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "商品化指図ﾃﾞｰﾀ(子)")
                        End If
                        
                        sts = BtErrEOF
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, com, "商品化指図票ﾃﾞｰﾀ(子)")
                    GoTo Abort_Tran
            End Select
        
        
        Loop
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        
        Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, P_CANCEL_ON)                     'ｷｬﾝｾﾙﾌﾗｸﾞ
                                                                                    'ｷｬﾝｾﾙ日時
        Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                    
                                                                                    '更新日時
        Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        If com = BtOpUpdate Then
                            sts = BTRV(BtOpUnlock, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "商品化指図ﾃﾞｰﾀ(子)")
                            End If
                        End If
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "商品化指図ﾃﾞｰﾀ(子)")
                    GoTo Abort_Tran
            End Select
        
        Loop
        
        
        
        com = BtOpGetNext
    
    Loop
    

End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    Call Input_UnLock
    
    Cancel_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  商品化指図票ﾃﾞｰﾀ/受入履歴更新
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim SEQNO       As Integer



Dim i           As Integer
Dim j           As Integer


    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    
    
    
    '---------------------------------------------------    '指図票データ更新
    
    '指図票データ(ﾍｯﾀﾞｰ)処理
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Text1(ptxSHIJI_NO).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            
        Select Case sts
            Case BtNoErr
            
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "商品化指図票ﾃﾞｰﾀ(親)")
                GoTo Abort_Tran
        End Select

    Loop
                                                                            '取引先ｺｰﾄﾞ
    Call UniCode_Conv(P_SSHIJI_O_REC.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            MsgBox "手配先情報が他で変更されました。更新処理を中止します。"
            GoTo Abort_Tran
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            Exit Function
    
    End Select
                                                                                    
    Call UniCode_Conv(P_SSHIJI_O_REC.S_CLASS_CODE, Text1(ptxS_CLASS_CODE))  '商品化ｸﾗｽ
    Call UniCode_Conv(P_SSHIJI_O_REC.F_CLASS_CODE, Text1(ptxF_CLASS_CODE))  '付加ｸﾗｽ
    Call UniCode_Conv(P_SSHIJI_O_REC.N_CLASS_CODE, Text1(ptxN_CLASS_CODE))  '内職ｸﾗｽ
    '取引先区分
    Call UniCode_Conv(P_SSHIJI_O_REC.TORI_KBN, StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))
    
    If CDbl(Text1(ptxMI_QTY).Text) = 0 Then
        Call UniCode_Conv(P_SSHIJI_O_REC.KAN_F, P_KAN_ON)                   '完了ﾌﾗｸﾞ
        Call UniCode_Conv(P_SSHIJI_O_REC.KAN_DT, Format(Now, "YYYYMMDD"))   '完了日
        If CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) = 0 Then     '分納回数
        Else
            Call UniCode_Conv(P_SSHIJI_O_REC.BUNNOU_CNT, Format(CInt(CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) + 1), "00"))
        End If
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.BUNNOU_CNT, Format(CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) + 1, "00"))
    End If
    
                                                                            '収単／担当者ｸﾗｽ
    Call UniCode_Conv(P_SSHIJI_O_REC.S_TANTO, Right(Combo1(pcmbS_TANTO).Text, 2))
    
    
''    j = 0
''
''    For i = ptxNIN01 To ptxNIN06 Step 3
''
''        If IsNumeric(Text1(i).Text) Then        '人
''            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(j).NIN, Format(CDbl(Text1(i).Text), "0.0"))
''        Else
''            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(j).NIN, "0.0")
''        End If
''        If IsNumeric(Text1(i + 1).Text) Then      '時間
''            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(j).TIMES, Format(CDbl(Text1(i + 1).Text), "000.00"))
''        Else
''            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(j).TIMES, "000.00")
''        End If
''
''        If CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(j).NIN, vbUnicode)) = 0 And _
''             CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(j).TIMES, vbUnicode)) = 0 Then
''        Else
''            If CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(j).NIN, vbUnicode)) = 0 Then
''                Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(j).TIMES, "1.0")
''            End If
''
''            If CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(j).TIMES, vbUnicode)) = 0 Then
''                Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(j).TIMES, "001.00")
''            End If
''        End If
''
''
''        j = j + 1
''    Next i
                                                
                                                
                                                
                                                
'    '準備１ 指図票/ﾗﾍﾞﾙ発行
'    If IsNumeric(Text1(ptxNIN01).Text) Then        '人
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(3).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)) + _
'                                                            CDbl(Text1(ptxNIN01).Text), "0.0"))
'    End If
'
'    If IsNumeric(Text1(ptxTIMES01).Text) Then      '時間
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(3).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)) + _
'                                                            CDbl(Text1(ptxTIMES01).Text), "000.00"))
'    End If
'
'    If Trim(Text1(ptxNIN01).Text) = "" And _
'         Trim(Text1(ptxTIMES01).Text) = "" Then
'    Else
'        If Trim(Text1(ptxNIN01).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(3).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)) + 1, "0.0"))
'        End If
'
'        If Trim(Text1(ptxTIMES01).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(3).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)) + 1, "000.00"))
'        End If
'    End If
'
'
'    '準備２ 部品搬入
'    If IsNumeric(Text1(ptxNIN02).Text) Then        '人
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) + CDbl(Text1(ptxNIN02).Text), "0.0"))
'    End If
'
'    If IsNumeric(Text1(ptxTIMES02).Text) Then      '時間
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES02).Text), "000.00"))
'    End If
'
'    If Trim(Text1(ptxNIN02).Text) = "" And _
'         Trim(Text1(ptxTIMES02).Text) = "" Then
'    Else
'        If Trim(Text1(ptxNIN02).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) + 1, "0.0"))
'        End If
'
'        If Trim(Text1(ptxTIMES02).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) + 1, "000.00"))
'        End If
'    End If
'
'    '準備３ 資材出庫
'    If IsNumeric(Text1(ptxNIN03).Text) Then        '人
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) + CDbl(Text1(ptxNIN03).Text), "0.0"))
'    End If
'
'    If IsNumeric(Text1(ptxTIMES03).Text) Then      '時間
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES03).Text), "000.00"))
'    End If
'
'    If Trim(Text1(ptxNIN03).Text) = "" And _
'         Trim(Text1(ptxTIMES03).Text) = "" Then
'    Else
'        If Trim(Text1(ptxNIN03).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) + 1, "0.0"))
'        End If
'
'        If Trim(Text1(ptxTIMES03).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) + 1, "000.00"))
'        End If
'    End If
'
'    '準備４ 同梱取説出庫など
'    If IsNumeric(Text1(ptxNIN04).Text) Then        '人
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) + CDbl(Text1(ptxNIN04).Text), "0.0"))
'    End If
'
'    If IsNumeric(Text1(ptxTIMES04).Text) Then      '時間
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(6).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES04).Text), "000.00"))
'    End If
'
'    If Trim(Text1(ptxNIN04).Text) = "" And _
'         Trim(Text1(ptxTIMES04).Text) = "" Then
'    Else
'        If Trim(Text1(ptxNIN04).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) + 1, "0.0"))
'        End If
'
'        If Trim(Text1(ptxTIMES04).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(6).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) + 1, "000.00"))
'        End If
'    End If
'
'    '作業１ ﾗﾍﾞﾙ貼り
'    If IsNumeric(Text1(ptxNIN05).Text) Then        '人
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) + CDbl(Text1(ptxNIN05).Text), "0.0"))
'    End If
'
'    If IsNumeric(Text1(ptxTIMES05).Text) Then      '時間
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES05).Text), "000.00"))
'    End If
'
'    If Trim(Text1(ptxNIN05).Text) = "" And _
'         Trim(Text1(ptxTIMES05).Text) = "" Then
'    Else
'        If Trim(Text1(ptxNIN05).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) + 1, "0.0"))
'        End If
'
'        If Trim(Text1(ptxTIMES05).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) + 1, "000.00"))
'        End If
'    End If
'
'    '作業２
'    If IsNumeric(Text1(ptxNIN06).Text) Then        '人
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) + CDbl(Text1(ptxNIN06).Text), "0.0"))
'    End If
'
'    If IsNumeric(Text1(ptxTIMES06).Text) Then      '時間
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES06).Text), "000.00"))
'    End If
'
'    If Trim(Text1(ptxNIN06).Text) = "" And _
'         Trim(Text1(ptxTIMES06).Text) = "" Then
'    Else
'        If Trim(Text1(ptxNIN06).Text) = "" Then
'
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) + 1, "0.0"))
'        End If
'
'        If Trim(Text1(ptxTIMES06).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) + 1, "000.00"))
'        End If
'    End If
'
'    '作業３
'    If IsNumeric(Text1(ptxNIN07).Text) Then        '人
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) + CDbl(Text1(ptxNIN07).Text), "0.0"))
'    End If
'
'    If IsNumeric(Text1(ptxTIMES07).Text) Then      '時間
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES07).Text), "000.00"))
'    End If
'
'    If Trim(Text1(ptxNIN07).Text) = "" And _
'         Trim(Text1(ptxTIMES07).Text) = "" Then
'    Else
'        If Trim(Text1(ptxNIN07).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) + 1, "0.0"))
'        End If
'
'        If Trim(Text1(ptxTIMES07).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) + 1, "000.00"))
'        End If
'    End If
'
'    '後片１　部品搬出
'    If IsNumeric(Text1(ptxNIN08).Text) Then        '人
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) + CDbl(Text1(ptxNIN08).Text), "0.0"))
'    End If
'
'    If IsNumeric(Text1(ptxTIMES08).Text) Then      '時間
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES08).Text), "000.00"))
'    End If
'
'    If Trim(Text1(ptxNIN08).Text) = "" And _
'         Trim(Text1(ptxTIMES08).Text) = "" Then
'    Else
'        If Trim(Text1(ptxNIN08).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) + 1, "0.0"))
'        End If
'
'        If Trim(Text1(ptxTIMES08).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) + 1, "000.00"))
'        End If
'    End If
'
'    '後片２　指図票完了登録
'    If IsNumeric(Text1(ptxNIN09).Text) Then        '人
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + CDbl(Text1(ptxNIN09).Text), "0.0"))
'    End If
'
'    If IsNumeric(Text1(ptxTIMES09).Text) Then      '時間
'        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES09).Text), "000.00"))
'    End If
'
'    If Trim(Text1(ptxNIN09).Text) = "" And _
'         Trim(Text1(ptxTIMES09).Text) = "" Then
'    Else
'        If Trim(Text1(ptxNIN09).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + 1, "0.0"))
'        End If
'
'        If Trim(Text1(ptxTIMES09).Text) = "" Then
'            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + 1, "000.00"))
'        End If
'    End If
                                                
                                                
                                                
                                                
                                                
                                                
    '準備１ 指図票/ﾗﾍﾞﾙ発行
    If IsNumeric(Text1(ptxNIN01).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) + _
                                                            CDbl(Text1(ptxNIN01).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES01).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)) + _
                                                            CDbl(Text1(ptxTIMES01).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN01).Text) = "" And _
         Trim(Text1(ptxTIMES01).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN01).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES01).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(4).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
                                                
    '準備２ 部品搬入
    If IsNumeric(Text1(ptxNIN02).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) + CDbl(Text1(ptxNIN02).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES02).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES02).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN02).Text) = "" And _
         Trim(Text1(ptxTIMES02).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN02).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES02).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
    '準備３ 資材出庫
    If IsNumeric(Text1(ptxNIN03).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) + CDbl(Text1(ptxNIN03).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES03).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(6).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES03).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN03).Text) = "" And _
         Trim(Text1(ptxTIMES03).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN03).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES03).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(5).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
    
                                                
    '作業１ ﾗﾍﾞﾙ貼り
    If IsNumeric(Text1(ptxNIN04).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) + CDbl(Text1(ptxNIN04).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES04).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES04).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN04).Text) = "" And _
         Trim(Text1(ptxTIMES04).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN04).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES04).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(0).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
    
    '作業２
    If IsNumeric(Text1(ptxNIN05).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) + CDbl(Text1(ptxNIN05).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES05).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES05).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN05).Text) = "" And _
         Trim(Text1(ptxTIMES05).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN05).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES05).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(1).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
    
    '作業３
    If IsNumeric(Text1(ptxNIN06).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) + CDbl(Text1(ptxNIN06).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES06).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES06).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN06).Text) = "" And _
         Trim(Text1(ptxTIMES06).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN06).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES06).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(2).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
    '後片１　部品搬出
    If IsNumeric(Text1(ptxNIN07).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) + CDbl(Text1(ptxNIN07).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES07).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES07).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN07).Text) = "" And _
         Trim(Text1(ptxTIMES07).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN07).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES07).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(7).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
    '後片２　指図票完了登録
    If IsNumeric(Text1(ptxNIN08).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + CDbl(Text1(ptxNIN08).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES08).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES08).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN08).Text) = "" And _
         Trim(Text1(ptxTIMES08).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN08).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES08).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(8).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
                                                
    '後片２　指図票完了登録
    If IsNumeric(Text1(ptxNIN09).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(9).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(9).NIN, vbUnicode)) + CDbl(Text1(ptxNIN09).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES09).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(9).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES09).Text), "000.00"))
    End If
    
    If Trim(Text1(ptxNIN09).Text) = "" And _
         Trim(Text1(ptxTIMES09).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN09).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(9).NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(9).NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES09).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(9).TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.GENKA_TBL(9).NIN, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                
                                                
                                                
                                                
                                                
                                                
                                                '自責　名称
    Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NAME, Combo1(pcmbJISEKI).Text)
    If IsNumeric(Text1(ptxNIN10).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) + CDbl(Text1(ptxNIN10).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES10).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES10).Text), "000.00"))
    End If
    
    
    If Trim(Text1(ptxNIN10).Text) = "" And _
         Trim(Text1(ptxTIMES10).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN10).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES10).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) + 1, "000.00"))
        End If
    End If
    
                                                '他責　名称
    Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NAME, Combo1(pcmbTASEKI).Text)
    If IsNumeric(Text1(ptxNIN11).Text) Then        '人
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) + CDbl(Text1(ptxNIN11).Text), "0.0"))
    End If

    If IsNumeric(Text1(ptxTIMES11).Text) Then      '時間
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) + CDbl(Text1(ptxTIMES11).Text), "000.00"))
    End If
                                                        
    If Trim(Text1(ptxNIN11).Text) = "" And _
         Trim(Text1(ptxTIMES11).Text) = "" Then
    Else
        If Trim(Text1(ptxNIN11).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NIN, Format(CDbl(StrConv(P_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) + 1, "0.0"))
        End If
    
        If Trim(Text1(ptxTIMES11).Text) = "" Then
            Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_TIMES, Format(CDbl(StrConv(P_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) + 1, "000.00"))
        End If
    End If
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        '受入数量
    Call UniCode_Conv(P_SSHIJI_O_REC.UKEIRE_QTY, Format(CDbl(CDbl(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) + CDbl(Text1(ptxUKEIRE_QTY).Text)), "00000000.00"))
                                                        '更新日時
    Call UniCode_Conv(P_SSHIJI_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    If com = BtOpUpdate Then
                        sts = BTRV(BtOpUnlock, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "商品化指図ﾃﾞｰﾀ(親)")
                        End If
                    End If
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "商品化指図ﾃﾞｰﾀ(親)")
                GoTo Abort_Tran
        End Select
    
    Loop
    
    
    SEQNO = 0
    
    
    
    '受入履歴ﾃﾞｰﾀ処理
    Call UniCode_Conv(K0_P_SUKEIRE.SHIJI_No, Text1(ptxSHIJI_NO).Text)
    Call UniCode_Conv(K0_P_SUKEIRE.SEQNO, "")
    
    com = BtOpGetGreater
    
    Do
    
        sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_SUKEIRE_REC.SHIJI_No, vbUnicode) <> Text1(ptxSHIJI_NO).Text Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "商品化受入履歴")
                GoTo Abort_Tran
        End Select
        
        
        
        SEQNO = SEQNO + 1
        
        
        com = BtOpGetNext
        
    Loop
        
        
    Call UniCode_Conv(P_SUKEIRE_REC.SHIJI_No, Text1(ptxSHIJI_NO).Text)          '指図票№
                                                                                '仕向け先ｺｰﾄﾞ
    Call UniCode_Conv(P_SUKEIRE_REC.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                                                                                '受入日
    Call UniCode_Conv(P_SUKEIRE_REC.UKEIRE_DT, Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD"))
                                                                                '受入数量
    Call UniCode_Conv(P_SUKEIRE_REC.UKEIRE_QTY, Format(CDbl(Text1(ptxUKEIRE_QTY).Text), "00000000.00"))
        
        
        
    '前1
    If IsNumeric(Text1(ptxNIN01).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, Format(CDbl(Text1(ptxNIN01).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES01).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, Format(CDbl(Text1(ptxTIMES01).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, "001.00")
        End If
    End If
        
        
    '前2
    If IsNumeric(Text1(ptxNIN02).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, Format(CDbl(Text1(ptxNIN02).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES02).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, Format(CDbl(Text1(ptxTIMES02).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, "001.00")
        End If
    End If
        
        
    '前3
    If IsNumeric(Text1(ptxNIN03).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, Format(CDbl(Text1(ptxNIN03).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES03).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, Format(CDbl(Text1(ptxTIMES03).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, "001.00")
        End If
    End If
        
    '作業1
    If IsNumeric(Text1(ptxNIN04).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).NIN, Format(CDbl(Text1(ptxNIN04).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES04).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, Format(CDbl(Text1(ptxTIMES04).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(0).TIMES, "001.00")
        End If
    End If
        
    '作業2
    If IsNumeric(Text1(ptxNIN05).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).NIN, Format(CDbl(Text1(ptxNIN05).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES05).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, Format(CDbl(Text1(ptxTIMES05).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(1).TIMES, "001.00")
        End If
    End If
        
    '作業3
    If IsNumeric(Text1(ptxNIN06).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).NIN, Format(CDbl(Text1(ptxNIN06).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES06).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, Format(CDbl(Text1(ptxTIMES06).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, "001.00")
        End If
    End If
        
    '後1
    If IsNumeric(Text1(ptxNIN07).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, Format(CDbl(Text1(ptxNIN07).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES07).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, Format(CDbl(Text1(ptxTIMES07).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(2).TIMES, "001.00")
        End If
    End If
        
    '後2
    If IsNumeric(Text1(ptxNIN08).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, Format(CDbl(Text1(ptxNIN08).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES08).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, Format(CDbl(Text1(ptxTIMES08).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, "001.00")
        End If
    End If
        
    '加工1
    If IsNumeric(Text1(ptxNIN09).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, Format(CDbl(Text1(ptxNIN09).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES09).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, Format(CDbl(Text1(ptxTIMES09).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, "000.00")
    End If
    
    If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, "001.00")
        End If
    End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_NAME, Combo1(pcmbJISEKI).Text)
    If IsNumeric(Text1(ptxNIN10).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_NIN, Format(CDbl(Text1(ptxNIN10).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES10).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_TIMES, Format(CDbl(Text1(ptxTIMES10).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_TIMES, "000.00")
    End If
    
    
    If CDbl(StrConv(P_SUKEIRE_REC.JISEKI_NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.JISEKI_TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.JISEKI_NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_NIN, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.JISEKI_TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_TIMES, "001.00")
        End If
    End If
    
    
    
    
                                                '他責　名称
    Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_NAME, Combo1(pcmbTASEKI).Text)
    If IsNumeric(Text1(ptxNIN11).Text) Then        '人
        Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_NIN, Format(CDbl(Text1(ptxNIN11).Text), "0.0"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_NIN, "0.0")
    End If

    If IsNumeric(Text1(ptxTIMES11).Text) Then      '時間
        Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_TIMES, Format(CDbl(Text1(ptxTIMES11).Text), "000.00"))
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_TIMES, "000.00")
    End If
        
    If CDbl(StrConv(P_SUKEIRE_REC.TASEKI_NIN, vbUnicode)) = 0 And _
         CDbl(StrConv(P_SUKEIRE_REC.TASEKI_TIMES, vbUnicode)) = 0 Then
    Else
        If CDbl(StrConv(P_SUKEIRE_REC.TASEKI_NIN, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_NIN, "1.0")
        End If
    
        If CDbl(StrConv(P_SUKEIRE_REC.TASEKI_TIMES, vbUnicode)) = 0 Then
            Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_TIMES, "001.00")
        End If
    End If
        
        
        
        
    If CDbl(Text1(ptxMI_QTY).Text) = 0 Then
        Call UniCode_Conv(P_SUKEIRE_REC.LAST_F, P_UKEIRE_END)
    Else
        Call UniCode_Conv(P_SUKEIRE_REC.LAST_F, P_UKEIRE_CON)
    End If
                                                '取引先
    Call UniCode_Conv(P_SUKEIRE_REC.TORI_CODE, StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
            
            
                                                '備考   2010.09.21
    Call UniCode_Conv(P_SUKEIRE_REC.BIKOU, Text1(ptxBIKOU).Text)
        
    Call UniCode_Conv(P_SUKEIRE_REC.FILLER, "")
                                                        '更新日時
    Call UniCode_Conv(P_SUKEIRE_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
            
        SEQNO = SEQNO + 1
                                                '追番
        Call UniCode_Conv(P_SUKEIRE_REC.SEQNO, Format(SEQNO, "000"))
            
        DoEvents
            
        sts = BTRV(BtOpInsert, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrDuplicates
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "商品化受入履歴")
                GoTo Abort_Tran
        End Select
        
    Loop

End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Call Input_UnLock
                                        '印刷に対象指図票№を通知
    Taget_Key = Text1(ptxSHIJI_NO).Text
    BUNNOU_CNT = SEQNO
    
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbSHIMUKE        '仕向け先
        Case pcmbUKEHARAI       '手配先
            Text1(ptxUKEHARAI_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).Text, 5))
    End Select
    
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbSHIMUKE        '仕向け先
        Case pcmbUKEHARAI       '手配先
            Text1(ptxUKEHARAI_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).Text, 5))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim rpt         As New PI00020F1
Dim f           As New PI000202

Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd        '更新
            
            
            For i = ptxUKEIRE_DT To ptxTOTAL
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.05.07
'            If HINBAN_CHECK = 1 Then
'                If Trim(lblHinban_Check1(1).Caption) = "" Then
'
'
'                    ans = MsgBox("品番未チェックです。処理を継続しますか？", vbYesNo + vbQuestion, "確認入力")
'
'                    If ans = vbNo Then
'
'                        Text1(ptxUKEIRE_DT).SetFocus
'                        Exit Sub
'                    End If
'                End If
'            End If



            Select Case HINBAN_COMPO_CHECK
                Case 0
                Case 1
                
                    If HINBAN_CHECK = 1 Or COMPO_CHECK = 1 Then
                    
                    
                        If Trim(lblHinban_Check1(0).Caption) = "未チェック" And Trim(lblCOMPO_Check1(0).Caption) = "未チェック" Then
                    
                            ans = MsgBox("品番/構成未チェックです。処理を継続しますか？", vbYesNo + vbQuestion, "確認入力")
        
                            If ans = vbNo Then
        
                                Text1(ptxUKEIRE_DT).SetFocus
                                Exit Sub
                            End If
                        
                        Else
                            If Trim(lblHinban_Check1(0).Caption) = "未チェック" And lblCOMPO_Check1(0).ForeColor = vbRed Then
                        
                        
                                ans = MsgBox("構成チェック未完了です。" & Chr(10) & "処理を継続しますか？", vbYesNo + vbQuestion, "確認入力")
            
                                If ans = vbNo Then
            
                                    Text1(ptxUKEIRE_DT).SetFocus
                                    Exit Sub
                                End If
                            End If
                        
                        
                        End If
                
                
                    End If
                
                Case 2
            
            
                    If HINBAN_CHECK = 1 Then
            
                        If lblHinban_Check1(0).ForeColor = vbRed Then
                    
                            ans = MsgBox("品番未チェックです。処理を継続しますか？", vbYesNo + vbQuestion, "確認入力")
        
                            If ans = vbNo Then
        
                                Text1(ptxUKEIRE_DT).SetFocus
                                Exit Sub
                            End If
                        End If
            
                    End If
            
            
                    If COMPO_CHECK = 1 Then
            
            
                        If lblCOMPO_Check1(0).ForeColor = vbRed Then
                
                            ans = MsgBox("構成未チェックです。処理を継続しますか？", vbYesNo + vbQuestion, "確認入力")
        
                            If ans = vbNo Then
        
                                Text1(ptxUKEIRE_DT).SetFocus
                                Exit Sub
                            End If
                        End If
                End If
            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.05.07
            
            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
                
                Taget_Key = Text1(ptxSHIJI_NO).Text
                
                If CDbl(Text1(ptxMI_QTY).Text) <> 0 Then
                
                    Set rpt = New PI00020F1
                
                    'レポートを印刷します。（true：印刷ダイアログあり false：なし）
                    rpt.PrintReport False
                
                    Set rpt = Nothing


'                    f.RunReport rpt
'                    f.Show
                End If
                
                
                If Init_Proc(1) Then        '2007.09.28 引数追加
                    Unload Me
                End If
            
                Text1(ptxSHIJI_NO).SetFocus
            
            
            Else
                Text1(ptxUKEIRE_DT).SetFocus
            End If
        Case P_CMD_DEL                      '削除
        
            '指図票ﾃﾞｰﾀ（親）
            Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Text1(ptxSHIJI_NO))
            sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
                
            Select Case sts
                Case BtNoErr
                                   
                    If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                        MsgBox "完了登録済です。"
                        Text1(ptxSHIJI_NO).SetFocus
                        Exit Sub
                    End If
                
                    If CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
                        MsgBox "完了実績が有ります。"
                        Text1(ptxSHIJI_NO).SetFocus
                        Exit Sub
                    End If
                
                    If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                        MsgBox "キャンセル済です。"
                        Text1(ptxSHIJI_NO).SetFocus
                        Exit Sub
                    End If
                
                Case BtErrKeyNotFound
                    MsgBox "商品化指図未登録です。"
                    Text1(ptxSHIJI_NO).SetFocus
                    Exit Sub
                Case Else
                    Unload Me
            End Select
                    
        
        
            Beep
            ans = MsgBox("キャンセルしますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Cancel_Proc() Then
                    Unload Me
                End If
                
                '20011.04.17
'''                If Init_Proc() Then
                If Init_Proc(1) Then
                    Unload Me
                End If
            
                Text1(ptxSHIJI_NO).SetFocus
            
            
            Else
                Text1(ptxUKEIRE_DT).SetFocus
            End If
        
        Case P_CMD_DSP                      '検索/表示
        
            If List_Disp_Proc() Then
                Exit Sub
            End If
        
            'ｿｰﾄ情報の初期化
            For i = 0 To UBound(Sort_Tbl)
                Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
            Next i
        
            Sort_Tbl(colHIN_NAME) = 9       'ｿｰﾄ除外
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
            
            
            
            
        Case P_CMD_End                      '終了
            Unload Me
    End Select

End Sub


Private Sub Form_DblClick()
'    PrintForm                      '2017.07.22
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c       As String * 128
Dim sts     As Integer
Dim i       As Integer


'    If App.PrevInstance Then
'        Beep
'        MsgBox "同一プログラム実行中です。"
'        End
'    End If
                                
                                
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]商品化単価見積作成処理", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.02.03 P_SYS.INI --> PI00020.INI
                                
                                'POSｼｽﾃﾑ有無の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", App.EXEName, c) Then
        POS_UMU = False
    Else
        If RTrim(c) = "0" Then
            POS_UMU = False
        Else
            POS_UMU = True
        End If
    End If
                                'ﾊﾞｰｺｰﾄﾞ印字
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "BCR", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "BCR", App.EXEName, c) Then
        PRI_MAIN_BCR = False
    Else
        If RTrim(c) = "0" Then
            PRI_MAIN_BCR = False
        Else
            If Not POS_UMU Then
                MsgBox "ＰＯＳｼｽﾃﾑが未設定です。処理を中止します。"
                End
            End If
            PRI_MAIN_BCR = True
        End If
    End If
                                    '明細備考印字内容
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "DET_BIKOU", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DET_BIKOU", App.EXEName, c) Then
        PRI_BIKOU_BCR = False
    Else
        If RTrim(c) = "0" Then
            PRI_BIKOU_BCR = False
        Else
            If Not POS_UMU Then
                MsgBox "ＰＯＳｼｽﾃﾑが未設定です。処理を中止します。"
                End
            End If
            PRI_BIKOU_BCR = True
        End If
    End If
                                '収単／担当者の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "S_TANTO", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "S_TANTO", App.EXEName, c) Then
        PRI_S_TANTO = False
    Else
        If RTrim(c) = "0" Then
            PRI_S_TANTO = False
        Else
            PRI_S_TANTO = True
        End If
    End If
                                
    Label1(54).Visible = PRI_S_TANTO
    Combo1(pcmbS_TANTO).Visible = PRI_S_TANTO
    
    TDBGrid1.Columns(colS_TANTO).Visible = PRI_S_TANTO
    
    
                                '作業日／数量／担当 2007.05.22
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "SAGYO_DAY", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "SAGYO_DAY", App.EXEName, c) Then
        PRI_SAGYO_DAY = False
    Else
        If RTrim(c) = "0" Then
            PRI_SAGYO_DAY = False
        Else
            PRI_SAGYO_DAY = True
        End If
    End If
    
    
                                
                                
                                '商品検査　同梱の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "DOUKON", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DOUKON", App.EXEName, c) Then
        PRI_DOUKON = False
    Else
        If RTrim(c) = "0" Then
            PRI_DOUKON = False
        Else
            PRI_DOUKON = True
        End If
    End If
                                '入庫完了印の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "NYUKO_IN", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "NYUKO_IN", App.EXEName, c) Then
        PRI_NYUKO_IN = False
    Else
        If RTrim(c) = "0" Then
            PRI_NYUKO_IN = False
        Else
            PRI_NYUKO_IN = True
        End If
    End If
                                '入力完了印の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "INPUT_IN", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "INPUT_IN", App.EXEName, c) Then
        PRI_INPUT_IN = False
    Else
        If RTrim(c) = "0" Then
            PRI_INPUT_IN = False
        Else
            PRI_INPUT_IN = True
        End If
    End If
                                
    '下部　品番／№／数量   2007.05.22
    If PRI_NYUKO_IN Or PRI_INPUT_IN Then
    Else
'        If GetIni(StrConv(App.EXEName, vbUpperCase), "HINBAN_BIKOU", "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), "HINBAN_BIKOU", App.EXEName, c) Then
            PRI_HINBAN_BIKOU = False
        Else
            If RTrim(c) = "0" Then
                PRI_HINBAN_BIKOU = False
            Else
                PRI_HINBAN_BIKOU = True
            End If
        End If
    End If
                                
                                '自責
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", App.EXEName, c) Then
        JISEKI_TITLE = ""
    Else
        JISEKI_TITLE = Split(Trim(c), ",", -1)
    End If
    
                                '他責
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", App.EXEName, c) Then
        TASEKI_TITLE = ""
    Else
        TASEKI_TITLE = Split(Trim(c), ",", -1)
    End If

                                
                                
                                
    '秒／分取り込み 2008.08.19
'   If GetIni(StrConv(App.EXEName, vbUpperCase), "JISSEKI_DSP", "P_SYS", c) Then
   If GetIni(StrConv(App.EXEName, vbUpperCase), "JISSEKI_DSP", App.EXEName, c) Then
        JISSEKI_DSP = "m"
    Else
        JISSEKI_DSP = Trim(c)
        If JISSEKI_DSP <> "m" And JISSEKI_DSP <> "s" Then
            JISSEKI_DSP = "m"
        End If
    End If
                                
                                
                                
    '最大読込み件数 2010.09.21
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "LOAD_LIMIT", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "LOAD_LIMIT", App.EXEName, c) Then
        LOAD_LIMIT = -1
    Else
        If Not IsNumeric(Trim(c)) Then
            LOAD_LIMIT = -1
        Else
            LOAD_LIMIT = Val(Trim(c))
        End If
    End If
                                
                                
                                
    '品番処理チェックコントロール 2010.09.21
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "HINBAN_CHECK", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "HINBAN_CHECK", App.EXEName, c) Then
        HINBAN_CHECK = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            HINBAN_CHECK = 0
        Else
        
            Select Case Trim(c)
            
                Case "0", "1", "2"
                    HINBAN_CHECK = Val(Trim(c))
                Case Else
                    HINBAN_CHECK = 0
            
            End Select
        
        End If
    End If
                                
    Label1(22).ToolTipText = "HINBAN_CHECK=" & HINBAN_CHECK
                                    
    '構成処理チェックコントロール 2011.05.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "COMPO_CHECK", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "COMPO_CHECK", App.EXEName, c) Then
        COMPO_CHECK = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            COMPO_CHECK = 0
        Else
        
            Select Case Trim(c)
            
                Case "0", "1", "2"
                    COMPO_CHECK = Val(Trim(c))
                Case Else
                    COMPO_CHECK = 0
            
            End Select
        
        End If
    End If
    Label1(23).ToolTipText = "COMPO_CHECK=" & COMPO_CHECK
    '構成処理チェックコントロール 2011.05.07
                                    
                                    
    '品番処理チェック/構成処理チェックコントロール 2011.05.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "HINBAN_COMPO_CHECK", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "HINBAN_COMPO_CHECK", App.EXEName, c) Then
        HINBAN_COMPO_CHECK = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            HINBAN_COMPO_CHECK = 0
        Else
        
            Select Case Trim(c)
            
                Case "0", "1", "2"
                    HINBAN_COMPO_CHECK = Val(Trim(c))
                Case Else
                    HINBAN_COMPO_CHECK = 0
            
            End Select
        
        End If
    End If
    '品番処理チェック/構成処理チェックコントロール 2011.05.07
                                
                                
                                
                                
    '外装ﾁｪｯｸ　表示/非表示 2015.11.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "GAISOU_CHECK", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "GAISOU_CHECK", App.EXEName, c) Then
        GAISOU_CHECK = False
    Else
        
            Select Case Trim(c)
            
                Case "1"
                    GAISOU_CHECK = True
                Case Else
                    GAISOU_CHECK = False
            
            End Select
        
    End If

    'チェックOKの設定   2016.03.09
    If GetIni(StrConv(App.EXEName, vbUpperCase), "CHECK_OK", App.EXEName, c) Then
        CHECK_OK = 0
    Else
            Select Case Trim(c)
            
                Case "1"
                    CHECK_OK = 1
                Case Else
                    CHECK_OK = 0
            
            End Select
        
    End If


    'DEF_INP_SURYOの設定   2017.08.09
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DEF_INP_SURYO", App.EXEName, c) Then
        DEF_INP_SURYO = 0
    Else
        If Trim(c) = "1" Then
            DEF_INP_SURYO = 1
        Else
            DEF_INP_SURYO = 0
        End If
    End If


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.02.03 P_SYS.INI --> PI00020.INI
                                
                                
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'クラスマスタＯＰＥＮ
    If P_Class_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指図（子）ﾃﾞｰﾀＯＰＥＮ
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指図（親）ﾃﾞｰﾀＯＰＥＮ
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指図受入履歴ﾃﾞｰﾀＯＰＥＮ
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '原産国ＯＰＥＮ     '2016.01.06
    If GENSAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '商品化指図（親）ﾜｰｸＯＰＥＮ
    If wP_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If SE_LOC_TANKA_M_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    
    
'    Load PI000201
    
    
    PI000201.Caption = PI000201.Caption & LAST_UPDATE_DAY   '2017.07.22
    
    
    '管理マスタの読み込み
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)

    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            If P_KANRI_MAKE_Proc() Then
                Unload Me
            End If
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ")
            Unload Me
    End Select
        
    
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    
    '収単／担当者のセット
    If Code_Set_Proc(pcmbS_TANTO, P_KBN05_CD, 1) Then
        Unload Me
    End If
    
    
    '受払先
    If Ukeharai_Set_Proc() Then
        Unload Me
    End If
    
    
    
    Doukon_Tbl_No(0) = "①"
    Doukon_Tbl_No(1) = "②"
    Doukon_Tbl_No(2) = "③"
    Doukon_Tbl_No(3) = "④"
    Doukon_Tbl_No(4) = "⑤"
    Doukon_Tbl_No(5) = "⑥"
    Doukon_Tbl_No(6) = "⑦"
    Doukon_Tbl_No(7) = "⑧"
    Doukon_Tbl_No(8) = "⑨"
    Doukon_Tbl_No(9) = "⑩"
    Doukon_Tbl_No(10) = "⑪"
    Doukon_Tbl_No(11) = "⑫"
    Doukon_Tbl_No(12) = "⑬"
    Doukon_Tbl_No(13) = "⑭"
    Doukon_Tbl_No(14) = "⑮"
    Doukon_Tbl_No(15) = "⑯"
    Doukon_Tbl_No(16) = "⑰"
    Doukon_Tbl_No(17) = "⑱"
    Doukon_Tbl_No(18) = "⑲"
    Doukon_Tbl_No(19) = "⑳"
    
    
    
    
    
    '画面初期設定
    If Init_Proc(1) Then
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
    
                                            'クラスマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "クラスマスタ")
        End If
    End If
    
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
    
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
                                            '商品化指図ﾃﾞｰﾀ(親)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図ﾃﾞｰﾀ(親)")
        End If
    End If
                                            '商品化指図ﾃﾞｰﾀ(子)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図ﾃﾞｰﾀ(子)")
        End If
    End If
    
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
    
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
                                            '受入履歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受入履歴")
        End If
    End If
                                            '在庫ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫ﾃﾞｰﾀ")
        End If
    End If
                                            '商品化指図ﾜｰｸ(親)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), K0_wP_SSHIJI_O, Len(K0_wP_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図(親)ﾜｰｸ")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI000201 = Nothing
    Set PI000202 = Nothing

    End
End Sub





Private Sub TDBGrid1_DblClick()
Dim sts As Integer
    
    Text1(ptxSHIJI_NO).Text = SSHIJI(TDBGrid1.Bookmark, colSHIJI_NO)
    '指図票ﾃﾞｰﾀのﾁｪｯｸ
    sts = P_SSHIJI_Read_Proc()
    Select Case sts
        Case False, BtNoErr
                    
            If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                MsgBox "完了登録済みです。"
                TDBGrid1.SetFocus
                Exit Sub
            End If
        
            Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, "")
            Save_UKEIRE_QTY = 0
        
        
        Case BtErrKeyNotFound
            MsgBox "入力した項目はエラーです。"
            TDBGrid1.SetFocus
            Exit Sub
        Case Else
            Exit Sub
    End Select
    
        
    

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)


    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        SSHIJI.QuickSort Min_Row, SSHIJI.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = SSHIJI
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If



End Sub



Private Sub Text1_Change(Index As Integer)
        Select Case Index
        
        
            Case ptxTIMES01, ptxTIMES02, ptxTIMES03, ptxTIMES04, ptxTIMES05, ptxTIMES06, ptxTIMES07, ptxTIMES08, ptxTIMES09, ptxTIMES10, ptxTIMES11 '時間
        
        
                Time_Input_F = True
        
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
        
        
        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.07.22
    Select Case Index
        Case ptxUKEHARAI_CODE
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
    
    
    

        
        Case ptxS_CLASS_CODE To ptxN_CLASS_CODE                             '2017.08.05
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)     '2017.08.05
    
    
    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.07.22
        
        
        
        
        
    If JISSEKI_DSP = "s" Then


        If Time_Input_F Then

            Select Case Index
            
            
                Case ptxTIMES01, ptxTIMES02, ptxTIMES03, ptxTIMES04, ptxTIMES05, ptxTIMES06, ptxTIMES07, ptxTIMES08, ptxTIMES09, ptxTIMES10, ptxTIMES11 '時間
            
            
                    If IsNumeric(Text1(Index).Text) Then
                        Text1(Index).Text = Format(ToHalfAdjust(CCur(CInt(Text1(Index).Text) / 60), 0), "#0")
                    End If
            
                    Time_Input_F = False
            
            
            End Select
        
    
        End If
    
    End If
        
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
        
    If JISSEKI_DSP = "s" Then

        If Time_Input_F Then

            Select Case Index
            
            
                Case ptxTIMES01, ptxTIMES02, ptxTIMES03, ptxTIMES04, ptxTIMES05, ptxTIMES06, ptxTIMES07, ptxTIMES08, ptxTIMES09, ptxTIMES10, ptxTIMES11 '時間
            
            
                    Time_Input_F = False
            
            End Select
        
    
        End If
    
    End If
        
        
        
    Call Tab_Ctrl(Shift)        '移動
End Sub
Private Function Init_Proc(Optional Mode As Integer = 0) As Integer
'----------------------------------------------------------------------------
'                   入力画面の初期設定
'   引数追加        2007.09.28
'----------------------------------------------------------------------------
Dim i       As Integer
Dim sts     As Integer


    Init_Proc = True
    
    
    
    For i = ptxSHIJI_NO To ptxBIKOU
        Text1(i).Text = ""
    Next i

    '2011.04.17
    lblHinban_Check1(0).Caption = ""
    lblHinban_Check1(0).ForeColor = vbBlack
    lblHinban_Check1(1).Caption = ""
    lblHinban_Check1(1).ForeColor = vbBlack
    '2011.04.17


    '2012.05.07
    lblCOMPO_Check1(0).Caption = ""
    lblCOMPO_Check1(0).ForeColor = vbBlack
    lblCOMPO_Check1(1).Caption = ""
    lblCOMPO_Check1(1).ForeColor = vbBlack
    '2011.05.07


    For i = pcmbSHIMUKE To pcmbS_TANTO
        
        Combo1(i).ListIndex = -1
    
    Next i


    If JISSEKI_DSP = "s" Then           '2008.08.19
    
        Label1(24).Caption = "秒"
    
    Else
    
        Label1(24).Caption = "分"
    
    
    End If



    '実績
    
    Combo1(pcmbJISEKI).Clear
    Combo1(pcmbJISEKI).AddItem CStr(JISEKI_TITLE(0))
    Combo1(pcmbJISEKI).AddItem CStr(JISEKI_TITLE(1))
    Combo1(pcmbJISEKI).ListIndex = -1
    '他責
    Combo1(pcmbTASEKI).Clear
    Combo1(pcmbTASEKI).AddItem CStr(TASEKI_TITLE(0))
    Combo1(pcmbTASEKI).AddItem CStr(TASEKI_TITLE(1))
    Combo1(pcmbTASEKI).ListIndex = -1


    '読み込みﾊﾞｯﾌｧｰｸﾘｱｰ
    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, "")


    If Mode = 0 Then        '2007.09.28
        If List_Disp_Proc() Then
            Exit Function
        End If
    
    
        'ｿｰﾄ情報の初期化
        For i = 0 To UBound(Sort_Tbl)
            Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
        Next i
    
        Sort_Tbl(colHIN_NAME) = 9       'ｿｰﾄ除外

    End If

    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, "")
    Save_UKEIRE_QTY = 0
    
    Time_Input_F = False
    
    
    Init_Proc = False

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



Private Function P_SSHIJI_Read_Proc() As Integer
'----------------------------------------------------------------------------
'                   指図データの読み込み
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
    P_SSHIJI_Read_Proc = True
    
    
    '指図票ﾃﾞｰﾀ（親）
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Text1(ptxSHIJI_NO))
    sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        
        
        Case Else
            P_SSHIJI_Read_Proc = sts
            Exit Function
    
    End Select
    
    
    If Item_Disp_Proc() Then
        Exit Function
    End If
    
    P_SSHIJI_Read_Proc = False
        
    

End Function
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           指図票ﾃﾞｰﾀの表示
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

Dim Row     As Long

Dim LOAD_LIMIT_Cnt  As Long     '2010.09.21


    List_Disp_Proc = True
    PI000201.MousePointer = vbHourglass
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[商品化完了登録]未完了の指示票データを読込中です...(最大=" & Format(LOAD_LIMIT, "#0") & ")", Me.hwnd, 0)
    
    
    
    Set SSHIJI = Nothing
    
    
'    com = BtOpGetFirst
    com = BtOpGetLast
    
    Row = Min_Row - 1
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化指図ﾃﾞｰﾀ(親)")
                Exit Function
        End Select
    
    
        If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = P_KAN_ON Or _
            StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
        Else
            
            If LOAD_LIMIT <> -1 Then
                If Row >= LOAD_LIMIT Then
                    Exit Do
                End If
            End If
            
            Row = Row + 1
            
            
            
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        
'            Row = Row + 1
'            If Grid_Set_Proc(Row) Then
'                Exit Function
'            End If
        
        
        com = BtOpGetPrev
    
    Loop
    
    Set TDBGrid1.Array = SSHIJI
            
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[商品化完了登録]" & Format(Row, "#0") & "件　読込ました。", Me.hwnd, 0)
    
    
    PI000201.MousePointer = vbDefault
    List_Disp_Proc = False
    


End Function

Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           指図票ﾃﾞｰﾀの内容をｸﾞﾘｯﾄﾞにｾｯﾄする
'----------------------------------------------------------------------------
Dim sts As Integer

    Grid_Set_Proc = True
    
    SSHIJI.ReDim Min_Row, Row, Min_Col, Max_Col


    '発行日時
    If Trim(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode)) = "" Then
        SSHIJI(Row, colHAKKO_DT) = Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2)
    Else
        SSHIJI(Row, colHAKKO_DT) = Mid(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode), 7, 2) & " " & _
                                    Mid(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode), 9, 2) & ":" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode), 11, 2)
    End If
    '指図票№
    SSHIJI(Row, colSHIJI_NO) = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)
    '仕向け先
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Function
    End Select
    '仕向け先
    SSHIJI(Row, colSHIMUKE) = StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) & " " & _
                                StrConv(P_CODEREC.C_RNAME, vbUnicode)
        
    '手配先
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            Exit Function
    End Select
    '手配先
    SSHIJI(Row, colUKEHARAI) = StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode) & " " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
    '品番
    SSHIJI(Row, colHIN_GAI) = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    '品名
    SSHIJI(Row, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    '予定数
    SSHIJI(Row, colSHIJI_QTY) = Format(CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)), "#0")
    '完了数
    SSHIJI(Row, colUKEIRE_QTY) = Format(CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "#0")
    '未完数
    SSHIJI(Row, colMI_QTY) = Format(CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)) - _
                                        CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "#0")
        
    '収単／担当者
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN05_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SSHIJI_O_REC.S_TANTO, vbUnicode))
    
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Function
    End Select
    
    
    
    SSHIJI(Row, colS_TANTO) = Trim(StrConv(P_SSHIJI_O_REC.S_TANTO, vbUnicode)) & " " & Trim(StrConv(P_CODEREC.C_NAME, vbUnicode))
    
    
    
    Grid_Set_Proc = False

End Function

Public Function wP_SSHIJI_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              商品化指図(親)ワーク  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    wP_SSHIJI_O_Open = True
                                            '商品化指図(親)ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_O]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "商品化指図(親)ﾜｰｸ")
                Exit Function
        End Select
    Loop
    
    wP_SSHIJI_O_Open = False

End Function

Private Sub Text1_LostFocus(Index As Integer)

    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.07.22
    Select Case Index
        Case ptxUKEHARAI_CODE
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
    
    
    

        
        Case ptxS_CLASS_CODE To ptxN_CLASS_CODE                             '2017.08.05
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)     '2017.08.05
    
    
    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.07.22
    
    
    
    
    If JISSEKI_DSP <> "s" Then
        Exit Sub
    End If

    If Time_Input_F Then

        Select Case Index
        
        
            Case ptxTIMES01, ptxTIMES02, ptxTIMES03, ptxTIMES04, ptxTIMES05, ptxTIMES06, ptxTIMES07, ptxTIMES08, ptxTIMES09, ptxTIMES10, ptxTIMES11 '時間
        
        
                If IsNumeric(Text1(Index).Text) Then
                    Text1(Index).Text = Format(ToHalfAdjust(CCur(CInt(Text1(Index).Text) / 60), 0), "#0")
                End If
            
        End Select
        
        Time_Input_F = False
        
    End If

End Sub

' ------------------------------------------------------------------------
'       指定した精度の数値に四捨五入します。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に四捨五入された数値。
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function

