VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PR000501 
   Caption         =   "商品化実績集計表"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
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
   ScaleHeight     =   10305
   ScaleWidth      =   15150
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   8880
      MaxLength       =   10
      TabIndex        =   76
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2160
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   6135
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   10821
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "仕向け先"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ｸﾗｽ"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "【内部】件数"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "【内部】数量"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "【外部】件数"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "【外部】数量"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "【合計】件数"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "【合計】数量"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "【合計】単価"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "【合計】金額"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "【資材】単価"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "【資材】金額"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "【工料】単価"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "【工料】金額"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "【他】単価"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "【他】金額"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "【原価】個装"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "【原価】外装"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "【原価】工料"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   19
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=19"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1588"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1482"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2381"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2275"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2381"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2381"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2275"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2381"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2275"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2381"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2275"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2381"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2275"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2381"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2275"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2381"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2275"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=2381"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=2275"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=2381"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=2275"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(61)=   "Column(12).Width=2381"
      Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=2275"
      Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(66)=   "Column(13).Width=2381"
      Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=2275"
      Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=2"
      Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(71)=   "Column(14).Width=2381"
      Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=2275"
      Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(76)=   "Column(15).Width=2381"
      Splits(0)._ColumnProps(77)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(15)._WidthInPix=2275"
      Splits(0)._ColumnProps(79)=   "Column(15)._ColStyle=2"
      Splits(0)._ColumnProps(80)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(81)=   "Column(16).Width=2381"
      Splits(0)._ColumnProps(82)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(16)._WidthInPix=2275"
      Splits(0)._ColumnProps(84)=   "Column(16)._ColStyle=2"
      Splits(0)._ColumnProps(85)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(86)=   "Column(17).Width=2381"
      Splits(0)._ColumnProps(87)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(88)=   "Column(17)._WidthInPix=2275"
      Splits(0)._ColumnProps(89)=   "Column(17)._ColStyle=2"
      Splits(0)._ColumnProps(90)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(91)=   "Column(18).Width=2381"
      Splits(0)._ColumnProps(92)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(18)._WidthInPix=2275"
      Splits(0)._ColumnProps(94)=   "Column(18)._ColStyle=2"
      Splits(0)._ColumnProps(95)=   "Column(18).Order=19"
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
      Caption         =   "生産集計明細"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
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
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=110,.parent=43"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=44"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=45"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=47"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=16,.parent=43,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=44"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=45"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=47"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=28,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(53)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(54)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=44"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=45"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=47"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(59)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(60)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=32,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(65)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(66)  =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=21,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=22,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=23,.parent=47"
      _StyleDefs(82)  =   "Splits(0).Columns(9).Style:id=78,.parent=43,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=62,.parent=43,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=47"
      _StyleDefs(90)  =   "Splits(0).Columns(11).Style:id=70,.parent=43,.alignment=1"
      _StyleDefs(91)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=44"
      _StyleDefs(92)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=45"
      _StyleDefs(93)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=47"
      _StyleDefs(94)  =   "Splits(0).Columns(12).Style:id=82,.parent=43,.alignment=1"
      _StyleDefs(95)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=44"
      _StyleDefs(96)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=45"
      _StyleDefs(97)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=47"
      _StyleDefs(98)  =   "Splits(0).Columns(13).Style:id=86,.parent=43,.alignment=1"
      _StyleDefs(99)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=44"
      _StyleDefs(100) =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=45"
      _StyleDefs(101) =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=47"
      _StyleDefs(102) =   "Splits(0).Columns(14).Style:id=90,.parent=43,.alignment=1"
      _StyleDefs(103) =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=44"
      _StyleDefs(104) =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=45"
      _StyleDefs(105) =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=47"
      _StyleDefs(106) =   "Splits(0).Columns(15).Style:id=94,.parent=43,.alignment=1"
      _StyleDefs(107) =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=44"
      _StyleDefs(108) =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=45"
      _StyleDefs(109) =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=47"
      _StyleDefs(110) =   "Splits(0).Columns(16).Style:id=98,.parent=43,.alignment=1"
      _StyleDefs(111) =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=44"
      _StyleDefs(112) =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=45"
      _StyleDefs(113) =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=47"
      _StyleDefs(114) =   "Splits(0).Columns(17).Style:id=102,.parent=43,.alignment=1"
      _StyleDefs(115) =   "Splits(0).Columns(17).HeadingStyle:id=99,.parent=44"
      _StyleDefs(116) =   "Splits(0).Columns(17).FooterStyle:id=100,.parent=45"
      _StyleDefs(117) =   "Splits(0).Columns(17).EditorStyle:id=101,.parent=47"
      _StyleDefs(118) =   "Splits(0).Columns(18).Style:id=106,.parent=43,.alignment=1"
      _StyleDefs(119) =   "Splits(0).Columns(18).HeadingStyle:id=103,.parent=44"
      _StyleDefs(120) =   "Splits(0).Columns(18).FooterStyle:id=104,.parent=45"
      _StyleDefs(121) =   "Splits(0).Columns(18).EditorStyle:id=105,.parent=47"
      _StyleDefs(122) =   "Named:id=33:Normal"
      _StyleDefs(123) =   ":id=33,.parent=0"
      _StyleDefs(124) =   "Named:id=34:Heading"
      _StyleDefs(125) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(126) =   ":id=34,.wraptext=-1"
      _StyleDefs(127) =   "Named:id=35:Footing"
      _StyleDefs(128) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(129) =   "Named:id=36:Selected"
      _StyleDefs(130) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(131) =   "Named:id=37:Caption"
      _StyleDefs(132) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(133) =   "Named:id=38:HighlightRow"
      _StyleDefs(134) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(135) =   "Named:id=39:EvenRow"
      _StyleDefs(136) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(137) =   "Named:id=40:OddRow"
      _StyleDefs(138) =   ":id=40,.parent=33"
      _StyleDefs(139) =   "Named:id=41:RecordSelector"
      _StyleDefs(140) =   ":id=41,.parent=34"
      _StyleDefs(141) =   "Named:id=42:FilterBar"
      _StyleDefs(142) =   ":id=42,.parent=33"
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
      TabIndex        =   15
      Top             =   9720
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9720
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印 刷"
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
      TabIndex        =   12
      Top             =   9720
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9720
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9720
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "検 索"
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
      TabIndex        =   8
      Top             =   9720
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
      Index           =   3
      Left            =   2760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9720
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9720
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9720
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
      Index           =   0
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "〜"
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   75
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "仕向け先"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   74
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   11640
      TabIndex        =   73
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   11640
      TabIndex        =   72
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   11640
      TabIndex        =   71
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   11640
      TabIndex        =   70
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "合計　　　�@＋�A"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   9120
      TabIndex        =   69
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "外注工料　　　�A"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   9120
      TabIndex        =   68
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "�@計"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   10800
      TabIndex        =   67
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "外装"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   10800
      TabIndex        =   66
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   19
      Left            =   9120
      TabIndex        =   65
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "個装"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   10800
      TabIndex        =   64
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   11640
      TabIndex        =   63
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "仕入原価"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   9120
      TabIndex        =   62
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "　　　　　　　消費資材"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   9120
      TabIndex        =   61
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   7440
      TabIndex        =   60
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   7440
      TabIndex        =   59
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   7440
      TabIndex        =   58
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   7440
      TabIndex        =   57
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   7440
      TabIndex        =   56
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   7440
      TabIndex        =   55
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "価格構成"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   7440
      TabIndex        =   54
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   5760
      TabIndex        =   53
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   5760
      TabIndex        =   52
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   5760
      TabIndex        =   51
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   5760
      TabIndex        =   50
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   5760
      TabIndex        =   49
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   5760
      TabIndex        =   48
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   5760
      TabIndex        =   47
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   5760
      TabIndex        =   46
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   4080
      TabIndex        =   45
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   4080
      TabIndex        =   44
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   4080
      TabIndex        =   43
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4080
      TabIndex        =   42
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   41
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   4080
      TabIndex        =   40
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   4080
      TabIndex        =   39
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   38
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   37
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   36
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   35
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   34
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   33
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   32
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   31
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   30
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "合　　計"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   5760
      TabIndex        =   29
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "Ｂ外部生産"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   28
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "Ａ内部生産"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   27
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '実線
      Caption         =   "�Bその他"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   26
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '実線
      Caption         =   "�A工料の部"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   25
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '実線
      Caption         =   "�@資材の部"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1080
      TabIndex        =   24
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "内　訳"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   720
      TabIndex        =   23
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BorderStyle     =   1  '実線
      Caption         =   "（構成比率）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   22
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "生産金額"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   21
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BorderStyle     =   1  '実線
      Caption         =   "（構成比率）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   20
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "生産数量"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   19
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "生産件数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   18
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "項目/生産内訳"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "対象年月日"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   16
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "PR000501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'テキスト用添字
Private Const ptxSHIMUKE_CODE% = 0          '仕向け先
Private Const ptxS_YMD% = 1                 '開始　対象年月日
Private Const ptxE_YMD% = 2                 '終了　対象年月日
'コンボ用添字
Private Const pcmbSHIMUKE_CODE% = 0         '仕向け先



'表示用ラベル
Private Const plblNAI_CNT% = 0              '内部生産　生産件数
Private Const plblNAI_SURYO% = 1            '内部生産　生産数量
Private Const plblNAI_SU_RITU% = 2          '内部生産　生産数量構成率
Private Const plblNAI_KIN% = 3              '内部生産　生産金額
Private Const plblNAI_KIN_RITU% = 4         '内部生産　生産金額構成率

Private Const plblNAI_UCHI_SHIZAI% = 5      '内部生産  内訳　資材
Private Const plblNAI_UCHI_KOURYO% = 6      '内部生産  内訳　工料
Private Const plblNAI_UCHI_ETC% = 7         '内部生産  内訳　その他

Private Const plblGAI_CNT% = 8              '外部生産　生産件数
Private Const plblGAI_SURYO% = 9            '外部生産　生産数量
Private Const plblGAI_SU_RITU% = 10         '外部生産　生産数量構成率
Private Const plblGAI_KIN% = 11             '外部生産　生産金額
Private Const plblGAI_KIN_RITU% = 12        '外部生産　生産金額構成率

Private Const plblGAI_UCHI_SHIZAI% = 13     '外部生産  内訳　資材
Private Const plblGAI_UCHI_KOURYO% = 14     '外部生産  内訳　工料
Private Const plblGAI_UCHI_ETC% = 15        '外部生産  内訳　その他

Private Const plblGK_CNT% = 16              '合計　生産件数
Private Const plblGK_SURYO% = 17            '合計　生産数量
Private Const plblGK_SU_RITU% = 18          '合計　生産数量構成率
Private Const plblGK_KIN% = 19              '合計　生産金額
Private Const plblGK_KIN_RITU% = 20         '合計　生産金額構成率

Private Const plblGK_UCHI_SHIZAI% = 21      '合計  内訳　資材
Private Const plblGK_UCHI_KOURYO% = 22      '合計  内訳　工料
Private Const plblGK_UCHI_ETC% = 23         '合計  内訳　その他


Private Const plblKAKAKU_RITU% = 24         '価格構成　生産金額
Private Const plblSHIZAI_RITU% = 25         '価格構成　資材
Private Const plblKOURYO_RITU% = 26         '価格構成　工料
Private Const plblETC_RITU% = 27            '価格構成　その他

Private Const plblGENKA_KOSOU% = 28         '仕入原価　個装
Private Const plblGENKA_GAISOU% = 29        '仕入原価　外装
Private Const plblGENKA_SHIZAI% = 30        '仕入原価　消費資材計
Private Const plblGENKA_KOURYO% = 31        '仕入原価　工料
Private Const plblGENKA_GK% = 32            '仕入原価　合計





'Glid用環境---------------------------------

'仕入明細
Private Const pGridDETAIL% = 0


Private SEISAN      As New XArrayDB


Private Const Min_Row% = 1                  '最小行数
Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 18                 '最大列数

Private Const colSHIMUKE_CODE% = 0          '仕向け先
Private Const colCLASS_CODE% = 1            'ｸﾗｽｺｰﾄﾞ

Private Const colNAI_CNT% = 2               '内部　件数
Private Const colNAI_SURYO% = 3             '内部　数量

Private Const colGAI_CNT% = 4               '外部　件数
Private Const colGAI_SURYO% = 5             '外部　数量

Private Const colGK_CNT% = 6                '合計　件数
Private Const colGK_SURYO% = 7              '合計　数量
Private Const colGK_TANKA% = 8              '合計　単価
Private Const colGK_KIN% = 9                '合計　金額

Private Const colSHIZAI_TANKA% = 10         '資材　単価
Private Const colSHIZAI_KIN% = 11           '資材　金額
Private Const colKOURYO_TANKA% = 12         '工料　単価
Private Const colKOURYO_KIN% = 13           '工料　金額
Private Const colETC_TANKA% = 14            'その他　単価
Private Const colETC_KIN% = 15              'その他　金額

Private Const colGENKA_KOSOU% = 16          '仕入原価　個装
Private Const colGENKA_GAISOU% = 17         '仕入原価　外装
Private Const colGENKA_KOURYO% = 18         '仕入原価　工料




Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
Private Tbl_Set_F   As Boolean
'-------------------------------------  合計集計用
Dim GK_NAI_CNT      As Integer          '内部生産　件数
Dim GK_NAI_SURYO    As Double           '内部生産  数量
Dim GK_GAI_CNT      As Integer          '内部生産　件数
Dim GK_GAI_SURYO    As Double           '内部生産  数量
Dim GK_TANKA        As Double
    
Dim NAI_TANKA(0 To 2)   As Double       '内訳　内部生産単価
Dim NAI_KIN(0 To 2)     As Double       '内訳　内部生産金額
Dim GAI_TANKA(0 To 2)   As Double       '内訳　外部生産単価
Dim GAI_KIN(0 To 2)     As Double       '内訳　内部生産金額
    
    
Dim KO_GENKA        As Double           '個装　原価
Dim GA_GENKA        As Double           '外装　原価
Dim GK_GENKA        As Double           '外注工料

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PR000501.MousePointer = vbHourglass

    Call Ctrl_Lock(PR000501)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PR000501)


    PR000501.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts     As Integer
Dim com     As Integer
    
Dim i       As Integer
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        Case ptxSHIMUKE_CODE    '仕向け先ｺｰﾄﾞ
        
           
           Combo1(pcmbSHIMUKE_CODE).ListIndex = -1
           For i = 0 To Combo1(pcmbSHIMUKE_CODE).ListCount - 1
               If Text1(ptxSHIMUKE_CODE).Text = Left(Right(Combo1(pcmbSHIMUKE_CODE).List(i), 4), 2) Then
                   Combo1(pcmbSHIMUKE_CODE).ListIndex = i
                   Exit For
               End If
           
           Next i
        
        
        
        Case ptxS_YMD           '対象年月日
        
            
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0000/01/01"
            End If
            
            If Not IsDate(Text1(Mode).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")
            
            End If
        
        Case ptxE_YMD           '対象年月日
        
            
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "9999/12/31"
            End If
            
            If Not IsDate(Text1(Mode).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")
            
            End If
        
        
    End Select
        
        
    Error_Check_Proc = False
    

End Function


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    Select Case Index
        Case pcmbSHIMUKE_CODE       '仕向け先ｺｰﾄﾞ
        
            Text1(ptxSHIMUKE_CODE).Text = Trim(Left(Right(Combo1(pcmbSHIMUKE_CODE).Text, 4), 2))
    End Select
    
    
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    Select Case Index
        Case pcmbSHIMUKE_CODE       '仕向け先ｺｰﾄﾞ
        
            Text1(ptxSHIMUKE_CODE).Text = Trim(Left(Right(Combo1(pcmbSHIMUKE_CODE).Text, 4), 2))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim Data_Flg    As Boolean

Dim rpt             As New PR00050F1
Dim f               As New PR000502


Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd          '更新
        
        Case P_CMD_DEL          '削除
        
        Case P_CMD_DSP                      '検索/表示
        
            For i = ptxS_YMD To ptxE_YMD
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            If SUM_Make_Proc(Data_Flg) Then
                Exit Sub
            End If
            
            
            If Not Data_Flg Then
                MsgBox "対象ﾃﾞｰﾀがありません"
                Exit Sub
            End If
        
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
            Text1(ptxS_YMD).SetFocus
        
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        
        Case P_CMD_PRT                      '印刷
 
            For i = ptxS_YMD To ptxE_YMD
                                            'エラーチェック
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            Next i
            
            If SUM_Make_Proc(Data_Flg) Then
                Exit Sub
            End If
            
            
            If Not Data_Flg Then
                MsgBox "対象ﾃﾞｰﾀがありません"
                Exit Sub
            End If
                
            ans = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                Set rpt = New PR00050F1
            
                'レポートを印刷します。（true：印刷ダイアログあり false：なし）
                rpt.PrintReport False
            
                Set rpt = Nothing
                
                
'                f.RunReport rpt
'                f.Show
            
            End If
            
            Text1(ptxS_YMD).SetFocus
            
            
        Case P_CMD_End                      '終了
    
            Unload Me
    
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

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
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
                                'クラスマスタＯＰＥＮ
    If P_Class_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
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
                                '生産実績集計ﾃﾞｰﾀＯＰＥＮ
    If P_SEISAN_SUM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指示(親)ﾃﾞｰﾀＯＰＥＮ
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指示(子)ﾃﾞｰﾀＯＰＥＮ
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指示受入履歴ﾃﾞｰﾀＯＰＥＮ
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    Load PR000502
    
    
    
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
    
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE_CODE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    
    
    '画面初期設定
    If Init_Proc() Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            'クラスマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "クラスマスタ")
        End If
    End If
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
    
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "クラスマスタ")
        End If
    End If
                                            '生産実績集計ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "生産実績集計ﾃﾞｰﾀ")
        End If
    End If
                                            '商品化指示（親）ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指示（親）ﾃﾞｰﾀ")
        End If
    End If
                                            '商品化指示（子）ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指示（親）ﾃﾞｰﾀ")
        End If
    End If
                                            '商品化受入履歴ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化受入履歴ﾃﾞｰﾀ")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PR000501 = Nothing
    Set PR000502 = Nothing


    End
End Sub





Private Sub TDBGrid1_HeadClick(Index As Integer, ByVal ColIndex As Integer)



    Select Case Index
        
        Case pGridDETAIL        '生産実績明細
    
    
            If Sort_Tbl(ColIndex) = 0 Then
                Sort_Tbl(ColIndex) = 1
            Else
                If Sort_Tbl(ColIndex) = 1 Then
                    Sort_Tbl(ColIndex) = 0
                End If
            
            End If
            
            If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                            
                SEISAN.QuickSort Min_Row, SEISAN.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
                
                Set TDBGrid1(Index).Array = SEISAN
                
                TDBGrid1(Index).ReBind
                TDBGrid1(Index).Update
                TDBGrid1(Index).MoveFirst
        
        
            End If
    
    
    
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
        
        
    If Error_Check_Proc(Index) Then    'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動
End Sub
Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   入力画面の初期設定
'----------------------------------------------------------------------------
Dim i       As Integer
Dim sts     As Integer


    Init_Proc = True
    
    
    
    For i = ptxS_YMD To ptxE_YMD
        Text1(i).Text = ""
    Next i
    '処理年月日＝当日
    Text1(ptxS_YMD).Text = Format(Now, "YYYY/MM/DD")
    Text1(ptxE_YMD).Text = Format(Now, "YYYY/MM/DD")
    
    For i = pcmbSHIMUKE_CODE To pcmbSHIMUKE_CODE
        
        Combo1(i).ListIndex = -1
    
    Next i



    
    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0               'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    Init_Proc = False

End Function



Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           資材受入データの表示
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Row             As Long


Dim wkValue         As Double
Dim i               As Integer



    List_Disp_Proc = True
    PR000501.MousePointer = vbHourglass
    
    
    Call UniCode_Conv(K0_P_SEISAN_SUM.SHIMUKE_CODE, Text1(ptxSHIMUKE_CODE))
    Call UniCode_Conv(K0_P_SEISAN_SUM.CLASS_CODE, P_ClassSum_Key)
    
    sts = BTRV(BtOpGetEqual, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
        
            MsgBox "対象ﾃﾞｰﾀがありません"
            List_Disp_Proc = False
            Exit Function
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "生産実績集計ﾃﾞｰﾀ")
            Exit Function
    End Select
    
    
    
        
                                            '内部生産　生産件数
    lblItem(plblNAI_CNT).Caption = Format(CInt(StrConv(P_SEISAN_SUM_REC.GK_NAI_CNT, vbUnicode)), "#0")
                                            '外部生産　生産件数
    lblItem(plblGAI_CNT).Caption = Format(CInt(StrConv(P_SEISAN_SUM_REC.GK_GAI_CNT, vbUnicode)), "#0")
                                            '合計　 　 生産件数
    lblItem(plblGK_CNT).Caption = Format(CInt(lblItem(plblNAI_CNT).Caption) + CInt(lblItem(plblGAI_CNT).Caption), "#0")
                                            
                                            '内部生産　生産数量
    lblItem(plblNAI_SURYO).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.GK_NAI_SURYO, vbUnicode)), "#0")
                                            '外部生産　生産数量
    lblItem(plblGAI_SURYO).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.GK_GAI_SURYO, vbUnicode)), "#0")
                                            '合計　 　 生産数量
    lblItem(plblGK_SURYO).Caption = Format(CDbl(lblItem(plblNAI_SURYO).Caption) + CDbl(lblItem(plblGAI_SURYO).Caption), "#0")
                                            
                                            '内部生産  構成比率
    wkValue = CDbl(lblItem(plblNAI_SURYO).Caption) / (CDbl(lblItem(plblNAI_SURYO).Caption) + CDbl(lblItem(plblGAI_SURYO).Caption)) * 100
    lblItem(plblNAI_SU_RITU).Caption = Format(wkValue, "#0.00") & "%"
                                            
                                            '外部生産  構成比率
    lblItem(plblGAI_SU_RITU).Caption = Format(100 - wkValue, "#0.00") & "%"
                                            '構成  構成比率
    lblItem(plblGK_SU_RITU).Caption = "100.00%"
    
                                            '内部生産　生産金額
    lblItem(plblNAI_KIN).Caption = Format(wkValue, "#,##0")
                                            '外部生産　生産金額
    wkValue = CDbl(lblItem(plblGK_SURYO).Caption) * CDbl(StrConv(P_SEISAN_SUM_REC.GK_TANKA, vbUnicode))
    lblItem(plblGAI_KIN).Caption = Format(wkValue, "#,##0")
                                            '合計　生産金額
    lblItem(plblGK_KIN).Caption = Format(CDbl(lblItem(plblNAI_KIN).Caption) + CDbl(lblItem(plblGAI_KIN).Caption), "#,##0")
        
                                            '内部生産  構成比率
    If CDbl(lblItem(plblNAI_KIN).Caption) + CDbl(lblItem(plblGAI_KIN).Caption) = 0 Then
        wkValue = 0
    Else
        wkValue = CDbl(lblItem(plblNAI_KIN).Caption) / (CDbl(lblItem(plblNAI_KIN).Caption) + CDbl(lblItem(plblGAI_KIN).Caption))
    End If
    lblItem(plblNAI_KIN_RITU).Caption = Format(wkValue, "#0.00") & "%"
                                            
                                            '外部生産  構成比率
    lblItem(plblGAI_KIN_RITU).Caption = Format(100 - wkValue, "#0.00") & "%"
                                            '構成  構成比率
    lblItem(plblGK_KIN_RITU).Caption = "100.00%"
        
                                            '内部生産　内訳　資材
    lblItem(plblNAI_UCHI_SHIZAI).Caption = Format(CDbl(lblItem(plblNAI_SURYO).Caption) * _
                                                    CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).NAI_TANKA, vbUnicode)), "#,##0")
                                            '内部生産　内訳　工料
    lblItem(plblNAI_UCHI_KOURYO).Caption = Format(CDbl(lblItem(plblNAI_SURYO).Caption) * _
                                                    CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(1).NAI_TANKA, vbUnicode)), "#,##0")
                                            '内部生産　内訳　その他
    lblItem(plblNAI_UCHI_ETC).Caption = Format(CDbl(lblItem(plblNAI_SURYO).Caption) * _
                                                    CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(2).NAI_TANKA, vbUnicode)), "#,##0")
        
                                            '外部生産　内訳　資材
    lblItem(plblGAI_UCHI_SHIZAI).Caption = Format(CDbl(lblItem(plblGAI_SURYO).Caption) * _
                                                    CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).GAI_TANKA, vbUnicode)), "#,##0")
                                            '外部生産　内訳　工料
    lblItem(plblGAI_UCHI_KOURYO).Caption = Format(CDbl(lblItem(plblGAI_SURYO).Caption) * _
                                                    CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(1).GAI_TANKA, vbUnicode)), "#,##0")
                                            '外部生産　内訳　その他
    lblItem(plblGAI_UCHI_ETC).Caption = Format(CDbl(lblItem(plblGAI_SURYO).Caption) * _
                                                    CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(2).GAI_TANKA, vbUnicode)), "#,##0")
        
                                            '合計　内訳　資材
    lblItem(plblGK_UCHI_SHIZAI).Caption = Format(CDbl(lblItem(plblNAI_UCHI_SHIZAI).Caption) + CDbl(lblItem(plblGAI_UCHI_SHIZAI).Caption), "#,##0")
                                            '合計　内訳　工料
    lblItem(plblGK_UCHI_KOURYO).Caption = Format(CDbl(lblItem(plblNAI_UCHI_KOURYO).Caption) + CDbl(lblItem(plblGAI_UCHI_KOURYO).Caption), "#,##0")
                                            '合計　内訳　その他
    lblItem(plblGK_UCHI_ETC).Caption = Format(CDbl(lblItem(plblNAI_UCHI_ETC).Caption) + CDbl(lblItem(plblGAI_UCHI_ETC).Caption), "#,##0")
        
                                            '価格構成比
    lblItem(plblKAKAKU_RITU).Caption = "100.00%"
    If (CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + CDbl(lblItem(plblGK_UCHI_ETC).Caption)) = 0 Then
        lblItem(plblSHIZAI_RITU).Caption = "0.00"
    Else
        lblItem(plblSHIZAI_RITU).Caption = Format(CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) / (CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_ETC).Caption)) * 100, "#0.00") & "%"
    End If
    If (CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + CDbl(lblItem(plblGK_UCHI_ETC).Caption)) = 0 Then
        lblItem(plblKOURYO_RITU).Caption = "0.00"
    Else
        lblItem(plblKOURYO_RITU).Caption = Format(CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) / (CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_ETC).Caption)) * 100, "#0.00") & "%"
    End If
    
    If ((CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + CDbl(lblItem(plblGK_UCHI_ETC).Caption))) = 0 Then
        lblItem(plblETC_RITU).Caption = "0.00"
    Else
    
        lblItem(plblETC_RITU).Caption = Format(CDbl(lblItem(plblGK_UCHI_ETC).Caption) / (CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_ETC).Caption)) * 100, "#0.00") & "%"
    End If
                                            '消費資材
    
    
    lblItem(plblGENKA_KOSOU).Caption = Format(CDbl(lblItem(plblGK_SURYO).Caption) * CDbl(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode)), "#,##0")
    lblItem(plblGENKA_GAISOU).Caption = Format(CDbl(lblItem(plblGK_SURYO).Caption) * CDbl(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode)), "#,##0")
    lblItem(plblGENKA_SHIZAI).Caption = Format(CDbl(lblItem(plblGENKA_KOSOU).Caption) + CDbl(lblItem(plblGENKA_GAISOU).Caption), "#,##0")
    lblItem(plblGENKA_KOURYO).Caption = Format(CDbl(lblItem(plblGK_SURYO).Caption) * CDbl(StrConv(P_SEISAN_SUM_REC.GK_GENKA, vbUnicode)), "#,##0")
    lblItem(plblGENKA_GK).Caption = Format(CDbl(lblItem(plblGENKA_KOURYO).Caption) + CDbl(lblItem(plblGENKA_SHIZAI).Caption), "#,##0")
        
    
    '-------------------------------------  '実績明細のｾｯﾄ
    Set SEISAN = Nothing
    
    Row = Min_Row - 1
    
    Call UniCode_Conv(K0_P_SEISAN_SUM.SHIMUKE_CODE, Text1(ptxSHIMUKE_CODE))
    Call UniCode_Conv(K0_P_SEISAN_SUM.CLASS_CODE, P_ClassSum_Key)
    
    
    com = BtOpGetGreater
    
    
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "生産実績集計ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        If Trim(StrConv(P_SEISAN_SUM_REC.CLASS_CODE, vbUnicode)) = "" Then
        Else
    
            Row = Row + 1
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        
        com = BtOpGetNext
    
    
    
    Loop
    
    
    Set TDBGrid1(pGridDETAIL).Array = SEISAN
    TDBGrid1(pGridDETAIL).ReBind
    TDBGrid1(pGridDETAIL).Update
    TDBGrid1(pGridDETAIL).MoveFirst
    
    
    PR000501.MousePointer = vbDefault
    
    
    List_Disp_Proc = False
    


End Function


Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           生産実績の内容をｸﾞﾘｯﾄﾞにｾｯﾄする
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer

Dim wkValue     As Double


    Grid_Set_Proc = True
    
    
    


    
    
    
    
    SEISAN.ReDim Min_Row, Row, Min_Col, Max_Col


    '仕向け先
    SEISAN(Row, colSHIMUKE_CODE) = StrConv(P_SEISAN_SUM_REC.SHIMUKE_CODE, vbUnicode)
    'ｸﾗｽ
    SEISAN(Row, colCLASS_CODE) = StrConv(P_SEISAN_SUM_REC.CLASS_CODE, vbUnicode)
    '内部生産 件数
    SEISAN(Row, colNAI_CNT) = Format(CInt(StrConv(P_SEISAN_SUM_REC.GK_NAI_CNT, vbUnicode)), "#0")
    '内部生産 数量
    SEISAN(Row, colNAI_SURYO) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.GK_NAI_SURYO, vbUnicode)), "#0")
    '外部生産 件数
    SEISAN(Row, colGAI_CNT) = Format(CInt(StrConv(P_SEISAN_SUM_REC.GK_GAI_CNT, vbUnicode)), "#0")
    '外部生産 数量
    SEISAN(Row, colGAI_SURYO) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.GK_GAI_SURYO, vbUnicode)), "#0")
    '外部生産 件数
    SEISAN(Row, colGAI_CNT) = Format(CInt(StrConv(P_SEISAN_SUM_REC.GK_GAI_CNT, vbUnicode)), "#0")
    '合計 件数
    SEISAN(Row, colGK_CNT) = Format(CInt(SEISAN(Row, colNAI_CNT)) + CInt(SEISAN(Row, colGAI_CNT)), "#0")
    '合計 数量
    SEISAN(Row, colGK_SURYO) = Format(CDbl(SEISAN(Row, colNAI_SURYO)) + CDbl(SEISAN(Row, colGAI_SURYO)), "#0")
    
    '合計 単価
    SEISAN(Row, colGK_TANKA) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.GK_TANKA, vbUnicode)), "#,##0.00")
    '合計　金額
    SEISAN(Row, colGK_KIN) = Format(CDbl(SEISAN(Row, colGK_TANKA)) * CDbl(SEISAN(Row, colGK_SURYO)), "#,##0")
    
    '資材　単価
    SEISAN(Row, colSHIZAI_TANKA) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).NAI_TANKA, vbUnicode)) + _
                                            CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).GAI_TANKA, vbUnicode)), "#,##0.00")
    '資材　金額
    SEISAN(Row, colSHIZAI_KIN) = Format(CDbl(SEISAN(Row, colSHIZAI_TANKA)) * (CLng(SEISAN(Row, colNAI_SURYO) + _
                                            CLng(SEISAN(Row, colGAI_SURYO)))), "#,##0")
    
    '工料　単価
    SEISAN(Row, colKOURYO_TANKA) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(1).NAI_TANKA, vbUnicode)) + _
                                            CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(1).GAI_TANKA, vbUnicode)), "#,##0.00")
    '工料　金額
    SEISAN(Row, colKOURYO_KIN) = Format(CDbl(SEISAN(Row, colKOURYO_TANKA)) * (CLng(SEISAN(Row, colNAI_SURYO) + _
                                            CLng(SEISAN(Row, colGAI_SURYO)))), "#,##0")
    'その他　単価
    SEISAN(Row, colETC_TANKA) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(2).NAI_TANKA, vbUnicode)) + _
                                            CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(2).GAI_TANKA, vbUnicode)), "#,##0.00")
    'その他　金額
    SEISAN(Row, colETC_KIN) = Format(CDbl(SEISAN(Row, colETC_TANKA)) * (CLng(SEISAN(Row, colNAI_SURYO) + _
                                            CLng(SEISAN(Row, colGAI_SURYO)))), "#,##0")
    
    
    '仕入原価　個装
    SEISAN(Row, colGENKA_KOSOU) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode)) * CDbl(SEISAN(Row, colGK_SURYO)), "#,##0")
    '仕入原価　外装
    SEISAN(Row, colGENKA_GAISOU) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode)) * CDbl(SEISAN(Row, colGK_SURYO)), "#,##0")
    '仕入原価　工料
    SEISAN(Row, colGENKA_KOURYO) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.GK_GENKA, vbUnicode)) * CDbl(SEISAN(Row, colGK_SURYO)), "#,##0")
    
    
    
    
    
    
    
    
    Grid_Set_Proc = False

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

Private Function SUM_Make_Proc(Data_Flg As Boolean) As Integer
'----------------------------------------------------------------------------
'                   生産実績集計ﾃﾞｰﾀ作成
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim upd_com         As Integer
    
Dim Shizai_com      As Integer
    
Dim SKIP_Flg        As Boolean
    
Dim wkYMD           As String * 8
    
Dim wkValue         As Double
Dim wkSuryo         As Double
    
Dim wkURI_TANKA     As Double
Dim wkSHI_TANKA     As Double
    
    
    
Dim i               As Integer
    
    
    SUM_Make_Proc = True
    PR000501.MousePointer = vbHourglass

    '-----------------------------------------  集計ﾃﾞｰﾀ全件削除


    com = BtOpGetFirst



    Do
    
    
        sts = BTRV(com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "生産実績集計ﾃﾞｰﾀ")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "生産実績集計ﾃﾞｰﾀ")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
        
    '-----------------------------------------  集計処理開始
    
    Data_Flg = False
    Call UniCode_Conv(K1_P_SUKEIRE.SHIMUKE_CODE, Text1(ptxSHIMUKE_CODE).Text)

    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K1_P_SUKEIRE, Len(K1_P_SUKEIRE), 1)
            
        Select Case sts
            Case BtNoErr
                '仕向け先ｺｰﾄﾞ
                If Trim(Text1(ptxSHIMUKE_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxSHIMUKE_CODE).Text) <> Trim(StrConv(P_SUKEIRE_REC.SHIMUKE_CODE, vbUnicode)) Then
                        Exit Do
                    End If
                End If
                '受入年月日のﾌﾞﾚｰｸ
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化指示受入履歴")
                Exit Function
        End Select
        
        SKIP_Flg = False
        
        
        If StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) < Format(CDate(Text1(ptxS_YMD).Text), "YYYYMMDD") Or _
            StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) > Format(CDate(Text1(ptxE_YMD).Text), "YYYYMMDD") Then
            SKIP_Flg = True
        End If
        
        
        If Not SKIP_Flg Then
        
            '指示ﾃﾞｰﾀ読み込み
            Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    SKIP_Flg = True
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "資材注文ﾃﾞｰﾀ")
                    Exit Function
            End Select
                
                
            If Not SKIP_Flg Then
                
                Data_Flg = True
                '生産実績集計ﾃﾞｰﾀ読み込み
                    
                If Trim(Text1(ptxSHIMUKE_CODE).Text) = "" Then
                    Call UniCode_Conv(K0_P_SEISAN_SUM.SHIMUKE_CODE, "")
                Else
                    Call UniCode_Conv(K0_P_SEISAN_SUM.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                End If
                Call UniCode_Conv(K0_P_SEISAN_SUM.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                
                sts = BTRV(BtOpGetEqual, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
                Select Case sts
                    Case BtNoErr
                        upd_com = BtOpUpdate
                    Case BtErrKeyNotFound
                        upd_com = BtOpInsert
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "生産実績集計ﾃﾞｰﾀ")
                        Exit Function
                End Select
                
                
                If upd_com = BtOpInsert Then
                
                
                    'ｸﾗｽﾏｽﾀ読み込み
                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            SKIP_Flg = True
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ｸﾗｽﾏｽﾀ")
                            Exit Function
                    End Select
                
                    If Trim(Text1(ptxSHIMUKE_CODE).Text) = "" Then
                        Call UniCode_Conv(P_SEISAN_SUM_REC.SHIMUKE_CODE, "")
                    Else
                        Call UniCode_Conv(P_SEISAN_SUM_REC.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                    End If
                    
                    Call UniCode_Conv(P_SEISAN_SUM_REC.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                
                
                    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_NAI_CNT, "00000")
                    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_NAI_SURYO, "00000000.00")
                    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GAI_CNT, "00000")
                    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GAI_SURYO, "00000000.00")
                
                
                
                    For i = 0 To 2
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(i).NAI_TANKA, "00000000.00")
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(i).GAI_TANKA, "00000000.00")
                    Next i
                
                                
                    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_TANKA, StrConv(P_CLASSREC.TANKA, vbUnicode))
                
                
                
                
                    Call UniCode_Conv(P_SEISAN_SUM_REC.KO_GENKA, "00000000000")
                    Call UniCode_Conv(P_SEISAN_SUM_REC.GA_GENKA, "00000000000")
                    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GENKA, "00000000000")
                
                
                End If
                
                
                Select Case StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode)
                    Case P_TORI_SYANAI
                        '内部生産　生産件数
                        Call UniCode_Conv(P_SEISAN_SUM_REC.GK_NAI_CNT, Format(CInt(StrConv(P_SEISAN_SUM_REC.GK_NAI_CNT, vbUnicode)) + 1, "00000"))
                        '内部生産　生産数量
                        Call UniCode_Conv(P_SEISAN_SUM_REC.GK_NAI_SURYO, Format(CDbl(StrConv(P_SEISAN_SUM_REC.GK_NAI_SURYO, vbUnicode)) + _
                                                                            CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "00000000.00"))
                        '工料
                        
Debug.Print StrConv(P_SEISAN_SUM_REC.GK_NAI_CNT, vbUnicode)
Debug.Print StrConv(P_SEISAN_SUM_REC.GK_NAI_SURYO, vbUnicode)
                        
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(1).NAI_TANKA, StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                        '内部生産　資材内訳
                        Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
                        Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
                        Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
                        
                        Shizai_com = BtOpGetGreater
                        
                        
                        Do
                            sts = BTRV(Shizai_com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                            Select Case sts
                                Case BtNoErr
                                    If StrConv(P_SSHIJI_K_REC.SHIJI_NO, vbUnicode) <> StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode) Then
                                        Exit Do
                                    End If
                                Case BtErrEOF
                                    Exit Do
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "商品化指示(子)ﾃﾞｰﾀ")
                                    Exit Function
                            End Select
                        
                        
                            Select Case StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode)
                                Case P_KOSOU    '個装資材
                                
                                    If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then
                                
                                        wkSuryo = CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
                                
                                
                                        '品目マスタ読み込み
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                                                                                                        
                                
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                                            
                                                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                                Exit Function
                                        End Select
                                
                                        '単価（売）
                                        wkValue = CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).NAI_TANKA, vbUnicode))
                                        wkValue = wkValue + (CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) * wkSuryo)
                                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).NAI_TANKA, _
                                                                            Format(wkValue, "00000000.00"))
                                
                                        '単価（仕）
                                        wkValue = CDbl(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode))
                                        wkValue = wkValue + (CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) * wkSuryo)
                                        Call UniCode_Conv(P_SEISAN_SUM_REC.KO_GENKA, Format(wkValue, "00000000.00"))
                                
                                    End If
                                
                                
                                Case P_GAISOU   '外装資材
                                
                                    If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then
                                                                    
                                        '数量
                                
'                                        wkSuryo = CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
'                                        wkSuryo = CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode))
                                
                                    
                                        wkSuryo = 1 = Int(CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)) / CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)))
                                
                                
                                
                                        '品目マスタ読み込み
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                                                                                                        
                                
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                                            
                                                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                                Exit Function
                                        End Select
                                
                                        '単価（売）
                                        wkValue = CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).NAI_TANKA, vbUnicode))
                                        wkValue = wkValue + (CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) * wkSuryo)
                                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).NAI_TANKA, _
                                                                            Format(wkValue, "00000000.00"))
                                
                                        '単価（仕）
                                        wkValue = CDbl(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode))
                                        wkValue = wkValue + (CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) * wkSuryo)
                                        Call UniCode_Conv(P_SEISAN_SUM_REC.GA_GENKA, Format(wkValue, "00000000.00"))
                                
                                    End If
                                
                                
                                
                                Case P_DOUKON   '同梱・構成
                            
                            
                                    If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then
                                                                    
                                        '数量
                                
                                        wkSuryo = CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
                                
                                
                                        '品目マスタ読み込み
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                                                                                                        
                                
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                                            
                                                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                                Exit Function
                                        End Select
                                
                                        '単価（売）
                                        wkValue = CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).NAI_TANKA, vbUnicode))
                                        wkValue = wkValue + (CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) * wkSuryo)
                                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).NAI_TANKA, _
                                                                            Format(wkValue, "00000000.00"))
                                
                                    End If
                            End Select
                        
                            Shizai_com = BtOpGetNext
                        
                        
                        Loop
                    
                    
                    
                    
                    
                    
                                        
                    
                    
                    
                    
                    
                    
                    Case P_TORI_GENERAL, P_TORI_NAISYOKU, P_TORI_GENKIN
                        '外部生産　生産件数
                        Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GAI_CNT, Format(CInt(StrConv(P_SEISAN_SUM_REC.GK_GAI_CNT, vbUnicode)) + 1, "00000"))
                        '外部生産　生産数量
                        Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GAI_SURYO, Format(CDbl(StrConv(P_SEISAN_SUM_REC.GK_GAI_SURYO, vbUnicode)) + _
                                                                            CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "00000000.00"))
                
                        '工料
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(1).GAI_TANKA, StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                        '外注工料
                        Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GENKA, StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                        
                        '外部生産　資材内訳
                        Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
                        Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
                        Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
                        
                        Shizai_com = BtOpGetGreater
                        
                        
                        Do
                            sts = BTRV(Shizai_com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                            Select Case sts
                                Case BtNoErr
                                    If StrConv(P_SSHIJI_K_REC.SHIJI_NO, vbUnicode) <> StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode) Then
                                        Exit Do
                                    End If
                                Case BtErrEOF
                                    Exit Do
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "商品化指示(子)ﾃﾞｰﾀ")
                                    Exit Function
                            End Select
                        
                        
                            Select Case StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode)
                                Case P_KOSOU    '個装資材
                                
                                    If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then
                                                                    
                                        '数量
                                        wkSuryo = CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
                                
                                        '品目マスタ読み込み
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                                                                                                        
                                
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                                            
                                                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                                Exit Function
                                        End Select
                                
                                        '単価（売）
                                        wkValue = CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).GAI_TANKA, vbUnicode))
                                        wkValue = wkValue + CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) * wkSuryo
                                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).GAI_TANKA, _
                                                                            Format(wkValue, "00000000.00"))
                                
                                        '単価（仕）
                                        wkValue = CDbl(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode))
                                        wkValue = wkValue + CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) * wkSuryo
                                        Call UniCode_Conv(P_SEISAN_SUM_REC.KO_GENKA, Format(wkValue, "00000000.00"))
                                
                                    End If
                                
                                
                                Case P_GAISOU   '外装資材
                                
                                    If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then
                                                                    
                                        '数量
                                
                                        wkSuryo = CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
                                
                                
                                        '品目マスタ読み込み
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                                                                                                        
                                
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                                            
                                                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                                Exit Function
                                        End Select
                                
                                        '単価（売）
                                        wkValue = CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).GAI_TANKA, vbUnicode))
                                        wkValue = wkValue + CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) * wkSuryo
                                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).GAI_TANKA, _
                                                                            Format(wkValue, "00000000.00"))
                                
                                        '単価（仕）
                                        wkValue = CDbl(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode))
                                        wkValue = wkValue + CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) * wkSuryo
                                        Call UniCode_Conv(P_SEISAN_SUM_REC.GA_GENKA, Format(wkValue, "00000000.00"))
                                
                                    
                                    End If
                                
                                
                                
                                Case P_DOUKON   '同梱・構成
                            
                            
                                    If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then
                                                                    
                                        '数量
                                
                                        wkSuryo = CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
                                
                                
                                        '品目マスタ読み込み
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                                                                                                        
                                
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                                            
                                                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                                Exit Function
                                        End Select
                                
                                        '単価（売）
                                        wkValue = CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).GAI_TANKA, vbUnicode))
                                        wkValue = wkValue + CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) * wkSuryo
                                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(0).NAI_TANKA, _
                                                                            Format(wkValue, "00000000.00"))
                                
                                    End If
                            
                            
                            End Select
                        
                            Shizai_com = BtOpGetNext
                        Loop
                
                
                End Select
            
                sts = BTRV(upd_com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, upd_com, "生産実績集計ﾃﾞｰﾀ")
                        Exit Function
                End Select
    
            End If
        End If
        
        com = BtOpGetNext
    
    Loop


    '-----------------------------------------  合計集計及び合計ﾚｺｰﾄﾞ作成
    GK_NAI_CNT = 0
    GK_NAI_SURYO = 0
    GK_GAI_CNT = 0
    GK_GAI_SURYO = 0
    
    For i = 0 To 2
        NAI_TANKA(i) = 0
        GAI_TANKA(i) = 0
    Next i
    
    
    KO_GENKA = 0
    GA_GENKA = 0
    GK_GENKA = 0
    
    
    
    
    
    
    
    com = BtOpGetFirst



    Do
        
        DoEvents
        
        sts = BTRV(com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "生産実績集計ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        
        GK_NAI_CNT = GK_NAI_CNT + CInt(StrConv(P_SEISAN_SUM_REC.GK_NAI_CNT, vbUnicode))
        GK_NAI_SURYO = GK_NAI_SURYO + CDbl(StrConv(P_SEISAN_SUM_REC.GK_NAI_SURYO, vbUnicode))
        GK_GAI_CNT = GK_GAI_CNT + CInt(StrConv(P_SEISAN_SUM_REC.GK_GAI_CNT, vbUnicode))
        GK_GAI_SURYO = GK_GAI_SURYO + CDbl(StrConv(P_SEISAN_SUM_REC.GK_GAI_SURYO, vbUnicode))
    
        For i = 0 To 2
        
            NAI_TANKA(i) = NAI_TANKA(i) + CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(i).NAI_TANKA, vbUnicode))
            GAI_TANKA(i) = GAI_TANKA(i) + CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(i).GAI_TANKA, vbUnicode))
        
        
        Next i
    
    
        KO_GENKA = KO_GENKA + CDbl(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode))
        GA_GENKA = GA_GENKA + CDbl(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode))
        GK_GENKA = GK_GENKA + CDbl(StrConv(P_SEISAN_SUM_REC.GK_GENKA, vbUnicode))
    
    
    
        com = BtOpGetNext
    
    Loop


    If Sum_Total_Make_Proc() Then
        Exit Function
    End If
    

    PR000501.MousePointer = vbDefault

   SUM_Make_Proc = False

End Function






Private Function Sum_Total_Make_Proc() As Integer
'----------------------------------------------------------------------------
'           合計ﾚｺｰﾄﾞ出力
'----------------------------------------------------------------------------
Dim i   As Integer
Dim sts As Integer
    
    Sum_Total_Make_Proc = True

    
    If Trim(Text1(ptxSHIMUKE_CODE).Text) = "" Then
        Call UniCode_Conv(P_SEISAN_SUM_REC.SHIMUKE_CODE, "")
    Else
        Call UniCode_Conv(P_SEISAN_SUM_REC.SHIMUKE_CODE, Trim(Text1(ptxSHIMUKE_CODE).Text))
    End If
    
    Call UniCode_Conv(P_SEISAN_SUM_REC.CLASS_CODE, P_ClassSum_Key)

    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_NAI_CNT, Format(GK_NAI_CNT, "00000"))
    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_NAI_SURYO, Format(GK_NAI_SURYO, "00000000.00"))
    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GAI_CNT, Format(GK_GAI_CNT, "00000"))
    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GAI_SURYO, Format(GK_GAI_SURYO, "00000000.00"))


    For i = 0 To 2
        
        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE_TBL(i).NAI_TANKA, Format(NAI_TANKA(i), "00000000.00"))
        
    Next i


    Call UniCode_Conv(P_SEISAN_SUM_REC.KO_GENKA, Format(KO_GENKA, "00000000.00"))
    Call UniCode_Conv(P_SEISAN_SUM_REC.GA_GENKA, Format(GA_GENKA, "00000000.00"))
    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GENKA, Format(GK_GENKA, "00000000.00"))


    sts = BTRV(BtOpInsert, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpInsert, "生産実績集計ﾃﾞｰﾀ")
            Exit Function
    End Select

    Sum_Total_Make_Proc = False

End Function


