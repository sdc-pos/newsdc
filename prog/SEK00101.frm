VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEK00101 
   Caption         =   "積水ハウス注文ﾃﾞｰﾀ　変換"
   ClientHeight    =   10848
   ClientLeft      =   2028
   ClientTop       =   -3216
   ClientWidth     =   15264
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
   ScaleHeight     =   10848
   ScaleWidth      =   15264
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Left            =   2040
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   5
      Top             =   600
      Width           =   8772
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8892
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   14652
      _ExtentX        =   25845
      _ExtentY        =   15685
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "処理結果"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "　データ　　　作成日"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "　データ　　　作成時刻"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "連番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "受注日"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "　　納入　　受入場所"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "　　納入　　　　　受入場所名"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "得意先　　　コード"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "直納先　　　コード"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "得意先品番　　■品番（上）"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "品番　　　　　■品番（下）"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "注文№　　　　　　■指図№（上）"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "出荷順番　　　　　　■指図№（下・左）"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "邸名　　　　　　　■指図№（下・右）"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "受注数量"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "出荷確定日"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "納入日"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "件管№　　　　　　■管理№（上）"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "品管№　　　　　■管理№（下）"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "単品区分"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "邸別ﾗﾍﾞﾙID"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "箱№"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   22
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   720
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=22"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2180"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2074"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2392"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2286"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2392"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2286"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1376"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1270"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1588"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1482"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2053"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1947"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2836"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2731"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=1947"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1842"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=1947"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1842"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2582"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2477"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=2582"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2477"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=3281"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=3556"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3450"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(53)=   "Column(13).Width=3344"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=3239"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=2752"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(60)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(61)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(62)=   "Column(15).Width=2519"
      Splits(0)._ColumnProps(63)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(15)._WidthInPix=2413"
      Splits(0)._ColumnProps(65)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(66)=   "Column(16).Width=2053"
      Splits(0)._ColumnProps(67)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(16)._WidthInPix=1947"
      Splits(0)._ColumnProps(69)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(70)=   "Column(17).Width=3260"
      Splits(0)._ColumnProps(71)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(17)._WidthInPix=3154"
      Splits(0)._ColumnProps(73)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(74)=   "Column(18).Width=2858"
      Splits(0)._ColumnProps(75)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(18)._WidthInPix=2752"
      Splits(0)._ColumnProps(77)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(78)=   "Column(19).Width=1291"
      Splits(0)._ColumnProps(79)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(19)._WidthInPix=1185"
      Splits(0)._ColumnProps(81)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(82)=   "Column(20).Width=3493"
      Splits(0)._ColumnProps(83)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(84)=   "Column(20)._WidthInPix=3387"
      Splits(0)._ColumnProps(85)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(86)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(87)=   "Column(21).Width=3493"
      Splits(0)._ColumnProps(88)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(89)=   "Column(21)._WidthInPix=3387"
      Splits(0)._ColumnProps(90)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(91)=   "Column(21).Order=22"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(27)  =   ":id=14,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(28)  =   ":id=14,.fontname=ＭＳ ゴシック"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=110,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=107,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=108,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=109,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=3"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=3"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=78,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=82,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=86,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=90,.parent=13,.alignment=1"
      _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=94,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=98,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=17"
      _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=102,.parent=13"
      _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=99,.parent=14"
      _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=100,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=101,.parent=17"
      _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=106,.parent=13"
      _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=103,.parent=14"
      _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=104,.parent=15"
      _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=105,.parent=17"
      _StyleDefs(114) =   "Splits(0).Columns(19).Style:id=114,.parent=13"
      _StyleDefs(115) =   "Splits(0).Columns(19).HeadingStyle:id=111,.parent=14"
      _StyleDefs(116) =   "Splits(0).Columns(19).FooterStyle:id=112,.parent=15"
      _StyleDefs(117) =   "Splits(0).Columns(19).EditorStyle:id=113,.parent=17"
      _StyleDefs(118) =   "Splits(0).Columns(20).Style:id=118,.parent=13"
      _StyleDefs(119) =   "Splits(0).Columns(20).HeadingStyle:id=115,.parent=14"
      _StyleDefs(120) =   "Splits(0).Columns(20).FooterStyle:id=116,.parent=15"
      _StyleDefs(121) =   "Splits(0).Columns(20).EditorStyle:id=117,.parent=17"
      _StyleDefs(122) =   "Splits(0).Columns(21).Style:id=122,.parent=13"
      _StyleDefs(123) =   "Splits(0).Columns(21).HeadingStyle:id=119,.parent=14"
      _StyleDefs(124) =   "Splits(0).Columns(21).FooterStyle:id=120,.parent=15"
      _StyleDefs(125) =   "Splits(0).Columns(21).EditorStyle:id=121,.parent=17"
      _StyleDefs(126) =   "Named:id=33:Normal"
      _StyleDefs(127) =   ":id=33,.parent=0"
      _StyleDefs(128) =   "Named:id=34:Heading"
      _StyleDefs(129) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(130) =   ":id=34,.wraptext=-1"
      _StyleDefs(131) =   "Named:id=35:Footing"
      _StyleDefs(132) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(133) =   "Named:id=36:Selected"
      _StyleDefs(134) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(135) =   "Named:id=37:Caption"
      _StyleDefs(136) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(137) =   "Named:id=38:HighlightRow"
      _StyleDefs(138) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(139) =   "Named:id=39:EvenRow"
      _StyleDefs(140) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(141) =   "Named:id=40:OddRow"
      _StyleDefs(142) =   ":id=40,.parent=33"
      _StyleDefs(143) =   "Named:id=41:RecordSelector"
      _StyleDefs(144) =   ":id=41,.parent=34"
      _StyleDefs(145) =   "Named:id=42:FilterBar"
      _StyleDefs(146) =   ":id=42,.parent=33"
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
      Index           =   2
      Left            =   3450
      TabIndex        =   2
      ToolTipText     =   "商品化構成を保存します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登 録"
      Enabled         =   0   'False
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
      Left            =   1845
      TabIndex        =   1
      ToolTipText     =   "商品化構成を読み込みます（Ｆ5）"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "読 込"
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
      Left            =   420
      TabIndex        =   0
      Top             =   0
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   13440
      TabIndex        =   7
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "件数"
      Height          =   252
      Index           =   1
      Left            =   12840
      TabIndex        =   6
      Top             =   720
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "ファイル名"
      Height          =   252
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   1332
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
Attribute VB_Name = "SEK00101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Y_Syuka_TEI     As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 21             '最大列数

Private Const colSHORI% = 0             '処理結果
Private Const colSND_YMD% = 1           'データ作成日
Private Const colSND_HMS% = 2           'データ作成時刻
Private Const colSEQ_NO% = 3            '連番
Private Const colJUC_YMD% = 4           '受注日
Private Const colNOU_CD% = 5            '納入受入場
Private Const colNOU_NM% = 6            '納入受入場名
Private Const colTOK_CD% = 7            '得意先ｺｰﾄﾞ
Private Const colCHO_CD% = 8            '直納先ｺｰﾄﾞ
Private Const colTHINB_CD% = 9          '得意先品番　■品番(上)
Private Const colHINB_CD% = 10          '品番　      ■品番(下)
Private Const colCHU_CD% = 11           '注文№　    ■指図№(上)
Private Const colSYU_JUN% = 12          '出荷順番　  ■指図№(下・左)
Private Const colTEI_NM% = 13           '邸名　      ■指図№(下・右)
Private Const colJUC_SUU% = 14          '受注数量
Private Const colSYU_YMD% = 15          '出荷確定日
Private Const colNOU_YMD% = 16          '納入日
Private Const colKEN_NO% = 17           '件管№　　　■管理№(上)
Private Const colHIN_NO% = 18           '件管№　　　■管理№(下)
Private Const colTANP_KB% = 19          '単品区分

Private Const colTEI_LABELID% = 20      '邸別ﾗﾍﾞﾙID
Private Const colHAKO_NO% = 21          '箱№





Private Const LAST_UPDATE_DAY$ = "[SEK0010] 2011.05.24 09:00"

Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '読込み

            'ﾌｧｲﾙ名ﾁｪｯｸ
''            sWk = Text1.Text
''            For i = Len(sWk) To 1 Step -1
''                If Mid(sWk, i, 1) = "\" Then
''                    Exit For
''                End If
''            Next i
''
''            i = i + 1
''
''            If Mid(sWk, i, 9) <> "C_22T826_" Then
''                MsgBox "エラー" & vbCrLf & vbCrLf & "ファイル名が違います。", vbExclamation
''                Exit Sub
''            End If
''
''            j = InStr(i, sWk, ".")
''            If StrConv(Right(sWk, Len(sWk) - j), vbLowerCase) <> "txt" Then
''                MsgBox "エラー" & vbCrLf & vbCrLf & "拡張子が「txt」以外です。", vbExclamation
''                Exit Sub
''            End If


            '取込みﾃﾞｰﾀ表示
            If List_Disp_Proc() Then
                Unload Me
            End If


            If Y_Syuka_TEI.Count(1) > 0 Then
                Command1(1).Enabled = True
            End If


        Case 1          '登録





            If Update_Proc() Then
                Unload Me
            End If



        Case 2          '終了

            Unload Me
    End Select



    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    
    
    If Shift = vbAltMask Then
        
        If TDBGrid1.AllowUpdate Then
        
            TDBGrid1.AllowUpdate = False
            TDBGrid1.AllowAddNew = False
            TDBGrid1.AllowDelete = False
    
    
            TDBGrid1.Columns(colTEI_LABELID).Visible = False
            TDBGrid1.Columns(colHAKO_NO).Visible = False
    
            TDBGrid1.Columns(colTEI_LABELID).Locked = True
            TDBGrid1.Columns(colHAKO_NO).Locked = True
    
    
    
        Else
    
    
            TDBGrid1.AllowUpdate = True
            TDBGrid1.AllowAddNew = True
            TDBGrid1.AllowDelete = True
    
    
            TDBGrid1.Columns(colTEI_LABELID).Visible = True
            TDBGrid1.Columns(colHAKO_NO).Visible = True
    
        
            TDBGrid1.Columns(colTEI_LABELID).Locked = False
            TDBGrid1.Columns(colHAKO_NO).Locked = False
        
        End If
    
    End If
    
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128



    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "積水ハウス注文ﾃﾞｰﾀ　変換", Me.hwnd, 0)
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

    SEK00101.Caption = SEK00101.Caption & " " & LAST_UPDATE_DAY

                                '邸別注文ﾃﾞｰﾀＯＰＥＮ
    If Y_SYU_TEI_Open(BtOpenNomal) Then
        Unload Me
    End If





End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Text = Trim(Data.Files(1))

'    Text1.Text = Data.GetData(vbCFText)

    Command1(0).Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "邸別注文ﾃﾞｰﾀ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set SEK00101 = Nothing



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

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   「注文データ」登録処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
    
Dim Upd_Com         As Integer
Dim Skip_Flg        As Integer
    
Dim INS_NOW         As String * 14
    
    


Dim Row             As Long

    If Y_Syuka_TEI.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "注文データ登録処理　処理開始！！", Me.hwnd, 0)

                                    
                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
                                    
                                    'テーブルリセット
    
    Skip_Flg = True
    For Row = 1 To Y_Syuka_TEI.UpperBound(1)
        
        
        DoEvents
        
        
        Call UniCode_Conv(K0_Y_SYU_TEI.SND_YMD, Y_Syuka_TEI(Row, colSND_YMD))
        Call UniCode_Conv(K0_Y_SYU_TEI.SND_HMS, Y_Syuka_TEI(Row, colSND_HMS))
        Call UniCode_Conv(K0_Y_SYU_TEI.SEQ_NO, Y_Syuka_TEI(Row, colSEQ_NO))
        
        
        sts = BTRV(BtOpGetEqual, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
        Select Case sts
            Case BtNoErr
            
                If Skip_Flg Then
         
                    Beep
                    ans = MsgBox("「邸別注文データ取り込み済みです。」" & Chr(13) & Chr(10) & _
                    "作成日:" & Y_Syuka_TEI(Row, colSND_YMD) & Chr(13) & Chr(10) & _
                    "作成時刻:" & Y_Syuka_TEI(Row, colSND_HMS) & Chr(13) & Chr(10) & _
                    "総件数:" & Format(Val(Label2.Caption), "#0") & Chr(13) & Chr(10) & _
                    "再取り込みしますか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")
                    If ans = vbNo Then
                    Call Input_UnLock
                        Update_Proc = False
                        Exit Function
                    Else
                        Skip_Flg = False
                        Upd_Com = BtOpUpdate
                    End If
                
                End If
            Case BtErrKeyNotFound
                Upd_Com = BtOpInsert
                Skip_Flg = False
            
            Case Else
                Call File_Error(sts, BtOpInsert, "邸別注文データ")
                Call Input_UnLock
                Exit Function
        End Select
        
        
        If Not Skip_Flg Then
        
        
            Call UniCode_Conv(Y_SYU_TEI_REC.SND_YMD, Y_Syuka_TEI(Row, colSND_YMD))          'データ作成日
            Call UniCode_Conv(Y_SYU_TEI_REC.SND_HMS, Y_Syuka_TEI(Row, colSND_HMS))          'データ作成時刻
            Call UniCode_Conv(Y_SYU_TEI_REC.SEQ_NO, Y_Syuka_TEI(Row, colSEQ_NO))            '連番
            Call UniCode_Conv(Y_SYU_TEI_REC.JUC_YMD, Y_Syuka_TEI(Row, colJUC_YMD))          '受注日
            Call UniCode_Conv(Y_SYU_TEI_REC.NOU_CD, Y_Syuka_TEI(Row, colNOU_CD))            '納入受入場
            Call UniCode_Conv(Y_SYU_TEI_REC.NOU_NM, Y_Syuka_TEI(Row, colNOU_NM))            '納入受入場名
            Call UniCode_Conv(Y_SYU_TEI_REC.TOK_CD, Y_Syuka_TEI(Row, colTOK_CD))            '得意先ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYU_TEI_REC.CHO_CD, Y_Syuka_TEI(Row, colCHO_CD))            '直納先ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYU_TEI_REC.THINB_CD, Y_Syuka_TEI(Row, colTHINB_CD))        '得意先品番　■品番(上)
            Call UniCode_Conv(Y_SYU_TEI_REC.HINB_CD, Y_Syuka_TEI(Row, colHINB_CD))          '品番　■品番(下)
            Call UniCode_Conv(Y_SYU_TEI_REC.CHU_CD, Y_Syuka_TEI(Row, colCHU_CD))            '注文№　    ■指図№(上)
            Call UniCode_Conv(Y_SYU_TEI_REC.SYU_JUN, Y_Syuka_TEI(Row, colSYU_JUN))          '出荷順番　  ■指図№(下・左)
            Call UniCode_Conv(Y_SYU_TEI_REC.TEI_NM, Y_Syuka_TEI(Row, colTEI_NM))            '邸名　      ■指図№(下・右)
                                                                                            '受注数量
            Call UniCode_Conv(Y_SYU_TEI_REC.JUC_SUU, Format(Val(Y_Syuka_TEI(Row, colJUC_SUU)), "00000000"))
            Call UniCode_Conv(Y_SYU_TEI_REC.SYU_YMD, Y_Syuka_TEI(Row, colSYU_YMD))          '出荷確定日
            Call UniCode_Conv(Y_SYU_TEI_REC.NOU_YMD, Y_Syuka_TEI(Row, colNOU_YMD))          '納入日
            Call UniCode_Conv(Y_SYU_TEI_REC.KEN_NO, Y_Syuka_TEI(Row, colKEN_NO))            '件管№　　　■管理№(上)
            Call UniCode_Conv(Y_SYU_TEI_REC.HIN_NO, Y_Syuka_TEI(Row, colHIN_NO))            '件管№　　　■管理№(下)
            Call UniCode_Conv(Y_SYU_TEI_REC.TANP_KB, Y_Syuka_TEI(Row, colTANP_KB))          '単品区分
            Call UniCode_Conv(Y_SYU_TEI_REC.YOBI1_NM, "")                                   '予備
            Call UniCode_Conv(Y_SYU_TEI_REC.GSEQ_NO, Format(Val(Label2.Caption), "00000"))  '総件数
                    
            
            
            
            
            If TDBGrid1.Columns(colTEI_LABELID).Visible Then                                '邸別ﾗﾍﾞﾙID(注文№■指図№(上)+箱№)
               Call UniCode_Conv(Y_SYU_TEI_REC.TEI_LABELID, Y_Syuka_TEI(Row, colTEI_LABELID))
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.TEI_LABELID, "")
            End If
            
            
            If TDBGrid1.Columns(colHAKO_NO).Visible Then                                    '箱№
                Call UniCode_Conv(Y_SYU_TEI_REC.HAKO_NO, Y_Syuka_TEI(Row, colHAKO_NO))
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.HAKO_NO, "")
            End If
            
            Call UniCode_Conv(Y_SYU_TEI_REC.JITU_SUU, "")                                   '実出庫数(梱包場への出庫数 現在未使用)
            Call UniCode_Conv(Y_SYU_TEI_REC.JITU_TANTO, "")                                 '出庫　担当者(現在未使用)
            Call UniCode_Conv(Y_SYU_TEI_REC.JITU_DATETIME, "")                              '出庫　日時(現在未使用)
            Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_TANTO, "")                                '梱包　担当者
            Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_DATETIME, "")                             '梱包　日時
            
            Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_TANTO, "")                                '照合　担当者
            Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_DATETIME, "")                             '照合　日時
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.L_KENKAN, "")                                   '件管末番 long
            Call UniCode_Conv(Y_SYU_TEI_REC.L_TEI_NAME, "")                                 '邸名2
            Call UniCode_Conv(Y_SYU_TEI_REC.L_TOK_NAME, "")                                 '得意先名
            Call UniCode_Conv(Y_SYU_TEI_REC.L_SOTO_NO, "")                                  '外箱番号
            Call UniCode_Conv(Y_SYU_TEI_REC.L_UCHI_NO, "")                                  '内箱番号
            Call UniCode_Conv(Y_SYU_TEI_REC.L_WIDTH, "")                                    '長さ(幅)
            Call UniCode_Conv(Y_SYU_TEI_REC.L_HEIGHT, "")                                   '高さ
            Call UniCode_Conv(Y_SYU_TEI_REC.L_CONTENT, "")                                  '体積
            Call UniCode_Conv(Y_SYU_TEI_REC.L_KNo, "")                                      '工場No 2
            Call UniCode_Conv(Y_SYU_TEI_REC.L_SERIES1, "")                                  '品番シリーズ
            Call UniCode_Conv(Y_SYU_TEI_REC.L_SERIES2, "")                                  '品番シリーズ2
            Call UniCode_Conv(Y_SYU_TEI_REC.L_PAGE, "")                                     'ページ番号
            
            Call UniCode_Conv(Y_SYU_TEI_REC.KUTI_SU, "0000")                                '口数
            Call UniCode_Conv(Y_SYU_TEI_REC.SAI_SU, "000.00")                               '才数
            
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_ID, "")                                   '梱包ID
    
    
            Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_TANTO, "")                               '検品担当者
            Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_DATETIME, "")                            '検品日時
    
    
            Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, "")                          '集合梱包担当者
            Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, "")                       '集合梱包日時
            
            
            
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.FILLER, "")                                     'FILLER
            Call UniCode_Conv(Y_SYU_TEI_REC.INS_TANTO, StrConv(App.EXEName, vbUpperCase))   '追加担当者
            Call UniCode_Conv(Y_SYU_TEI_REC.Ins_DateTime, INS_NOW)                          '追加担当者
            
            
Debug.Print StrConv(Y_SYU_TEI_REC.Ins_DateTime, vbUnicode)
            
            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, "")                                  '更新担当者
            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_DATETIME, "")                               '更新担当者
                    
                    
            Do
                sts = BTRV(Upd_Com, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        
                        Beep
                        ans = MsgBox("「邸別注文データ」他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Call Input_UnLock
                            Exit Function
                        End If
                    
                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, Upd_Com, "邸別注文データ")
                        Exit Function
                End Select
            
            Loop
            
            Y_Syuka_TEI(Row, colSHORI) = "済"                           'データ出力
    
            Set TDBGrid1.Array = Y_Syuka_TEI
            TDBGrid1.ReBind
            
            TDBGrid1.Update
            TDBGrid1.Bookmark = Row
        
        End If

    Next Row


    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "注文データ登録処理　処理終了！！", Me.hwnd, 0)





    Update_Proc = False
    Call Input_UnLock
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description
    Call Input_UnLock

End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「注文データ」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
Dim fileName        As String
Dim FileNo          As Long
    
Dim wkText          As String
Dim Length          As Integer


Dim SND_YMD         As String * 8           'データ作成日
Dim SND_HMS         As String * 6           'データ作成時刻
Dim SEQ_NO          As String * 5           '連番
Dim JUC_YMD         As String * 8           '受注日
Dim NOU_CD          As String * 4           '納入受入場
Dim NOU_NM          As String * 20          '納入受入場名
Dim TOK_CD          As String * 8           '得意先ｺｰﾄﾞ
Dim CHO_CD          As String * 8           '直納先ｺｰﾄﾞ
Dim THINB_CD        As String * 20          '得意先品番　■品番(上)
Dim HINB_CD         As String * 20          '品番　      ■品番(下)
Dim CHU_CD          As String * 10          '注文№　    ■指図№(上)
Dim SYU_JUN         As String * 10          '出荷順番　  ■指図№(下・左)
Dim TEI_NM          As String * 30          '邸名　      ■指図№(下・右)
Dim JUC_SUU         As String * 8           '受注数量
Dim SYU_YMD         As String * 8           '出荷確定日
Dim NOU_YMD         As String * 8           '納入日
Dim KEN_NO          As String * 6           '件管№　　　■管理№(上)
Dim HIN_NO          As String * 6           '件管№　　　■管理№(下)
Dim TANP_KB         As String * 1           '単品区分
Dim YOBI1_NM        As String * 55          '予備
Dim GSEQ_NO         As String * 5           'ﾃﾞｰﾀ総件数


Dim Row             As Long


    List_Disp_Proc = True

    Call Input_Lock

    FileNo = FreeFile
    fileName = Trim(Text1.Text)
    On Error GoTo Error_Proc

    Open fileName For Input As #FileNo

    On Error GoTo 0

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "注文データ読込み処理　処理開始！！", Me.hwnd, 0)

                                    'テーブルリセット
    Set Y_Syuka_TEI = Nothing
    Row = Min_Row - 1
    Label2.Caption = ""

    Do Until EOF(FileNo)
        Line Input #FileNo, wkText
    
        If LenB(StrConv(wkText, vbFromUnicode)) <> 254 Then
            
            
            Exit Do
        End If
    
    
        DoEvents
        
        Length = 1                                                  'データ作成日
        SND_YMD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SND_YMD)), vbUnicode)
                                                                    
                                                                    
        Length = Length + Len(SND_YMD)                              'データ作成時刻
        SND_HMS = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SND_HMS)), vbUnicode)

        Length = Length + Len(SND_HMS)                              '連番
        SEQ_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SEQ_NO)), vbUnicode)

        Length = Length + Len(SEQ_NO)                              '受注日
        JUC_YMD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(JUC_YMD)), vbUnicode)

        Length = Length + Len(JUC_YMD)                              '納入受入場
        NOU_CD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(NOU_CD)), vbUnicode)

        Length = Length + Len(NOU_CD)                               '納入受入場名
        NOU_NM = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(NOU_NM)), vbUnicode)

        Length = Length + Len(NOU_NM)                               '得意先ｺｰﾄﾞ
        TOK_CD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(TOK_CD)), vbUnicode)

        Length = Length + Len(TOK_CD)                               '直納先ｺｰﾄﾞ
        CHO_CD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(CHO_CD)), vbUnicode)

        Length = Length + Len(CHO_CD)                               '得意先品番　■品番(上)
        THINB_CD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(THINB_CD)), vbUnicode)

        Length = Length + Len(THINB_CD)                             '品番　■品番(下)
        HINB_CD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HINB_CD)), vbUnicode)

        Length = Length + Len(HINB_CD)                              '注文№　    ■指図№(上)
        CHU_CD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(CHU_CD)), vbUnicode)

        Length = Length + Len(CHU_CD)                               '出荷順序　  ■指図№(下・左)
        SYU_JUN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SYU_JUN)), vbUnicode)

        Length = Length + Len(SYU_JUN)                              '邸名　      ■指図№(下・右)
        TEI_NM = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(TEI_NM)), vbUnicode)

        Length = Length + Len(TEI_NM)                               '受注数量
        JUC_SUU = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(JUC_SUU)), vbUnicode)
        
        Length = Length + Len(JUC_SUU)                              '出荷確定日
        SYU_YMD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SYU_YMD)), vbUnicode)
        
        Length = Length + Len(SYU_YMD)                              '納入日
        NOU_YMD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(NOU_YMD)), vbUnicode)
        
        Length = Length + Len(NOU_YMD)                              '件管№　　　■管理№(上)
        KEN_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(KEN_NO)), vbUnicode)

        Length = Length + Len(KEN_NO)                               '件管№　　　■管理№(下)
        HIN_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NO)), vbUnicode)

        Length = Length + Len(HIN_NO)                               '単品区分
        TANP_KB = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(TANP_KB)), vbUnicode)

        Length = Length + Len(TANP_KB)                              '予備
        YOBI1_NM = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOBI1_NM)), vbUnicode)
        
        Length = Length + Len(YOBI1_NM)                             '総件数
        GSEQ_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(GSEQ_NO)), vbUnicode)




        Row = Row + 1
        Y_Syuka_TEI.ReDim Min_Row, Row, Min_Col, Max_Col
        
        Y_Syuka_TEI(Row, colSND_YMD) = SND_YMD                          'データ作成日
        Y_Syuka_TEI(Row, colSND_HMS) = SND_HMS                          'データ作成時刻
        Y_Syuka_TEI(Row, colSEQ_NO) = SEQ_NO                            '連番
        Y_Syuka_TEI(Row, colJUC_YMD) = JUC_YMD                          '受注日
        Y_Syuka_TEI(Row, colNOU_CD) = NOU_CD                            '納入受入場
        Y_Syuka_TEI(Row, colNOU_NM) = NOU_NM                            '納入受入場名
        Y_Syuka_TEI(Row, colTOK_CD) = TOK_CD                            '得意先ｺｰﾄﾞ
        Y_Syuka_TEI(Row, colCHO_CD) = CHO_CD                            '直納先ｺｰﾄﾞ
        Y_Syuka_TEI(Row, colTHINB_CD) = THINB_CD                        '得意先品番　■品番(上)
        Y_Syuka_TEI(Row, colHINB_CD) = HINB_CD                         '品番　■品番(下)
        Y_Syuka_TEI(Row, colCHU_CD) = CHU_CD                            '注文№　    ■指図№(上)
        Y_Syuka_TEI(Row, colSYU_JUN) = SYU_JUN                          '出荷順番　  ■指図№(下・左)
        Y_Syuka_TEI(Row, colTEI_NM) = TEI_NM                            '邸名　      ■指図№(下・右)
        Y_Syuka_TEI(Row, colJUC_SUU) = Format(Val(JUC_SUU), "#0")       '受注数量
        Y_Syuka_TEI(Row, colSYU_YMD) = SYU_YMD                          '出荷確定日
        Y_Syuka_TEI(Row, colNOU_YMD) = NOU_YMD                          '納入日
        Y_Syuka_TEI(Row, colKEN_NO) = KEN_NO                            '件管№　　　■管理№(上)
        Y_Syuka_TEI(Row, colHIN_NO) = HIN_NO                            '件管№　　　■管理№(下)
        Y_Syuka_TEI(Row, colTANP_KB) = TANP_KB                          '単品区分








        If Trim(Label2.Caption) = "" Then
            Label2.Caption = Format(Val(GSEQ_NO), "#0")
        End If












    Loop


    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "注文データ読込み処理　処理終了！！", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 ファイル名または番号が不正です。
        '53 ファイルが見つかりません。
        '54 ファイル モードが不正です。
        '55 ファイルは既に開かれています。
        '57 デバイス I/O エラーです。
        '59 レコード長が一致しません。
        '61 ディスクの空き容量が不足しています。
        '62 ファイルにこれ以上データがありません。
        '63 レコード番号が不正です。
        '68 デバイスが準備されていません。
        '70 書き込みできません。
        '71 ディスクが準備されていません。
        '75 パス名が無効です。
        '76 パスが見つかりません。
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "指定のファイルが見つかりません。" & Chr(13) & Chr(10) & "正しいファイル名を入力してください。"
            
            
            
            List_Disp_Proc = False      '





        Case Else
    End Select
    Call Input_UnLock

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    SEK00101.MousePointer = vbHourglass

    Call Ctrl_Lock(SEK00101)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(SEK00101)


    SEK00101.MousePointer = vbDefault

End Sub

Private Sub Text1_OLESetData(Data As DataObject, DataFormat As Integer)
'    If DataFormat = vbCFText Then
'        Data.SetData Text1.SelText, vbCFText
'    End If
End Sub
