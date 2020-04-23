VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PLN00301 
   Caption         =   "[商品化計画システム]商品化予定データ登録"
   ClientHeight    =   9795
   ClientLeft      =   2025
   ClientTop       =   -4470
   ClientWidth     =   15585
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
   ScaleHeight     =   9795
   ScaleWidth      =   15585
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "全表示"
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   13560
      TabIndex        =   9
      ToolTipText     =   "商品化構成を読み込みます（Ｆ5）"
      Top             =   0
      Width           =   1380
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
      Index           =   3
      Left            =   3480
      TabIndex        =   8
      ToolTipText     =   "処理を終了します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      Left            =   1680
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   7
      Top             =   600
      Width           =   6975
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8295
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   14631
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "削除"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ＢＵ"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "標準棚番"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "対外品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "商品化予定日"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "商品化　　予定数"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "在庫数　　（済）"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "在庫数　　（未）"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "月平均出荷数"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "事前 　       商品化状況(%)"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "事前　　　　商品化必要数"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "部品　　　入荷予定日"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "部品　　　入荷予定数"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "商品化予定日（元）"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "商品化予定数（元）"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "見積工数(分/個)"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "標準時間(元)"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "資材(箱№)"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "外装品番"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "外装使用枚数"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "別置１　倉庫"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "別置１　列"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "別置１　連"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "別置１　段"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "別置１　数量"
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "別置２　倉庫"
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "別置２　列"
      Columns(26).DataField=   ""
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "別置２　連"
      Columns(27).DataField=   ""
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "別置２　段"
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "別置２　数量"
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "実績工数"
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "作業工数"
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "国内供給区分"
      Columns(32).DataField=   ""
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "海外供給区分"
      Columns(33).DataField=   ""
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "商品化完了手配先"
      Columns(34).DataField=   ""
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).Caption=   "KEY_NO"
      Columns(35).DataField=   ""
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).Caption=   "入荷予定KEY_NO"
      Columns(36).DataField=   ""
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   37
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=37"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=503"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1085"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=979"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=2805"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2699"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=8193"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=3096"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2990"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8192"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=2328"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2223"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8193"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=2037"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1931"
      Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=8194"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=1667"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=1561"
      Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=8194"
      Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(35)=   "Column(7).Width=1746"
      Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=1640"
      Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=8194"
      Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(40)=   "Column(8).Width=2223"
      Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2117"
      Splits(0)._ColumnProps(43)=   "Column(8)._ColStyle=8194"
      Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(45)=   "Column(9).Width=2381"
      Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=2275"
      Splits(0)._ColumnProps(48)=   "Column(9)._ColStyle=8194"
      Splits(0)._ColumnProps(49)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(50)=   "Column(10).Width=2090"
      Splits(0)._ColumnProps(51)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(10)._WidthInPix=1984"
      Splits(0)._ColumnProps(53)=   "Column(10)._ColStyle=8194"
      Splits(0)._ColumnProps(54)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(55)=   "Column(11).Width=1958"
      Splits(0)._ColumnProps(56)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(11)._WidthInPix=1852"
      Splits(0)._ColumnProps(58)=   "Column(11)._ColStyle=8196"
      Splits(0)._ColumnProps(59)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(60)=   "Column(12).Width=1879"
      Splits(0)._ColumnProps(61)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(12)._WidthInPix=1773"
      Splits(0)._ColumnProps(63)=   "Column(12)._ColStyle=8194"
      Splits(0)._ColumnProps(64)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(65)=   "Column(13).Width=3281"
      Splits(0)._ColumnProps(66)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(13)._WidthInPix=3175"
      Splits(0)._ColumnProps(68)=   "Column(13)._ColStyle=8193"
      Splits(0)._ColumnProps(69)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(71)=   "Column(14).Width=3281"
      Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=3175"
      Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=8196"
      Splits(0)._ColumnProps(75)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(76)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(77)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(78)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(80)=   "Column(15)._ColStyle=8196"
      Splits(0)._ColumnProps(81)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(82)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(83)=   "Column(16).Width=3281"
      Splits(0)._ColumnProps(84)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(16)._WidthInPix=3175"
      Splits(0)._ColumnProps(86)=   "Column(16)._ColStyle=8196"
      Splits(0)._ColumnProps(87)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(88)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(89)=   "Column(17).Width=3281"
      Splits(0)._ColumnProps(90)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(17)._WidthInPix=3175"
      Splits(0)._ColumnProps(92)=   "Column(17)._ColStyle=8196"
      Splits(0)._ColumnProps(93)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(94)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(95)=   "Column(18).Width=3281"
      Splits(0)._ColumnProps(96)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(97)=   "Column(18)._WidthInPix=3175"
      Splits(0)._ColumnProps(98)=   "Column(18)._ColStyle=8196"
      Splits(0)._ColumnProps(99)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(100)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(101)=   "Column(19).Width=3281"
      Splits(0)._ColumnProps(102)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(19)._WidthInPix=3175"
      Splits(0)._ColumnProps(104)=   "Column(19)._ColStyle=8196"
      Splits(0)._ColumnProps(105)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(106)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(107)=   "Column(20).Width=3281"
      Splits(0)._ColumnProps(108)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(109)=   "Column(20)._WidthInPix=3175"
      Splits(0)._ColumnProps(110)=   "Column(20)._ColStyle=8196"
      Splits(0)._ColumnProps(111)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(112)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(113)=   "Column(21).Width=3281"
      Splits(0)._ColumnProps(114)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(21)._WidthInPix=3175"
      Splits(0)._ColumnProps(116)=   "Column(21)._ColStyle=8196"
      Splits(0)._ColumnProps(117)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(118)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(119)=   "Column(22).Width=3281"
      Splits(0)._ColumnProps(120)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(22)._WidthInPix=3175"
      Splits(0)._ColumnProps(122)=   "Column(22)._ColStyle=8196"
      Splits(0)._ColumnProps(123)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(124)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(125)=   "Column(23).Width=3281"
      Splits(0)._ColumnProps(126)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(23)._WidthInPix=3175"
      Splits(0)._ColumnProps(128)=   "Column(23)._ColStyle=8196"
      Splits(0)._ColumnProps(129)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(130)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(131)=   "Column(24).Width=3281"
      Splits(0)._ColumnProps(132)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(133)=   "Column(24)._WidthInPix=3175"
      Splits(0)._ColumnProps(134)=   "Column(24)._ColStyle=8196"
      Splits(0)._ColumnProps(135)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(136)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(137)=   "Column(25).Width=3281"
      Splits(0)._ColumnProps(138)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(25)._WidthInPix=3175"
      Splits(0)._ColumnProps(140)=   "Column(25)._ColStyle=8196"
      Splits(0)._ColumnProps(141)=   "Column(25).Visible=0"
      Splits(0)._ColumnProps(142)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(143)=   "Column(26).Width=3281"
      Splits(0)._ColumnProps(144)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(145)=   "Column(26)._WidthInPix=3175"
      Splits(0)._ColumnProps(146)=   "Column(26)._ColStyle=8196"
      Splits(0)._ColumnProps(147)=   "Column(26).Visible=0"
      Splits(0)._ColumnProps(148)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(149)=   "Column(27).Width=3281"
      Splits(0)._ColumnProps(150)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(151)=   "Column(27)._WidthInPix=3175"
      Splits(0)._ColumnProps(152)=   "Column(27)._ColStyle=8196"
      Splits(0)._ColumnProps(153)=   "Column(27).Visible=0"
      Splits(0)._ColumnProps(154)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(155)=   "Column(28).Width=3281"
      Splits(0)._ColumnProps(156)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(157)=   "Column(28)._WidthInPix=3175"
      Splits(0)._ColumnProps(158)=   "Column(28)._ColStyle=8196"
      Splits(0)._ColumnProps(159)=   "Column(28).Visible=0"
      Splits(0)._ColumnProps(160)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(161)=   "Column(29).Width=3281"
      Splits(0)._ColumnProps(162)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(163)=   "Column(29)._WidthInPix=3175"
      Splits(0)._ColumnProps(164)=   "Column(29)._ColStyle=8196"
      Splits(0)._ColumnProps(165)=   "Column(29).Visible=0"
      Splits(0)._ColumnProps(166)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(167)=   "Column(30).Width=3281"
      Splits(0)._ColumnProps(168)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(169)=   "Column(30)._WidthInPix=3175"
      Splits(0)._ColumnProps(170)=   "Column(30)._ColStyle=8196"
      Splits(0)._ColumnProps(171)=   "Column(30).Visible=0"
      Splits(0)._ColumnProps(172)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(173)=   "Column(31).Width=3281"
      Splits(0)._ColumnProps(174)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(31)._WidthInPix=3175"
      Splits(0)._ColumnProps(176)=   "Column(31)._ColStyle=8196"
      Splits(0)._ColumnProps(177)=   "Column(31).Visible=0"
      Splits(0)._ColumnProps(178)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(179)=   "Column(32).Width=3281"
      Splits(0)._ColumnProps(180)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(181)=   "Column(32)._WidthInPix=3175"
      Splits(0)._ColumnProps(182)=   "Column(32)._ColStyle=8196"
      Splits(0)._ColumnProps(183)=   "Column(32).Visible=0"
      Splits(0)._ColumnProps(184)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(185)=   "Column(33).Width=3281"
      Splits(0)._ColumnProps(186)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(187)=   "Column(33)._WidthInPix=3175"
      Splits(0)._ColumnProps(188)=   "Column(33)._ColStyle=8196"
      Splits(0)._ColumnProps(189)=   "Column(33).Visible=0"
      Splits(0)._ColumnProps(190)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(191)=   "Column(34).Width=3281"
      Splits(0)._ColumnProps(192)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(193)=   "Column(34)._WidthInPix=3175"
      Splits(0)._ColumnProps(194)=   "Column(34)._ColStyle=8196"
      Splits(0)._ColumnProps(195)=   "Column(34).Visible=0"
      Splits(0)._ColumnProps(196)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(197)=   "Column(35).Width=3281"
      Splits(0)._ColumnProps(198)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(199)=   "Column(35)._WidthInPix=3175"
      Splits(0)._ColumnProps(200)=   "Column(35)._ColStyle=8196"
      Splits(0)._ColumnProps(201)=   "Column(35).Visible=0"
      Splits(0)._ColumnProps(202)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(203)=   "Column(36).Width=3493"
      Splits(0)._ColumnProps(204)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(205)=   "Column(36)._WidthInPix=3387"
      Splits(0)._ColumnProps(206)=   "Column(36)._ColStyle=8196"
      Splits(0)._ColumnProps(207)=   "Column(36).Visible=0"
      Splits(0)._ColumnProps(208)=   "Column(36).Order=37"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowAddNew     =   -1  'True
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
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=16,.parent=67"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=68"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=69"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=71"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=67,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=71"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=67,.alignment=2,.locked=-1"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=68"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=69"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=71"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=94,.parent=67,.alignment=0,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=68"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=69"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=71"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=46,.parent=67,.alignment=2,.locked=-1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=68"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=69"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=71"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=98,.parent=67,.alignment=1,.locked=-1"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=95,.parent=68"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=96,.parent=69"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=97,.parent=71"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=102,.parent=67,.alignment=1,.locked=-1"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=99,.parent=68"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=100,.parent=69"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=101,.parent=71"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=114,.parent=67,.alignment=1,.locked=-1"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=111,.parent=68"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=112,.parent=69"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=113,.parent=71"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=20,.parent=67,.alignment=1,.locked=-1"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=17,.parent=68"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=18,.parent=69"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=19,.parent=71"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=24,.parent=67,.alignment=1,.locked=-1"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=21,.parent=68"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=22,.parent=69"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=23,.parent=71"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=54,.parent=67,.alignment=1,.locked=-1"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=68"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=69"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=71"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=28,.parent=67,.locked=-1"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=25,.parent=68"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=26,.parent=69"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=27,.parent=71"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=32,.parent=67,.alignment=1,.locked=-1"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=29,.parent=68"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=30,.parent=69"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=31,.parent=71"
      _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=170,.parent=67,.alignment=2,.locked=-1"
      _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=167,.parent=68"
      _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=168,.parent=69"
      _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=169,.parent=71"
      _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=174,.parent=67,.locked=-1"
      _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=171,.parent=68"
      _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=172,.parent=69"
      _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=173,.parent=71"
      _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=62,.parent=67,.locked=-1"
      _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=59,.parent=68"
      _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=60,.parent=69"
      _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=61,.parent=71"
      _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=66,.parent=67,.locked=-1"
      _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=63,.parent=68"
      _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=64,.parent=69"
      _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=65,.parent=71"
      _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=82,.parent=67,.locked=-1"
      _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=79,.parent=68"
      _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=80,.parent=69"
      _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=81,.parent=71"
      _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=86,.parent=67,.locked=-1"
      _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=83,.parent=68"
      _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=84,.parent=69"
      _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=85,.parent=71"
      _StyleDefs(114) =   "Splits(0).Columns(19).Style:id=90,.parent=67,.locked=-1"
      _StyleDefs(115) =   "Splits(0).Columns(19).HeadingStyle:id=87,.parent=68"
      _StyleDefs(116) =   "Splits(0).Columns(19).FooterStyle:id=88,.parent=69"
      _StyleDefs(117) =   "Splits(0).Columns(19).EditorStyle:id=89,.parent=71"
      _StyleDefs(118) =   "Splits(0).Columns(20).Style:id=106,.parent=67,.locked=-1"
      _StyleDefs(119) =   "Splits(0).Columns(20).HeadingStyle:id=103,.parent=68"
      _StyleDefs(120) =   "Splits(0).Columns(20).FooterStyle:id=104,.parent=69"
      _StyleDefs(121) =   "Splits(0).Columns(20).EditorStyle:id=105,.parent=71"
      _StyleDefs(122) =   "Splits(0).Columns(21).Style:id=110,.parent=67,.locked=-1"
      _StyleDefs(123) =   "Splits(0).Columns(21).HeadingStyle:id=107,.parent=68"
      _StyleDefs(124) =   "Splits(0).Columns(21).FooterStyle:id=108,.parent=69"
      _StyleDefs(125) =   "Splits(0).Columns(21).EditorStyle:id=109,.parent=71"
      _StyleDefs(126) =   "Splits(0).Columns(22).Style:id=118,.parent=67,.locked=-1"
      _StyleDefs(127) =   "Splits(0).Columns(22).HeadingStyle:id=115,.parent=68"
      _StyleDefs(128) =   "Splits(0).Columns(22).FooterStyle:id=116,.parent=69"
      _StyleDefs(129) =   "Splits(0).Columns(22).EditorStyle:id=117,.parent=71"
      _StyleDefs(130) =   "Splits(0).Columns(23).Style:id=122,.parent=67,.locked=-1"
      _StyleDefs(131) =   "Splits(0).Columns(23).HeadingStyle:id=119,.parent=68"
      _StyleDefs(132) =   "Splits(0).Columns(23).FooterStyle:id=120,.parent=69"
      _StyleDefs(133) =   "Splits(0).Columns(23).EditorStyle:id=121,.parent=71"
      _StyleDefs(134) =   "Splits(0).Columns(24).Style:id=126,.parent=67,.locked=-1"
      _StyleDefs(135) =   "Splits(0).Columns(24).HeadingStyle:id=123,.parent=68"
      _StyleDefs(136) =   "Splits(0).Columns(24).FooterStyle:id=124,.parent=69"
      _StyleDefs(137) =   "Splits(0).Columns(24).EditorStyle:id=125,.parent=71"
      _StyleDefs(138) =   "Splits(0).Columns(25).Style:id=130,.parent=67,.locked=-1"
      _StyleDefs(139) =   "Splits(0).Columns(25).HeadingStyle:id=127,.parent=68"
      _StyleDefs(140) =   "Splits(0).Columns(25).FooterStyle:id=128,.parent=69"
      _StyleDefs(141) =   "Splits(0).Columns(25).EditorStyle:id=129,.parent=71"
      _StyleDefs(142) =   "Splits(0).Columns(26).Style:id=134,.parent=67,.locked=-1"
      _StyleDefs(143) =   "Splits(0).Columns(26).HeadingStyle:id=131,.parent=68"
      _StyleDefs(144) =   "Splits(0).Columns(26).FooterStyle:id=132,.parent=69"
      _StyleDefs(145) =   "Splits(0).Columns(26).EditorStyle:id=133,.parent=71"
      _StyleDefs(146) =   "Splits(0).Columns(27).Style:id=138,.parent=67,.locked=-1"
      _StyleDefs(147) =   "Splits(0).Columns(27).HeadingStyle:id=135,.parent=68"
      _StyleDefs(148) =   "Splits(0).Columns(27).FooterStyle:id=136,.parent=69"
      _StyleDefs(149) =   "Splits(0).Columns(27).EditorStyle:id=137,.parent=71"
      _StyleDefs(150) =   "Splits(0).Columns(28).Style:id=142,.parent=67,.locked=-1"
      _StyleDefs(151) =   "Splits(0).Columns(28).HeadingStyle:id=139,.parent=68"
      _StyleDefs(152) =   "Splits(0).Columns(28).FooterStyle:id=140,.parent=69"
      _StyleDefs(153) =   "Splits(0).Columns(28).EditorStyle:id=141,.parent=71"
      _StyleDefs(154) =   "Splits(0).Columns(29).Style:id=146,.parent=67,.locked=-1"
      _StyleDefs(155) =   "Splits(0).Columns(29).HeadingStyle:id=143,.parent=68"
      _StyleDefs(156) =   "Splits(0).Columns(29).FooterStyle:id=144,.parent=69"
      _StyleDefs(157) =   "Splits(0).Columns(29).EditorStyle:id=145,.parent=71"
      _StyleDefs(158) =   "Splits(0).Columns(30).Style:id=150,.parent=67,.locked=-1"
      _StyleDefs(159) =   "Splits(0).Columns(30).HeadingStyle:id=147,.parent=68"
      _StyleDefs(160) =   "Splits(0).Columns(30).FooterStyle:id=148,.parent=69"
      _StyleDefs(161) =   "Splits(0).Columns(30).EditorStyle:id=149,.parent=71"
      _StyleDefs(162) =   "Splits(0).Columns(31).Style:id=154,.parent=67,.locked=-1"
      _StyleDefs(163) =   "Splits(0).Columns(31).HeadingStyle:id=151,.parent=68"
      _StyleDefs(164) =   "Splits(0).Columns(31).FooterStyle:id=152,.parent=69"
      _StyleDefs(165) =   "Splits(0).Columns(31).EditorStyle:id=153,.parent=71"
      _StyleDefs(166) =   "Splits(0).Columns(32).Style:id=158,.parent=67,.locked=-1"
      _StyleDefs(167) =   "Splits(0).Columns(32).HeadingStyle:id=155,.parent=68"
      _StyleDefs(168) =   "Splits(0).Columns(32).FooterStyle:id=156,.parent=69"
      _StyleDefs(169) =   "Splits(0).Columns(32).EditorStyle:id=157,.parent=71"
      _StyleDefs(170) =   "Splits(0).Columns(33).Style:id=162,.parent=67,.locked=-1"
      _StyleDefs(171) =   "Splits(0).Columns(33).HeadingStyle:id=159,.parent=68"
      _StyleDefs(172) =   "Splits(0).Columns(33).FooterStyle:id=160,.parent=69"
      _StyleDefs(173) =   "Splits(0).Columns(33).EditorStyle:id=161,.parent=71"
      _StyleDefs(174) =   "Splits(0).Columns(34).Style:id=166,.parent=67,.locked=-1"
      _StyleDefs(175) =   "Splits(0).Columns(34).HeadingStyle:id=163,.parent=68"
      _StyleDefs(176) =   "Splits(0).Columns(34).FooterStyle:id=164,.parent=69"
      _StyleDefs(177) =   "Splits(0).Columns(34).EditorStyle:id=165,.parent=71"
      _StyleDefs(178) =   "Splits(0).Columns(35).Style:id=178,.parent=67,.locked=-1"
      _StyleDefs(179) =   "Splits(0).Columns(35).HeadingStyle:id=175,.parent=68"
      _StyleDefs(180) =   "Splits(0).Columns(35).FooterStyle:id=176,.parent=69"
      _StyleDefs(181) =   "Splits(0).Columns(35).EditorStyle:id=177,.parent=71"
      _StyleDefs(182) =   "Splits(0).Columns(36).Style:id=182,.parent=67,.locked=-1"
      _StyleDefs(183) =   "Splits(0).Columns(36).HeadingStyle:id=179,.parent=68"
      _StyleDefs(184) =   "Splits(0).Columns(36).FooterStyle:id=180,.parent=69"
      _StyleDefs(185) =   "Splits(0).Columns(36).EditorStyle:id=181,.parent=71"
      _StyleDefs(186) =   "Named:id=33:Normal"
      _StyleDefs(187) =   ":id=33,.parent=0"
      _StyleDefs(188) =   "Named:id=34:Heading"
      _StyleDefs(189) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(190) =   ":id=34,.wraptext=-1"
      _StyleDefs(191) =   "Named:id=35:Footing"
      _StyleDefs(192) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(193) =   "Named:id=36:Selected"
      _StyleDefs(194) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(195) =   "Named:id=37:Caption"
      _StyleDefs(196) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(197) =   "Named:id=38:HighlightRow"
      _StyleDefs(198) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(199) =   "Named:id=39:EvenRow"
      _StyleDefs(200) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(201) =   "Named:id=40:OddRow"
      _StyleDefs(202) =   ":id=40,.parent=33"
      _StyleDefs(203) =   "Named:id=41:RecordSelector"
      _StyleDefs(204) =   ":id=41,.parent=34"
      _StyleDefs(205) =   "Named:id=42:FilterBar"
      _StyleDefs(206) =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "追 加"
      Enabled         =   0   'False
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
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "商品化予定データを追加登録します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新 規"
      Enabled         =   0   'False
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
      Left            =   408
      TabIndex        =   1
      ToolTipText     =   "商品化予定データを全件削除後に、新規登録します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "読　込"
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
      Left            =   420
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "ファイル名"
      Height          =   252
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   13320
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "読込件数"
      Height          =   255
      Index           =   1
      Left            =   12240
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "読込"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu SHORI 
         Caption         =   "新規"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "追加"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   3
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "PLN00301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim PLN_S_YOTEI     As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 36             '最大列数

Private Const colSHORI% = 0             '処理
Private Const colJGYOBU% = 1            'BU
Private Const colST_TANABAN% = 2        '標準棚番
Private Const colHIN_GAI% = 3           '対外品番
Private Const colYOTEI_DT% = 4          '商品化予定日
Private Const colYOTEI_QTY% = 5         '商品化予定数
Private Const colSumi_QTY% = 6          '在庫数(済)
Private Const colMi_QTY% = 7            '在庫数(未)
Private Const colAVE_SYUKA% = 8         '月平均出荷数
Private Const colSUMI_PERCENT% = 9      '事前商品化状況
Private Const colSUMI_GOODS_QTY% = 10   '事前商品化必要数
Private Const colN_YOTEI_DT% = 11       '部品入荷予定日
Private Const colN_YOTEI_QTY% = 12      '部品入荷予定数


'---------------------------------------    非表示
Private Const colYOTEI_DT_X% = 13       '商品化予定日
Private Const colYOTEI_QTY_X% = 14      '商品化予定数
Private Const colS_KOUSU_X% = 15        '見積工数（分/個）
Private Const colS_JIKAN_X% = 16        '標準時間(元)
Private Const colSIZAI% = 17            '資材（箱№）
Private Const colGAISO_HINBAN% = 18     '外装品番
Private Const colGAISO_MAISU% = 19      '外装使用枚数
Private Const colBETU1_SOKO% = 20       '別置１　倉庫
Private Const colBETU1_RETU% = 21       '別置１　列
Private Const colBETU1_REN% = 22        '別置１　連
Private Const colBETU1_DAN% = 23        '別置１　段
Private Const colBETU1_QTY% = 24        '別置１　数量
Private Const colBETU2_SOKO% = 25       '別置２　倉庫
Private Const colBETU2_RETU% = 26       '別置２　列
Private Const colBETU2_REN% = 27        '別置２　連
Private Const colBETU2_DAN% = 28        '別置２　段
Private Const colBETU2_QTY% = 29        '別置２　数量
Private Const colJITU_KOUSU% = 30       '実績工数
Private Const colSAGYOU_KOUSU% = 31     '作業工数
Private Const colNAI_BUHIN% = 32        '国内供給部品区分
Private Const colGAI_BUHIN% = 33        '海外供給部品区分
Private Const colTEHAISAKI% = 34        '商品化完了手配先
Private Const colKEY_NO% = 35           'KEY_NO

Private Const colY_NYUKA_KEY_NO% = 36   '入荷予定KEY_NO


Private KEY_NO  As String * 8
    


'Private Const LAST_UPDATE_DAY$ = "[PLN0030] 2012.09.29 09:30"
Private Const LAST_UPDATE_DAY$ = "[PLN0030] 2018.04.20 14:50"

Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '読込み

            'ﾌｧｲﾙ名ﾁｪｯｸ


            '取込みﾃﾞｰﾀ表示
            If List_Disp_Proc() Then
                Unload Me
            End If


            If PLN_S_YOTEI.Count(1) < 1 Then
                Command1(1).Enabled = False
                Command1(2).Enabled = False
                Command1(4).Enabled = False
                
                
                SHORI(1).Enabled = False
                SHORI(2).Enabled = False
                
                Command1(3).SetFocus
                Exit Sub
            Else
                Command1(1).Enabled = True
                Command1(2).Enabled = True
                Command1(4).Enabled = True
                
                SHORI(1).Enabled = True
                SHORI(2).Enabled = True
                
                Command1(1).SetFocus
            End If


        Case 1, 2       '新規／上書き
            
            If Update_Proc(Index) Then
                Unload Me
            End If


            If PLN_S_YOTEI.Count(1) < 1 Then
                Command1(1).Enabled = False
                Command1(2).Enabled = False
                Command1(4).Enabled = False
                
                SHORI(1).Enabled = False
                SHORI(2).Enabled = False
                
                
                Command1(3).SetFocus
                Exit Sub
            Else
                Command1(1).Enabled = True
                Command1(2).Enabled = True
                Command1(4).Enabled = True
                
                SHORI(1).Enabled = True
                SHORI(2).Enabled = True
                
                Command1(Index).SetFocus
            End If



        Case 3          '終了

            Unload Me
    
    
        Case 4
    
            For i = 15 To 35
                If TDBGrid1.Columns(i).Visible Then
                    TDBGrid1.Columns(i).Visible = False
                Else
                    TDBGrid1.Columns(i).Visible = True
                End If
            Next i
    
    End Select



'    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128



    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[商品化計画システム]商品化計画支援データ読込み処理", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)

                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If







    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)

    PLN00301.Caption = PLN00301.Caption & " " & LAST_UPDATE_DAY

                                '商品化予定ファイルＯＰＥＮ
    If PLN_S_YOTEI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '商品化用入荷予定ファイルＯＰＥＮ
    If PLN_Y_NYUKA_Open(BtOpenRead) Then
        Unload Me
    End If
                                
                                
                                '月平均出荷数ＯＰＥＮ
    If AVE_SYUKA_Open(BtOpenRead) Then
        Unload Me
    End If

''2011.10.04                                '商品化集計ファイルＯＰＥＮ
''    If GOODS_Open(BtOpenRead, 1) Then
''        Unload Me
''    End If

                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenRead) Then
        Unload Me
    End If


                                '品目マスタＯＰＥＮ 2012.09.29
    If ITEM_Open(BtOpenRead) Then
        Unload Me
    End If




End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Text = Trim(Data.Files(1))
    
    Command1(0).Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化予定ファイル")
        End If
    End If
    
    sts = BTRV(BtOpClose, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K0_PLN_Y_NYUKA, Len(K0_PLN_Y_NYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化入荷予定ファイル")
        End If
    End If
    
    
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "月平均出荷数")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set PLN00301 = Nothing



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
        Case 3
            Command1(3).Value = True
    End Select



End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    
    Set TDBGrid1.Array = PLN_S_YOTEI
    TDBGrid1.Update

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

Private Function Update_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   「商品化予定ファイル」登録処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
    
Dim Upd_Com         As Integer
Dim Skip_Flg        As Integer
    
Dim INS_NOW         As String * 14

Dim Row             As Long

Dim KEY_NO          As String * 8

    If PLN_S_YOTEI.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定ファイル登録処理　処理開始！！", Me.hwnd, 0)

                                    
                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
                                    
                                    
                                    
                                    
    If Mode = 1 Then
        If Delete_Proc() Then
            Exit Function
        End If
    End If
                                    
    sts = BTRV(BtOpGetLast, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K3_PLN_S_YOTEI, Len(K3_PLN_S_YOTEI), 3)
    Select Case sts
        Case BtNoErr
            If Val(StrConv(PLN_S_YOTEI_R.KEY_NO, vbUnicode)) = 99999999 Then
                KEY_NO = "00000001"
            Else
                KEY_NO = Format(Val(StrConv(PLN_S_YOTEI_R.KEY_NO, vbUnicode)) + 1, "00000000")
            End If
        Case BtErrEOF
            KEY_NO = "00000001"
        Case Else
            Call File_Error(sts, BtOpGetLast, "商品化予定ファイル")
            Call Input_UnLock
            Exit Function
    End Select
                                    
                                    
                                    
                                    'テーブルリセット
    Skip_Flg = True
    
    
    For Row = 1 To PLN_S_YOTEI.UpperBound(1)
        
        DoEvents
        
        
        
        
        
        
        
        
        Skip_Flg = False
        
        
        
        
        
        If Trim(PLN_S_YOTEI(Row, colKEY_NO)) = "" Then
            If PLN_S_YOTEI(Row, colSHORI) Then
                Upd_Com = BtOpDelete
            Else
                Upd_Com = BtOpInsert
            End If
        Else
            Call UniCode_Conv(K3_PLN_S_YOTEI.KEY_NO, PLN_S_YOTEI(Row, colSHORI))
            
            
            sts = BTRV(BtOpGetEqual, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K3_PLN_S_YOTEI, Len(K3_PLN_S_YOTEI), 3)
            Select Case sts
                Case BtNoErr
                    Upd_Com = BtOpUpdate
                
                    If PLN_S_YOTEI(Row, colSHORI) Then
                        Upd_Com = BtOpDelete
                    End If
                
                Case BtErrKeyNotFound
                    Upd_Com = BtOpInsert
                
                    If PLN_S_YOTEI(Row, colSHORI) Then
                        Skip_Flg = True
                    End If
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "商品化予定ファイル")
                    Call Input_UnLock
                    Exit Function
            End Select
        End If
        
        
        If Not Skip_Flg Then
            If Upd_Com <> BtOpDelete Then
                If Upd_Com = BtOpInsert Then
                   
                    '取込み日付
                    Call UniCode_Conv(PLN_S_YOTEI_R.TORIKOMI_DT, Format(Now, "YYYYMMDD"))
                    '事業部区分
                    Call UniCode_Conv(PLN_S_YOTEI_R.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
                    '国内外
                    Call UniCode_Conv(PLN_S_YOTEI_R.NAIGAI, "1")
                    '外部品番
                    Call UniCode_Conv(PLN_S_YOTEI_R.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
                    '商品化予定リスト印刷日時
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_LIST_DateTime, "")
                    '商品化指図票印刷日時
                    Call UniCode_Conv(PLN_S_YOTEI_R.SASIZU_DateTime, "")
                    '商品化完了登録日時
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_KAN_DateTime, "")
                    '所要量展開日時
                    Call UniCode_Conv(PLN_S_YOTEI_R.TENKAI_DateTime, "")
                    '月平均出荷情報
                    Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
                    Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, "1")
                    Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
                    sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            
                            '総出荷件数
                            Call UniCode_Conv(PLN_S_YOTEI_R.TOTAL_CNT, StrConv(AVE_SYUKAREC.TOTAL_CNT, vbUnicode))
                            
                            '平均総出荷件数
                            Call UniCode_Conv(PLN_S_YOTEI_R.TOTAL_AVE_CNT, StrConv(AVE_SYUKAREC.TOTAL_AVE_CNT, vbUnicode))
                            
                            '生産計画出荷件数(1)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_SYUKA_QTY1, StrConv(AVE_SYUKAREC.S_SYUKA_QTY1, vbUnicode))
                            '生産計画出荷件数(1)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_SYUKA_CNT1, StrConv(AVE_SYUKAREC.S_SYUKA_CNT1, vbUnicode))
                            '平均生産計画出荷数(1)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_AVE_SYUKA_QTY1, StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, vbUnicode))
                            '平均生産計画出荷件数(1)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_AVE_SYUKA_CNT1, StrConv(AVE_SYUKAREC.S_AVE_SYUKA_CNT1, vbUnicode))
                            '生産計画出荷件数(2)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_SYUKA_QTY2, StrConv(AVE_SYUKAREC.S_SYUKA_QTY2, vbUnicode))
                            '生産計画出荷件数(2)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_SYUKA_CNT2, StrConv(AVE_SYUKAREC.S_SYUKA_CNT2, vbUnicode))
                            '平均生産計画出荷数(2)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_AVE_SYUKA_QTY2, StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY2, vbUnicode))
                            '平均生産計画出荷件数(2)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_AVE_SYUKA_CNT2, StrConv(AVE_SYUKAREC.S_AVE_SYUKA_CNT2, vbUnicode))
                        
                        Case BtErrKeyNotFound
                        
                            '総出荷件数
                            Call UniCode_Conv(PLN_S_YOTEI_R.TOTAL_CNT, "00000000")
                            '平均総出荷件数
                            Call UniCode_Conv(PLN_S_YOTEI_R.TOTAL_AVE_CNT, "000000.0")
                            
                            '生産計画出荷件数(1)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_SYUKA_QTY1, "00000000")
                            '生産計画出荷件数(1)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_SYUKA_CNT1, "00000000")
                            '平均生産計画出荷数(1)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_AVE_SYUKA_QTY1, "000000.0")
                            '平均生産計画出荷件数(1)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_AVE_SYUKA_CNT1, "000000.0")
                            '生産計画出荷件数(2)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_SYUKA_QTY2, "00000000")
                            '生産計画出荷件数(2)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_SYUKA_CNT2, "00000000")
                            '平均生産計画出荷数(2)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_AVE_SYUKA_QTY2, "000000.0")
                            '平均生産計画出荷件数(2)
                            Call UniCode_Conv(PLN_S_YOTEI_R.S_AVE_SYUKA_CNT2, "000000.0")
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "月平均出荷数")
                            Call Input_UnLock
                            Exit Function
                    End Select
                    '在庫数(未)
                    Call UniCode_Conv(PLN_S_YOTEI_R.Z_QTY_MI, Format(CLng(PLN_S_YOTEI(Row, colMi_QTY)), "00000000"))
                    '在庫数(済)
                    Call UniCode_Conv(PLN_S_YOTEI_R.Z_QTY_S, Format(CLng(PLN_S_YOTEI(Row, colSumi_QTY)), "00000000"))
                    '事前商品化状況
                    Call UniCode_Conv(PLN_S_YOTEI_R.JIZEN, Format(Val(PLN_S_YOTEI(Row, colSUMI_PERCENT)), "000"))
                    '商品化用部品入荷予定日
                    Call UniCode_Conv(PLN_S_YOTEI_R.NYUKA_YOTEI_DT, Format(PLN_S_YOTEI(Row, colN_YOTEI_DT), "YYYYMMDD"))
                    '商品化用部品入荷予定数
                    If IsNumeric(PLN_S_YOTEI(Row, colN_YOTEI_QTY)) Then
                        Call UniCode_Conv(PLN_S_YOTEI_R.NYUKA_YOTEI_QTY, Format(CLng(PLN_S_YOTEI(Row, colN_YOTEI_QTY)), "00000000"))
                    Else
                        Call UniCode_Conv(PLN_S_YOTEI_R.NYUKA_YOTEI_QTY, "00000000")
                    End If
                    '見積工数
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_KOUSU_X, Format(CDbl(PLN_S_YOTEI(Row, colS_KOUSU_X)), "000000.0"))
                    '商品化　標準時間
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_JIKAN_X, Format(CDbl(PLN_S_YOTEI(Row, colS_JIKAN_X)), "000000.0"))
                    
                    '商品化予定日(元)
                    Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_DT_X, PLN_S_YOTEI(Row, colYOTEI_DT_X))
                    '商品化予定数(元)
                    If IsNumeric(PLN_S_YOTEI(Row, colYOTEI_QTY_X)) Then
                        Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_QTY_X, Format(CLng(PLN_S_YOTEI(Row, colYOTEI_QTY_X)), "00000000"))
                    Else
                        Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_QTY_X, "00000000")
                    End If
                    '資材（箱№）
                    Call UniCode_Conv(PLN_S_YOTEI_R.SIZAI, PLN_S_YOTEI(Row, colSIZAI))
                    '外装品番
                    Call UniCode_Conv(PLN_S_YOTEI_R.GAISO_HINBAN, PLN_S_YOTEI(Row, colGAISO_HINBAN))
                    '外装使用枚数
                    Call UniCode_Conv(PLN_S_YOTEI_R.GAISO_MAISU, Format(Val(PLN_S_YOTEI(Row, colGAISO_MAISU)), "0000"))
                    '外装品番
                    Call UniCode_Conv(PLN_S_YOTEI_R.GAISO_HINBAN, PLN_S_YOTEI(Row, colGAISO_HINBAN))
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   標準棚番を品目マスタから設定するように変更  2012.09.29
                    '標準入庫棚
'                    If Len(Trim(PLN_S_YOTEI(Row, colST_TANABAN))) >= 11 Then
'                        Call UniCode_Conv(PLN_S_YOTEI_R.ST_SOKO, Mid(PLN_S_YOTEI(Row, colST_TANABAN), 1, 2))
'                        Call UniCode_Conv(PLN_S_YOTEI_R.ST_RETU, Mid(PLN_S_YOTEI(Row, colST_TANABAN), 4, 2))
'                        Call UniCode_Conv(PLN_S_YOTEI_R.ST_REN, Mid(PLN_S_YOTEI(Row, colST_TANABAN), 7, 2))
'                        Call UniCode_Conv(PLN_S_YOTEI_R.ST_DAN, Mid(PLN_S_YOTEI(Row, colST_TANABAN), 10, 2))
'                    Else
'                        Call UniCode_Conv(PLN_S_YOTEI_R.ST_SOKO, "")
'                        Call UniCode_Conv(PLN_S_YOTEI_R.ST_RETU, "")
'                        Call UniCode_Conv(PLN_S_YOTEI_R.ST_REN, "")
'                        Call UniCode_Conv(PLN_S_YOTEI_R.ST_DAN, "")
'                    End If

                    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))

                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                    Select Case sts
                        Case BtNoErr

                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                    
                        Case BtErrKeyNotFound

                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_SOKO, "")
                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_RETU, "")
                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_REN, "")
                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_DAN, "")

                        Case Else

                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                            Call Input_UnLock
                            Exit Function

                    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   標準棚番を品目マスタから設定するように変更  2012.09.29
                    '別置１ 棚番
                    Call UniCode_Conv(PLN_S_YOTEI_R.BETU1_SOKO, PLN_S_YOTEI(Row, colBETU1_SOKO))
                    Call UniCode_Conv(PLN_S_YOTEI_R.BETU1_RETU, PLN_S_YOTEI(Row, colBETU1_RETU))
                    Call UniCode_Conv(PLN_S_YOTEI_R.BETU1_REN, PLN_S_YOTEI(Row, colBETU1_REN))
                    Call UniCode_Conv(PLN_S_YOTEI_R.BETU1_DAN, PLN_S_YOTEI(Row, colBETU1_DAN))
                    '別置１ 在庫
                    Call UniCode_Conv(PLN_S_YOTEI_R.BETU1_QTY, Format(CLng(PLN_S_YOTEI(Row, colBETU1_QTY)), "00000000"))
                    '別置２ 棚番
                    Call UniCode_Conv(PLN_S_YOTEI_R.BETU2_SOKO, PLN_S_YOTEI(Row, colBETU2_SOKO))
                    Call UniCode_Conv(PLN_S_YOTEI_R.BETU2_RETU, PLN_S_YOTEI(Row, colBETU2_RETU))
                    Call UniCode_Conv(PLN_S_YOTEI_R.BETU2_REN, PLN_S_YOTEI(Row, colBETU2_REN))
                    Call UniCode_Conv(PLN_S_YOTEI_R.BETU2_DAN, PLN_S_YOTEI(Row, colBETU2_DAN))
                    '別置２ 在庫
                    Call UniCode_Conv(PLN_S_YOTEI_R.BETU2_QTY, Format(CLng(PLN_S_YOTEI(Row, colBETU2_QTY)), "00000000"))
                    '事前商品化必要数
                    Call UniCode_Conv(PLN_S_YOTEI_R.JIZEN_NEEDS_QTY, Format(CLng(PLN_S_YOTEI(Row, colSUMI_GOODS_QTY)), "0000000"))
                    '実績工数
                    Call UniCode_Conv(PLN_S_YOTEI_R.JITU_KOUSU, Format(CDbl(PLN_S_YOTEI(Row, colJITU_KOUSU)), "000000.0"))
                    '作業工数
                    Call UniCode_Conv(PLN_S_YOTEI_R.SAGYOU_KOUSU, Format(CDbl(PLN_S_YOTEI(Row, colSAGYOU_KOUSU)), "000000.0"))
                    '国内供給部品区分
                    Call UniCode_Conv(PLN_S_YOTEI_R.NAI_BUHIN, PLN_S_YOTEI(Row, colNAI_BUHIN))
                    '海外供給部品区分
                    Call UniCode_Conv(PLN_S_YOTEI_R.GAI_BUHIN, PLN_S_YOTEI(Row, colGAI_BUHIN))
                    '商品化完了手配先
                    Call UniCode_Conv(PLN_S_YOTEI_R.TEHAISAKI, PLN_S_YOTEI(Row, colTEHAISAKI))
                
                    'KEY_NO
                    Call UniCode_Conv(PLN_S_YOTEI_R.KEY_NO, KEY_NO)
                    KEY_NO = Format(Val(KEY_NO) + 1, "00000000")
                    
                    '商品化用部品入荷予定日(入力)
                    If IsDate(PLN_S_YOTEI(Row, colN_YOTEI_DT)) Then
                        Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, Format(PLN_S_YOTEI(Row, colN_YOTEI_DT), "YYYYMMDD"))
                    Else
                        Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, "")
                    End If
                    '商品化用部品入荷予定数(入力)
                    If IsNumeric(PLN_S_YOTEI(Row, colN_YOTEI_QTY)) Then
                        Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_QTY, Format(CLng(PLN_S_YOTEI(Row, colN_YOTEI_QTY)), "00000000"))
                    Else
                        Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_QTY, "00000000")
                    End If
                    '入荷予定KEYNO
                    Call UniCode_Conv(PLN_S_YOTEI_R.Y_NYUKA_KEY_NO, PLN_S_YOTEI(Row, colY_NYUKA_KEY_NO))
                    
                    Call UniCode_Conv(PLN_S_YOTEI_R.FILLER, "")
                    
                    '追加　担当者
                    Call UniCode_Conv(PLN_S_YOTEI_R.INS_TANTO, App.EXEName)
                    '追加　日時
                    Call UniCode_Conv(PLN_S_YOTEI_R.Ins_DateTime, INS_NOW)
                    '更新　担当者
                    Call UniCode_Conv(PLN_S_YOTEI_R.UPD_TANTO, "")
                    '更新  日時
                    Call UniCode_Conv(PLN_S_YOTEI_R.UPD_DATETIME, "")

                
                End If
                '商品化予定日付
                If IsDate(PLN_S_YOTEI(Row, colYOTEI_DT)) Then
                    Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_DT, Format(PLN_S_YOTEI(Row, colYOTEI_DT), "YYYYMMDD"))
                Else
                    Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_DT, "")
                End If
                '商品化予定数
                If IsNumeric(PLN_S_YOTEI(Row, colYOTEI_QTY)) Then
                    Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_QTY, Format(CLng(PLN_S_YOTEI(Row, colYOTEI_QTY)), "00000000"))
                Else
                    Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_QTY, "00000000")
                End If
                '商品化　標準工数
                If IsNumeric(PLN_S_YOTEI(Row, colS_KOUSU_X)) Then
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_KOUSU, Format(CDbl(PLN_S_YOTEI(Row, colS_KOUSU_X)), "000000.0"))
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_KOUSU_X, Format(CDbl(PLN_S_YOTEI(Row, colS_KOUSU_X)), "000000.0"))
                Else
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_KOUSU, "000000.0")
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_KOUSU_X, "000000.0")
                End If
                '商品化　標準時間   YOTEI_QTY × S_KOUSU
                If IsNumeric(PLN_S_YOTEI(Row, colS_JIKAN_X)) Then
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_JIKAN, Format(CDbl(PLN_S_YOTEI(Row, colS_JIKAN_X)), "000000.0"))
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_JIKAN_X, Format(CDbl(PLN_S_YOTEI(Row, colS_JIKAN_X)), "000000.0"))
                Else
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_JIKAN_X, "000000.0")
                End If
                '更新担当／日時
                If Upd_Com = BtOpUpdate Then
                    '更新　担当者
                    Call UniCode_Conv(PLN_S_YOTEI_R.UPD_TANTO, App.EXEName)
                    '更新  日時
                    Call UniCode_Conv(PLN_S_YOTEI_R.UPD_DATETIME, INS_NOW)
                End If
            
            
            
            End If
            Do
                sts = BTRV(Upd_Com, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K3_PLN_S_YOTEI, Len(K3_PLN_S_YOTEI), 3)
                Select Case sts
                    Case BtNoErr
                        
                        
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        
                        Beep
                        ans = MsgBox("商品化予定ファイル」他端末でデータ使用中です。<PLN_S_YOTEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Call Input_UnLock
                            Exit Function
                        End If
                    
                    Case Else
                        
                        
                        
                        
                        Call Input_UnLock
                        Call File_Error(sts, Upd_Com, "商品化予定ファイル")
                        Exit Function
                End Select
            
            Loop
            
        End If
            
        Set TDBGrid1.Array = PLN_S_YOTEI
        TDBGrid1.ReBind
        
        TDBGrid1.Update
        TDBGrid1.Bookmark = Row
        

    Next Row


    Set TDBGrid1.Array = PLN_S_YOTEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定ファイル登録処理　処理終了！！", Me.hwnd, 0)




    Call Input_UnLock

    Update_Proc = False
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description
    Call Input_UnLock

End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「商品化予定ファイル」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim INS_NOW         As String * 14
    
    
Dim FileName        As String
Dim FileNo          As Long
    

Dim wkBuf           As String
Dim wkText          As Variant

Dim wkDATE          As String * 8

Dim Skip_Flg        As Integer


Dim JGYOBU          As String * 1       'ＢＵ
Dim ST_TANABAN      As String * 11      '標準棚番("-"付き)
Dim HIN_GAI         As String * 20      '対外品番
Dim YOTEI_DT        As String * 8       '商品化予定日(YYYYMMDD)
Dim YOTEI_QTY       As String * 8       '商品化予定数
Dim Sumi_QTY        As String * 8       '在庫数(済)
Dim Mi_QTY          As String * 8       '在庫数(未)
Dim AVE_SYUKA       As String * 8       '月平均出荷数
Dim SUMI_PERCENT    As String * 8       '事前商品化状況
Dim SUMI_GOODS_QTY  As String * 8       '事前商品化必要数
Dim N_YOTEI_DT      As String * 8       '部品入荷予定日(YYYYMMDD)
Dim N_YOTEI_QTY     As String * 8       '部品入荷予定数


Dim Row             As Long
Dim i               As Integer



    List_Disp_Proc = True

    Call Input_Lock

    FileNo = FreeFile
    FileName = Trim(Text1.Text)
    On Error GoTo Error_Proc

    Open FileName For Input As #FileNo

    On Error GoTo 0

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定ファイル　表示処理開始！！", Me.hwnd, 0)

                                    'テーブルリセット
    Set PLN_S_YOTEI = Nothing
    Row = Min_Row - 1
    Label2.Caption = ""


    sts = BTRV(BtOpGetLast, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K1_PLN_S_YOTEI, Len(K1_PLN_S_YOTEI), 1)
    Select Case sts
        Case BtNoErr
            KEY_NO = Format(Val(StrConv(PLN_S_YOTEI_R.KEY_NO, vbUnicode)), "00000000")
        Case BtErrEOF
            KEY_NO = "00000000"
        Case Else
            Call File_Error(sts, BtOpGetLast, "商品化予定ファイル")
            Call Input_UnLock
            Exit Function
    End Select




    Do Until EOF(FileNo)
        
        
        DoEvents
        
        Line Input #FileNo, wkBuf
    
    
    
    
        wkText = Split(wkBuf, vbTab, -1)
    
    
        If UBound(wkText) < 25 Then
            
            Exit Do
        End If
    
        Skip_Flg = True
        For i = 0 To UBound(JGYOBU_T)
            If wkText(0) = JGYOBU_T(i).CODE Then
                Skip_Flg = False
                Exit For
            End If
        Next i
    
        DoEvents
        
        If Skip_Flg Then
        Else
            
            
            
            On Error GoTo Error_Proc2   '2018.04.20
            
            
            
            
            Row = Row + 1
            PLN_S_YOTEI.ReDim Min_Row, Row, Min_Col, Max_Col
        
            '削除ﾌﾗｸﾞ
            PLN_S_YOTEI(Row, colSHORI) = False
            'BU
            For i = 0 To UBound(JGYOBU_T)
                If wkText(0) = JGYOBU_T(i).CODE Then
                    PLN_S_YOTEI(Row, colJGYOBU) = JGYOBU_T(i).NAME & "          " & JGYOBU_T(i).CODE
                    Exit For
                End If
            Next i
''            PLN_S_YOTEI(Row, colJGYOBU) = wkText(0)
            '標準棚番
            PLN_S_YOTEI(Row, colST_TANABAN) = wkText(1)
            '対外品番
            PLN_S_YOTEI(Row, colHIN_GAI) = wkText(2)
            
            
            
            '商品化予定日                       12--->11    '2012.05.15
            If IsDate(wkText(11)) Then
                PLN_S_YOTEI(Row, colYOTEI_DT) = Format(wkText(11), "YYYY/MM/DD")
            Else
                If Len(wkText(11)) = 8 Then
                    wkDATE = wkText(11)
                    PLN_S_YOTEI(Row, colYOTEI_DT) = Mid(wkDATE, 1, 4) & "/" & Mid(wkDATE, 5, 2) & "/" & Mid(wkDATE, 7, 2)
                Else
                    PLN_S_YOTEI(Row, colYOTEI_DT) = ""
                End If
            End If
            '商品化予定数
            If IsNumeric(wkText(3)) Then
                PLN_S_YOTEI(Row, colYOTEI_QTY) = Format(Val(wkText(3)), "#,##0")
            Else
                PLN_S_YOTEI(Row, colYOTEI_QTY) = "0"
            End If
            '在庫数(済)
            If Left(wkText(4), 1) = """" Then
                wkText(4) = Mid(wkText(4), 2, Len(wkText(4)) - 2)
                wkText(4) = Trim(wkText(4))
            End If
            If IsNumeric(wkText(4)) Then
                PLN_S_YOTEI(Row, colSumi_QTY) = Format(CLng(wkText(4)), "#,##0")
            Else
                PLN_S_YOTEI(Row, colSumi_QTY) = "0"
            End If
            '在庫数(未)
            If Left(wkText(5), 1) = """" Then
                wkText(5) = Mid(wkText(5), 2, Len(wkText(5)) - 2)
                wkText(5) = Trim(wkText(5))
            End If
            If IsNumeric(wkText(5)) Then
                PLN_S_YOTEI(Row, colMi_QTY) = Format(CLng(wkText(5)), "#,##0")
            Else
                PLN_S_YOTEI(Row, colMi_QTY) = "0"
            End If
            '月平均出荷数
            If Left(wkText(6), 1) = """" Then
                wkText(6) = Mid(wkText(6), 2, Len(wkText(6)) - 2)
                wkText(6) = Trim(wkText(6))
            End If
            If IsNumeric(wkText(6)) Then
                PLN_S_YOTEI(Row, colAVE_SYUKA) = Format(CDbl(wkText(6)), "#,##0.0")
            Else
                PLN_S_YOTEI(Row, colAVE_SYUKA) = Format(0, "#0.0")
            End If
            '事前商品化状況(%)
            If IsNumeric(Left(wkText(7), Len(wkText(7)) - 1)) Then
                PLN_S_YOTEI(Row, colSUMI_PERCENT) = Format(Val(Left(wkText(7), Len(wkText(7)) - 1)), "#0") & "%"
            Else
                PLN_S_YOTEI(Row, colSUMI_PERCENT) = Format(0, "#0")
            End If
            '事前商品化必要数                                       20--->18    '2012.05.15
            If Left(wkText(18), 1) = """" Then
                wkText(18) = Mid(wkText(18), 2, Len(wkText(18)) - 2)
                wkText(18) = Trim(wkText(18))
            End If
            If IsNumeric(wkText(18)) Then
                PLN_S_YOTEI(Row, colSUMI_GOODS_QTY) = Format(CLng(wkText(18)), "#,##0")
            Else
                PLN_S_YOTEI(Row, colSUMI_GOODS_QTY) = Format(0, "#0")
            End If
            '部品入荷予定日
            If IsDate(wkText(8)) Then
                PLN_S_YOTEI(Row, colN_YOTEI_DT) = Format(wkText(8), "YYYY/MM/DD")
            Else
                If Len(wkText(8)) = 8 Then
                    wkDATE = wkText(8)
                    PLN_S_YOTEI(Row, colN_YOTEI_DT) = Mid(wkDATE, 1, 4) & "/" & Mid(wkDATE, 5, 2) & "/" & Mid(wkDATE, 7, 2)
                Else
                    PLN_S_YOTEI(Row, colN_YOTEI_DT) = ""
                End If
            End If
            '部品入荷予定数
            If IsNumeric(wkText(9)) Then
                PLN_S_YOTEI(Row, colN_YOTEI_QTY) = Format(Val(wkText(9)), "#,##0")
            Else
                PLN_S_YOTEI(Row, colN_YOTEI_QTY) = Format(0, "#0")
            End If
            
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<   非表示
            
            
            '商品化予定日                       '12---->11  2012.05.15
            If Len(wkText(11)) = 8 Then
                wkDATE = wkText(11)
                PLN_S_YOTEI(Row, colYOTEI_DT_X) = Mid(wkDATE, 1, 4) & "/" & Mid(wkDATE, 5, 2) & "/" & Mid(wkDATE, 7, 2)
            Else
                PLN_S_YOTEI(Row, colYOTEI_DT_X) = ""
            End If
            '商品化予定数
            If IsNumeric(wkText(3)) Then
                PLN_S_YOTEI(Row, colYOTEI_QTY_X) = Format(Val(wkText(3)), "#,##0")
            Else
                PLN_S_YOTEI(Row, colYOTEI_QTY_X) = ""
            End If
            
            
            '見積工数（分/個）                  '10---->19  2012.05.15
            If IsNumeric(wkText(19)) Then
                PLN_S_YOTEI(Row, colS_KOUSU_X) = Format(Val(wkText(19)), "#0.0")
            Else
                PLN_S_YOTEI(Row, colS_KOUSU_X) = 0
            End If
            
            '作業工数                           '11---->10  2012.05.15

            If IsNumeric(wkText(10)) Then
                PLN_S_YOTEI(Row, colS_JIKAN_X) = Format(Val(wkText(10)), "#0.0")
            Else
                PLN_S_YOTEI(Row, colS_JIKAN_X) = 0
            End If
            
            
            
            '資材（箱№）                       '13---->12  2012.05.15
            PLN_S_YOTEI(Row, colSIZAI) = wkText(12)
            '外装品番                           '14---->21  2012.05.15
            PLN_S_YOTEI(Row, colGAISO_HINBAN) = wkText(21)
            '外装使用枚数                       '15---->22  2012.05.15
            PLN_S_YOTEI(Row, colGAISO_MAISU) = wkText(12)
            '別置１　倉庫/列/連/段              '16---->13  2012.05.15
            If Len(wkText(13)) = 11 Then
                PLN_S_YOTEI(Row, colBETU1_SOKO) = Mid(wkText(13), 1, 2)
                PLN_S_YOTEI(Row, colBETU1_RETU) = Mid(wkText(13), 4, 2)
                PLN_S_YOTEI(Row, colBETU1_REN) = Mid(wkText(13), 7, 2)
                PLN_S_YOTEI(Row, colBETU1_DAN) = Mid(wkText(13), 10, 2)
            End If
            '別置１　数量                       '17---->14  2012.05.15
            If Left(wkText(14), 1) = """" Then
                wkText(14) = Mid(wkText(14), 2, Len(wkText(14)) - 2)
                wkText(14) = Trim(wkText(14))
            End If
            If IsNumeric(wkText(14)) Then
                PLN_S_YOTEI(Row, colBETU1_QTY) = Format(CLng(wkText(14)), "#,##0")
            Else
                PLN_S_YOTEI(Row, colBETU1_QTY) = 0
            End If
            '別置２　倉庫/列/連/段              '18---->15  2012.05.15
            If Len(wkText(15)) = 11 Then
                PLN_S_YOTEI(Row, colBETU2_SOKO) = Mid(wkText(15), 1, 2)
                PLN_S_YOTEI(Row, colBETU2_RETU) = Mid(wkText(15), 4, 2)
                PLN_S_YOTEI(Row, colBETU2_REN) = Mid(wkText(15), 7, 2)
                PLN_S_YOTEI(Row, colBETU2_DAN) = Mid(wkText(15), 10, 2)
            End If
            '別置２　数量                       '19---->16  2012.05.15
            If Left(wkText(16), 1) = """" Then
                wkText(16) = Mid(wkText(16), 2, Len(wkText(16)) - 2)
                wkText(16) = Trim(wkText(16))
            End If
            If IsNumeric(wkText(16)) Then
                PLN_S_YOTEI(Row, colBETU2_QTY) = Format(CLng(wkText(16)), "#,##0")
            Else
                PLN_S_YOTEI(Row, colBETU2_QTY) = 0
            End If
            '実績工数                           '21---->19  2012.05.15
            If IsNumeric(wkText(21)) Then
                PLN_S_YOTEI(Row, colJITU_KOUSU) = Format(Val(wkText(19)), "#0.0")
            Else
                PLN_S_YOTEI(Row, colJITU_KOUSU) = 0
            End If
            '作業工数                           '22---->17  2012.05.15
            If IsNumeric(wkText(17)) Then
                PLN_S_YOTEI(Row, colSAGYOU_KOUSU) = Format(Val(wkText(17)), "#0.0")
            Else
                PLN_S_YOTEI(Row, colSAGYOU_KOUSU) = 0
            End If
            '国内供給部品区分
            PLN_S_YOTEI(Row, colNAI_BUHIN) = wkText(23)
            '海外供給部品区分
            PLN_S_YOTEI(Row, colGAI_BUHIN) = wkText(24)
            '商品化完了手配先           2012.08.28 25-->26
            PLN_S_YOTEI(Row, colTEHAISAKI) = wkText(26)
            'KEY_NO
            KEY_NO = Format(Val(KEY_NO) + 1, "00000000")
            PLN_S_YOTEI(Row, colKEY_NO) = KEY_NO
        
        
            On Error GoTo 0             '2018.04.20
                    
        
        
''2011.10.04            '商品化用入荷予定ﾌｧｲﾙKEY_NO獲得
''            PLN_S_YOTEI(Row, colY_NYUKA_KEY_NO) = ""
''            Call UniCode_Conv(K2_GOODS.JGYOBU, PLN_S_YOTEI(Row, colJGYOBU))   '事業部区分
''            Call UniCode_Conv(K2_GOODS.NAIGAI, "1")                           '国内外
''            Call UniCode_Conv(K2_GOODS.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI)) '品番（外部）
''
''            sts = BTRV(BtOpGetEqual, GOODS_POS, GOODSREC, Len(GOODSREC), K2_GOODS, Len(K2_GOODS), 2)
''            Select Case sts
''
''                Case BtNoErr
''                    If IsNumeric(StrConv(GOODSREC.N_YOTEI_QTY, vbUnicode)) And IsNumeric(PLN_S_YOTEI(Row, colN_YOTEI_QTY)) Then
''                        If CLng(StrConv(GOODSREC.N_YOTEI_QTY, vbUnicode)) = CLng(PLN_S_YOTEI(Row, colN_YOTEI_QTY)) Then
''                            PLN_S_YOTEI(Row, colY_NYUKA_KEY_NO) = StrConv(GOODSREC.N_YOTEI_KEY_NO, vbUnicode)
''                        End If
''                    End If
''                Case BtErrKeyNotFound
''
''                Case Else
''                    Call File_Error(sts, BtOpGetEqual, "商品化集計ファイル")
''                    Exit Function
''            End Select
''
        End If



    Loop

    Label2.Caption = Format(Row, "#0") & "件"
    DoEvents

    Set TDBGrid1.Array = PLN_S_YOTEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定ファイル　表示処理終了！！", Me.hwnd, 0)



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
    Exit Function           '2018.04.20

Error_Proc2:
    Select Case Err.Number
        Case 9
    
            MsgBox "指定ファイルは項目数が正しくありません。"
            List_Disp_Proc = False      '
    
        Case Else
        
            MsgBox "エラーが発生しました。ERR.NUMBER=" & Err.Number
        
        
    End Select
    Call Input_UnLock
        
End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    PLN00301.MousePointer = vbHourglass

    Call Ctrl_Lock(PLN00301)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PLN00301)


    PLN00301.MousePointer = vbDefault

End Sub

Private Sub Text1_OLESetData(Data As DataObject, DataFormat As Integer)
'    If DataFormat = vbCFText Then
'        Data.SetData Text1.SelText, vbCFText
'    End If
End Sub

Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   「商品化予定ファイル」削除処理
'----------------------------------------------------------------------------
Dim com     As Integer
Dim sts     As Integer
Dim Row     As Long

    Delete_Proc = True
    
    
    com = BtOpGetFirst
    
    
    Do
        DoEvents
        sts = BTRV(com, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "商品化予定ファイル")
                Exit Function
        End Select
    
        sts = BTRV(BtOpDelete, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
        If sts Then
            Call Input_UnLock
            Call File_Error(sts, BtOpDelete, "商品化予定ファイル")
            Exit Function
        End If
    
        com = BtOpGetNext
    
    Loop
    
    
    Delete_Proc = False

End Function
