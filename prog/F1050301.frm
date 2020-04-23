VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1050301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入出庫実績照会"
   ClientHeight    =   11145
   ClientLeft      =   795
   ClientTop       =   -450
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11145
   ScaleWidth      =   15960
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "画面印刷"
      Height          =   495
      Left            =   13440
      TabIndex        =   33
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   8760
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8655
      Left            =   105
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1080
      Width           =   15750
      _ExtentX        =   27781
      _ExtentY        =   15266
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "品番（外部）"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "品名"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "要因"
      Columns(2).DataField=   ""
      Columns(2).DefaultValue=   "１２３４５"
      Columns(2).DefaultValue.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "伝票日付"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "伝票№"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "入庫数"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "出庫数"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "在訂"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "移動"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "総在庫数"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "向け先"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "担当者"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "メモ"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "実績日付"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "実績時刻"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "移動元"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "移動先"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "入荷日"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "品番（内部）"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "直送先"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "特"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "伝票ＩＤ"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   24
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).Size  =   310
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=24"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3122"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3016"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=4339"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2328"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2223"
      Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2037"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1931"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=1323"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=1217"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=1746"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=1640"
      Splits(0)._ColumnProps(25)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(27)=   "Column(6).Width=1746"
      Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=1640"
      Splits(0)._ColumnProps(30)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(32)=   "Column(7).Width=1746"
      Splits(0)._ColumnProps(33)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(7)._WidthInPix=1640"
      Splits(0)._ColumnProps(35)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(8).Width=1746"
      Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=1640"
      Splits(0)._ColumnProps(40)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(41)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(42)=   "Column(9).Width=1746"
      Splits(0)._ColumnProps(43)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(9)._WidthInPix=1640"
      Splits(0)._ColumnProps(45)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(46)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(47)=   "Column(10).Width=2117"
      Splits(0)._ColumnProps(48)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(10)._WidthInPix=2011"
      Splits(0)._ColumnProps(50)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(51)=   "Column(11).Width=1799"
      Splits(0)._ColumnProps(52)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(11)._WidthInPix=1693"
      Splits(0)._ColumnProps(54)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(55)=   "Column(12).Width=3942"
      Splits(0)._ColumnProps(56)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(12)._WidthInPix=3836"
      Splits(0)._ColumnProps(58)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(59)=   "Column(13).Width=2037"
      Splits(0)._ColumnProps(60)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(13)._WidthInPix=1931"
      Splits(0)._ColumnProps(62)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(63)=   "Column(14).Width=1879"
      Splits(0)._ColumnProps(64)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(14)._WidthInPix=1773"
      Splits(0)._ColumnProps(66)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(67)=   "Column(15).Width=2090"
      Splits(0)._ColumnProps(68)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(15)._WidthInPix=1984"
      Splits(0)._ColumnProps(70)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(71)=   "Column(16).Width=2090"
      Splits(0)._ColumnProps(72)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(16)._WidthInPix=1984"
      Splits(0)._ColumnProps(74)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(75)=   "Column(17).Width=2037"
      Splits(0)._ColumnProps(76)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(77)=   "Column(17)._WidthInPix=1931"
      Splits(0)._ColumnProps(78)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(79)=   "Column(18).Width=3122"
      Splits(0)._ColumnProps(80)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(18)._WidthInPix=3016"
      Splits(0)._ColumnProps(82)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(83)=   "Column(19).Width=3810"
      Splits(0)._ColumnProps(84)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(19)._WidthInPix=3704"
      Splits(0)._ColumnProps(86)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(87)=   "Column(20).Width=476"
      Splits(0)._ColumnProps(88)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(89)=   "Column(20)._WidthInPix=370"
      Splits(0)._ColumnProps(90)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(91)=   "Column(21).Width=2619"
      Splits(0)._ColumnProps(92)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(21)._WidthInPix=2514"
      Splits(0)._ColumnProps(94)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(95)=   "Column(22).Width=2514"
      Splits(0)._ColumnProps(96)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(97)=   "Column(22)._WidthInPix=2408"
      Splits(0)._ColumnProps(98)=   "Column(22)._ColStyle=0"
      Splits(0)._ColumnProps(99)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(100)=   "Column(23).Width=3069"
      Splits(0)._ColumnProps(101)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(102)=   "Column(23)._WidthInPix=2963"
      Splits(0)._ColumnProps(103)=   "Column(23)._ColStyle=2"
      Splits(0)._ColumnProps(104)=   "Column(23).Order=24"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).RepeatGridHeader=   -1  'True
      PrintInfos(0).VariableRowHeight=   -1  'True
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&HFFFF&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=1050,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFF&"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35,.bgcolor=&HFF80&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36,.bgcolor=&HFFFF&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=77,.parent=2,.namedParent=79"
      _StyleDefs(23)  =   "FilterBarStyle:id=80,.parent=1,.namedParent=82"
      _StyleDefs(24)  =   "Splits(0).Style:id=109,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=118,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=110,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=111,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=112,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=114,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=113,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=115,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=116,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=117,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=78,.parent=77"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=81,.parent=80"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=122,.parent=109"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=119,.parent=110"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=120,.parent=111"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=121,.parent=113"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=126,.parent=109"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=123,.parent=110"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=124,.parent=111"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=125,.parent=113"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=130,.parent=109,.alignment=2,.locked=0,.bold=0"
      _StyleDefs(45)  =   ":id=130,.fontsize=1020,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(46)  =   ":id=130,.fontname=ＭＳ ゴシック"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=127,.parent=110,.alignment=3"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=128,.parent=111,.alignment=3"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=129,.parent=113"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=134,.parent=109"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=131,.parent=110"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=132,.parent=111"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=133,.parent=113"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=150,.parent=109"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=147,.parent=110"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=148,.parent=111"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=149,.parent=113"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=158,.parent=109,.alignment=1,.locked=0"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=155,.parent=110,.alignment=3"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=156,.parent=111,.alignment=3"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=157,.parent=113"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=162,.parent=109,.alignment=1,.locked=0"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=159,.parent=110,.alignment=3"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=160,.parent=111,.alignment=3"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=161,.parent=113"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=166,.parent=109,.alignment=1,.locked=0"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=163,.parent=110,.alignment=3"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=164,.parent=111,.alignment=3"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=165,.parent=113"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=170,.parent=109,.alignment=1,.locked=0"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=167,.parent=110,.alignment=3"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=168,.parent=111,.alignment=3"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=169,.parent=113"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=64,.parent=109,.alignment=1,.locked=0"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=61,.parent=110,.alignment=3"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=62,.parent=111,.alignment=3"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=63,.parent=113"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=186,.parent=109"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=183,.parent=110"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=184,.parent=111"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=185,.parent=113"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=190,.parent=109"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=187,.parent=110"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=188,.parent=111"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=189,.parent=113"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=198,.parent=109"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=195,.parent=110"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=196,.parent=111"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=197,.parent=113"
      _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=24,.parent=109"
      _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=21,.parent=110"
      _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=22,.parent=111"
      _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=23,.parent=113"
      _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=28,.parent=109"
      _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=25,.parent=110"
      _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=26,.parent=111"
      _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=27,.parent=113"
      _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=40,.parent=109"
      _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=37,.parent=110"
      _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=38,.parent=111"
      _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=39,.parent=113"
      _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=44,.parent=109"
      _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=41,.parent=110"
      _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=42,.parent=111"
      _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=43,.parent=113"
      _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=48,.parent=109"
      _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=45,.parent=110"
      _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=46,.parent=111"
      _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=47,.parent=113"
      _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=52,.parent=109"
      _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=49,.parent=110"
      _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=50,.parent=111"
      _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=51,.parent=113"
      _StyleDefs(114) =   "Splits(0).Columns(19).Style:id=56,.parent=109"
      _StyleDefs(115) =   "Splits(0).Columns(19).HeadingStyle:id=53,.parent=110"
      _StyleDefs(116) =   "Splits(0).Columns(19).FooterStyle:id=54,.parent=111"
      _StyleDefs(117) =   "Splits(0).Columns(19).EditorStyle:id=55,.parent=113"
      _StyleDefs(118) =   "Splits(0).Columns(20).Style:id=60,.parent=109"
      _StyleDefs(119) =   "Splits(0).Columns(20).HeadingStyle:id=57,.parent=110"
      _StyleDefs(120) =   "Splits(0).Columns(20).FooterStyle:id=58,.parent=111"
      _StyleDefs(121) =   "Splits(0).Columns(20).EditorStyle:id=59,.parent=113"
      _StyleDefs(122) =   "Splits(0).Columns(21).Style:id=68,.parent=109"
      _StyleDefs(123) =   "Splits(0).Columns(21).HeadingStyle:id=65,.parent=110"
      _StyleDefs(124) =   "Splits(0).Columns(21).FooterStyle:id=66,.parent=111"
      _StyleDefs(125) =   "Splits(0).Columns(21).EditorStyle:id=67,.parent=113"
      _StyleDefs(126) =   "Splits(0).Columns(22).Style:id=72,.parent=109,.alignment=0,.locked=0"
      _StyleDefs(127) =   "Splits(0).Columns(22).HeadingStyle:id=69,.parent=110,.alignment=3"
      _StyleDefs(128) =   "Splits(0).Columns(22).FooterStyle:id=70,.parent=111,.alignment=3"
      _StyleDefs(129) =   "Splits(0).Columns(22).EditorStyle:id=71,.parent=113,.alignment=3"
      _StyleDefs(130) =   "Splits(0).Columns(23).Style:id=76,.parent=109,.alignment=1,.locked=0"
      _StyleDefs(131) =   "Splits(0).Columns(23).HeadingStyle:id=73,.parent=110,.alignment=3"
      _StyleDefs(132) =   "Splits(0).Columns(23).FooterStyle:id=74,.parent=111,.alignment=3"
      _StyleDefs(133) =   "Splits(0).Columns(23).EditorStyle:id=75,.parent=113,.alignment=1"
      _StyleDefs(134) =   "Named:id=29:Normal"
      _StyleDefs(135) =   ":id=29,.parent=0"
      _StyleDefs(136) =   "Named:id=30:Heading"
      _StyleDefs(137) =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(138) =   ":id=30,.wraptext=-1"
      _StyleDefs(139) =   "Named:id=31:Footing"
      _StyleDefs(140) =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(141) =   "Named:id=32:Selected"
      _StyleDefs(142) =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(143) =   "Named:id=33:Caption"
      _StyleDefs(144) =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(145) =   "Named:id=34:HighlightRow"
      _StyleDefs(146) =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(147) =   "Named:id=35:EvenRow"
      _StyleDefs(148) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(149) =   "Named:id=36:OddRow"
      _StyleDefs(150) =   ":id=36,.parent=29"
      _StyleDefs(151) =   "Named:id=79:RecordSelector"
      _StyleDefs(152) =   ":id=79,.parent=30"
      _StyleDefs(153) =   "Named:id=82:FilterBar"
      _StyleDefs(154) =   ":id=82,.parent=29"
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   300
      Index           =   0
      Left            =   1080
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2760
      MaxLength       =   20
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
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
      Left            =   10320
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   9480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   8640
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "データ"
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
      Left            =   7800
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "再表示"
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
      Left            =   6480
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   5640
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   5
      Left            =   4800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   4
      Left            =   3960
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "逆  順"
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
      Left            =   2640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   1800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   960
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "正  順"
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
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label lblDateTime 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11400
      TabIndex        =   34
      Top             =   9840
      Width           =   2295
   End
   Begin VB.Label lblcnt 
      Height          =   375
      Left            =   9960
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   735
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
      TabIndex        =   30
      Top             =   10320
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   3960
      TabIndex        =   29
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   3480
      TabIndex        =   28
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   2640
      TabIndex        =   27
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   2160
      TabIndex        =   26
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   1680
      TabIndex        =   25
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日付"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   24
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番"
      Height          =   252
      Index           =   0
      Left            =   2040
      TabIndex        =   22
      Top             =   240
      Width           =   612
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1050301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const pcmbNAIGAI% = 0           '国内外

Private Const ptxHin_Gai% = 0           '品番（外部）
Private Const ptxHin_Name% = 1          '品名
Private Const ptxST_DT_YY% = 2          '開始日付 年
Private Const ptxST_DT_MM% = 3          '開始日付 月
Private Const ptxST_DT_DD% = 4          '開始日付 日
Private Const ptxEN_DT_YY% = 5          '終了日付 年
Private Const ptxEN_DT_MM% = 6          '終了日付 月
Private Const ptxEN_DT_DD% = 7          '終了日付 日

Private Const Text_Max% = 7

Dim IDO     As New XArrayDB

Private Const Min_Row% = 1              '最小行数
'Private Const Max_Row& = 2000           '最大行数
Dim Max_Row     As Long                 'リストボックス最大表示件数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 23             '最大列数

Private Const ColHin_Gai% = 0           '列 品番（外部）
Private Const ColHin_Name% = 1          '列 品名
Private Const ColRIRK% = 2              '列 履歴
Private Const ColDEN_DT% = 3            '列 伝票日付
Private Const ColDEN_No% = 4            '列 伝票№
Private Const ColNyuko_Qty% = 5         '列 入庫数
Private Const ColSyuko_Qty% = 6         '列 出庫数
Private Const ColZAITEI_Qty% = 7        '列 在庫訂正数
Private Const ColIDO_Qty% = 8           '列 移動数
Private Const ColHin_Zaiko_Qty% = 9     '列 品目別在庫数
Private Const ColMUKE_DNAME% = 10       '列 向け先
Private Const ColTANTO_NAME% = 11       '列 ID
Private Const ColMEMO% = 12             '列 メモ
Private Const ColJITU_DT% = 13          '列 実績日付
Private Const ColJITU_TM% = 14          '列 実績時刻
Private Const ColFrom_Location% = 15    '列 From棚
Private Const ColTO_Location% = 16      '列 To棚
Private Const ColNYUKA_DT% = 17         '列 入荷日
Private Const ColHin_Nai% = 18          '列 品番（内部）
Private Const ColSS_Name% = 19          '列 直送先名
Private Const ColTOKU_MARK% = 20        '列 特売りマーク
Private Const ColID_NO% = 21            '列 伝票ＩＤ

Private Const ColSHIIRE_CODE% = 22      '列 仕入先ｺｰﾄﾞ
Private Const ColSHIIRE_TANKA% = 23     '列 SHIIRE_TANKA


'Private Const colPRG_ID% = 24           '列 出力元プログラム


Private Sort_Tbl(ColHin_Gai To ColSHIIRE_TANKA) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順

Dim Excel_Put_Flg       As Boolean      '棚札出力有無


Dim Excel_Template      As String       '棚札 ﾃﾝﾌﾟﾚｰﾄ(ﾌﾙ･ﾊﾟｽ)
Dim Excel_PutPath       As String       '棚札 書き込み先ﾊﾟｽ

Dim Excel_Put_Yoin_IN   As Variant      '棚札 入庫対象要因配列
Dim Excel_Put_Yoin_OUT  As Variant      '棚札 入庫対象要因配列


Dim Excel_Bin_Mei       As Variant      '棚札 便名称配列
Dim ExcelApp            As Excel.Application
Dim Excelbook           As Excel.Workbook
Dim ExcelWorkSheet      As Excel.Worksheet




'---------------------------------------------------------  画面ﾘｻｲｽﾞ対応   2012.02.10
Private Clipped         As Boolean
Private ctls            As Collection
Private clpScaleWidth   As Double
Private clpScaleHeight  As Double
'---------------------------------------------------------　画面ﾘｻｲｽﾞ対応   2012.02.10




Private Function List_Disp_Proc(Mode As Integer) As Integer
                                    '画面表示内容設定
                                    'Mode = 0:昇順
                                    'mode = 1:降順
Dim sts         As Integer
Dim com         As Integer
Dim Key_Mode    As Integer
Dim NAIGAI      As String * 1

Dim ans         As Integer
Dim i           As Integer
Dim Row         As Long
    
Dim SKIP_Flg    As Boolean  '2004.07.16
    
    List_Disp_Proc = True
                                    'エラーチェック
    sts = Item_Read_Proc()
    Select Case sts
        Case False
        Case True
            ans = MsgBox("品目マスタは登録されてません。 処理を継続しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbNo Then
                Text(ptxHin_Gai).SetFocus
                List_Disp_Proc = False
                Exit Function
            End If
        Case Else
            Exit Function
    End Select
        
    For i = ptxST_DT_YY To ptxEN_DT_DD
        Select Case i
            Case ptxST_DT_YY, ptxEN_DT_YY
                If Len(Trim(Text(i).Text)) = 0 Then
                    If i = ptxST_DT_YY Then
                        Text(i).Text = "0000"
                    Else
                        Text(i).Text = "9999"
                    End If
                Else
                    If Not IsNumeric(Text(i).Text) Then
                    Else
                        Text(i).Text = Format(CInt(Text(i).Text), "0000")
                    End If
                End If
            Case Else
                If Len(Trim(Text(i).Text)) = 0 Then
                    If i = ptxST_DT_MM Or i = ptxST_DT_DD Then
                                
                        Text(i).Text = "00"
                    Else
                        Text(i).Text = "99"
                 End If
            Else
                If Not IsNumeric(Text(i).Text) Then
                Else
                    Text(i).Text = Format(CInt(Text(i).Text), "00")
                End If
            End If
        End Select
    Next i
    
    If (Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text) > _
        (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
        MsgBox "検索開始日付が先日付です"
        Text(ptxST_DT_YY).SetFocus
        List_Disp_Proc = False              '2015.03.13
        Exit Function
    End If
                                    
    Call Input_Lock
                                    
                                    'テーブルリセット
    Set IDO = Nothing
    
    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI1                '国内
            NAIGAI = NAIGAI_NAI
            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
                Key_Mode = 0
            Else
                Key_Mode = 1
            End If
        Case NAIGAI2                '海外
            NAIGAI = NAIGAI_GAI
            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
                Key_Mode = 0
            Else
                Key_Mode = 1
            End If
        Case NAIGAI0                '内外指定なし
            Key_Mode = 0
    End Select
                                    
                                    
                                    '在庫移動歴読み込み開始
    If Key_Mode = 0 Then
                                    '時系列で読み込み
        Call UniCode_Conv(K0_IDO.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
        If Mode = 0 Then
            Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
            Call UniCode_Conv(K0_IDO.JITU_TM, "")           '昇順
        Else
            Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text)
            Call UniCode_Conv(K0_IDO.JITU_TM, "zzzzzzzz")   '降順
        End If
                                    '列表示あり 品番／品名
        TDBGrid1.Columns(ColHin_Gai).Visible = True
        TDBGrid1.Columns(ColHin_Name).Visible = True
    
    Else
                                    '品番＞時系列で読み込む
        Call UniCode_Conv(K1_IDO.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K1_IDO.HIN_GAI, Text(ptxHin_Gai).Text)
        If Mode = 0 Then
            Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
            Call UniCode_Conv(K1_IDO.JITU_TM, "")           '昇順
        Else
            Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text)
            Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzzzz")   '降順
        End If
                                    '列表示なし 品番／品名
        TDBGrid1.Columns(ColHin_Gai).Visible = False
        TDBGrid1.Columns(ColHin_Name).Visible = False
    End If
    
    
    Row = Min_Row - 1
        
    If Mode = 0 Then
        com = BtOpGetGreater        '昇順
    Else
        com = BtOpGetLess           '降順
    End If
    Do
        If Key_Mode = 0 Then
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Else
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        End If
    
        SKIP_Flg = False
    
        Select Case sts
            Case BtNoErr
        
                If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = "A" Or _
                    Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = "C" Then
                    SKIP_Flg = True
                End If
        
                '2018.11.21
                If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = "S" Or _
                    Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = "T" Then
                    SKIP_Flg = True
                End If
                '2018.11.21
        
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫移動歴")
                List_Disp_Proc = SYS_ERR
        End Select
                                
        If Not SKIP_Flg Then
                                    
                                    '事業部 KEYﾌﾞﾚｰｸ
            If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Exit Do
            End If
                                    '日付範囲外
            If Mode = 0 Then
                                    '昇順
                If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
                    Exit Do
                End If
            Else
                                    '降順
                If StrConv(IDOREC.JITU_DT, vbUnicode) < (Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text) Then
                    Exit Do
                End If
            End If
            
            If Key_Mode = 1 Then
                                    '品番指定時、品番ﾌﾞﾚｰｸ
                If StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Text(ptxHin_Gai).Text) Then
                    Exit Do
                End If
            End If
        
        
            If Key_Mode = 0 Then
                If StrConv(IDOREC.NAIGAI, vbUnicode) = NAIGAI Then
                    Row = Row + 1
                    If Row > Max_Row Then
                        Beep
                        MsgBox "最大表示行数を超えました。"
                        Exit Do
                    End If
                    Call Grid_Set_Proc(Row)
                End If
            Else
                Row = Row + 1
                If Row > Max_Row Then
                    Beep
                    MsgBox "最大表示行数を超えました。"
                    Exit Do
                End If
                    
                Call Grid_Set_Proc(Row)
            End If
        
        End If
        
        If Mode = 0 Then
            com = BtOpGetNext   '昇順
        Else
            com = BtOpGetPrev   '降順
        End If
        DoEvents
    Loop
                                'DBテーブルリンク
    Set TDBGrid1.Array = IDO
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    
    lblDateTime.Caption = Format(Now, "yyyy/mm/dd HH:MM")       '2018.10.02
    
    
lblcnt.Caption = Row
    
    Call Input_UnLock
    
    
    
    
    Text(ptxHin_Gai).SetFocus
    
    List_Disp_Proc = False

End Function
                                    '画面初期状態を設定する
Private Sub Clear_Field(Mode As Integer)
Dim i As Integer
   
    For i = Mode To Text_Max
        Text(i).Text = ""
    Next i
    
'    Text(ptxST_DT_YY).Text = Mid(Format(Now, "YYYYMMDD"), 1, 4)      '2020/01/10 検索日付初期表示を空白から今日日付に変更
'    Text(ptxST_DT_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)      '2020/01/10 検索日付初期表示を空白から今日日付に変更
'    Text(ptxST_DT_DD).Text = Mid(Format(Now, "YYYYMMDD"), 7, 2)      '2020/01/10 検索日付初期表示を空白から今日日付に変更
'    Text(ptxEN_DT_YY).Text = Mid(Format(Now, "YYYYMMDD"), 1, 4)      '2020/01/10 検索日付初期表示を空白から今日日付に変更
'    Text(ptxEN_DT_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)      '2020/01/10 検索日付初期表示を空白から今日日付に変更
'    Text(ptxEN_DT_DD).Text = Mid(Format(Now, "YYYYMMDD"), 7, 2)      '2020/01/10 検索日付初期表示を空白から今日日付に変更
    
End Sub
                                    '品目マスタより各項目を表示する
Private Function Item_Read_Proc() As Integer

Dim sts     As Integer
Dim NAIGAI  As String * 1

    Item_Read_Proc = True
                                                '国内外の判定
    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI1
            NAIGAI = NAIGAI_NAI
        Case NAIGAI2
            NAIGAI = NAIGAI_GAI
        Case NAIGAI0
            Text(ptxHin_Gai).Text = ""
            Text(ptxHin_Name).Text = ""
            Item_Read_Proc = False
            Exit Function
    End Select
                                                
    If Len(Text(ptxHin_Gai).Text) = 0 Then
        Text(ptxHin_Name).Text = ""
        Item_Read_Proc = False
        Exit Function
    End If
                                                'まず外部品番で読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxHin_Gai).Text)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        
        Case BtErrKeyNotFound
                                                '内部品番で再度読み込み
            Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI)
            Call UniCode_Conv(K2_ITEM.HIN_NAI, Text(ptxHin_Gai).Text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
            Select Case sts
                Case BtNoErr
                    
                    Text(ptxHin_Gai).Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    
        
                Case BtErrKeyNotFound
        
                    Exit Function
        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Item_Read_Proc = SYS_ERR
            End Select
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Item_Read_Proc = SYS_ERR
    End Select
            
    Item_Read_Proc = False

End Function

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        Case pcmbNAIGAI
            Call Clear_Field(0)
            
            If Combo(Index).Text = NAIGAI0 Then
                Text(ptxHin_Gai).Text = ""
                Text(ptxHin_Name).Text = ""
                Text(ptxST_DT_YY).SetFocus
            Else
                Text(ptxHin_Gai).SetFocus
            End If
    End Select

End Sub
Private Sub Command_Click(Index As Integer)
Dim sts As Integer
    
On Error Resume Next
    Select Case Index
        Case 0                           '正順表示
            Text(ptxHin_Gai).Text = RTrim(StrConv(Text(ptxHin_Gai).Text, vbUpperCase))
            If List_Disp_Proc(0) Then
                Unload Me
            End If
'            IDO.QuickSort Min_Row, (IDO.UpperBound(1)), 4, XORDER_ASCEND, XTYPE_DATE, 5, XORDER_ASCEND, XTYPE_DATE
'            TDBGrid1.Refresh
'            Exit Sub
        Case 3                              '逆順表示
            Text(ptxHin_Gai).Text = RTrim(StrConv(Text(ptxHin_Gai).Text, vbUpperCase))
            If List_Disp_Proc(1) Then
                Unload Me
            End If
'            IDO.QuickSort Min_Row, (IDO.UpperBound(1)), 4, XORDER_DESCEND, XTYPE_DATE, 5, XORDER_DESCEND, XTYPE_DATE
'            TDBGrid1.Refresh
'            Exit Sub
        
        
        
        
        Case 4                             '棚札        2007.5.15
            If Tana_Fuda_Put() Then
                Unload Me
            End If
        
        
        
        
        
        Case 7                             '再表示
            Text(ptxHin_Gai).Text = RTrim(StrConv(Text(ptxHin_Gai).Text, vbUpperCase))
            If List_Disp_Proc(0) Then
                Unload Me
            End If
        
        Case 8                             'ﾃﾞｰﾀ出力
            Text(ptxHin_Gai).Text = RTrim(StrConv(Text(ptxHin_Gai).Text, vbUpperCase))
                        
            Call Select_Set_Proc
            
            F1050302.Show vbModal
        
        
        
        
        Case 9
        
            If MsgBox("単価更新しますか？", vbYesNo + vbDefaultButton2, "確認") = vbYes Then
            
                If Update_Proc() Then
                    Unload Me
                End If
            
            End If
        
        
        
        Case 11                            '終了
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Command1_Click()

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "入出庫実績照会 画面印刷を開始しました ", Me.hwnd, 0)


Call Form_HCopy_Win7_NEW(Picture1, vbPRPSA4, vbPRORLandscape)       '2017.04.27


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "入出庫実績照会 画面印刷が終了しました ", Me.hwnd, 0)

End Sub

Private Sub Form_DblClick()
    
'Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)           '2013.11.15
Call Form_HCopy_Win7(Picture1, vbPRPSA4, vbPRORLandscape)       '2013.11.15


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    
        Case vbKeyZ
            If Shift = vbShiftMask Then
                    
                Command(9).Enabled = True
                Command(9).Caption = "更 新"
            
                TDBGrid1.AllowUpdate = True
            
            
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
    
'    If App.PrevInstance Then
'        Beep
'        MsgBox "同一プログラム実行中です。"
'        End
'    End If
    
    Show
                                
                                
                                
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "入出庫実績照会", Me.hwnd, 0)
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
                    
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.02.15 SYS.INI --> F105030.INI
                    
                    '最大表示件数の獲得
'    If GetIni(App.EXEName, "LISTMAX", "SYS", c) Then
    If GetIni(App.EXEName, "LISTMAX", App.EXEName, c) Then
        Beep
        MsgBox "最大表示件数の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    Max_Row = CLng(RTrim(c))
                                
                                
                                
    '棚札用定義情報取り込み 2007.05.15      ここから

                                            
                                            
                    '棚札出力有無
'    If GetIni(App.EXEName, "Excel_Put", "SYS", c) Then
    If GetIni(App.EXEName, "Excel_Put", App.EXEName, c) Then
        Excel_Put_Flg = False
    Else
        If Trim(c) = "1" Then
            Excel_Put_Flg = True
        Else
            Excel_Put_Flg = False
        End If
    End If
                                            
                                            
    If Excel_Put_Flg Then
                                                'ﾃﾝﾌﾟﾚｰﾄ(ﾌﾙ･ﾊﾟｽ)
'        If GetIni(App.EXEName, "F105030_EXCEL_TEMPLATE", "SYS", c) Then
        If GetIni(App.EXEName, "F105030_EXCEL_TEMPLATE", App.EXEName, c) Then
            Beep
            MsgBox "ﾃﾝﾌﾟﾚｰﾄ(ﾌﾙ･ﾊﾟｽ)の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
        Excel_Template = Trim(c)
                                                '書き込み先ﾊﾟｽ
'        If GetIni(App.EXEName, "F105030_EXCEL_OUTPUT", "SYS", c) Then
        If GetIni(App.EXEName, "F105030_EXCEL_OUTPUT", App.EXEName, c) Then
            Beep
            MsgBox "書き込み先ﾊﾟｽの獲得に失敗しました。処理を中止して下さい。"
            End
        End If
        Excel_PutPath = Trim(c)
                                                '対象入庫要因配列
'        If GetIni(App.EXEName, "YOIN_IN", "SYS", c) Then
        If GetIni(App.EXEName, "YOIN_IN", App.EXEName, c) Then
            Beep
            MsgBox "対象入庫要因配列の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
        Excel_Put_Yoin_IN = Split(Trim(c), ",", -1)
                                                '対象出庫要因配列
'        If GetIni(App.EXEName, "YOIN_OUT", "SYS", c) Then
        If GetIni(App.EXEName, "YOIN_OUT", App.EXEName, c) Then
            Beep
            MsgBox "対象出庫要因配列の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
        Excel_Put_Yoin_OUT = Split(Trim(c), ",", -1)
                                                '便名称配列
'        If GetIni("F105030", "BIN", "SYS", c) Then
        If GetIni(App.EXEName, "BIN", App.EXEName, c) Then
            Beep
            MsgBox "便名称配列の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
        Excel_Bin_Mei = Split(Trim(c), ",", -1)
    
    
        Command(4).Enabled = True
        Command(4).Caption = "棚 札"
    
    
    Else
    
        Command(4).Enabled = False
        Command(4).Caption = ""
    End If
    '棚札用定義情報取り込み 2007.05.15      ここまで
                                
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.02.15 SYS.INI --> F105030.INI
                                
                                
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1050301.Caption = "入出庫実績照会(" + RTrim(JGYOBU_T(i).NAME) + ")" & " " & Last_Update_Day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                '国内外取り込み
    Combo(pcmbNAIGAI).AddItem NAIGAI1
    Combo(pcmbNAIGAI).AddItem NAIGAI2
    Combo(pcmbNAIGAI).AddItem NAIGAI0
    Combo(pcmbNAIGAI).Text = NAIGAI1
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
'---------------------------------------------------------　画面ﾘｻｲｽﾞ対応   2012.02.10
    Call ClipControl
'---------------------------------------------------------　画面ﾘｻｲｽﾞ対応   2012.02.10
                                
    Load F1050302
    Load SDC_FLD_F
                                '画面初期設定
    Call Clear_Field(0)
        
    TDBGrid1.Columns(ColHin_Gai).Visible = False
    TDBGrid1.Columns(ColHin_Name).Visible = False
    
    TDBGrid1.style.Locked = True
    
    Combo(pcmbNAIGAI).SetFocus
    
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
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1050301 = Nothing
    Set F1050302 = Nothing
    Set SDC_FLD_F = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1050301.Caption = "入出庫実績照会(" + RTrim(JGYOBU_T(Index).NAME) + ")" & " " & Last_Update_Day
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
'    Text(0).Text = "" '2020/04/10 事業部切り替え時に品番をクリア
'    Text(1).Text = "" '2020/04/10 事業部切り替え時に品名をクリア
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
                            
    If IDO.Count(1) <= 0 Then       '2012.10.12
        Exit Sub                    '2012.10.12
    End If                          '2012.10.12
    
    
    
    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        IDO.QuickSort Min_Row, IDO.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = IDO
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim i   As Integer
Dim sts As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub

    Select Case Index
        Case ptxHin_Gai             '品番
            
            If (Combo(pcmbNAIGAI).Text = NAIGAI0 Or _
                Len(Trim(Text(ptxHin_Gai).Text)) = 0) Then
            Else
                
                Text(Index).Text = RTrim(StrConv(Text(Index).Text, vbUpperCase))
                
                sts = Item_Read_Proc()
                Select Case sts
                    Case False
                    Case True
                        Text(ptxHin_Name).Text = ""
                    Case SYS_ERR
                        Unload Me
                End Select
            End If
                        
        Case ptxST_DT_YY, ptxEN_DT_YY
            If Len(Trim(Text(Index).Text)) = 0 Then
                If Index = ptxST_DT_YY Then
                    Text(Index).Text = "0000"
                Else
                    Text(Index).Text = "9999"
                End If
            Else
                If Not IsNumeric(Text(Index).Text) Then
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "0000")
                End If
            End If
    
        Case ptxST_DT_MM, ptxST_DT_DD, ptxEN_DT_MM, ptxEN_DT_DD
            If Len(Trim(Text(Index).Text)) = 0 Then
                If Index = ptxST_DT_MM Or Index = ptxST_DT_DD Then
                                
                    Text(Index).Text = "00"
                Else
                    Text(Index).Text = "99"
                 End If
            Else
                If Not IsNumeric(Text(Index).Text) Then
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
    
            If Index = ptxEN_DT_DD Then
                If List_Disp_Proc(0) Then
                    Unload Me
                End If
            End If
    
    End Select
        
    For i = Index + 1 To Text_Max
        If Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1050301.MousePointer = vbHourglass

    Call Ctrl_Lock(F1050301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1050301)


    F1050301.MousePointer = vbDefault

End Sub


Private Sub Grid_Set_Proc(Row As Long)


    IDO.ReDim Min_Row, Row, Min_Col, Max_Col
                                            '品目（外部）
    IDO(Row, ColHin_Gai) = StrConv(IDOREC.HIN_GAI, vbUnicode)       '品目（外部）
                                            '品名
    IDO(Row, ColHin_Name) = StrConv(IDOREC.HIN_NAME, vbUnicode)     '品目名称
                                            '履歴（要因）
    IDO(Row, ColRIRK) = StrConv(IDOREC.RIRK_NAME, vbUnicode)        '要因名称
                                            '特売りマーク
    IDO(Row, ColTOKU_MARK) = StrConv(IDOREC.TOKU_MARK, vbUnicode)   '特売りマーク
                                            '実績日付
    IDO(Row, ColJITU_DT) = Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" _
                            & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" _
                            & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2)
                                            '実績時刻
    IDO(Row, ColJITU_TM) = Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 1, 2) & ":" _
                            & Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 3, 2) & ":" _
                            & Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 5, 2)
                                            '伝票日付
    If Len(Trim(StrConv(IDOREC.DEN_DT, vbUnicode))) <> 0 Then
        IDO(Row, ColDEN_DT) = Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 1, 4) & "/" _
                                & Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 5, 2) & "/" _
                                & Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 7, 2)
    End If
                                            '伝票№
    IDO(Row, ColDEN_No) = StrConv(IDOREC.DEN_NO, vbUnicode)
                                            '品目別在庫数
    IDO(Row, ColHin_Zaiko_Qty) = Format(CLng(StrConv(IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode)) + CLng(StrConv(IDOREC.MI_HIN_Zaiko_Qty, vbUnicode)), "#,##0")
                                            '実績数
    Select Case StrConv(IDOREC.SUM_KBN, vbUnicode)
        Case SUM_KBN_IN
                                            '入庫数
            IDO(Row, ColNyuko_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
        Case SUM_KBN_OT
                                            '出庫数
            IDO(Row, ColSyuko_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
                
        Case SUM_KBN_ZT
                If Mid(StrConv(IDOREC.RIRK_ID, vbUnicode), 1, 1) = ACT_ZAITEI_IN Then
                                            '在訂（＋）
                    IDO(Row, ColZAITEI_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
                Else
                                            '在訂（－）
                    IDO(Row, ColZAITEI_Qty) = Format((CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))) * -1), "#,##0")
                End If
        
        Case SUM_KBN_MV
                                            '移動数
                IDO(Row, ColIDO_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
    End Select
                                            'FROM棚
    If Len(Trim(StrConv(IDOREC.FROM_SOKO, vbUnicode))) = 0 Then
    Else
        IDO(Row, ColFrom_Location) = StrConv(IDOREC.FROM_SOKO, vbUnicode) & "-" _
                                    & StrConv(IDOREC.FROM_RETU, vbUnicode) & "-" _
                                    & StrConv(IDOREC.FROM_REN, vbUnicode) & "-" _
                                    & StrConv(IDOREC.FROM_DAN, vbUnicode)
    End If
                                            'TO棚
    If Len(Trim(StrConv(IDOREC.TO_SOKO, vbUnicode))) = 0 Then
    Else
        IDO(Row, ColTO_Location) = StrConv(IDOREC.TO_SOKO, vbUnicode) & "-" _
                                    & StrConv(IDOREC.TO_RETU, vbUnicode) & "-" _
                                    & StrConv(IDOREC.TO_REN, vbUnicode) & "-" _
                                    & StrConv(IDOREC.TO_DAN, vbUnicode)
    End If
                                            '入荷日
    If Len(Trim(StrConv(IDOREC.NYUKA_DT, vbUnicode))) = 0 Then
    Else
        IDO(Row, ColNYUKA_DT) = Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 1, 4) & "/" _
                                & Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" _
                                & Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 7, 2)
    End If
                                            '向け先
    IDO(Row, ColMUKE_DNAME) = StrConv(IDOREC.MUKE_DNAME, vbUnicode)
                                            '担当者
    IDO(Row, ColTANTO_NAME) = StrConv(IDOREC.TANTO_NAME, vbUnicode)
                                            '品番（内部）
    IDO(Row, ColHin_Nai) = StrConv(IDOREC.HIN_NAI, vbUnicode)
                                            'メモ
    IDO(Row, ColMEMO) = StrConv(IDOREC.MEMO, vbUnicode)
                                            
                                            
                                            'メモ
    If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_BINNO Then
    
        IDO(Row, ColSS_Name) = StrConv(IDOREC.SS_CODE, vbUnicode)
    End If
                                            
                                            '伝票ＩＤ
    IDO(Row, ColID_NO) = StrConv(IDOREC.ID_NO, vbUnicode)
                            
                            
                            
                                            '仕入先
    IDO(Row, ColSHIIRE_CODE) = Trim(StrConv(IDOREC.SHIIRE_CODE, vbUnicode))
                                            '仕入単価
    If IsNumeric(StrConv(IDOREC.SHIIRE_TANKA, vbUnicode)) Then
    
        IDO(Row, ColSHIIRE_TANKA) = Format(StrConv(IDOREC.SHIIRE_TANKA, vbUnicode), "#0.00")
    Else
        IDO(Row, ColSHIIRE_TANKA) = ""
                            
    End If
'    TDBGrid1.Update
End Sub



Private Sub Select_Set_Proc()

    F1050302.Combo(pcmbNAIGAI).Text = Combo(pcmbNAIGAI).Text
    F1050302.Text(ptxHin_Gai).Text = Text(ptxHin_Gai).Text
    F1050302.Text(ptxHin_Name).Text = Text(ptxHin_Name).Text
    F1050302.Text(ptxST_DT_YY).Text = Text(ptxST_DT_YY).Text
    F1050302.Text(ptxST_DT_MM).Text = Text(ptxST_DT_MM).Text
    F1050302.Text(ptxST_DT_DD).Text = Text(ptxST_DT_DD).Text
    If Len(Trim(Text(ptxEN_DT_YY).Text)) = 0 Then
        F1050302.Text(ptxEN_DT_YY).Text = "9999"
    Else
        F1050302.Text(ptxEN_DT_YY).Text = Text(ptxEN_DT_YY).Text
    End If
        
    If Len(Trim(Text(ptxEN_DT_MM).Text)) = 0 Then
        F1050302.Text(ptxEN_DT_MM).Text = "99"
    Else
        F1050302.Text(ptxEN_DT_MM).Text = Text(ptxEN_DT_MM).Text
    End If
    
    If Len(Trim(Text(ptxEN_DT_DD).Text)) = 0 Then
        F1050302.Text(ptxEN_DT_DD).Text = "99"
    Else
        F1050302.Text(ptxEN_DT_DD).Text = Text(ptxEN_DT_DD).Text
    End If
End Sub

Private Sub Text_LostFocus(Index As Integer)

    If Index = ptxHin_Gai Then
        Text(Index).Text = RTrim(StrConv(Text(Index).Text, vbUpperCase))
    End If

End Sub

Private Function Tana_Fuda_Put() As Integer

'   部材流通現品棚札　作成                  2007.5.15

Dim strExelFile     As String
Dim Rec_Cnt         As Long
Dim Page_Offset     As Long
Dim posG            As Long

Dim sts             As Integer
Dim com             As Integer
Dim Key_Mode        As Integer
Dim NAIGAI          As String * 1
Dim ans             As Integer
Dim i               As Integer
Dim SKIP_Flg        As Boolean

'On Error GoTo ERR_PRT


    Tana_Fuda_Put = True

    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI1
            NAIGAI = NAIGAI_NAI
        Case NAIGAI2
            NAIGAI = NAIGAI_GAI
        
        
        
        Case NAIGAI0
            MsgBox "国内外は省略できません。", vbExclamation
            Text(ptxHin_Gai).SetFocus
            Tana_Fuda_Put = False
            Exit Function
    
        
    
    End Select
                                    'エラーチェック
    If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
        MsgBox "品番は省略できません。", vbExclamation
        Text(ptxHin_Gai).SetFocus
        Tana_Fuda_Put = False
        Exit Function
    End If

    sts = Item_Read_Proc()
    Select Case sts
        Case False
        Case True
            ans = MsgBox("品目マスタは登録されていません。処理を継続しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbNo Then
                Text(ptxHin_Gai).SetFocus
                Tana_Fuda_Put = False
                Exit Function
            End If
        Case Else
            Exit Function
    End Select

    If Err_Chk_Proc Then            '入力ｴﾗｰﾁｪｯｸ    2007.5.15 (ｻﾌﾞﾌﾟﾛｼｰｼﾞｬ化)
        Exit Function
    End If


    Call Input_Lock

                                    '出力ﾌｧｲﾙ名編集
    strExelFile = Excel_PutPath & Trim(Text(ptxHin_Gai).Text) & ".xls"

    'Excelｱﾌﾟﾘｹｰｼｮﾝｵﾌﾞｼﾞｪｸﾄ取得
    Set ExcelApp = CreateObject("Excel.Application")

    Set Excelbook = ExcelApp.Workbooks.Open(Excel_Template)         'ﾃﾝﾌﾟﾚｰﾄﾌﾞｯｸを開く
'    Set Excelbook = ExcelApp.Workbooks.Add
    
    Set ExcelWorkSheet = Excelbook.Worksheets(1)                    '１ｼｰﾄ目を選択

    '品番
    ExcelWorkSheet.Application.Cells(3, 2).Value = Trim(Text(ptxHin_Gai).Text)
    '発行日
    ExcelWorkSheet.Application.Cells(1, 8).Value = Format(Now, "yyyy/m/d")

                                    '品番＞時系列で読み込む
    Call UniCode_Conv(K1_IDO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K1_IDO.HIN_GAI, Text(ptxHin_Gai).Text)

    Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
    Call UniCode_Conv(K1_IDO.JITU_TM, "")

    Rec_Cnt = 0
    Page_Offset = 6
    posG = 6

    com = BtOpGetGreater
    Do
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)

        SKIP_Flg = False

        Select Case sts
            Case BtNoErr


            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫移動歴")
                Tana_Fuda_Put = SYS_ERR
        End Select

        If Not SKIP_Flg Then
                                    '事業部 KEYﾌﾞﾚｰｸ
            If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Exit Do
            End If
                                    '日付範囲外
            If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
                Exit Do
            End If

                                    '品番指定時、品番ﾌﾞﾚｰｸ
            If StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Text(ptxHin_Gai).Text) Then
                Exit Do
            End If


            For i = 0 To UBound(Excel_Put_Yoin_IN)
                If Trim(StrConv(IDOREC.RIRK_ID, vbUnicode)) = Excel_Put_Yoin_IN(i) Then
                    Call TanaFuda_Set_Proc(1, posG, Page_Offset)
                    Rec_Cnt = Rec_Cnt + 1
                    Exit For
                End If
            Next i
        
        
            For i = 0 To UBound(Excel_Put_Yoin_OUT)
                If Trim(StrConv(IDOREC.RIRK_ID, vbUnicode)) = Excel_Put_Yoin_OUT(i) Then
                    Call TanaFuda_Set_Proc(2, posG, Page_Offset)
                    Rec_Cnt = Rec_Cnt + 1
                    Exit For
                End If
            Next i
        
        End If

        com = BtOpGetNext
        DoEvents
    Loop

    '当該ページの残り行をクリア
    If posG <= Page_Offset + 35 Then
        Call UniCode_Conv(IDOREC.JITU_DT, "")
        Call UniCode_Conv(IDOREC.BIN_NO, "")
        Call UniCode_Conv(IDOREC.DEN_NO, "")
        Call UniCode_Conv(IDOREC.SUM_KBN, "")
        Call UniCode_Conv(IDOREC.TANTO_NAME, "")
        Call UniCode_Conv(IDOREC.RIRK_NAME, "")
        Do
            If posG > Page_Offset + 35 Then
                Exit Do
            End If
            Call TanaFuda_Set_Proc(0, posG, Page_Offset)        '１行分編集
        Loop
    End If



    '編集時にﾜｰｸｼｰﾄの先頭が表示される様に「A1」をｱｸﾃｨﾌﾞにする
    ExcelWorkSheet.Application.Range("A1").Activate

    ExcelApp.DisplayAlerts = False              'ﾏｸﾛ実行ｴﾗｰは表示しない


    If Rec_Cnt > 0 Then
        On Error Resume Next
     '   Kill strExelFile
        ExcelWorkSheet.SaveAs strExelFile
'        On Error GoTo 0
    End If


    ExcelApp.Visible = False
    ExcelApp.Workbooks.Close                                        'ﾜｰｸﾌﾞｯｸ閉じる
    ExcelApp.Quit

    Set ExcelWorkSheet = Nothing                                    'ﾜｰｸｼｰﾄ開放
    Set Excelbook = Nothing                                         'ﾜｰｸﾌﾞｯｸ開放

    Set ExcelApp = Nothing                                         'ﾜｰｸﾌﾞｯｸ開放


    Call Input_UnLock

    Text(ptxHin_Gai).SetFocus


    Tana_Fuda_Put = False

End Function

Private Sub TanaFuda_Set_Proc(InOut As Integer, posG As Long, Page_Offset As Long)


'InOut =0(DUMMY) =1(In) =2()


Dim c   As String * 128


    '１頁分編集完了⇒次頁分のフォーマットをコピー
    If posG > Page_Offset + 35 Then
        ExcelWorkSheet.Application.Range(Page_Offset & ":" & Page_Offset + 35).Copy
        ExcelWorkSheet.Application.Range(Page_Offset + 36 & ":" & Page_Offset + 71).Select
        ExcelWorkSheet.Paste

        Page_Offset = Page_Offset + 36
        posG = Page_Offset
    End If

                                            '実績日付
    If Len(Trim(StrConv(IDOREC.JITU_DT, vbUnicode))) <> 0 Then
        ExcelWorkSheet.Application.Cells(posG, 1).Value = _
                                          Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" _
                                        & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" _
                                        & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2)
    Else
        ExcelWorkSheet.Application.Cells(posG, 1).Value = ""
    End If
                                            '便
    Select Case StrConv(IDOREC.BIN_NO, vbUnicode)
        Case "01"         '１便
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Excel_Bin_Mei(0)
        Case "02"         '２便
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Excel_Bin_Mei(1)
        Case "03"         '３便
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Excel_Bin_Mei(2)
        Case Else
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Trim(StrConv(IDOREC.BIN_NO, vbUnicode))
    End Select
                                            '伝票№
    If InOut = 1 Then
        ExcelWorkSheet.Application.Cells(posG, 3).Value = Trim(StrConv(IDOREC.DEN_NO, vbUnicode))
    End If
                                            '実績数
    ExcelWorkSheet.Application.Cells(posG, 4).Value = ""
    ExcelWorkSheet.Application.Cells(posG, 5).Value = ""
    ExcelWorkSheet.Application.Cells(posG, 6).Value = ""
    Select Case InOut
        Case 1         '入庫数
            ExcelWorkSheet.Application.Cells(posG, 4).Value = _
                Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
        
        
                                            '品目別在庫数
            ExcelWorkSheet.Application.Cells(posG, 6).Value = _
                Val(StrConv(IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode)) + Val(StrConv(IDOREC.MI_HIN_Zaiko_Qty, vbUnicode))
        
        Case 2         '出庫数
            ExcelWorkSheet.Application.Cells(posG, 5).Value = _
                Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
    
                                            '品目別在庫数
            ExcelWorkSheet.Application.Cells(posG, 6).Value = _
                Val(StrConv(IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode)) + Val(StrConv(IDOREC.MI_HIN_Zaiko_Qty, vbUnicode))
    
    End Select
                                            '担当者
    ExcelWorkSheet.Application.Cells(posG, 7).Value = Trim(StrConv(IDOREC.TANTO_NAME, vbUnicode))
                                        
                                        '履歴（要因名称）
    If GetIni(App.EXEName, StrConv(IDOREC.RIRK_ID, vbUnicode), "SYS", c) Then
        ExcelWorkSheet.Application.Cells(posG, 8).Value = ""
    Else
        ExcelWorkSheet.Application.Cells(posG, 8).Value = Trim(c)
    End If
    
    

    posG = posG + 1

End Sub

Private Function Err_Chk_Proc() As Integer

'日付範囲入力エラーチェック    2007.5.15 (ｻﾌﾞﾌﾟﾛｼｰｼﾞｬ化)

Dim sts         As Integer
Dim ans         As Integer
Dim i           As Integer


    Err_Chk_Proc = True


    For i = ptxST_DT_YY To ptxEN_DT_DD
        Select Case i
            Case ptxST_DT_YY, ptxEN_DT_YY
                If Len(Trim(Text(i).Text)) = 0 Then
                    If i = ptxST_DT_YY Then
                        Text(i).Text = "0000"
                    Else
                        Text(i).Text = "9999"
                    End If
                Else
                    If Not IsNumeric(Text(i).Text) Then
                    Else
                        Text(i).Text = Format(CInt(Text(i).Text), "0000")
                    End If
                End If
            Case Else
                If Len(Trim(Text(i).Text)) = 0 Then
                    If i = ptxST_DT_MM Or i = ptxST_DT_DD Then

                        Text(i).Text = "00"
                    Else
                        Text(i).Text = "99"
                 End If
            Else
                If Not IsNumeric(Text(i).Text) Then
                Else
                    Text(i).Text = Format(CInt(Text(i).Text), "00")
                End If
            End If
        End Select
    Next i

    If (Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text) > _
        (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
        MsgBox "検索開始日付が先日付です"
        Text(ptxST_DT_YY).SetFocus
'        Exit Function                      2015.03.13
    End If


    Err_Chk_Proc = False

End Function



Private Function Update_Proc() As Integer

Dim sts As Integer
Dim com As Integer

Dim i   As Integer


    Update_Proc = True
                                     
    Set TDBGrid1.Array = IDO
    TDBGrid1.Refresh
    
    TDBGrid1.Update
                                     
    If IDO.Count(1) < 1 Then
        Update_Proc = False
        Exit Function
    End If
                                    
    For i = 1 To IDO.Count(1)
                                    
        Call UniCode_Conv(K1_IDO.JGYOBU, Last_JGYOBU)
            
        Select Case Combo(pcmbNAIGAI).Text
            Case NAIGAI1                '国内
                Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI_NAI)
            Case NAIGAI2                '海外
                Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI_GAI)
        End Select
            
            
        Call UniCode_Conv(K1_IDO.HIN_GAI, Text(ptxHin_Gai).Text)
        Call UniCode_Conv(K1_IDO.JITU_DT, Format(IDO(i, ColJITU_DT), "YYYYMMDD"))
        Call UniCode_Conv(K1_IDO.JITU_TM, Format(IDO(i, ColJITU_TM), "HHMMSS"))
            
            
                    
            
            
            
        sts = BTRV(BtOpGetEqual, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        Select Case sts
            Case BtNoErr
            
            
                Call UniCode_Conv(IDOREC.SHIIRE_CODE, IDO(i, ColSHIIRE_CODE))
                    
                If IsNumeric(IDO(i, ColSHIIRE_TANKA)) Then
                                        
                    Call UniCode_Conv(IDOREC.SHIIRE_TANKA, Format(IDO(i, ColSHIIRE_TANKA), "00000000.00"))
                Else
                    Call UniCode_Conv(IDOREC.SHIIRE_TANKA, "00000000.00")
                End If
            
            
            
                sts = BTRV(BtOpUpdate, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "在庫移動歴")
                        Exit Function
                End Select
            
            
            
            
            
            
            
            
            
            
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "在庫移動歴")
                Exit Function
        End Select
    
    Next i
                                    
                                    
                                    
                                        
                                        
    
    
    Update_Proc = False
    
    Exit Function


End Function

'---------------------------------------------------------　画面ﾘｻｲｽﾞ対応   2012.02.10
Private Function ClipControl()
    'コントロールの現在の状態をクリップする
    Dim ctl             As Control
    Dim ctlst           As Class1
    
    On Error Resume Next
    
    Set ctls = New Collection
    clpScaleWidth = Me.ScaleWidth
    clpScaleHeight = Me.ScaleHeight
    For Each ctl In Me.Controls
        Set ctlst = New Class1
        With ctlst
            Set .csControl = ctl
            .csLeft = ctl.Left
            .csTop = ctl.Top
            .csWidth = ctl.Width
            .csHeight = ctl.Height
            .csFontSize = ctl.FontSize
        End With
        Call ctls.Add(ctlst)
    Next
    Clipped = True
End Function
'---------------------------------------------------------　画面ﾘｻｲｽﾞ対応   2012.02.10

'---------------------------------------------------------　画面ﾘｻｲｽﾞ対応   2012.02.10
Private Sub Form_Resize()
    'クリップしたコントロールをリサイズする
    Dim ctlst           As Class1
    Dim ratScaleWidth   As Double
    Dim ratScaleHeight  As Double
    
    If Clipped Then
        On Error Resume Next
        '水平、垂直方向の拡大率を決定する
        ratScaleWidth = Me.ScaleWidth / clpScaleWidth
        ratScaleHeight = Me.ScaleHeight / clpScaleHeight
        'それぞれのコントロールを拡大する
        For Each ctlst In ctls
            With ctlst
                .csControl.Top = .csTop * ratScaleHeight
                .csControl.Left = .csLeft * ratScaleWidth
                .csControl.Width = .csWidth * ratScaleWidth
                .csControl.Height = .csHeight * ratScaleHeight
                .csControl.FontSize = .csFontSize * ratScaleWidth  'フォントサイズの拡大幅は適当です
            End With
        Next
    End If
End Sub
'---------------------------------------------------------　画面ﾘｻｲｽﾞ対応   2012.02.10

