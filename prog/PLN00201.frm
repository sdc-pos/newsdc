VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PLN00201 
   Caption         =   "[商品化計画システム]入荷予定データメンテナンス"
   ClientHeight    =   10275
   ClientLeft      =   2025
   ClientTop       =   -5235
   ClientWidth     =   15210
   ClipControls    =   0   'False
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
   ScaleHeight     =   10275
   ScaleWidth      =   15210
   StartUpPosition =   2  '画面の中央
   Begin TrueDBGrid80.TDBDropDown TDBDropDown2 
      Height          =   2055
      Left            =   1800
      TabIndex        =   20
      Top             =   5520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3625
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   714
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4366"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4366"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
      DataField       =   ""
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2055
      Left            =   1800
      TabIndex        =   19
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3625
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   714
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4366"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4366"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
      DataField       =   ""
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      ItemData        =   "PLN00201.frx":0000
      Left            =   1560
      List            =   "PLN00201.frx":000A
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   9
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全削除"
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
      Index           =   3
      Left            =   6240
      TabIndex        =   17
      ToolTipText     =   "表示対象のデータを全削除します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      Index           =   3
      Left            =   9120
      TabIndex        =   7
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   6
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      ItemData        =   "PLN00201.frx":001A
      Left            =   1560
      List            =   "PLN00201.frx":0024
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7575
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   13361
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
      Columns(1).Caption=   "取込み日付"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   1
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ＢＵ"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "TDBDropDown1"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   1
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "データ　　区分"
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "TDBDropDown2"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "対外品番"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "対内品番"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "予定日"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "実績日"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "予定数"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "実績数"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "SEQNo."
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "KEY_NO"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=635"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(9)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1561"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1455"
      Splits(0)._ColumnProps(14)=   "Column(2).Button=1"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2540"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2434"
      Splits(0)._ColumnProps(19)=   "Column(3).Button=1"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=4974"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=4868"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=0"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=5159"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=5054"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=0"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2170"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2064"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=0"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2170"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2064"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=0"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2090"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1984"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2170"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2064"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=1296"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=1191"
      Splits(0)._ColumnProps(54)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=3281"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=3175"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=8196"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
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
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H8000000D&"
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
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=16,.parent=67,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=68"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=69"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=71"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=24,.parent=67"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=71"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=86,.parent=67"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=83,.parent=68"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=84,.parent=69"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=85,.parent=71"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=90,.parent=67"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=87,.parent=68"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=88,.parent=69"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=89,.parent=71"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=94,.parent=67,.alignment=0"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=91,.parent=68"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=92,.parent=69"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=93,.parent=71"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=98,.parent=67,.alignment=0"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=95,.parent=68"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=96,.parent=69"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=97,.parent=71"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=102,.parent=67,.alignment=0"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=99,.parent=68"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=100,.parent=69"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=101,.parent=71"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=106,.parent=67,.alignment=0"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=103,.parent=68"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=104,.parent=69"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=105,.parent=71"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=114,.parent=67,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=111,.parent=68"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=112,.parent=69"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=113,.parent=71"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=118,.parent=67,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=115,.parent=68"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=116,.parent=69"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=117,.parent=71"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=20,.parent=67"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=17,.parent=68"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=18,.parent=69"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=19,.parent=71"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=28,.parent=67,.locked=-1"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=25,.parent=68"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=26,.parent=69"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=27,.parent=71"
      _StyleDefs(86)  =   "Named:id=33:Normal"
      _StyleDefs(87)  =   ":id=33,.parent=0"
      _StyleDefs(88)  =   "Named:id=34:Heading"
      _StyleDefs(89)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(90)  =   ":id=34,.wraptext=-1"
      _StyleDefs(91)  =   "Named:id=35:Footing"
      _StyleDefs(92)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(93)  =   "Named:id=36:Selected"
      _StyleDefs(94)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=37:Caption"
      _StyleDefs(96)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(97)  =   "Named:id=38:HighlightRow"
      _StyleDefs(98)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(99)  =   "Named:id=39:EvenRow"
      _StyleDefs(100) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(101) =   "Named:id=40:OddRow"
      _StyleDefs(102) =   ":id=40,.parent=33"
      _StyleDefs(103) =   "Named:id=41:RecordSelector"
      _StyleDefs(104) =   ":id=41,.parent=34"
      _StyleDefs(105) =   "Named:id=42:FilterBar"
      _StyleDefs(106) =   ":id=42,.parent=33"
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
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      ToolTipText     =   "処理を終了します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更 新"
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
      Left            =   1845
      TabIndex        =   1
      ToolTipText     =   "データ更新を行います"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "検 索"
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
      ToolTipText     =   "データを検索し表示します"
      Top             =   0
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "データ区分"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "対内品番"
      Height          =   252
      Index           =   5
      Left            =   8040
      TabIndex        =   16
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "対外品番"
      Height          =   252
      Index           =   4
      Left            =   3600
      TabIndex        =   15
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "～"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   14
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "予定日"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ＢＵ"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblDisp_Count 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   13560
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "表示件数"
      Height          =   255
      Index           =   1
      Left            =   12480
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "検索"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "更新"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   2
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "PLN00201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Const pcmbBU% = 0               'ＢＵ
Private Const pcmbDATA_KB% = 1          'データ区分


Private Const ptxN_YOTEI_DT_S% = 0      '予定日　開始
Private Const ptxN_YOTEI_DT_E% = 1      '予定日　終了
Private Const ptxHIN_GAI% = 2           '対外品番
Private Const ptxHIN_NAI% = 3           '対内品番




Dim PLN_Y_NYUKA         As New XArrayDB
Dim TDB_BU              As New XArrayDB
Dim TDB_DATA_KB         As New XArrayDB




Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 11             '最大列数

Private Const colSHORI% = 0             '削除ﾌﾗｸﾞ
Private Const colTORIKOMI_DT% = 1
Private Const colJGYOBU% = 2            '事業部
Private Const colDATA_KB% = 3           'データ区分
Private Const colHIN_GAI% = 4           '対外品番
Private Const colHIN_NAI% = 5           '対内品番
Private Const colN_YOTEI_DT% = 6        '予定日
Private Const colJ_NYUKA_DT% = 7        '実績日
Private Const colN_YOTEI_QTY% = 8       '予定数
Private Const colJ_NYUKA_QTY% = 9       '実績数
Private Const colSEQ_NO% = 10           'SEQ_NO
Private Const colKEY_NO% = 11           'KEY_NO

Private Z_Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順




Private DATA_KB             As Variant  'データ区分
Private HIN_CHECK           As Integer  '品番チェック
Private N_YOTEI_DT_CHECK    As Integer  '入荷予定日ﾁｪｯｸ
Private J_NYUKA_DISP        As Integer  '実績ﾁｪｯｸ
Private J_NYUKA_INPUT       As Integer  '実績入力ﾁｪｯｸ

Private LIST_DISP_FLG       As Boolean  '表示中

Private Const LAST_UPDATE_DAY$ = "[PLN0020] 2011.11.09 08:30"
Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long

Dim yn          As Integer


    Select Case Index



        Case 0          '読込み



            '取込みﾃﾞｰﾀ表示
            
            LIST_DISP_FLG = True
            If List_Disp_Proc() Then
                Unload Me
            End If
            LIST_DISP_FLG = False


            If PLN_Y_NYUKA.Count(1) > 0 Then
                Command1(1).Enabled = True
                Command1(3).Enabled = True
                SHORI(1).Enabled = True
            
'                TDBGrid1.Col = colHIN_GAI
'                TDBGrid1.Row = Min_Row - 1
'                TDBGrid1.SetFocus
'                Exit Sub
            Else
                Command1(1).Enabled = False
                Command1(3).Enabled = False
                SHORI(1).Enabled = False
            End If




        Case 1          '登録

            If Error_Check_Proc Then
                Exit Sub
            End If

            If Update_Proc() Then
                Unload Me
            End If


            LIST_DISP_FLG = True
            If List_Disp_Proc() Then
                Unload Me
            End If

            LIST_DISP_FLG = False

            If PLN_Y_NYUKA.Count(1) > 0 Then
                Command1(1).Enabled = True
                Command1(3).Enabled = True
                SHORI(1).Enabled = True
            
'                TDBGrid1.Col = colHIN_GAI
'                TDBGrid1.Row = Min_Row - 1
'                TDBGrid1.SetFocus
'                Exit Sub
            Else
                Command1(1).Enabled = False
                Command1(3).Enabled = False
                SHORI(1).Enabled = False
            End If




        Case 2          '終了

            Unload Me
    
    
        Case 3          '全削除
            
            
            yn = MsgBox("全て削除してよろしいですか？", vbYesNo + vbDefaultButton2, "確認入力")
            If yn = vbYes Then
            
                If Delete_Proc() Then
                    Unload Me
                End If
        
        
                LIST_DISP_FLG = True
                If List_Disp_Proc() Then
                    Unload Me
                End If
                LIST_DISP_FLG = False
    
    
                If PLN_Y_NYUKA.Count(1) < 1 Then
                    Command1(1).Enabled = False
                    Command1(3).Enabled = False
                    SHORI(1).Enabled = False
                    Command1(0).SetFocus
                    Exit Sub
                Else
'                    Command1(1).Enabled = True
'                    Command1(3).Enabled = True
'                    TDBGrid1.Col = colHIN_GAI
'                    TDBGrid1.Row = Min_Row - 1
'                    TDBGrid1.SetFocus
                End If
            End If
    
    End Select



'    Command1(Index).SetFocus


End Sub


Private Sub Form_Load()


Dim c       As String * 128



    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[商品化計画システム]商品化用入荷予定データメンテナンス", Me.hwnd, 0)
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

    PLN00201.Caption = PLN00201.Caption & " " & LAST_UPDATE_DAY


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
                                'データ区分
    If GetIni(App.EXEName, "DATA_KB", App.EXEName, c) Then
        Beep
        MsgBox "データ区分の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    DATA_KB = Split(Trim(c), ",", -1)


                                '品番チェック
    If GetIni(App.EXEName, "HIN_CHECK", App.EXEName, c) Then
        HIN_CHECK = 0
    Else
        If Not IsNumeric(c) Then
            HIN_CHECK = 0
        Else
            HIN_CHECK = Val(Trim(c))
        End If
    End If
                                '入荷予定日チェック
    If GetIni(App.EXEName, "N_YOTEI_DT_CHECK", App.EXEName, c) Then
        N_YOTEI_DT_CHECK = 0
    Else
        If Not IsNumeric(c) Then
            N_YOTEI_DT_CHECK = 0
        Else
            N_YOTEI_DT_CHECK = Val(Trim(c))
        End If
    End If

                                '実績ﾁｪｯｸ
    If GetIni(App.EXEName, "J_NYUKA_DISP", App.EXEName, c) Then
        J_NYUKA_DISP = 0
    Else
        If Not IsNumeric(c) Then
            J_NYUKA_DISP = 0
        Else
            J_NYUKA_DISP = Val(Trim(c))
        End If
    End If
                                '実績入力ﾁｪｯｸ
    If GetIni(App.EXEName, "J_NYUKA_INPUT", App.EXEName, c) Then
        J_NYUKA_INPUT = 0
    Else
        If Not IsNumeric(c) Then
            J_NYUKA_INPUT = 0
        Else
            J_NYUKA_INPUT = Val(Trim(c))
        End If
    End If


    LIST_DISP_FLG = True
    
    Call Bu_Set_Proc
    Call Data_Kb_Set_Proc

    Call TDB_BU_Set_Proc
    Call TDB_DATA_KB_Set_Proc
    
    LIST_DISP_FLG = False

    
    


    
    '商品化用入荷予定ファイル　ＯＰＥＮ
    If PLN_Y_NYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If

    '品目マスタ　ＯＰＥＮ
    If ITEM_Open(BtOpenRead) Then
        Unload Me
    End If


    'ＰＮマスタ　ＯＰＥＮ
    If HIN_CHECK = 1 Then
        If PN_M_Open(BtOpenRead) Then
            Unload Me
        End If
    End If

    
    
''    If PLN_Y_NYUKA.Count(1) < 1 Then
''        Text1(ptxN_YOTEI_DT_S).SetFocus
''    Else
''        TDBGrid1.Col = colHIN_GAI
''        TDBGrid1.Row = Min_Row
''        TDBGrid1.SetFocus
''    End If
    
    Command1(0).SetFocus
End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    
    
    Call PLN_Y_NYUKA_CLOSE
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set PLN00201 = Nothing



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


Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   「商品化用入荷予定」登録処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim Upd_Com         As Integer
Dim Skip_Flg        As Integer
    
Dim INS_NOW         As String * 14
Dim svSeq_No        As String * 3
Dim KEY_NO          As String * 8


Dim Row             As Long

    If PLN_Y_NYUKA.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化用入荷予定ファイル　[更新]処理開始！！", Me.hwnd, 0)

                                    
                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
                                    
                                    'テーブルリセット
    
    Skip_Flg = True
    For Row = 1 To PLN_Y_NYUKA.UpperBound(1)
        
        
        DoEvents
        
        
        sts = BTRV(BtOpGetLast, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
        Select Case sts
            Case BtNoErr
            
                KEY_NO = Format(Val(StrConv(PLN_Y_NYUKA_R.KEY_NO, vbUnicode)) + 1, "00000000")
            
            Case BtErrEOF
                KEY_NO = "00000001"
            Case Else
        
        
                Call File_Error(sts, BtOpGetEqual, "商品化用入荷予定ファイル")
                Call Input_UnLock
                
                
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "商品化用入荷予定ファイル　[更新]処理異常停止！！", Me.hwnd, 0)
                
                
                Exit Function
        
        
        
        End Select
        
        
        
        
        Skip_Flg = False
        
                
                
                
        If Trim(PLN_Y_NYUKA(Row, colKEY_NO)) <> "" Then
            KEY_NO = PLN_Y_NYUKA(Row, colKEY_NO)
        End If

        Call UniCode_Conv(K4_PLN_Y_NYUKA.KEY_NO, KEY_NO)
        sts = BTRV(BtOpGetEqual, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
        Select Case sts
            Case BtNoErr
                Upd_Com = BtOpUpdate
            Case BtErrKeyNotFound
                Upd_Com = BtOpInsert
            Case Else
                Call File_Error(sts, BtOpGetEqual, "商品化用入荷予定ファイル")
                Call Input_UnLock


                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "商品化用入荷予定ファイル　[更新]処理異常停止！！", Me.hwnd, 0)


                Exit Function

        End Select




'
'        Call UniCode_Conv(K3_PLN_Y_NYUKA.JGYOBU, Right(PLN_Y_NYUKA(Row, colJGYOBU), 1))
'        Call UniCode_Conv(K3_PLN_Y_NYUKA.DATA_KB, Right(PLN_Y_NYUKA(Row, colDATA_KB), 2))
'        Call UniCode_Conv(K3_PLN_Y_NYUKA.NAIGAI, "1")
'        Call UniCode_Conv(K3_PLN_Y_NYUKA.HIN_GAI, PLN_Y_NYUKA(Row, colHIN_GAI))
'        Call UniCode_Conv(K3_PLN_Y_NYUKA.N_YOTEI_DT, PLN_Y_NYUKA(Row, colN_YOTEI_DT))
'        Call UniCode_Conv(K3_PLN_Y_NYUKA.SEQ_NO, "000")
'
'
'        sts = BTRV(BtOpGetEqual, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K3_PLN_Y_NYUKA, Len(K3_PLN_Y_NYUKA), 3)
'        Select Case sts
'            Case BtNoErr
'
'
'                If Trim(PLN_Y_NYUKA(Row, colSEQ_NO)) = "" Then
'
'                    svSeq_No = "001"
'                    Do
'
'                        DoEvents
'
'                        sts = BTRV(BtOpGetNext, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K3_PLN_Y_NYUKA, Len(K3_PLN_Y_NYUKA), 3)
'                        Select Case sts
'                            Case BtNoErr
'
'                                If Right(PLN_Y_NYUKA(Row, colJGYOBU), 1) = StrConv(PLN_Y_NYUKA_R.JGYOBU, vbUnicode) And _
'                                    Right(PLN_Y_NYUKA(Row, colDATA_KB), 2) = StrConv(PLN_Y_NYUKA_R.DATA_KB, vbUnicode) And _
'                                    "1" = StrConv(PLN_Y_NYUKA_R.DATA_KB, vbUnicode) And _
'                                    PLN_Y_NYUKA(Row, colHIN_GAI) = StrConv(PLN_Y_NYUKA_R.HIN_GAI, vbUnicode) And _
'                                    PLN_Y_NYUKA(Row, colN_YOTEI_DT) = StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode) Then
'
'                                    svSeq_No = Format(Val(svSeq_No) + 1, "000")
'                                Else
'                                    Upd_Com = BtOpInsert
'                                    Exit Do
'                                End If
'
'                            Case BtErrEOF
'                                Upd_Com = BtOpInsert
'                                Exit Do
'                            Case Else
'
'                        End Select
'
'
'
'
'
'                    Loop
'
'
'                Else
'
'                    If PLN_Y_NYUKA(Row, colSHORI) = "1" Then
'                        Upd_Com = BtOpDelete
'                    Else
'                        Upd_Com = BtOpUpdate
'                    End If
'
'
'                End If
'
'
'            Case BtErrKeyNotFound
'
'
'                Upd_Com = BtOpInsert
'                svSeq_No = "000"
'
'                If PLN_Y_NYUKA(Row, colSHORI) = "1" Then
'                    Skip_Flg = True
'                End If
'            Case Else
'                Call File_Error(sts, BtOpGetEqual, "商品化用入荷予定ファイル")
'                Call Input_UnLock
'
'
'                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
'                    "商品化用入荷予定ファイル　[更新]処理異常停止！！", Me.hwnd, 0)
'
'
'                Exit Function
'        End Select
        
        
        If Upd_Com = BtOpInsert Then
            
            Call PLN_Y_NYUKA_CLR
            
            
            Call UniCode_Conv(PLN_Y_NYUKA_R.TORIKOMI_DT, Format(Now, "YYYYMMDD"))
            
            Call UniCode_Conv(PLN_Y_NYUKA_R.KEY_NO, KEY_NO)
            Call UniCode_Conv(PLN_Y_NYUKA_R.INS_TANTO, App.EXEName)                                 '追加担当者
            Call UniCode_Conv(PLN_Y_NYUKA_R.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))            '追加日時
                    
            Call UniCode_Conv(PLN_Y_NYUKA_R.SEQ_NO, "000")
        Else
        
            Call UniCode_Conv(K4_PLN_Y_NYUKA.KEY_NO, KEY_NO)
        
            sts = BTRV(BtOpGetEqual, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
            Select Case sts
                Case BtNoErr
                
                    Call UniCode_Conv(PLN_Y_NYUKA_R.UPD_TANTO, App.EXEName)                         '更新担当者
                    Call UniCode_Conv(PLN_Y_NYUKA_R.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))    '更新日時
                
                Case BtErrEOF
                
                    Call PLN_Y_NYUKA_CLR
                    
                    
                    Call UniCode_Conv(PLN_Y_NYUKA_R.TORIKOMI_DT, Format(Now, "YYYYMMDD"))
                    
                    Call UniCode_Conv(PLN_Y_NYUKA_R.KEY_NO, KEY_NO)
                    Call UniCode_Conv(PLN_Y_NYUKA_R.INS_TANTO, App.EXEName)                                 '追加担当者
                    Call UniCode_Conv(PLN_Y_NYUKA_R.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))            '追加日時
                
                    Call UniCode_Conv(PLN_Y_NYUKA_R.SEQ_NO, "000")
                
                
                    Upd_Com = BtOpInsert
                
                
                Case Else
            
            
                    Call File_Error(sts, BtOpGetEqual, "商品化用入荷予定ファイル")
                    Call Input_UnLock
                    
                    
                    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                        "商品化用入荷予定ファイル　[更新]処理異常停止！！", Me.hwnd, 0)
                    
                    
                    Exit Function
            
            
            
            End Select
        
        
        
        End If
        
        
        
        If PLN_Y_NYUKA(Row, colSHORI) Then
            Upd_Com = BtOpDelete
        End If
        
        
        Call UniCode_Conv(PLN_Y_NYUKA_R.JGYOBU, Right(PLN_Y_NYUKA(Row, colJGYOBU), 1))
        Call UniCode_Conv(PLN_Y_NYUKA_R.NAIGAI, "1")
        Call UniCode_Conv(PLN_Y_NYUKA_R.DATA_KB, Right(PLN_Y_NYUKA(Row, colDATA_KB), 2))
        Call UniCode_Conv(PLN_Y_NYUKA_R.HIN_GAI, PLN_Y_NYUKA(Row, colHIN_GAI))
        Call UniCode_Conv(PLN_Y_NYUKA_R.HIN_NAI, PLN_Y_NYUKA(Row, colHIN_NAI))
        Call UniCode_Conv(PLN_Y_NYUKA_R.N_YOTEI_DT, Format(PLN_Y_NYUKA(Row, colN_YOTEI_DT), "YYYYMMDD"))
        Call UniCode_Conv(PLN_Y_NYUKA_R.N_YOTEI_QTY, Format(Val(PLN_Y_NYUKA(Row, colN_YOTEI_QTY)), "00000000"))
        
        If Trim(PLN_Y_NYUKA(Row, colJ_NYUKA_DT)) = "" Then
        Else
            Call UniCode_Conv(PLN_Y_NYUKA_R.J_NYUKA_DT, Format(PLN_Y_NYUKA(Row, colJ_NYUKA_DT), "YYYYMMDD"))
        End If
        
        
        If Trim(PLN_Y_NYUKA(Row, colJ_NYUKA_QTY)) = "" Then
        Else
            Call UniCode_Conv(PLN_Y_NYUKA_R.J_NYUKA_QTY, Format(PLN_Y_NYUKA(Row, colJ_NYUKA_QTY), "00000000"))
        End If
        
'        If Trim(PLN_Y_NYUKA(Row, colSEQ_NO)) <> "" Then
'            Call UniCode_Conv(PLN_Y_NYUKA_R.SEQ_NO, PLN_Y_NYUKA(Row, colSEQ_NO))
'        Else
'            Call UniCode_Conv(PLN_Y_NYUKA_R.SEQ_NO, svSeq_No)
'        End If
        
        Call UniCode_Conv(PLN_Y_NYUKA_R.SEQ_NO, "000")
        
        
        Do
            sts = BTRV(Upd_Com, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    
                    Beep
                    ans = MsgBox("「商品化用入荷予定ファイル」他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Call Input_UnLock
                        Exit Function
                    End If
                
                Case BtErrKeyNotFound
                
                
                Case Else
                    If Upd_Com <> BtOpDelete Then
                        Call Input_UnLock
                        Call File_Error(sts, Upd_Com, "商品化用入荷予定ファイル")
                        Exit Function
                    End If
            End Select
        
        Loop
            
    

    Next Row





hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化用入荷予定ファイル　[更新]処理終了！！", Me.hwnd, 0)




    Update_Proc = False
    Call Input_UnLock
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description
    Call Input_UnLock

End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「商品化用入荷予定ファイル」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer

Dim Row             As Long


Dim i               As Long

Dim Skip_Flg        As Boolean
Dim Y_Len           As Long
    
    
    List_Disp_Proc = True

    Call Input_Lock


    LIST_DISP_FLG = True

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化用入荷予定ファイル　[検索]処理開始！！", Me.hwnd, 0)

                                    'テーブルリセット
    Set PLN_Y_NYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""


    If IsDate(Text1(ptxN_YOTEI_DT_S).Text) Then
        Call UniCode_Conv(K2_PLN_Y_NYUKA.N_YOTEI_DT, Format(Text1(ptxN_YOTEI_DT_S).Text, "YYYYMMDD"))
    Else
        Call UniCode_Conv(K2_PLN_Y_NYUKA.N_YOTEI_DT, "")
    End If
    Call UniCode_Conv(K2_PLN_Y_NYUKA.SEQ_NO, "")
    Call UniCode_Conv(K2_PLN_Y_NYUKA.DATA_KB, "")
    Call UniCode_Conv(K2_PLN_Y_NYUKA.JGYOBU, "")
    Call UniCode_Conv(K2_PLN_Y_NYUKA.NAIGAI, "")
    Call UniCode_Conv(K2_PLN_Y_NYUKA.HIN_GAI, "")


    com = BtOpGetGreaterEqual



    Do
        DoEvents
        sts = BTRV(com, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K2_PLN_Y_NYUKA, Len(K2_PLN_Y_NYUKA), 2)
        Select Case sts
            Case BtNoErr
            
                If IsDate(Text1(ptxN_YOTEI_DT_E).Text) Then
                    If StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode) > Format(Text1(ptxN_YOTEI_DT_E).Text, "YYYYMMDD") Then
                        Exit Do
                    End If
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                LIST_DISP_FLG = False
                Call File_Error(sts, com, "商品化用入荷予定ファイル")
                Exit Function
        End Select
    
        Skip_Flg = False
        If Right(Combo1(pcmbDATA_KB).Text, 2) <> "  " Then
            If Right(Combo1(pcmbDATA_KB).Text, 2) <> StrConv(PLN_Y_NYUKA_R.DATA_KB, vbUnicode) Then
                Skip_Flg = True
            End If
        End If
        
        If Right(Combo1(pcmbBU).Text, 1) <> " " Then
            If Right(Combo1(pcmbBU).Text, 1) <> StrConv(PLN_Y_NYUKA_R.JGYOBU, vbUnicode) Then
                Skip_Flg = True
            End If
        End If
        
        
        If Trim(Text1(ptxHIN_GAI).Text) <> "" Then
            Y_Len = Len(Trim(Text1(ptxHIN_GAI).Text))
            If Trim(Text1(ptxHIN_GAI).Text) <> Left(Trim(StrConv(PLN_Y_NYUKA_R.HIN_GAI, vbUnicode)), Y_Len) Then
                Skip_Flg = True
            End If
        End If
        
        If Trim(Text1(ptxHIN_NAI).Text) <> "" Then
            Y_Len = Len(Trim(Text1(ptxHIN_NAI).Text))
            If Trim(Text1(ptxHIN_NAI).Text) <> Left(Trim(StrConv(PLN_Y_NYUKA_R.HIN_NAI, vbUnicode)), Y_Len) Then
                Skip_Flg = True
            End If
        End If
        
        If Skip_Flg Then
        Else
            Row = Row + 1
            PLN_Y_NYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
            PLN_Y_NYUKA(Row, colSHORI) = False
            
            PLN_Y_NYUKA(Row, colTORIKOMI_DT) = StrConv(PLN_Y_NYUKA_R.TORIKOMI_DT, vbUnicode)
            
            
            
            
            For i = 0 To UBound(JGYOBU_T)
                If StrConv(PLN_Y_NYUKA_R.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
                    PLN_Y_NYUKA(Row, colJGYOBU) = JGYOBU_T(i).NAME & "          " & JGYOBU_T(i).CODE
                    Exit For
                End If
            Next i
            
            For i = 0 To UBound(DATA_KB) - 1 Step 2
                If StrConv(PLN_Y_NYUKA_R.DATA_KB, vbUnicode) = DATA_KB(i) Then
                    PLN_Y_NYUKA(Row, colDATA_KB) = DATA_KB(i + 1) & "                    " & DATA_KB(i)
                    Exit For
                End If
            Next i
            
            
            
            PLN_Y_NYUKA(Row, colHIN_GAI) = Trim(StrConv(PLN_Y_NYUKA_R.HIN_GAI, vbUnicode))
            PLN_Y_NYUKA(Row, colHIN_NAI) = Trim(StrConv(PLN_Y_NYUKA_R.HIN_NAI, vbUnicode))
            
            
            If Trim(StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode)) <> "" Then        '2011.11.09
                PLN_Y_NYUKA(Row, colN_YOTEI_DT) = Mid(StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode), 1, 4) & "/" & Mid(StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode), 5, 2) & "/" & Mid(StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode), 7, 2)
            Else
                PLN_Y_NYUKA(Row, colN_YOTEI_DT) = ""
            End If
                
            If Trim(StrConv(PLN_Y_NYUKA_R.J_NYUKA_DT, vbUnicode)) <> "" Then
                PLN_Y_NYUKA(Row, colJ_NYUKA_DT) = Mid(StrConv(PLN_Y_NYUKA_R.J_NYUKA_DT, vbUnicode), 1, 4) & "/" & Mid(StrConv(PLN_Y_NYUKA_R.J_NYUKA_DT, vbUnicode), 5, 2) & "/" & Mid(StrConv(PLN_Y_NYUKA_R.J_NYUKA_DT, vbUnicode), 7, 2)
            Else
                PLN_Y_NYUKA(Row, colJ_NYUKA_DT) = ""
            End If
        
            PLN_Y_NYUKA(Row, colN_YOTEI_QTY) = Format(Val(StrConv(PLN_Y_NYUKA_R.N_YOTEI_QTY, vbUnicode)), "#")
            PLN_Y_NYUKA(Row, colJ_NYUKA_QTY) = Format(Val(StrConv(PLN_Y_NYUKA_R.J_NYUKA_QTY, vbUnicode)), "#")
        
        
            PLN_Y_NYUKA(Row, colSEQ_NO) = StrConv(PLN_Y_NYUKA_R.SEQ_NO, vbUnicode)
            
            PLN_Y_NYUKA(Row, colKEY_NO) = StrConv(PLN_Y_NYUKA_R.KEY_NO, vbUnicode)
        
        
        End If

        com = BtOpGetNext


    Loop


    Set TDBGrid1.Array = PLN_Y_NYUKA
    
    
    TDBGrid1.Bookmark = Null
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic
   
    
    If PLN_Y_NYUKA.Count(1) > 0 Then
        TDBGrid1.MoveFirst
        TDBGrid1.Col = colHIN_GAI
        TDBGrid1.Row = Min_Row
    End If




    lblDisp_Count.Caption = Format(Row, "#0") & "件"


    For i = 0 To UBound(Z_Sort_Tbl)
        Z_Sort_Tbl(i) = 0                 'ﾃﾞﾌｫﾙﾄ昇順
    Next i


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化用入荷予定ファイル　[検索]処理終了！！", Me.hwnd, 0)

    LIST_DISP_FLG = False


    Call Input_UnLock


    List_Disp_Proc = False


End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    Call Ctrl_Lock(PLN00201)

    TDBGrid1.Enabled = False
    
    DoEvents
    PLN00201.MousePointer = vbHourglass



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PLN00201)

    TDBGrid1.Enabled = True

    DoEvents
    PLN00201.MousePointer = vbDefault

End Sub


Private Sub Bu_Set_Proc()
'----------------------------------------------------------------------------
'                   画面項目（ＢＵ）のセット
'----------------------------------------------------------------------------
Dim i   As Integer




    Combo1(pcmbBU).Clear

    Combo1(pcmbBU).AddItem "全　て" & "          " & " "

    



    For i = 0 To UBound(JGYOBU_T)
            
        Combo1(pcmbBU).AddItem JGYOBU_T(i).NAME & "          " & JGYOBU_T(i).CODE
            
            
    Next i

    Combo1(pcmbBU).ListIndex = 0
End Sub


Private Sub Data_Kb_Set_Proc()
'----------------------------------------------------------------------------
'                   画面項目（データ区分）のセット
'----------------------------------------------------------------------------
Dim i   As Integer




    Combo1(pcmbDATA_KB).Clear

    Combo1(pcmbDATA_KB).AddItem "全　て" & "                    " & " "

    



    For i = 0 To UBound(DATA_KB) - 1 Step 2
            
        Combo1(pcmbDATA_KB).AddItem DATA_KB(i + 1) & "                    " & DATA_KB(i)
            
            
    Next i

    Combo1(pcmbDATA_KB).ListIndex = 0

End Sub

Private Sub TDB_BU_Set_Proc()
'----------------------------------------------------------------------------
'                   COMP項目（BU）のセット
'----------------------------------------------------------------------------

Dim i           As Integer
Dim j           As Integer
    
    
    Set TDB_BU = Nothing
    j = 0
    
    For i = 0 To UBound(JGYOBU_T)
            
        j = j + 1
        TDB_BU.ReDim 1, j, 0, 0
        
        
        TDB_BU(j, 0) = JGYOBU_T(i).NAME & "          " & JGYOBU_T(i).CODE
        
        
            
            
    Next i
    
    
    

    Set TDBDropDown1.Array = TDB_BU
    TDBDropDown1.ReBind

    



End Sub



Private Sub TDB_DATA_KB_Set_Proc()
'----------------------------------------------------------------------------
'                   COMP項目（データ区分）のセット
'----------------------------------------------------------------------------

Dim i           As Integer
Dim j           As Integer
    
    
    Set TDB_DATA_KB = Nothing
    j = 0
    
    For i = 0 To UBound(DATA_KB) - 1 Step 2
            
        j = j + 1
        TDB_DATA_KB.ReDim 1, j, 0, 0
        
        
        TDB_DATA_KB(j, 0) = DATA_KB(i + 1) & "                    " & DATA_KB(i)
        
        
            
            
    Next i
    
    
    

    Set TDBDropDown2.Array = TDB_DATA_KB
    TDBDropDown2.ReBind

    



End Sub




Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
Dim sts         As Integer
Dim Bookmark    As Variant


Dim i           As Integer

    If LIST_DISP_FLG Then
        Exit Sub
    End If




    If TDBGrid1.Bookmark = Null Then
        Exit Sub
    End If

    If TDBGrid1.Bookmark < 0 Then
        Exit Sub
    End If


    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.Update

    If PLN_Y_NYUKA(TDBGrid1.Bookmark, colSHORI) Then
    Else
        Select Case ColIndex


            Case colJGYOBU

                'BU
                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJGYOBU)) = "" Then
                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(ＢＵ　必須入力)"
                    Exit Sub
                End If


            Case colDATA_KB

                'ﾃﾞｰﾀ区分
                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colDATA_KB)) = "" Then
                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(データ区分　必須入力)"
                    Exit Sub
                End If

            Case colHIN_GAI

                '品番
                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI)) <> "" Then
                    
                    '2011.12.01 小文字-->大文字
                    PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI) = Trim(StrConv(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI), vbUpperCase))
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJGYOBU), 1))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI)) '

                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr

                            PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_NAI) = Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode))

Debug.Print "colHIN_GAI =" & TDBGrid1.Bookmark



                        Case BtErrKeyNotFound
'                            MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(品番未登録)"
'                            Exit Sub
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                            Unload Me
                    End Select
                End If

            Case colHIN_NAI
                '品番
                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI)) = "" And Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_NAI)) = "" Then
                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(品番　必須入力)"
                    Exit Sub
                End If

                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI)) = "" Then
                    
                    
                    '2011.12.01 小文字-->大文字
                    PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_NAI) = Trim(StrConv(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_NAI), vbUpperCase))
                    
                    Call UniCode_Conv(K2_ITEM.JGYOBU, Right(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJGYOBU), 1))
                    Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K2_ITEM.HIN_NAI, PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_NAI))

                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                    Select Case sts
                        Case BtNoErr

                            PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))


                        Case BtErrKeyNotFound
'                            MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(品番未登録)"
'                            Exit Sub

                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                            Unload Me
                    End Select
                End If

            Case colN_YOTEI_DT

                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colN_YOTEI_DT)) <> "" Then           '2011.11.09
                    If Not IsDate(PLN_Y_NYUKA(TDBGrid1.Bookmark, colN_YOTEI_DT)) Then
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(予定日)"
                        Exit Sub
                    End If
                End If

           Case colN_YOTEI_QTY

                If Not IsNumeric(PLN_Y_NYUKA(TDBGrid1.Bookmark, colN_YOTEI_QTY)) Then
                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(予定数)"
                    Exit Sub
                End If

                If Val(PLN_Y_NYUKA(TDBGrid1.Bookmark, colN_YOTEI_QTY)) < 1 Then
                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(予定数≦０)"
                    Exit Sub
                End If

            Case colJ_NYUKA_DT


                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_DT)) <> "" Then
                    If Not IsDate(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_DT)) Then
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(実績日)"
                        Exit Sub
                    End If
                End If


            Case colJ_NYUKA_QTY


                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_QTY)) <> "" Then

                    If Not IsDate(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_DT)) Then
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(実績日)"
                        Exit Sub
                    End If

                    If Not IsNumeric(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_QTY)) Then
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(実績数)"
                        Exit Sub
                    End If

                    If Val(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_QTY)) < 0 Then
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(予定数≦０)"
                        Exit Sub
                    End If


                End If
        End Select
    End If

    Set TDBGrid1.Array = PLN_Y_NYUKA


'    TDBGrid1.Refresh
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic




'    TDBGrid1.SetFocus

End Sub

Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)
    
    PLN_Y_NYUKA.ReDim Min_Row, PLN_Y_NYUKA.Count(1), Min_Col, Max_Col

    Command1(1).Enabled = True

End Sub


Private Sub TDBGrid1_ColEdit(ByVal ColIndex As Integer)
'Dim sts         As Integer
'Dim Bookmark    As Variant
'
'
'Dim i           As Integer
'
'    If LIST_DISP_FLG Then
'        Exit Sub
'    End If
'
'
'
'
'    If TDBGrid1.Bookmark = Null Then
'        Exit Sub
'    End If
'
'    If TDBGrid1.Bookmark < 0 Then
'        Exit Sub
'    End If
'
'
'    Set TDBGrid1.Array = PLN_Y_NYUKA
'    TDBGrid1.Update
'
'    If PLN_Y_NYUKA(TDBGrid1.Bookmark, colSHORI) Then
'    Else
'        Select Case ColIndex
'
'
'            Case colJGYOBU
'
'                'BU
'                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJGYOBU)) = "" Then
'                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(ＢＵ　必須入力)"
'                    Exit Sub
'                End If
'
'
'            Case colDATA_KB
'
'                'ﾃﾞｰﾀ区分
'                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colDATA_KB)) = "" Then
'                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(データ区分　必須入力)"
'                    Exit Sub
'                End If
'
'            Case colHIN_GAI
'
'                '品番
'                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI)) <> "" Then
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJGYOBU), 1))
'                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
'                    Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI))'
'
'                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                    Select Case sts
'                        Case BtNoErr
'
'                            PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_NAI) = Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
'
'Debug.Print "colHIN_GAI =" & TDBGrid1.Bookmark
'
'
'
'                        Case BtErrKeyNotFound
''                            MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(品番未登録)"
''                            Exit Sub
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                            Unload Me
'                    End Select
'                End If
'
'            Case colHIN_NAI
'                '品番
'                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI)) = "" And Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_NAI)) = "" Then
'                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(品番　必須入力)"
'                    Exit Sub
'                End If
'
'                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI)) = "" Then
'                    Call UniCode_Conv(K2_ITEM.JGYOBU, Right(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJGYOBU), 1))
'                    Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
'                    Call UniCode_Conv(K2_ITEM.HIN_NAI, PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_NAI))
'
'                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
'                    Select Case sts
'                        Case BtNoErr
'
'                            PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
'
'
'                        Case BtErrKeyNotFound
''                            MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(品番未登録)"
''                            Exit Sub
'
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                            Unload Me
'                    End Select
'                End If
'
'            Case colN_YOTEI_DT
'
'                If Not IsDate(PLN_Y_NYUKA(TDBGrid1.Bookmark, colN_YOTEI_DT)) Then
'                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(予定日)"
'                    Exit Sub
'                End If
'
'            Case colN_YOTEI_QTY
'
'                If Not IsNumeric(PLN_Y_NYUKA(TDBGrid1.Bookmark, colN_YOTEI_QTY)) Then
'                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(予定数)"
'                    Exit Sub
'                End If
'
'                If Val(PLN_Y_NYUKA(TDBGrid1.Bookmark, colN_YOTEI_QTY)) < 1 Then
'                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(予定数≦０)"
'                    Exit Sub
'                End If
'
'            Case colJ_NYUKA_DT
'
'
'                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_DT)) <> "" Then
'                    If Not IsDate(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_DT)) Then
'                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(実績日)"
'                        Exit Sub
'                    End If
'                End If
'
'
'            Case colJ_NYUKA_QTY
'
'
'                If Trim(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_QTY)) <> "" Then
'
'                    If Not IsDate(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_DT)) Then
'                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(実績日)"
'                        Exit Sub
'                    End If
'
'                    If Not IsNumeric(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_QTY)) Then
'                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(実績数)"
'                        Exit Sub
'                    End If
'
'                    If Val(PLN_Y_NYUKA(TDBGrid1.Bookmark, colJ_NYUKA_QTY)) < 1 Then
'                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(予定数≦０)"
'                        Exit Sub
'                    End If
'
'
'                End If
'        End Select
'    End If
'
'    Set TDBGrid1.Array = PLN_Y_NYUKA
'
'
''    TDBGrid1.Refresh
'    TDBGrid1.Update
'    TDBGrid1.ScrollBars = dbgAutomatic
'
'
'
'
''    TDBGrid1.SetFocus

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)

    If PLN_Y_NYUKA.Count(1) <= 0 Then
        Exit Sub
    End If
    
    If Z_Sort_Tbl(ColIndex) = 0 Then
        Z_Sort_Tbl(ColIndex) = 1
    Else
        If Z_Sort_Tbl(ColIndex) = 1 Then
            Z_Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Z_Sort_Tbl(ColIndex) = 0 Or Z_Sort_Tbl(ColIndex) = 1 Then
                    
        PLN_Y_NYUKA.QuickSort Min_Row, PLN_Y_NYUKA.UpperBound(1), ColIndex, Z_Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = PLN_Y_NYUKA
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If

End Sub


Private Function Delete_Proc() As Integer


Dim sts As Integer
Dim i   As Long


    Delete_Proc = True





    If PLN_Y_NYUKA.Count(1) <= 0 Then
        Delete_Proc = False
        Exit Function
    End If

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化用入荷予定ファイル　[全削除]処理開始！！", Me.hwnd, 0)



    For i = 1 To PLN_Y_NYUKA.Count(1)
    
        Call UniCode_Conv(K4_PLN_Y_NYUKA.KEY_NO, PLN_Y_NYUKA(i, colKEY_NO))
        
            sts = BTRV(BtOpGetEqual, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
            Select Case sts
                Case BtNoErr
                
                    sts = BTRV(BtOpDelete, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
                    Select Case sts
                        Case BtNoErr
                        
                        Case Else
                    
                            Call File_Error(sts, BtOpGetEqual, "商品化用入荷予定ファイル")
                            Exit Function
                    
                    End Select
                                
                                
                
                Case BtErrKeyNotFound
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "商品化用入荷予定ファイル")
                    Exit Function
            End Select
    
    
    Next i

    Call Input_UnLock
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化用入荷予定ファイル　[全削除]処理終了！！", Me.hwnd, 0)

    Delete_Proc = False

End Function

Private Function Error_Check_Proc() As Integer

Dim sts As Integer
Dim i   As Long



    Error_Check_Proc = True



    Set TDBGrid1.Array = PLN_Y_NYUKA
    
    TDBGrid1.Update


    If PLN_Y_NYUKA.Count(1) <= 0 Then
        Error_Check_Proc = False
        Exit Function
    End If




    For i = 1 To PLN_Y_NYUKA.Count(1)
        
        
        If Trim(PLN_Y_NYUKA(i, colJGYOBU)) = "" And _
            Trim(PLN_Y_NYUKA(i, colDATA_KB)) = "" And _
            Trim(PLN_Y_NYUKA(i, colHIN_GAI)) = "" And _
            Trim(PLN_Y_NYUKA(i, colHIN_NAI)) = "" And _
            Trim(PLN_Y_NYUKA(i, colN_YOTEI_DT)) = "" And _
            Trim(PLN_Y_NYUKA(i, colN_YOTEI_QTY)) = "" And _
            Trim(PLN_Y_NYUKA(i, colJ_NYUKA_DT)) = "" And _
            Trim(PLN_Y_NYUKA(i, colJ_NYUKA_QTY)) = "" Then
            Exit For
        End If
                    
            
        If PLN_Y_NYUKA(i, colSHORI) Then
        Else
            
            
            
            'BU
            If Trim(PLN_Y_NYUKA(i, colJGYOBU)) = "" Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(ＢＵ　必須入力)"
                TDBGrid1.SetFocus
                Exit Function
            End If
        
        
        
            'ﾃﾞｰﾀ区分
            If Trim(PLN_Y_NYUKA(i, colDATA_KB)) = "" Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(データ区分　必須入力)"
                TDBGrid1.SetFocus
                Exit Function
            End If
        
            '品番
            If Trim(PLN_Y_NYUKA(i, colHIN_GAI)) <> "" Then
                
                '2011.12.01 小文字-->大文字
                PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI) = Trim(StrConv(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_GAI), vbUpperCase))
                
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Right(PLN_Y_NYUKA(i, colJGYOBU), 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_Y_NYUKA(i, colHIN_GAI))
            
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                        PLN_Y_NYUKA(i, colHIN_NAI) = Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                    
                    
                    Case BtErrKeyNotFound
'                        MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(品番未登録)"
'                        TDBGrid1.SetFocus
'                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                        Unload Me
                End Select
            End If
        
            '品番
            If Trim(PLN_Y_NYUKA(i, colHIN_GAI)) = "" And Trim(PLN_Y_NYUKA(i, colHIN_NAI)) = "" Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(品番　必須入力)"
                TDBGrid1.SetFocus
                Exit Function
            End If
        
            If Trim(PLN_Y_NYUKA(i, colHIN_GAI)) = "" Then
                
                '2011.12.01 小文字-->大文字
                PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_NAI) = Trim(StrConv(PLN_Y_NYUKA(TDBGrid1.Bookmark, colHIN_NAI), vbUpperCase))
                
                
                
                Call UniCode_Conv(K2_ITEM.JGYOBU, Right(PLN_Y_NYUKA(i, colJGYOBU), 1))
                Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                Call UniCode_Conv(K2_ITEM.HIN_NAI, PLN_Y_NYUKA(i, colHIN_NAI))
            
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                Select Case sts
                    Case BtNoErr
                    
                        PLN_Y_NYUKA(i, colHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                    
                    
                    Case BtErrKeyNotFound
'                        MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(品番未登録)"
'                        TDBGrid1.SetFocus
'                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                        Unload Me
                End Select
            End If
        
                
            If Trim(PLN_Y_NYUKA(i, colN_YOTEI_DT)) <> "" Then           '2011.11.09
                If Not IsDate(PLN_Y_NYUKA(i, colN_YOTEI_DT)) Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(予定日)"
                    TDBGrid1.SetFocus
                    Exit Function
                End If
            End If
        
            If Not IsNumeric(PLN_Y_NYUKA(i, colN_YOTEI_QTY)) Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(予定数)"
                TDBGrid1.SetFocus
                Exit Function
            End If
    
    
            If CLng(PLN_Y_NYUKA(i, colN_YOTEI_QTY)) < 1 Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(予定数≦０)"
                TDBGrid1.SetFocus
                Exit Function
            End If
        
        
        
            If Trim(PLN_Y_NYUKA(i, colJ_NYUKA_DT)) <> "" Then
                If Not IsDate(PLN_Y_NYUKA(i, colJ_NYUKA_DT)) Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(実績日)"
                    TDBGrid1.SetFocus
                    Exit Function
                End If
            End If
        
        
        
        
            If Trim(PLN_Y_NYUKA(i, colJ_NYUKA_QTY)) <> "" Then
                
                If Not IsDate(PLN_Y_NYUKA(i, colJ_NYUKA_DT)) Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(実績日)"
                    TDBGrid1.SetFocus
                    Exit Function
                End If
                
                If Not IsNumeric(PLN_Y_NYUKA(i, colJ_NYUKA_QTY)) Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(実績数)"
                    TDBGrid1.SetFocus
                    Exit Function
                End If
            
                If CLng(PLN_Y_NYUKA(i, colJ_NYUKA_QTY)) < 1 Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(実績数≦０)"
                    TDBGrid1.SetFocus
                    Exit Function
                End If
    
    
            End If
        End If
    Next i

    Set TDBGrid1.Array = PLN_Y_NYUKA
        
    
    TDBGrid1.Refresh
    TDBGrid1.Update

    Error_Check_Proc = False

End Function

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim sts         As Integer
Dim Bookmark    As Variant


Dim i           As Integer

    If LIST_DISP_FLG Then
        Exit Sub
    End If




    If IsNull(LastRow) Then
        Exit Sub
    End If

    If LastRow < 0 Then
        Exit Sub
    End If


    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.Update

    If PLN_Y_NYUKA(LastRow, colSHORI) Then
    Else
        Select Case LastCol


            Case colJGYOBU

                'BU
                If Trim(PLN_Y_NYUKA(LastRow, colJGYOBU)) = "" Then
                    MsgBox "[" & Format(LastRow, "0") & "]行目 入力した項目はエラーです。(ＢＵ　必須入力)"
                    Exit Sub
                End If


            Case colDATA_KB

                'ﾃﾞｰﾀ区分
                If Trim(PLN_Y_NYUKA(LastRow, colDATA_KB)) = "" Then
                    MsgBox "[" & Format(LastRow, "0") & "]行目 入力した項目はエラーです。(データ区分　必須入力)"
                    Exit Sub
                End If

            Case colHIN_GAI

                '品番
                If Trim(PLN_Y_NYUKA(LastRow, colHIN_GAI)) <> "" Then
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(PLN_Y_NYUKA(LastRow, colJGYOBU), 1))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_Y_NYUKA(LastRow, colHIN_GAI))

                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr

                            PLN_Y_NYUKA(LastRow, colHIN_NAI) = Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode))

Debug.Print "colHIN_GAI =" & LastRow


      
                        Case BtErrKeyNotFound
'                            MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(品番未登録)"
'                            Exit Sub
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                            Unload Me
                    End Select
                End If

            Case colHIN_NAI
                '品番
                If Trim(PLN_Y_NYUKA(LastRow, colHIN_GAI)) = "" And Trim(PLN_Y_NYUKA(LastRow, colHIN_NAI)) = "" Then
                    MsgBox "[" & Format(LastRow, "0") & "]行目 入力した項目はエラーです。(品番　必須入力)"
                    Exit Sub
                End If

                If Trim(PLN_Y_NYUKA(LastRow, colHIN_GAI)) = "" Then
                    Call UniCode_Conv(K2_ITEM.JGYOBU, Right(PLN_Y_NYUKA(LastRow, colJGYOBU), 1))
                    Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K2_ITEM.HIN_NAI, PLN_Y_NYUKA(LastRow, colHIN_NAI))

                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                    Select Case sts
                        Case BtNoErr

                            PLN_Y_NYUKA(LastRow, colHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))


                        Case BtErrKeyNotFound
'                            MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(品番未登録)"
'                            Exit Sub

                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                            Unload Me
                    End Select
                End If

            Case colN_YOTEI_DT

                If Not IsDate(PLN_Y_NYUKA(LastRow, colN_YOTEI_DT)) Then
                    MsgBox "[" & Format(LastRow, "0") & "]行目 入力した項目はエラーです。(予定日)"
                    Exit Sub
                End If

            Case colN_YOTEI_QTY

                If Not IsNumeric(PLN_Y_NYUKA(LastRow, colN_YOTEI_QTY)) Then
                    MsgBox "[" & Format(LastRow, "0") & "]行目 入力した項目はエラーです。(予定数)"
                    Exit Sub
                End If


                If Val(PLN_Y_NYUKA(LastRow, colN_YOTEI_QTY)) < 1 Then
                    MsgBox "[" & Format(LastRow, "0") & "]行目 入力した項目はエラーです。(予定数≦０)"
                    Exit Sub
                End If

            Case colJ_NYUKA_DT


                If Trim(PLN_Y_NYUKA(LastRow, colJ_NYUKA_DT)) <> "" Then
                    If Not IsDate(PLN_Y_NYUKA(LastRow, colJ_NYUKA_DT)) Then
                        MsgBox "[" & Format(LastRow, "0") & "]行目 入力した項目はエラーです。(実績日)"
                        Exit Sub
                    End If
                End If


            Case colJ_NYUKA_QTY


                If Trim(PLN_Y_NYUKA(LastRow, colJ_NYUKA_QTY)) <> "" Then

                    If Not IsDate(PLN_Y_NYUKA(LastRow, colJ_NYUKA_DT)) Then
                        MsgBox "[" & Format(LastRow, "0") & "]行目 入力した項目はエラーです。(実績日)"
                        Exit Sub
                    End If

                    If Not IsNumeric(PLN_Y_NYUKA(LastRow, colJ_NYUKA_QTY)) Then
                        MsgBox "[" & Format(LastRow, "0") & "]行目 入力した項目はエラーです。(実績数)"
                        Exit Sub
                    End If

                    If Val(PLN_Y_NYUKA(LastRow, colJ_NYUKA_QTY)) < 1 Then
                        MsgBox "[" & Format(LastRow, "0") & "]行目 入力した項目はエラーです。(予定数≦０)"
                        Exit Sub
                    End If


                End If
        End Select
    End If

    Set TDBGrid1.Array = PLN_Y_NYUKA


    TDBGrid1.Refresh
    TDBGrid1.Update
'    TDBGrid1.ReBind
    TDBGrid1.ScrollBars = dbgAutomatic




'    TDBGrid1.SetFocus


End Sub

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
        
    Call Tab_Ctrl(Shift)        '移動

End Sub
