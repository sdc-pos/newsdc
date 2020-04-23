VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1011251 
   Caption         =   "要因マスタメンテナンス"
   ClientHeight    =   9540
   ClientLeft      =   2025
   ClientTop       =   -4470
   ClientWidth     =   15210
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
   ScaleHeight     =   9540
   ScaleWidth      =   15210
   StartUpPosition =   2  '画面の中央
   Begin TrueDBGrid80.TDBDropDown TDBDropDown4 
      Height          =   1215
      Left            =   10800
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
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
   Begin TrueDBGrid80.TDBDropDown TDBDropDown3 
      Height          =   1455
      Left            =   1800
      TabIndex        =   10
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2566
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
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
   Begin TrueDBGrid80.TDBDropDown TDBDropDown2 
      Height          =   1695
      Left            =   7800
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2990
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
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
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   1920
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   8
      Top             =   600
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   252
      Left            =   12840
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   372
   End
   Begin TrueDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   1695
      Left            =   4440
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2990
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   714
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
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "商品化予定データの検索/表示を行います"
      Top             =   0
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   6855
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   12091
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
      Columns(1).Caption=   "主要因"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "詳細要因"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "名称"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   1
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "集計先"
      Columns(4).DataField=   ""
      Columns(4).DropDown=   "TDBDropDown1"
      Columns(4).DropDown.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   1
      Columns(5)._MaxComboItems=   4
      Columns(5).Caption=   "使用機器"
      Columns(5).DataField=   ""
      Columns(5).DropDown=   "TDBDropDown2"
      Columns(5).DropDown.vt=   8
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   1
      Columns(6)._MaxComboItems=   3
      Columns(6).Caption=   "なし/向け先/倉庫"
      Columns(6).DataField=   ""
      Columns(6).DropDown=   "TDBDropDown3"
      Columns(6).DropDown.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "指定倉庫"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "倉庫名"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   1
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "特殊処理"
      Columns(9).DataField=   ""
      Columns(9).DropDown=   "TDBDropDown4"
      Columns(9).DropDown.vt=   8
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "入出庫登録画面表示順"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=820"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=688"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1561"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1429"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1588"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1455"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3519"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3387"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2381"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2249"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Button=1"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(4).DropDownList=1"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2593"
      Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(32)=   "Column(5).Button=1"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(5).DropDownList=1"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=2752"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2619"
      Splits(0)._ColumnProps(38)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(39)=   "Column(6).Button=1"
      Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(41)=   "Column(6).DropDownList=1"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=900"
      Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=767"
      Splits(0)._ColumnProps(45)=   "Column(7)._ColStyle=8704"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=2170"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2037"
      Splits(0)._ColumnProps(50)=   "Column(8)._ColStyle=8704"
      Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(52)=   "Column(9).Width=1879"
      Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=1746"
      Splits(0)._ColumnProps(55)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(56)=   "Column(9).Button=1"
      Splits(0)._ColumnProps(57)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(58)=   "Column(10).Width=1879"
      Splits(0)._ColumnProps(59)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(10)._WidthInPix=1746"
      Splits(0)._ColumnProps(61)=   "Column(10)._ColStyle=512"
      Splits(0)._ColumnProps(62)=   "Column(10).Order=11"
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
      HeadLines       =   4
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=1125"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H8000000D&"
      _StyleDefs(18)  =   ":id=6,.fgcolor=&H8000000E&"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=67,.parent=1,.bgcolor=&HFFFF00&,.bold=0,.fontsize=1125"
      _StyleDefs(26)  =   ":id=67,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(27)  =   ":id=67,.fontname=ＭＳ ゴシック"
      _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(30)  =   ":id=68,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(31)  =   ":id=68,.fontname=ＭＳ ゴシック"
      _StyleDefs(32)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(35)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(36)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(37)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(38)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(39)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(40)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=16,.parent=67,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=68"
      _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=69"
      _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=71"
      _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=54,.parent=67,.locked=-1,.bold=0,.fontsize=1125"
      _StyleDefs(46)  =   ":id=54,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(47)  =   ":id=54,.fontname=ＭＳ ゴシック"
      _StyleDefs(48)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(49)  =   ":id=51,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(50)  =   ":id=51,.fontname=ＭＳ ゴシック"
      _StyleDefs(51)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=69"
      _StyleDefs(52)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=71"
      _StyleDefs(53)  =   "Splits(0).Columns(2).Style:id=94,.parent=67,.alignment=3,.bold=0,.fontsize=1125"
      _StyleDefs(54)  =   ":id=94,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(55)  =   ":id=94,.fontname=ＭＳ ゴシック"
      _StyleDefs(56)  =   "Splits(0).Columns(2).HeadingStyle:id=91,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(57)  =   ":id=91,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(58)  =   ":id=91,.fontname=ＭＳ ゴシック"
      _StyleDefs(59)  =   "Splits(0).Columns(2).FooterStyle:id=92,.parent=69"
      _StyleDefs(60)  =   "Splits(0).Columns(2).EditorStyle:id=93,.parent=71"
      _StyleDefs(61)  =   "Splits(0).Columns(3).Style:id=24,.parent=67,.locked=0,.bold=0,.fontsize=1125"
      _StyleDefs(62)  =   ":id=24,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(63)  =   ":id=24,.fontname=ＭＳ ゴシック"
      _StyleDefs(64)  =   "Splits(0).Columns(3).HeadingStyle:id=21,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(65)  =   ":id=21,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(66)  =   ":id=21,.fontname=ＭＳ ゴシック"
      _StyleDefs(67)  =   "Splits(0).Columns(3).FooterStyle:id=22,.parent=69"
      _StyleDefs(68)  =   "Splits(0).Columns(3).EditorStyle:id=23,.parent=71"
      _StyleDefs(69)  =   "Splits(0).Columns(4).Style:id=106,.parent=67,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(70)  =   ":id=106,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(71)  =   ":id=106,.fontname=ＭＳ ゴシック"
      _StyleDefs(72)  =   "Splits(0).Columns(4).HeadingStyle:id=103,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(73)  =   ":id=103,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(74)  =   ":id=103,.fontname=ＭＳ ゴシック"
      _StyleDefs(75)  =   "Splits(0).Columns(4).FooterStyle:id=104,.parent=69"
      _StyleDefs(76)  =   "Splits(0).Columns(4).EditorStyle:id=105,.parent=71"
      _StyleDefs(77)  =   "Splits(0).Columns(5).Style:id=20,.parent=67"
      _StyleDefs(78)  =   "Splits(0).Columns(5).HeadingStyle:id=17,.parent=68"
      _StyleDefs(79)  =   "Splits(0).Columns(5).FooterStyle:id=18,.parent=69"
      _StyleDefs(80)  =   "Splits(0).Columns(5).EditorStyle:id=19,.parent=71"
      _StyleDefs(81)  =   "Splits(0).Columns(6).Style:id=28,.parent=67"
      _StyleDefs(82)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=68"
      _StyleDefs(83)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=69"
      _StyleDefs(84)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=71"
      _StyleDefs(85)  =   "Splits(0).Columns(7).Style:id=58,.parent=67,.alignment=0,.locked=-1,.bold=0"
      _StyleDefs(86)  =   ":id=58,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(87)  =   ":id=58,.fontname=ＭＳ ゴシック"
      _StyleDefs(88)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(89)  =   ":id=55,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(90)  =   ":id=55,.fontname=ＭＳ ゴシック"
      _StyleDefs(91)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=69"
      _StyleDefs(92)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=71"
      _StyleDefs(93)  =   "Splits(0).Columns(8).Style:id=62,.parent=67,.alignment=0,.locked=-1,.bold=0"
      _StyleDefs(94)  =   ":id=62,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(95)  =   ":id=62,.fontname=ＭＳ ゴシック"
      _StyleDefs(96)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(97)  =   ":id=59,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(98)  =   ":id=59,.fontname=ＭＳ ゴシック"
      _StyleDefs(99)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=69"
      _StyleDefs(100) =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=71"
      _StyleDefs(101) =   "Splits(0).Columns(9).Style:id=98,.parent=67,.alignment=3,.bold=0,.fontsize=1125"
      _StyleDefs(102) =   ":id=98,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(103) =   ":id=98,.fontname=ＭＳ ゴシック"
      _StyleDefs(104) =   "Splits(0).Columns(9).HeadingStyle:id=95,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(105) =   ":id=95,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(106) =   ":id=95,.fontname=ＭＳ ゴシック"
      _StyleDefs(107) =   "Splits(0).Columns(9).FooterStyle:id=96,.parent=69"
      _StyleDefs(108) =   "Splits(0).Columns(9).EditorStyle:id=97,.parent=71"
      _StyleDefs(109) =   "Splits(0).Columns(10).Style:id=102,.parent=67,.alignment=0,.locked=0,.bold=0"
      _StyleDefs(110) =   ":id=102,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(111) =   ":id=102,.fontname=ＭＳ ゴシック"
      _StyleDefs(112) =   "Splits(0).Columns(10).HeadingStyle:id=99,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(113) =   ":id=99,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(114) =   ":id=99,.fontname=ＭＳ ゴシック"
      _StyleDefs(115) =   "Splits(0).Columns(10).FooterStyle:id=100,.parent=69"
      _StyleDefs(116) =   "Splits(0).Columns(10).EditorStyle:id=101,.parent=71"
      _StyleDefs(117) =   "Named:id=33:Normal"
      _StyleDefs(118) =   ":id=33,.parent=0"
      _StyleDefs(119) =   "Named:id=34:Heading"
      _StyleDefs(120) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(121) =   ":id=34,.wraptext=-1"
      _StyleDefs(122) =   "Named:id=35:Footing"
      _StyleDefs(123) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(124) =   "Named:id=36:Selected"
      _StyleDefs(125) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(126) =   "Named:id=37:Caption"
      _StyleDefs(127) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(128) =   "Named:id=38:HighlightRow"
      _StyleDefs(129) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(130) =   "Named:id=39:EvenRow"
      _StyleDefs(131) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(132) =   "Named:id=40:OddRow"
      _StyleDefs(133) =   ":id=40,.parent=33"
      _StyleDefs(134) =   "Named:id=41:RecordSelector"
      _StyleDefs(135) =   ":id=41,.parent=34"
      _StyleDefs(136) =   "Named:id=42:FilterBar"
      _StyleDefs(137) =   ":id=42,.parent=33"
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
      Left            =   3720
      TabIndex        =   1
      ToolTipText     =   "商品化予定データメンテナンス画面を閉じる"
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
      Left            =   2040
      TabIndex        =   0
      ToolTipText     =   "商品化予定データの書き込みを行います"
      Top             =   0
      Width           =   1380
   End
   Begin VB.Label lblWeek 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   11880
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      Caption         =   "処理内容"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   1095
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
Attribute VB_Name = "F1011251"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim YOIN            As New XArrayDB

Dim TDB_SUM_KBN     As New XArrayDB     '集計先
Dim TDB_REGI_F      As New XArrayDB     '使用機器
Dim TDB_PARAM_F     As New XArrayDB     '無/向け先/倉庫
Dim TDB_SYSTEM_F    As New XArrayDB     '特殊処理



Private Type TBL_Tag
    CODE    As String * 1
    Title   As String
End Type



Dim SUM_KBN_CODE()  As TBL_Tag          '集計先
Dim REGI_F()        As TBL_Tag          '使用機器
Dim PARAM_F()       As TBL_Tag          '無/向け先/倉庫
Dim SYSTEM_F()      As TBL_Tag          '特殊処理



Private Const Min_Row% = 1              '最小行数
Private Max_Row    As Integer           'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 10             '最大列数



Private Const colDEL% = 0               '削除ﾌﾗｸﾞ
Private Const colCODE_TYPE% = 1         '主要因
Private Const colYOIN_CODE% = 2         '詳細要因
Private Const colYOIN_DNAME% = 3        '名称
Private Const colSUM_KBN% = 4           '集計先
Private Const colREGI_F% = 5            '使用機器
Private Const colPARAM_F% = 6           '無/向け先/倉庫
Private Const colSoko_No% = 7           '指定倉庫
Private Const colSoko_NAME% = 8         '倉庫名
Private Const colSYSTEM_F% = 9          '特殊処理
Private Const colDSP_No% = 10           '入出庫登録順

Private Type MENU_TBL_Tag
    CODE    As String * 1
    NAME    As String * 20
    TYPE    As String * 1
    YOIN    As String * 1
End Type

Private MENU_TBL()  As MENU_TBL_Tag


Private Z_Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順



Private Const LAST_UPDATE_DAY$ = "[F101125] 2018.04.07 10:45"

Private Sub Combo1_Click()

    If List_Disp_Proc() Then
        Unload Me
    End If


    If YOIN.Count(1) > 0 Then
        Command1(1).Enabled = True
    
    Else
        Command1(1).Enabled = False
        Command1(0).SetFocus
        Exit Sub
    End If

    Combo1.SetFocus
End Sub

Private Sub Command1_Click(Index As Integer)


Dim sWk             As String
Dim i               As Long
Dim j               As Long


Dim NYUKA_FLG       As Boolean

Dim SelBks          As SelBookmarks

Dim TOTAL_JIKAN     As Double


Dim yn              As Integer          '2012.04.24


    Select Case Index



        Case 0          '読込み
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
        
        
        Case 1          '登録



            If Err_Check_Proc() Then
                Exit Sub
            End If


            If Update_Proc() Then
                Unload Me
            End If




            If List_Disp_Proc() Then
                Unload Me
            End If


            If YOIN.Count(1) > 0 Then
                Command1(1).Enabled = True
            
            Else
                Command1(1).Enabled = False
                Command1(0).SetFocus
                Exit Sub
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
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()



Dim c           As String * 128
Dim sts         As Integer
Dim i           As Integer
Dim wkVariant   As Variant



    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "要因マスタメンテナンス", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)



    Show
                                'ログファイル名取り込み
    If GetIni(App.EXEName, "LOGF", App.EXEName, c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)

    
    F1011251.Caption = F1011251.Caption & " " & LAST_UPDATE_DAY

                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If


    TDBGrid1.Columns(colDEL).HeadBackColor = vbBlue
    TDBGrid1.Columns(colDEL).HeadForeColor = vbWhite


    TDBGrid1.Columns(colYOIN_CODE).HeadBackColor = vbBlue
    TDBGrid1.Columns(colYOIN_CODE).HeadForeColor = vbWhite

    TDBGrid1.Columns(colYOIN_DNAME).HeadBackColor = vbBlue
    TDBGrid1.Columns(colYOIN_DNAME).HeadForeColor = vbWhite


    TDBGrid1.Columns(colSUM_KBN).HeadBackColor = vbBlue
    TDBGrid1.Columns(colSUM_KBN).HeadForeColor = vbWhite

    TDBGrid1.Columns(colREGI_F).HeadBackColor = vbBlue
    TDBGrid1.Columns(colREGI_F).HeadForeColor = vbWhite

    TDBGrid1.Columns(colPARAM_F).HeadBackColor = vbBlue
    TDBGrid1.Columns(colPARAM_F).HeadForeColor = vbWhite

    TDBGrid1.Columns(colSoko_No).HeadBackColor = vbBlue
    TDBGrid1.Columns(colSoko_No).HeadForeColor = vbWhite

    TDBGrid1.Columns(colSYSTEM_F).HeadBackColor = vbBlue
    TDBGrid1.Columns(colSYSTEM_F).HeadForeColor = vbWhite

    TDBGrid1.Columns(colDSP_No).HeadBackColor = vbBlue
    TDBGrid1.Columns(colDSP_No).HeadForeColor = vbWhite


                                '作業情報取り込み
    '作業タイプ設定
    i = 0
    Do
        If GetIni("ACTION", "ACTION_CD" & Format(i + 1, "00"), "SYS", c) Then
            Exit Do
        End If
        If Trim(c) = "NON" Then
            Exit Do
        End If
    
        ReDim Preserve MENU_TBL(i)
        MENU_TBL(i).CODE = Trim(c)
        
        If GetIni("ACTION", "ACTION_NM" & Format(i + 1, "00"), "SYS", c) Then
            MsgBox "作業情報の獲得に失敗しました。" & "[ACTION]" & "[ACTION_NM" & Format(i, "00") & "]"
            Exit Do
        End If
        MENU_TBL(i).NAME = Trim(c)
        
        If GetIni("ACTION", "ACTION_TYPE" & Format(i + 1, "00"), "SYS", c) Then
            MsgBox "作業情報の獲得に失敗しました。" & "[ACTION]" & "[ACTION_TYPE" & Format(i, "00") & "]"
            Exit Do
        End If
        MENU_TBL(i).TYPE = Trim(c)
        
        If GetIni("ACTION", "ACTION_YOIN" & Format(i + 1, "00"), "SYS", c) Then
            MsgBox "作業情報の獲得に失敗しました。" & "[ACTION]" & "[ACTION_YOIN" & Format(i, "00") & "]"
            Exit Do
        End If
        MENU_TBL(i).YOIN = Trim(c)
        i = i + 1
    Loop


    Combo1.Clear
    For i = 0 To UBound(MENU_TBL)
        Combo1.AddItem MENU_TBL(i).NAME & "     " & MENU_TBL(i).CODE
    Next i
    Combo1.ListIndex = -1


    Call TDBDropDown_Set_Proc

End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因ﾏｽﾀ")
        End If
    End If
    
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(YOINREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫ﾏｽﾀ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    
    Set F1011251 = Nothing



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

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
Dim sts         As Integer
Dim Bookmark    As Variant
    
    
Dim i           As Integer
    
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
    
Dim KEISAN_FLG  As Boolean      '2011.11.30
    
    
    Set TDBGrid1.Array = YOIN
    TDBGrid1.Update
    
    
    
    
    If TDBGrid1.Bookmark = Null Then
        Exit Sub
    End If
    
    If TDBGrid1.Bookmark < 0 Then
        Exit Sub
    End If
    
    
    i = TDBGrid1.Bookmark
    
    If YOIN(TDBGrid1.Bookmark, colDEL) Then
    Else
        
        '主要因
        YOIN(i, colCODE_TYPE) = Right(Combo1.Text, 1)
        
        
        Select Case ColIndex
            
            '詳細要因
            Case colYOIN_CODE
        
                If Trim(YOIN(i, colYOIN_CODE)) = "" Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(詳細要因　必須入力)"
                    TDBGrid1.SetFocus
                    Exit Sub
                End If
            
            
            
            
            '名称
            Case colYOIN_DNAME
            
                If Trim(YOIN(i, colYOIN_DNAME)) = "" Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(詳細要因　必須入力)"
                    TDBGrid1.SetFocus
                    Exit Sub
                End If
            
            
            
            
            '集計先
            Case colSUM_KBN
                If Trim(YOIN(i, colSUM_KBN)) = "" Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(集計先　必須入力)"
                    TDBGrid1.SetFocus
                    Exit Sub
                End If
            
            '使用機器
            Case colREGI_F
                If Trim(YOIN(i, colREGI_F)) = "" Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(使用機器　必須入力)"
                    TDBGrid1.SetFocus
                    Exit Sub
                End If
            '無/向け先/倉庫
            Case colPARAM_F
            
                If Trim(YOIN(i, colPARAM_F)) = "" Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(無/向け先/倉庫　必須入力)"
                    TDBGrid1.SetFocus
                    Exit Sub
                End If
            
            
            
            
            '指定倉庫
            Case colSoko_NAME
                If Right(YOIN(i, colPARAM_F), 1) <> "2" Then
                    YOIN(i, colSoko_No) = ""
                    YOIN(i, colSoko_NAME) = ""
                Else
                    Call UniCode_Conv(K0_SOKO.Soko_No, YOIN(i, colSoko_No))
                
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                    Case BtNoErr
                        YOIN(i, colSoko_NAME) = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                    
                    
                        If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
                        Else
                            MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(指定倉庫 仮想倉庫を指定して下さい)"
                            TDBGrid1.SetFocus
                            Exit Sub
                        End If
                    
                    
                    
                    
                    Case BtErrKeyNotFound
                        MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(指定倉庫)"
                        TDBGrid1.SetFocus
                        Exit Sub
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                        Unload Me
                    
                    End Select
        
        
                End If
        
            '特殊処理
            Case colSYSTEM_F
                If Trim(YOIN(i, colSYSTEM_F)) = "" Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(特殊処理　必須入力)"
                    TDBGrid1.SetFocus
                    Exit Sub
                End If
        
        
        
        End Select
    End If
        
    Set TDBGrid1.Array = YOIN
        
    
    TDBGrid1.Refresh
    TDBGrid1.Update
    TDBGrid1.SetFocus
    




End Sub




Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)
    
    YOIN.ReDim Min_Row, YOIN.Count(1), Min_Col, Max_Col

    Command1(1).Enabled = True

End Sub



Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    If YOIN.Count(1) <= 0 Then
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
                    
        
        If TDBGrid1.Columns(ColIndex).Alignment = dbgRight Then
            YOIN.QuickSort Min_Row, YOIN.UpperBound(1), ColIndex, Z_Sort_Tbl(ColIndex), XTYPE_LONG
        Else
            YOIN.QuickSort Min_Row, YOIN.UpperBound(1), ColIndex, Z_Sort_Tbl(ColIndex), XTYPE_STRING
        End If
        
        Set TDBGrid1.Array = YOIN
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If

End Sub
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   「要因マスタ」登録処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim Row             As Long


Dim i               As Integer
Dim j               As Integer

Dim Upd_Com         As Integer



    If YOIN.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "要因マスタ登録処理　処理開始！！", Me.hwnd, 0)
                                    'テーブルリセット
    
    For Row = 1 To YOIN.UpperBound(1)
        DoEvents
        
        
               
        Call UniCode_Conv(K0_YOIN.CODE_TYPE, YOIN(Row, colCODE_TYPE))
        Call UniCode_Conv(K0_YOIN.YOIN_CODE, YOIN(Row, colYOIN_CODE))
        
        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
        Select Case sts
            Case BtNoErr
                Upd_Com = BtOpUpdate
            Case BtErrKeyNotFound
                Upd_Com = BtOpInsert
            Case Else
                Call File_Error(sts, BtOpGetEqual, "要因マスタ")
                Unload Me
        End Select
        
        
        
        
        
        If YOIN(Row, colDEL) Then
        
        
            If Upd_Com = BtOpUpdate Then
            
                sts = BTRV(BtOpDelete, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, Upd_Com, "要因マスタ")
                        Unload Me
                End Select
            End If
        
        
        
        Else
        
                        
            Call UniCode_Conv(YOINREC.CODE_TYPE, YOIN(Row, colCODE_TYPE))
            Call UniCode_Conv(YOINREC.YOIN_CODE, YOIN(Row, colYOIN_CODE))
            Call UniCode_Conv(YOINREC.YOIN_DNAME, YOIN(Row, colYOIN_DNAME))
            Call UniCode_Conv(YOINREC.SUM_KBN, Right(YOIN(Row, colSUM_KBN), 1))
            Call UniCode_Conv(YOINREC.SYSTEM_F, Right(YOIN(Row, colSYSTEM_F), 1))
            Call UniCode_Conv(YOINREC.REGI_F, Right(YOIN(Row, colREGI_F), 1))
            Call UniCode_Conv(YOINREC.PARAM_F, Right(YOIN(Row, colPARAM_F), 1))
            Call UniCode_Conv(YOINREC.Soko_No, YOIN(Row, colSoko_No))
            Call UniCode_Conv(YOINREC.DSP_No, YOIN(Row, colDSP_No))
            Call UniCode_Conv(YOINREC.FILLER, "")
                        
                        
                        
                        
                        
                        
            sts = BTRV(Upd_Com, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                    
                
                Case Else
                    Call File_Error(sts, Upd_Com, "要因マスタ")
                    Unload Me
            End Select
            

            Set TDBGrid1.Array = YOIN
            TDBGrid1.ReBind

            TDBGrid1.Update
            TDBGrid1.Bookmark = Row
        
        End If



    Next Row
    
    
    



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "要因マスタ処理　処理終了！！", Me.hwnd, 0)




    Call Input_UnLock



    Update_Proc = False



End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「要因マスタ」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer

Dim Row             As Long
Dim i               As Integer



    List_Disp_Proc = True

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "要因マスタ読込み処理　処理開始", Me.hwnd, 0)

                                    'テーブルリセット
    Set YOIN = Nothing
    Row = Min_Row - 1


    
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, Right(Combo1.Text, 1))
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, "")
    
    
    
    com = BtOpGetGreater



    Do
        DoEvents
        sts = BTRV(com, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "要因ﾏｽﾀ")
                Exit Function
        End Select
    
        If StrConv(YOINREC.CODE_TYPE, vbUnicode) <> Right(Combo1.Text, 1) Then
            Exit Do
        End If
            
            
            
        Row = Row + 1
        YOIN.ReDim Min_Row, Row, Min_Col, Max_Col
            
            
            
        YOIN(Row, colDEL) = False
        YOIN(Row, colCODE_TYPE) = StrConv(YOINREC.CODE_TYPE, vbUnicode)
        YOIN(Row, colYOIN_CODE) = StrConv(YOINREC.YOIN_CODE, vbUnicode)
        YOIN(Row, colYOIN_DNAME) = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
        
        For i = 1 To TDB_SUM_KBN.UpperBound(1)
            If Right(TDB_SUM_KBN(i, 1), 1) = StrConv(YOINREC.SUM_KBN, vbUnicode) Then
                YOIN(Row, colSUM_KBN) = TDB_SUM_KBN(i, 1)
                Exit For
            End If
        Next i
            
        For i = 1 To TDB_REGI_F.UpperBound(1)
            If Right(TDB_REGI_F(i, 1), 1) = StrConv(YOINREC.REGI_F, vbUnicode) Then
                YOIN(Row, colREGI_F) = TDB_REGI_F(i, 1)
                Exit For
            End If
        Next i
            
        For i = 1 To TDB_PARAM_F.UpperBound(1)
            If Right(TDB_PARAM_F(i, 1), 1) = StrConv(YOINREC.PARAM_F, vbUnicode) Then
                YOIN(Row, colPARAM_F) = TDB_PARAM_F(i, 1)
                Exit For
            End If
        Next i
            
        If StrConv(YOINREC.PARAM_F, vbUnicode) = "2" Then
            YOIN(Row, colSoko_No) = StrConv(YOINREC.Soko_No, vbUnicode)
            Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(YOINREC.Soko_No, vbUnicode))
            
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                    YOIN(Row, colSoko_NAME) = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                    
                Case BtErrKeyNotFound
                    YOIN(Row, colSoko_NAME) = ""
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "倉庫ﾏｽﾀ")
                    Exit Function
            End Select
 
        Else
            YOIN(Row, colSoko_No) = ""
            YOIN(Row, colSoko_NAME) = ""
        End If
 
 
        For i = 1 To TDB_SYSTEM_F.UpperBound(1)
            If Right(TDB_SYSTEM_F(i, 1), 1) = StrConv(YOINREC.SYSTEM_F, vbUnicode) Then
                YOIN(Row, colSYSTEM_F) = TDB_SYSTEM_F(i, 1)
                Exit For
            End If
        Next i
 
        YOIN(Row, colDSP_No) = StrConv(YOINREC.DSP_No, vbUnicode)
 
 
        com = BtOpGetNext


    Loop


    Set TDBGrid1.Array = YOIN
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



    For i = 0 To UBound(Z_Sort_Tbl)
        Z_Sort_Tbl(i) = 0                 'ﾃﾞﾌｫﾙﾄ昇順
    Next i







hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "要因マスタ読込み処理　処理終了", Me.hwnd, 0)




    Call Input_UnLock


    List_Disp_Proc = False
    Exit Function

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    F1011251.MousePointer = vbHourglass

    Call Ctrl_Lock(F1011251)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(F1011251)


    F1011251.MousePointer = vbDefault

End Sub

Private Sub TDBDropDown_Set_Proc()
'----------------------------------------------------------------------------
'                   COMP項目（BU）のセット
'----------------------------------------------------------------------------

Dim i           As Integer
    
    
    
    
    
    
    Set TDB_SUM_KBN = Nothing
        
    
    TDB_SUM_KBN.ReDim 1, 5, 1, 1
    TDB_SUM_KBN(1, 1) = "入　庫" & Space(10) & "1"
    TDB_SUM_KBN(2, 1) = "出　庫" & Space(10) & "2"
    TDB_SUM_KBN(3, 1) = "在訂±" & Space(10) & "3"
    TDB_SUM_KBN(4, 1) = "移　動" & Space(10) & "4"
    TDB_SUM_KBN(5, 1) = "な　し" & Space(10) & "0"
    Set TDBDropDown1.Array = TDB_SUM_KBN
    TDBDropDown1.ReBind


    Set TDB_REGI_F = Nothing
    TDB_REGI_F.ReDim 1, 4, 1, 1
    TDB_REGI_F(1, 1) = "両方可" & Space(10) & "0"
    TDB_REGI_F(2, 1) = "ｽｷｬﾅのみ" & Space(10) & "1"
    TDB_REGI_F(3, 1) = "画面のみ" & Space(10) & "2"
    TDB_REGI_F(4, 1) = "両方不可" & Space(10) & "9"
    Set TDBDropDown2.Array = TDB_REGI_F
    TDBDropDown2.ReBind

    Set TDB_PARAM_F = Nothing
    TDB_PARAM_F.ReDim 1, 3, 1, 1
    TDB_PARAM_F(1, 1) = "なし　" & Space(10) & "0"
    TDB_PARAM_F(2, 1) = "向け先" & Space(10) & "1"
    TDB_PARAM_F(3, 1) = "倉庫　" & Space(10) & "2"
    Set TDBDropDown3.Array = TDB_PARAM_F
    TDBDropDown3.ReBind


    Set TDB_SYSTEM_F = Nothing
    TDB_SYSTEM_F.ReDim 1, 2, 1, 1
    TDB_SYSTEM_F(1, 1) = "一般" & Space(10) & "0"
    TDB_SYSTEM_F(2, 1) = "ｼｽﾃﾑ" & Space(10) & "1"
    Set TDBDropDown4.Array = TDB_SYSTEM_F
    TDBDropDown4.ReBind

    



End Sub




Private Function Err_Check_Proc() As Integer
Dim sts         As Integer
Dim Bookmark    As Variant
    
    
Dim i           As Integer
    
    
    
    Err_Check_Proc = True
    
    Set TDBGrid1.Array = YOIN
    TDBGrid1.Update
    
    
    
    For i = 1 To YOIN.UpperBound(1)
        
        If YOIN(i, colDEL) Then
        Else
            
            
                        
            '主要因
            YOIN(i, colCODE_TYPE) = Right(Combo1.Text, 1)
            '詳細要因
            If Trim(YOIN(i, colYOIN_CODE)) = "" Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(詳細要因　必須入力)"
                TDBGrid1.SetFocus
                Exit Function
            End If
            '名称
            If Trim(YOIN(i, colYOIN_DNAME)) = "" Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(詳細要因　必須入力)"
                TDBGrid1.SetFocus
                Exit Function
            End If
            '集計先
            If Trim(YOIN(i, colSUM_KBN)) = "" Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(集計先　必須入力)"
                TDBGrid1.SetFocus
                Exit Function
            End If
            '使用機器
            If Trim(YOIN(i, colREGI_F)) = "" Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(使用機器　必須入力)"
                TDBGrid1.SetFocus
                Exit Function
            End If
            '無/向け先/倉庫
            If Trim(YOIN(i, colPARAM_F)) = "" Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(無/向け先/倉庫　必須入力)"
                TDBGrid1.SetFocus
                Exit Function
            End If
            
            
            
            
            
            
            
            '指定倉庫
            If Right(YOIN(i, colPARAM_F), 1) <> "2" Then
                YOIN(i, colSoko_No) = ""
                YOIN(i, colSoko_NAME) = ""
            Else
                Call UniCode_Conv(K0_SOKO.Soko_No, YOIN(i, colSoko_No))
            
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        YOIN(i, colSoko_NAME) = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                    
                        If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
                        Else
                            MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(指定倉庫 仮想倉庫を指定して下さい)"
                            TDBGrid1.SetFocus
                            Exit Function
                        End If
                        
                    Case BtErrKeyNotFound
                        MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(指定倉庫)"
                        TDBGrid1.SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                        Unload Me
                
                End Select
            End If
    
    
            '特殊処理
            If Trim(YOIN(i, colSYSTEM_F)) = "" Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(特殊処理　必須入力)"
                TDBGrid1.SetFocus
                Exit Function
            End If
    
    
    
        End If
    
    Next i
        
    Set TDBGrid1.Array = YOIN
        
    
    TDBGrid1.Refresh
    TDBGrid1.Update
    TDBGrid1.SetFocus

    Err_Check_Proc = False

End Function

