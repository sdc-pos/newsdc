VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PLN00401 
   Caption         =   "[商品化計画システム]商品化予定データメンテナンス"
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
   Begin VB.CommandButton Command1 
      Caption         =   "ﾃﾞｰﾀ出力"
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
      Index           =   6
      Left            =   11400
      TabIndex        =   30
      ToolTipText     =   "商品化予定データメンテナンス画面を閉じる"
      Top             =   0
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Index           =   7
      Left            =   2760
      TabIndex        =   28
      Top             =   1440
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   252
      Left            =   12840
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton Command1 
      Caption         =   "画面印刷"
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
      Index           =   5
      Left            =   13080
      TabIndex        =   26
      ToolTipText     =   "商品化予定データメンテナンス画面を閉じる"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "解除"
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
      Left            =   10800
      TabIndex        =   25
      ToolTipText     =   "選択された行を解除します"
      Top             =   1440
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "商品化予定工数集計"
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
      Left            =   7440
      TabIndex        =   23
      ToolTipText     =   "選択された行の工数集計を行います"
      Top             =   1440
      Width           =   2340
   End
   Begin TrueDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2055
      Left            =   5160
      TabIndex        =   22
      Top             =   4200
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
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Index           =   6
      Left            =   13440
      TabIndex        =   17
      Top             =   960
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Index           =   5
      Left            =   11640
      TabIndex        =   16
      Top             =   960
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Index           =   4
      Left            =   8400
      TabIndex        =   14
      Top             =   960
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Index           =   3
      Left            =   6600
      TabIndex        =   12
      Top             =   960
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Index           =   2
      Left            =   3960
      TabIndex        =   9
      Top             =   960
      Width           =   492
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   492
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   372
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
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   12091
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "削除"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   1
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "BU"
      Columns(1).DataField=   ""
      Columns(1).DropDown=   "TDBDropDown1"
      Columns(1).DropDown.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "対外品番"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "標準棚番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "個装  資材"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "事前商品化状況％"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "事前商品化必要数"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "商品化予定数"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "在庫数 済"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "在庫数 未"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "商品化            予定日"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "商品  化工数"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "商品化   予定  工数"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "部品  　    入荷予定日"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "部品入荷予定数"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "指図票          　発行日"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "完了日"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "KEY_NO"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "入荷予定KEY_NO"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   21
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=21"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=476"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=344"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=926"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=794"
      Splits(0)._ColumnProps(8)=   "Column(1).Button=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=2672"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2540"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2514"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2381"
      Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=8196"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=1005"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=873"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=847"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=714"
      Splits(0)._ColumnProps(26)=   "Column(5)._ColStyle=8194"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=847"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=714"
      Splits(0)._ColumnProps(31)=   "Column(6)._ColStyle=8194"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=1244"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=1111"
      Splits(0)._ColumnProps(36)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(38)=   "Column(8).Width=1244"
      Splits(0)._ColumnProps(39)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(8)._WidthInPix=1111"
      Splits(0)._ColumnProps(41)=   "Column(8)._ColStyle=8194"
      Splits(0)._ColumnProps(42)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(43)=   "Column(9).Width=1244"
      Splits(0)._ColumnProps(44)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(9)._WidthInPix=1111"
      Splits(0)._ColumnProps(46)=   "Column(9)._ColStyle=8194"
      Splits(0)._ColumnProps(47)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(48)=   "Column(10).Width=2249"
      Splits(0)._ColumnProps(49)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(10)._WidthInPix=2117"
      Splits(0)._ColumnProps(51)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(52)=   "Column(11).Width=926"
      Splits(0)._ColumnProps(53)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(11)._WidthInPix=794"
      Splits(0)._ColumnProps(55)=   "Column(11)._ColStyle=8194"
      Splits(0)._ColumnProps(56)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(57)=   "Column(12).Width=1376"
      Splits(0)._ColumnProps(58)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(12)._WidthInPix=1244"
      Splits(0)._ColumnProps(60)=   "Column(12)._ColStyle=8194"
      Splits(0)._ColumnProps(61)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(62)=   "Column(13).Width=2249"
      Splits(0)._ColumnProps(63)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(13)._WidthInPix=2117"
      Splits(0)._ColumnProps(65)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(66)=   "Column(14).Width=1217"
      Splits(0)._ColumnProps(67)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(14)._WidthInPix=1085"
      Splits(0)._ColumnProps(69)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(70)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(71)=   "Column(15).Width=2249"
      Splits(0)._ColumnProps(72)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(15)._WidthInPix=2117"
      Splits(0)._ColumnProps(74)=   "Column(15)._ColStyle=8196"
      Splits(0)._ColumnProps(75)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(76)=   "Column(16).Width=2249"
      Splits(0)._ColumnProps(77)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(16)._WidthInPix=2117"
      Splits(0)._ColumnProps(79)=   "Column(16)._ColStyle=8196"
      Splits(0)._ColumnProps(80)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(81)=   "Column(17).Width=1852"
      Splits(0)._ColumnProps(82)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(17)._WidthInPix=1720"
      Splits(0)._ColumnProps(84)=   "Column(17)._ColStyle=8196"
      Splits(0)._ColumnProps(85)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(86)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(87)=   "Column(18).Width=3519"
      Splits(0)._ColumnProps(88)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(89)=   "Column(18)._WidthInPix=3387"
      Splits(0)._ColumnProps(90)=   "Column(18)._ColStyle=8196"
      Splits(0)._ColumnProps(91)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(92)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(93)=   "Column(19).Width=3519"
      Splits(0)._ColumnProps(94)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(19)._WidthInPix=3387"
      Splits(0)._ColumnProps(96)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(97)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(98)=   "Column(20).Width=3519"
      Splits(0)._ColumnProps(99)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(100)=   "Column(20)._WidthInPix=3387"
      Splits(0)._ColumnProps(101)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(102)=   "Column(20).Order=21"
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=16,.parent=67"
      _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(43)  =   ":id=13,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(44)  =   ":id=13,.fontname=ＭＳ ゴシック"
      _StyleDefs(45)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=69"
      _StyleDefs(46)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=71"
      _StyleDefs(47)  =   "Splits(0).Columns(1).Style:id=54,.parent=67,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(48)  =   ":id=54,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(49)  =   ":id=54,.fontname=ＭＳ ゴシック"
      _StyleDefs(50)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(51)  =   ":id=51,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(52)  =   ":id=51,.fontname=ＭＳ ゴシック"
      _StyleDefs(53)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=69"
      _StyleDefs(54)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=71"
      _StyleDefs(55)  =   "Splits(0).Columns(2).Style:id=94,.parent=67,.alignment=3,.bold=0,.fontsize=1125"
      _StyleDefs(56)  =   ":id=94,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(57)  =   ":id=94,.fontname=ＭＳ ゴシック"
      _StyleDefs(58)  =   "Splits(0).Columns(2).HeadingStyle:id=91,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(59)  =   ":id=91,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(60)  =   ":id=91,.fontname=ＭＳ ゴシック"
      _StyleDefs(61)  =   "Splits(0).Columns(2).FooterStyle:id=92,.parent=69"
      _StyleDefs(62)  =   "Splits(0).Columns(2).EditorStyle:id=93,.parent=71"
      _StyleDefs(63)  =   "Splits(0).Columns(3).Style:id=24,.parent=67,.locked=-1,.bold=0,.fontsize=1125"
      _StyleDefs(64)  =   ":id=24,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(65)  =   ":id=24,.fontname=ＭＳ ゴシック"
      _StyleDefs(66)  =   "Splits(0).Columns(3).HeadingStyle:id=21,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(67)  =   ":id=21,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(68)  =   ":id=21,.fontname=ＭＳ ゴシック"
      _StyleDefs(69)  =   "Splits(0).Columns(3).FooterStyle:id=22,.parent=69"
      _StyleDefs(70)  =   "Splits(0).Columns(3).EditorStyle:id=23,.parent=71"
      _StyleDefs(71)  =   "Splits(0).Columns(4).Style:id=106,.parent=67,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(72)  =   ":id=106,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(73)  =   ":id=106,.fontname=ＭＳ ゴシック"
      _StyleDefs(74)  =   "Splits(0).Columns(4).HeadingStyle:id=103,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(75)  =   ":id=103,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(76)  =   ":id=103,.fontname=ＭＳ ゴシック"
      _StyleDefs(77)  =   "Splits(0).Columns(4).FooterStyle:id=104,.parent=69"
      _StyleDefs(78)  =   "Splits(0).Columns(4).EditorStyle:id=105,.parent=71"
      _StyleDefs(79)  =   "Splits(0).Columns(5).Style:id=58,.parent=67,.alignment=1,.locked=-1,.bold=0"
      _StyleDefs(80)  =   ":id=58,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(81)  =   ":id=58,.fontname=ＭＳ ゴシック"
      _StyleDefs(82)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(83)  =   ":id=55,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(84)  =   ":id=55,.fontname=ＭＳ ゴシック"
      _StyleDefs(85)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=69"
      _StyleDefs(86)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=71"
      _StyleDefs(87)  =   "Splits(0).Columns(6).Style:id=62,.parent=67,.alignment=1,.locked=-1,.bold=0"
      _StyleDefs(88)  =   ":id=62,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(89)  =   ":id=62,.fontname=ＭＳ ゴシック"
      _StyleDefs(90)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(91)  =   ":id=59,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(92)  =   ":id=59,.fontname=ＭＳ ゴシック"
      _StyleDefs(93)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=69"
      _StyleDefs(94)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=71"
      _StyleDefs(95)  =   "Splits(0).Columns(7).Style:id=98,.parent=67,.alignment=1,.bold=0,.fontsize=1125"
      _StyleDefs(96)  =   ":id=98,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(97)  =   ":id=98,.fontname=ＭＳ ゴシック"
      _StyleDefs(98)  =   "Splits(0).Columns(7).HeadingStyle:id=95,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(99)  =   ":id=95,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(100) =   ":id=95,.fontname=ＭＳ ゴシック"
      _StyleDefs(101) =   "Splits(0).Columns(7).FooterStyle:id=96,.parent=69"
      _StyleDefs(102) =   "Splits(0).Columns(7).EditorStyle:id=97,.parent=71"
      _StyleDefs(103) =   "Splits(0).Columns(8).Style:id=102,.parent=67,.alignment=1,.locked=-1,.bold=0"
      _StyleDefs(104) =   ":id=102,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(105) =   ":id=102,.fontname=ＭＳ ゴシック"
      _StyleDefs(106) =   "Splits(0).Columns(8).HeadingStyle:id=99,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(107) =   ":id=99,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(108) =   ":id=99,.fontname=ＭＳ ゴシック"
      _StyleDefs(109) =   "Splits(0).Columns(8).FooterStyle:id=100,.parent=69"
      _StyleDefs(110) =   "Splits(0).Columns(8).EditorStyle:id=101,.parent=71"
      _StyleDefs(111) =   "Splits(0).Columns(9).Style:id=114,.parent=67,.alignment=1,.locked=-1,.bold=0"
      _StyleDefs(112) =   ":id=114,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(113) =   ":id=114,.fontname=ＭＳ ゴシック"
      _StyleDefs(114) =   "Splits(0).Columns(9).HeadingStyle:id=111,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(115) =   ":id=111,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(116) =   ":id=111,.fontname=ＭＳ ゴシック"
      _StyleDefs(117) =   "Splits(0).Columns(9).FooterStyle:id=112,.parent=69"
      _StyleDefs(118) =   "Splits(0).Columns(9).EditorStyle:id=113,.parent=71"
      _StyleDefs(119) =   "Splits(0).Columns(10).Style:id=20,.parent=67,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(120) =   ":id=20,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(121) =   ":id=20,.fontname=ＭＳ ゴシック"
      _StyleDefs(122) =   "Splits(0).Columns(10).HeadingStyle:id=17,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(123) =   ":id=17,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(124) =   ":id=17,.fontname=ＭＳ ゴシック"
      _StyleDefs(125) =   "Splits(0).Columns(10).FooterStyle:id=18,.parent=69"
      _StyleDefs(126) =   "Splits(0).Columns(10).EditorStyle:id=19,.parent=71"
      _StyleDefs(127) =   "Splits(0).Columns(11).Style:id=66,.parent=67,.alignment=1,.locked=-1,.bold=0"
      _StyleDefs(128) =   ":id=66,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(129) =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(130) =   "Splits(0).Columns(11).HeadingStyle:id=63,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(131) =   ":id=63,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(132) =   ":id=63,.fontname=ＭＳ ゴシック"
      _StyleDefs(133) =   "Splits(0).Columns(11).FooterStyle:id=64,.parent=69"
      _StyleDefs(134) =   "Splits(0).Columns(11).EditorStyle:id=65,.parent=71"
      _StyleDefs(135) =   "Splits(0).Columns(12).Style:id=82,.parent=67,.alignment=1,.locked=-1,.bold=0"
      _StyleDefs(136) =   ":id=82,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(137) =   ":id=82,.fontname=ＭＳ ゴシック"
      _StyleDefs(138) =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(139) =   ":id=79,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(140) =   ":id=79,.fontname=ＭＳ ゴシック"
      _StyleDefs(141) =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=69"
      _StyleDefs(142) =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=71"
      _StyleDefs(143) =   "Splits(0).Columns(13).Style:id=28,.parent=67,.locked=0,.bold=0,.fontsize=1125"
      _StyleDefs(144) =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(145) =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(146) =   "Splits(0).Columns(13).HeadingStyle:id=25,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(147) =   ":id=25,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(148) =   ":id=25,.fontname=ＭＳ ゴシック"
      _StyleDefs(149) =   "Splits(0).Columns(13).FooterStyle:id=26,.parent=69"
      _StyleDefs(150) =   "Splits(0).Columns(13).EditorStyle:id=27,.parent=71"
      _StyleDefs(151) =   "Splits(0).Columns(14).Style:id=32,.parent=67,.alignment=1,.locked=0,.bold=0"
      _StyleDefs(152) =   ":id=32,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(153) =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(154) =   "Splits(0).Columns(14).HeadingStyle:id=29,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(155) =   ":id=29,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(156) =   ":id=29,.fontname=ＭＳ ゴシック"
      _StyleDefs(157) =   "Splits(0).Columns(14).FooterStyle:id=30,.parent=69"
      _StyleDefs(158) =   "Splits(0).Columns(14).EditorStyle:id=31,.parent=71"
      _StyleDefs(159) =   "Splits(0).Columns(15).Style:id=46,.parent=67,.locked=-1,.bold=0,.fontsize=1125"
      _StyleDefs(160) =   ":id=46,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(161) =   ":id=46,.fontname=ＭＳ ゴシック"
      _StyleDefs(162) =   "Splits(0).Columns(15).HeadingStyle:id=43,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(163) =   ":id=43,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(164) =   ":id=43,.fontname=ＭＳ ゴシック"
      _StyleDefs(165) =   "Splits(0).Columns(15).FooterStyle:id=44,.parent=69"
      _StyleDefs(166) =   "Splits(0).Columns(15).EditorStyle:id=45,.parent=71"
      _StyleDefs(167) =   "Splits(0).Columns(16).Style:id=50,.parent=67,.locked=-1,.bold=0,.fontsize=1125"
      _StyleDefs(168) =   ":id=50,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(169) =   ":id=50,.fontname=ＭＳ ゴシック"
      _StyleDefs(170) =   "Splits(0).Columns(16).HeadingStyle:id=47,.parent=68,.bold=0,.fontsize=900"
      _StyleDefs(171) =   ":id=47,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(172) =   ":id=47,.fontname=ＭＳ ゴシック"
      _StyleDefs(173) =   "Splits(0).Columns(16).FooterStyle:id=48,.parent=69"
      _StyleDefs(174) =   "Splits(0).Columns(16).EditorStyle:id=49,.parent=71"
      _StyleDefs(175) =   "Splits(0).Columns(17).Style:id=86,.parent=67,.locked=-1"
      _StyleDefs(176) =   "Splits(0).Columns(17).HeadingStyle:id=83,.parent=68"
      _StyleDefs(177) =   "Splits(0).Columns(17).FooterStyle:id=84,.parent=69"
      _StyleDefs(178) =   "Splits(0).Columns(17).EditorStyle:id=85,.parent=71"
      _StyleDefs(179) =   "Splits(0).Columns(18).Style:id=90,.parent=67,.locked=-1"
      _StyleDefs(180) =   "Splits(0).Columns(18).HeadingStyle:id=87,.parent=68"
      _StyleDefs(181) =   "Splits(0).Columns(18).FooterStyle:id=88,.parent=69"
      _StyleDefs(182) =   "Splits(0).Columns(18).EditorStyle:id=89,.parent=71"
      _StyleDefs(183) =   "Splits(0).Columns(19).Style:id=110,.parent=67"
      _StyleDefs(184) =   "Splits(0).Columns(19).HeadingStyle:id=107,.parent=68"
      _StyleDefs(185) =   "Splits(0).Columns(19).FooterStyle:id=108,.parent=69"
      _StyleDefs(186) =   "Splits(0).Columns(19).EditorStyle:id=109,.parent=71"
      _StyleDefs(187) =   "Splits(0).Columns(20).Style:id=118,.parent=67"
      _StyleDefs(188) =   "Splits(0).Columns(20).HeadingStyle:id=115,.parent=68"
      _StyleDefs(189) =   "Splits(0).Columns(20).FooterStyle:id=116,.parent=69"
      _StyleDefs(190) =   "Splits(0).Columns(20).EditorStyle:id=117,.parent=71"
      _StyleDefs(191) =   "Named:id=33:Normal"
      _StyleDefs(192) =   ":id=33,.parent=0"
      _StyleDefs(193) =   "Named:id=34:Heading"
      _StyleDefs(194) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(195) =   ":id=34,.wraptext=-1"
      _StyleDefs(196) =   "Named:id=35:Footing"
      _StyleDefs(197) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(198) =   "Named:id=36:Selected"
      _StyleDefs(199) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(200) =   "Named:id=37:Caption"
      _StyleDefs(201) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(202) =   "Named:id=38:HighlightRow"
      _StyleDefs(203) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(204) =   "Named:id=39:EvenRow"
      _StyleDefs(205) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(206) =   "Named:id=40:OddRow"
      _StyleDefs(207) =   ":id=40,.parent=33"
      _StyleDefs(208) =   "Named:id=41:RecordSelector"
      _StyleDefs(209) =   ":id=41,.parent=34"
      _StyleDefs(210) =   "Named:id=42:FilterBar"
      _StyleDefs(211) =   ":id=42,.parent=33"
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
   Begin VB.Label Label1 
      Caption         =   "対外品番"
      Height          =   255
      Index           =   7
      Left            =   1560
      TabIndex        =   29
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblS_S_JIKAN 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9840
      TabIndex        =   24
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "表示件数"
      Height          =   255
      Index           =   6
      Left            =   12840
      TabIndex        =   21
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblDisp_Count 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   13920
      TabIndex        =   20
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblWeek 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   11880
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      Caption         =   "～"
      Height          =   255
      Index           =   9
      Left            =   13080
      TabIndex        =   18
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "部品入荷予定日"
      Height          =   255
      Index           =   8
      Left            =   9960
      TabIndex        =   15
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "～"
      Height          =   255
      Index           =   5
      Left            =   8040
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "商品化予定日"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "％"
      Height          =   252
      Index           =   3
      Left            =   4560
      TabIndex        =   10
      Top             =   1080
      Width           =   252
   End
   Begin VB.Label Label1 
      Caption         =   "％～"
      Height          =   252
      Index           =   2
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "事前商品化状況"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "標準棚番（倉庫№）"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   2175
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
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   3
      End
   End
End
Attribute VB_Name = "PLN00401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim PLN_S_YOTEI         As New XArrayDB
Dim TDB_BU              As New XArrayDB


Private Const ptxST_SOKO% = 0           '標準棚番(倉庫№)
Private Const ptxJIZEN_S% = 1           '事前商品化状況 開始
Private Const ptxJIZEN_E% = 2           '事前商品化状況 終了
Private Const ptxYOTEI_DT_S% = 3        '商品化予定日 開始
Private Const ptxYOTEI_DT_E% = 4        '商品化予定日 終了
Private Const ptxNYUKA_YOTEI_DT_S% = 5  '部品入荷予定日 開始
Private Const ptxNYUKA_YOTEI_DT_E% = 6  '部品入荷予定日 終了
Private Const ptxHIN_GAI% = 7           '対外品番       2011.12.19




Private Const Min_Row% = 1              '最小行数
Private Max_Row    As Integer           'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 20             '最大列数       18-->20 2011.11.30

Private Const colSHORI% = 0             '処理結果
Private Const colJGYOBU% = 1            '事業部
Private Const colHIN_GAI% = 2           '対外品番
Private Const colST_TANABAN% = 3        '標準棚番

Private Const colSIZAI% = 4             '個装資材   2011.11.10



Private Const colJIZEN% = 5             '事前商品化状況(%)
Private Const colJIZEN_NEEDS_QTY% = 6   '事前商品化必要数
Private Const colYOTEI_QTY% = 7         '商品化予定数
Private Const colZ_QTY_S% = 8           '在庫数(済)
Private Const colZ_QTY_MI% = 9          '在庫数(未)
Private Const colYOTEI_DT% = 10         '商品化予定日

Private Const colS_KOUSU% = 11          '標準工数
Private Const colS_JIKAN% = 12          '標準時間


Private Const colNYUKA_YOTEI_DT% = 13   '部品入荷予定日
Private Const colNYUKA_YOTEI_QTY% = 14  '部品入荷予定数
Private Const colSASIZU_DateTime% = 15  '商品化指図票印刷日時
Private Const colS_KAN_DateTime% = 16   '商品化完了登録日時

Private Const colKEY_No% = 17           'KEY_NO
Private Const colY_NYUKA_KEY_NO% = 18   '入荷予定KEY_NO

Private Const colBEF_YOTEI_QTY% = 19    '商品化予定数       2011.11.30
Private Const colBEF_S_KOUSU% = 20      '標準工数(変更前)   2011.11.30



Private List_Week   As Long             '表示するn週間
Private SAMPLE_QTY  As Integer          '見本除外数




Private Type KOUSEI_TBL
    KO_JGYOBU   As String * 1           '事業部
    KO_NAIGAI   As String * 1           '国内外
    KO_SYUBETSU As String * 2           '種別
    KO_HIN_GAI  As String * 20          '品番
    KO_QTY      As Double               '員数
    G_ST_SHITAN As Double               '仕入＠
    G_ST_URITAN As Double               '売上＠
    G_ST_SHIKIN As Double               '仕入金額
    G_ST_URIKIN As Double               '売上金額
    S_KOUSU     As Double               '作業時間
    SEI_SYU_KON As Double               '集合梱包
    G_ST_URIKIN_KUSATU As _
                    Double              '草津専用
End Type




Dim SHIZAI_T        As Variant          '資材対象
Dim DOUKON_T        As Variant          '同梱対象
Dim KAKOU_T         As Variant          '加工対象

Dim KUSATU_F                As Boolean  '対象センター　草津 OR 草津以外


Dim KOSOU_KBN       As String * 2       '個装区分
Dim GAISO_KBN       As String * 2       '外装区分


Dim KEY_NO          As String * 8
Dim NYUKA_KEY_NO          As String * 8


Dim PLN0040CSV      As String           '2012.04.24

Private Z_Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順


Private Const LAST_UPDATE_DAY$ = "[PLN0040] 2015.02.16 10:15"

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
            '取込みﾃﾞｰﾀ表示
            '--------------------------------   事前商品化状況
            If Trim(Text1(ptxJIZEN_S).Text) <> "" Then
                If Not IsNumeric(Text1(ptxJIZEN_S).Text) Then
                    MsgBox "入力した項目はエラーです。（事前商品化状況 開始）"
                    Text1(ptxJIZEN_S).SetFocus
                    Exit Sub
                End If
            Else
'2011.11.10                Text1(ptxJIZEN_S).Text = "000"
            End If
            
            If Trim(Text1(ptxJIZEN_E).Text) <> "" Then
                If Not IsNumeric(Text1(ptxJIZEN_E).Text) Then
                    MsgBox "入力した項目はエラーです。（事前商品化状況 終了）"
                    Text1(ptxJIZEN_E).SetFocus
                    Exit Sub
                End If
            Else
'2011.11.10                Text1(ptxJIZEN_E).Text = "999"
            End If
            If Text1(ptxJIZEN_S).Text > Text1(ptxJIZEN_E).Text Then
                MsgBox "入力した項目はエラーです。（事前商品化状況）"
                Text1(ptxJIZEN_S).SetFocus
                Exit Sub
            End If
            '--------------------------------   商品化予定日
            If Trim(Text1(ptxYOTEI_DT_S).Text) <> "" Then
                If Not IsDate(Text1(ptxYOTEI_DT_S).Text) Then
                    MsgBox "入力した項目はエラーです。（商品化予定日 開始）"
                    Text1(ptxYOTEI_DT_S).SetFocus
                    Exit Sub
                End If
            End If
            
            If Trim(Text1(ptxYOTEI_DT_E).Text) <> "" Then
                If Not IsDate(Text1(ptxYOTEI_DT_E).Text) Then
                    MsgBox "入力した項目はエラーです。（商品化予定日 終了）"
                    Text1(ptxYOTEI_DT_E).SetFocus
                    Exit Sub
                End If
            End If
            If Text1(ptxYOTEI_DT_S).Text > Text1(ptxYOTEI_DT_E).Text Then
                MsgBox "入力した項目はエラーです。（商品化予定日）"
                Text1(ptxYOTEI_DT_S).SetFocus
                Exit Sub
            End If
            '--------------------------------   部品入荷予定日
            If Trim(Text1(ptxNYUKA_YOTEI_DT_S).Text) <> "" Then
                If Not IsDate(Text1(ptxNYUKA_YOTEI_DT_S).Text) Then
                    MsgBox "入力した項目はエラーです。（部品入荷予定日 開始）"
                    Text1(ptxNYUKA_YOTEI_DT_S).SetFocus
                    Exit Sub
                End If
            End If
            
            If Trim(Text1(ptxNYUKA_YOTEI_DT_E).Text) <> "" Then
                If Not IsDate(Text1(ptxNYUKA_YOTEI_DT_E).Text) Then
                    MsgBox "入力した項目はエラーです。（部品入荷予定日 終了）"
                    Text1(ptxNYUKA_YOTEI_DT_E).SetFocus
                    Exit Sub
                End If
            End If
            If Text1(ptxNYUKA_YOTEI_DT_S).Text > Text1(ptxNYUKA_YOTEI_DT_E).Text Then
                MsgBox "入力した項目はエラーです。（部品入荷予定日）"
                Text1(ptxYOTEI_DT_S).SetFocus
                Exit Sub
            End If
            
                        
                        
                        
            
            
            
            If List_Disp_Proc(False) Then
                Unload Me
            End If


            If PLN_S_YOTEI.Count(1) > 0 Then
                Command1(1).Enabled = True
                Command1(3).Enabled = True
                Command1(4).Enabled = True
                SHORI(1).Enabled = True
            
                If RTrim(PLN0040CSV) <> "" Then     '2012.04.24
                    Command1(6).Enabled = True      '2012.04.24
                End If                              '2012.04.24
            Else
                Command1(1).Enabled = False
                Command1(3).Enabled = False
                Command1(4).Enabled = False
                SHORI(1).Enabled = False
            
                Command1(6).Enabled = False         '2012.04.24
            
            
                Command1(0).SetFocus
                Exit Sub
            End If


        Case 1          '登録



            If Err_Check_Proc() Then
                Exit Sub
            End If


            If Update_Proc(NYUKA_FLG) Then
                Unload Me
            End If




            If List_Disp_Proc(NYUKA_FLG) Then
                Unload Me
            End If


            If PLN_S_YOTEI.Count(1) > 0 Then
                Command1(1).Enabled = True
                Command1(3).Enabled = True
                Command1(4).Enabled = True
                
                SHORI(1).Enabled = True
            
            Else
                Command1(1).Enabled = False
                Command1(3).Enabled = False
                Command1(4).Enabled = False
                
                SHORI(1).Enabled = False
            
                Command1(0).SetFocus
                Exit Sub
            End If

        Case 2          '終了

            Unload Me
        Case 3          '複数行選択集計
            Set SelBks = TDBGrid1.SelBookmarks
            TOTAL_JIKAN = 0
                
            Set TDBGrid1.Array = PLN_S_YOTEI
            TDBGrid1.Update
                
                
            '未選択時は、全件集計
            If TDBGrid1.SelBookmarks.Count <= 0 Then        '2011.11.18
            
                For i = 1 To PLN_S_YOTEI.Count(1)
                    DoEvents
                        
                    If IsNumeric(PLN_S_YOTEI(i, colS_JIKAN)) Then
                        TOTAL_JIKAN = TOTAL_JIKAN + CDbl(PLN_S_YOTEI(i, colS_JIKAN))
                    End If
                Next i
            
            
            Else
                
                For i = 0 To TDBGrid1.SelBookmarks.Count - 1
                    DoEvents
                    If TDBGrid1.SelBookmarks.ITEM(i) <> 0 Then
                        
                        If IsNumeric(PLN_S_YOTEI(TDBGrid1.SelBookmarks.ITEM(i), colS_JIKAN)) Then
                            TOTAL_JIKAN = TOTAL_JIKAN + CDbl(PLN_S_YOTEI(TDBGrid1.SelBookmarks.ITEM(i), colS_JIKAN))
                        End If
                    End If
                Next i
            
            End If

            
            lblS_S_JIKAN = Format(TOTAL_JIKAN, "#0.0")



        Case 4          '複数行選択解除
            
            
            
            Set SelBks = TDBGrid1.SelBookmarks
            Do
                DoEvents
                If TDBGrid1.SelBookmarks.Count <> 0 Then
                    SelBks.Remove 0
                Else
                    Exit Do
                End If
            Loop

            lblS_S_JIKAN = ""
    
    
        Case 5          '画面印刷   2011.11.30
    
            'Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)           '2015.02.16
            Call Form_HCopy_Win7(Picture1, vbPRPSA4, vbPRORLandscape)       '2015.02.16

    
    
        Case 6          'ﾃﾞｰﾀ出力   2012.04.24
    
            yn = MsgBox("[商品化予定ﾃﾞｰﾀ]ﾃﾞｰﾀ出力しますか？", vbYesNo, "確認入力")
            If yn = vbYes Then
                If OutPut_Proc() Then
                    Unload Me
                End If
            End If
    
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


Dim c       As String * 128
Dim sts     As Integer


    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[商品化計画システム]商品化予定データメンテナンス処理", Me.hwnd, 0)
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

    Call TDB_BU_Set_Proc





                                '展開対象期間名取り込み
    If GetIni(App.EXEName, "WEEK", App.EXEName, c) Then
        List_Week = 2
    Else
        If Not IsNumeric(Trim(c)) Then
            List_Week = 2
        Else
            If Val(Trim(c)) < 1 Then
                List_Week = 2
            Else
                If Val(Trim(c)) > 8 Then
                    List_Week = 8
                Else
                    List_Week = Val(Trim(c))
                End If
            End If
        End If
    End If
        

                                'サンプル除外数
    If GetIni(App.EXEName, "Sample_QTY", App.EXEName, c) Then
        SAMPLE_QTY = 0
    Else
        If IsNumeric(Trim(c)) Then
            SAMPLE_QTY = CLng(Trim(c))
        Else
            SAMPLE_QTY = 0
        End If
    End If



                                '資材対象種別
    If GetIni("SEI0010", "SHIZAI", "SEI0010", c) Then
        
        c = "**"
        SHIZAI_T = Split(Trim(c), ",", -1)
        
    Else
        SHIZAI_T = Split(Trim(c), ",", -1)
    End If
                                
                                '同梱対象種別
    If GetIni("SEI0010", "DOUKON", "SEI0010", c) Then
        c = "**"
        DOUKON_T = Split(Trim(c), ",", -1)
    Else
        DOUKON_T = Split(Trim(c), ",", -1)
    End If
                                '加工対象種別
   If GetIni("SEI0010", "KAKOU", "SEI0010", c) Then
        c = "**"
        KAKOU_T = Split(Trim(c), ",", -1)
    Else
        KAKOU_T = Split(Trim(c), ",", -1)
    End If
                                'センターの識別
    If GetIni("SEI0010", "KUSATU", "SEI0010", c) Then
        KUSATU_F = False
    Else
        If Trim(c) = "1" Then
            KUSATU_F = True
        Else
            KUSATU_F = False
        End If
    End If
                                
                                
                                
                                '個装資材区分の獲得
    If GetIni("SEI0010", "KOSOU", "SEI0010", c) Then
        KOSOU_KBN = ""
    Else
        KOSOU_KBN = Trim(c)
    End If
                                '外装資材区分の獲得
    If GetIni("SEI0010", "GAISO", "SEI0010", c) Then
        GAISO_KBN = ""
    Else
        GAISO_KBN = Trim(c)
    End If


                                'CSV出力ファイル名の獲得    2012.04.24
    If GetIni("PLN0040", "PLN0040CSV", "PLN0040", c) Then
        PLN0040CSV = ""
    Else
        PLN0040CSV = Trim(c)
    End If




    PLN00401.Caption = PLN00401.Caption & " " & LAST_UPDATE_DAY

                                '商品化予定ファイルＯＰＥＮ
    If PLN_S_YOTEI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化用入荷予定ファイルＯＰＥＮ
    If PLN_Y_NYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenRead) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenRead) Then
        Unload Me
    End If
                                
                                '指図票データ(親)ＯＰＥＮ   2011.11.11
    If P_SSHIJI_O_Open(BtOpenRead) Then
        Unload Me
    End If
                                '商品化指図受入履歴データ ＯＰＥＮ   2011.11.11
    If P_SUKEIRE_Open(BtOpenRead) Then
        Unload Me
    End If
                                '管理マスタＯＰＥＮ＆ＲＥＡＤ
    If P_KANRI_Open(BtOpenRead) Then
        Unload Me
    End If
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ(KEY=01)")
        Unload Me
    End Select

    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_DEF_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC02, Len(P_KANRIREC02), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ(KEY=02)")
        Unload Me
    End Select
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '月平均出荷数ＯＰＥＮ
    If AVE_SYUKA_Open(BtOpenRead) Then
        Unload Me
    End If


    TDBGrid1.Columns(colYOTEI_QTY).HeadBackColor = vbBlue
    TDBGrid1.Columns(colYOTEI_QTY).HeadForeColor = vbWhite

    TDBGrid1.Columns(colYOTEI_DT).HeadBackColor = vbBlue
    TDBGrid1.Columns(colYOTEI_DT).HeadForeColor = vbWhite


'    TDBGrid1.Columns(colS_KOUSU).HeadBackColor = vbBlue
'    TDBGrid1.Columns(colS_KOUSU).HeadForeColor = vbWhite



    TDBGrid1.Columns(colNYUKA_YOTEI_DT).HeadBackColor = vbBlue
    TDBGrid1.Columns(colNYUKA_YOTEI_DT).HeadForeColor = vbWhite

    TDBGrid1.Columns(colNYUKA_YOTEI_QTY).HeadBackColor = vbBlue
    TDBGrid1.Columns(colNYUKA_YOTEI_QTY).HeadForeColor = vbWhite


End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化予定ファイル")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    
    Set PLN00401 = Nothing



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
            Command1(5).Value = True
    
    End Select



End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
Dim sts         As Integer
Dim Bookmark    As Variant
    
    
Dim i           As Integer
    
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
    
Dim KEISAN_FLG  As Boolean      '2011.11.30
    
    
    Set TDBGrid1.Array = PLN_S_YOTEI
    TDBGrid1.Update
    
    
    
    
    If TDBGrid1.Bookmark = Null Then
        Exit Sub
    End If
    
    If TDBGrid1.Bookmark < 0 Then
        Exit Sub
    End If
    
    If PLN_S_YOTEI(TDBGrid1.Bookmark, colSHORI) Then
    Else
        Select Case ColIndex
        
        
            Case colJGYOBU
        
                'BU
                If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colJGYOBU)) = "" Then
                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(ＢＵ　必須入力)"
                    Exit Sub
                End If
        
            Case colHIN_GAI
        
                '品番
        
                If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colHIN_GAI)) <> "" Then
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(PLN_S_YOTEI(TDBGrid1.Bookmark, colJGYOBU), 1))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_S_YOTEI(TDBGrid1.Bookmark, colHIN_GAI))
                
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            PLN_S_YOTEI(TDBGrid1.Bookmark, colST_TANABAN) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)

                        
                        
                            If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU)) = "" Then
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.05.07
                                'If IsNumeric(StrConv(ITEMREC.PLN_KOUSU, vbUnicode)) Then
                                '    PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU) = Format(StrConv(ITEMREC.PLN_KOUSU, vbUnicode), "#0.0")
                                If IsNumeric(StrConv(ITEMREC.PLN_SAGYOU_KOUSU, vbUnicode)) Then
                                    PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU) = Format(StrConv(ITEMREC.PLN_SAGYOU_KOUSU, vbUnicode), "#0.0")
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.05.07
                                Else
                                    PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU) = "0.0"
                                End If
                            End If
                        
                            If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                                Unload Me
                            End If
                        
                            PLN_S_YOTEI(TDBGrid1.Bookmark, colZ_QTY_S) = Format(SUMI_QTY, "#,##0")
                            PLN_S_YOTEI(TDBGrid1.Bookmark, colZ_QTY_MI) = Format(MI_QTY, "#,##0")
                                                    
                        
                        
                        Case BtErrKeyNotFound
                            MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(品番未登録)"
                            Exit Sub
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                            Unload Me
                    End Select
                End If
        
        
        
            Case colYOTEI_QTY
        
        
        
                If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) = "" Then
                Else
                    If Not IsNumeric(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) Then
                            
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(商品化予定数)"
                        Exit Sub
                    End If
            
            
                    If Val(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) < 1 Then
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(商品化予定数≦０)"
                        Exit Sub
                    End If
                End If
                If IsNumeric(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) Then
                    PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY) = Format(CLng(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)), "#,###")
''2011.11.30                    PLN_S_YOTEI(TDBGrid1.Bookmark, colS_JIKAN) = Format(Round(Val(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) * Val(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU)), 2), "#0.0")
                End If
        
        
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   工数時間の自動計算制御2011.11.30
                    
                If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) = Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colBEF_YOTEI_QTY)) And _
                    Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_JIKAN)) = Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_JIKAN)) Then
                                
                
                Else
                
                    If IsNumeric(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) And IsNumeric(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU)) Then
                        
                        PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY) = Format(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY), "#,##0")
                        
                        
                        PLN_S_YOTEI(TDBGrid1.Bookmark, colS_JIKAN) = Format(Round(CLng(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) * Val(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU)), 2), "#0.0")
                                    
                        PLN_S_YOTEI(TDBGrid1.Bookmark, colBEF_YOTEI_QTY) = PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)
                        PLN_S_YOTEI(TDBGrid1.Bookmark, colBEF_S_KOUSU) = PLN_S_YOTEI(TDBGrid1.Bookmark, colS_JIKAN)
                
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   工数時間の自動計算制御2011.11.30
        
        
            Case colYOTEI_DT
                
                If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_DT)) = "" Then
''2011.10.31                    If IsNumeric(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) Then
''2011.10.31                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。([商品化予定日]商品化予定数入力時、必須入力)"
''2011.10.31                        Exit Sub
''2011.10.31                    End If
                Else
                    If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) = "" Then
                        
''2011.10.31                        If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_DT)) <> "" Then
''2011.10.31                            MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。([商品化予定日]商品化予定数未入力時、入力不可)"
''2011.10.31                            Exit Sub
''2011.10.31                        End If
                        
                    Else
                        If Not IsDate(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_DT)) Then
                            MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(商品化予定日)"
                            Exit Sub
                        End If
                
                        If PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_DT) < Format(Now, "YYYY/MM/DD") Then
                            MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(商品化予定日＜本日)"
                            Exit Sub
                        End If
                
                    End If
                End If
        
        
        
            Case colS_KOUSU
                If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU)) = "" Then
                    PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU) = "0.0"
                End If
        
        
                If Not IsNumeric(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU)) Then
                
                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(標準工数)"
                    Exit Sub
                
                End If
        
                If Val(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU)) < 0 Then
                
                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(標準工数＜０)"
                    Exit Sub
                
                End If
        
        
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   工数時間の自動計算制御2011.11.30
                If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU)) = "" Then
                Else
                    If Not IsNumeric(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) Then
                            
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(商品化予定数)"
                        Exit Sub
                    End If
            
            
                    If Val(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) < 1 Then
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(商品化予定数≦０)"
                        Exit Sub
                    End If
                End If
                If IsNumeric(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) Then
                    PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY) = Format(CLng(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)), "#,###")
''2011.11.30                    PLN_S_YOTEI(TDBGrid1.Bookmark, colS_JIKAN) = Format(Round(Val(PLN_S_YOTEI(TDBGrid1.Bookmark, colYOTEI_QTY)) * Val(PLN_S_YOTEI(TDBGrid1.Bookmark, colS_KOUSU)), 2), "#0.0")
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   工数時間の自動計算制御2011.11.30
        
        
            Case colNYUKA_YOTEI_DT
        
                If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colNYUKA_YOTEI_DT)) <> "" Then
                    If Not IsDate(PLN_S_YOTEI(TDBGrid1.Bookmark, colNYUKA_YOTEI_DT)) Then
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(入荷予定日)"
                        Exit Sub
                    End If
            
''2011.10.31                    If PLN_S_YOTEI(TDBGrid1.Bookmark, colNYUKA_YOTEI_DT) < Format(Now, "YYYY/MM/DD") Then
''2011.10.31                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(入荷予定日＜本日)"
''2011.10.31                        Exit Sub
''2011.10.31                    End If
                End If
            Case colNYUKA_YOTEI_QTY
                
        
                If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colNYUKA_YOTEI_QTY)) = "" Then
''2011.10.31                    If IsDate(PLN_S_YOTEI(TDBGrid1.Bookmark, colNYUKA_YOTEI_DT)) Then
''2011.10.31                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。([入荷予定数]入荷予定日入力時、必須入力)"
''2011.10.31                        Exit Sub
''2011.10.31                    End If
                Else
                        
''2011.10.31                    If Trim(PLN_S_YOTEI(TDBGrid1.Bookmark, colNYUKA_YOTEI_DT)) = "" Then
''2011.10.31                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。([入荷予定数]入荷予定日未入力時、入力不可)"
''2011.10.31                        Exit Sub
''2011.10.31                    End If
                        
                    If Not IsNumeric(PLN_S_YOTEI(TDBGrid1.Bookmark, colNYUKA_YOTEI_QTY)) Then
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(入荷予定数)"
                        Exit Sub
                    End If
                
                    If Val(PLN_S_YOTEI(TDBGrid1.Bookmark, colNYUKA_YOTEI_QTY)) < 1 Then
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(入荷予定数≦０)"
                        Exit Sub
                    End If
                
                    PLN_S_YOTEI(TDBGrid1.Bookmark, colNYUKA_YOTEI_QTY) = Format(CLng(PLN_S_YOTEI(TDBGrid1.Bookmark, colNYUKA_YOTEI_QTY)), "#,###")
                
                End If
        
        End Select
    End If
        
    Set TDBGrid1.Array = PLN_S_YOTEI
        
    
    TDBGrid1.Refresh
    TDBGrid1.Update
    TDBGrid1.SetFocus
    




End Sub


Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)
    
    PLN_S_YOTEI.ReDim Min_Row, PLN_S_YOTEI.Count(1), Min_Col, Max_Col

    Command1(1).Enabled = True

End Sub



Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    If PLN_S_YOTEI.Count(1) <= 0 Then
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
            PLN_S_YOTEI.QuickSort Min_Row, PLN_S_YOTEI.UpperBound(1), ColIndex, Z_Sort_Tbl(ColIndex), XTYPE_LONG
        Else
            PLN_S_YOTEI.QuickSort Min_Row, PLN_S_YOTEI.UpperBound(1), ColIndex, Z_Sort_Tbl(ColIndex), XTYPE_STRING
        End If
        
        Set TDBGrid1.Array = PLN_S_YOTEI
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If

End Sub
Private Function Update_Proc(NYUKA_FLG As Boolean) As Integer
'----------------------------------------------------------------------------
'                   「商品化予定ファイル」登録処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim Upd_Com         As Integer
Dim Upd_NYUKA_Com   As Integer

Dim Skip_Flg        As Integer
    
Dim INS_NOW         As String * 14
Dim Row             As Long


Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim AVE_QTY         As Double

Dim MAIN_KOUTEI(0 To 9) _
                    As Long
Dim wkTANI          As Double
Dim wkQTY           As Double

Dim KOUSEI()        As KOUSEI_TBL
Dim i               As Integer
Dim j               As Integer
Dim KOUSEI_FLG      As Boolean

Dim wkInt           As Integer

Dim SHIMUKE_CODE    As String * 2
Dim c               As String * 128



    If PLN_S_YOTEI.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定ファイル登録処理　処理開始！！", Me.hwnd, 0)

                                    
                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
                                    
                                    'テーブルリセット
    
    Skip_Flg = False
    NYUKA_FLG = False
    For Row = 1 To PLN_S_YOTEI.UpperBound(1)
        
        
        DoEvents
        
        
        sts = BTRV(BtOpGetLast, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K3_PLN_S_YOTEI, Len(K3_PLN_S_YOTEI), 3)
        Select Case sts
            Case BtNoErr
                KEY_NO = Format(Val(StrConv(PLN_S_YOTEI_R.KEY_NO, vbUnicode)) + 1, "00000000")
            Case BtErrEOF
                KEY_NO = "00000001"
            Case Else
                Call File_Error(sts, BtOpGetLast, "商品化予定ファイル")
                Call Input_UnLock
                Exit Function
        End Select
        
        
''        sts = BTRV(BtOpGetLast, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
''        Select Case sts
''            Case BtNoErr
''                NYUKA_KEY_NO = Format(Val(StrConv(PLN_Y_NYUKA_R.KEY_NO, vbUnicode)) + 1, "00000000")
''            Case BtErrEOF
''                KEY_NO = "00000001"
''            Case Else
''                Call File_Error(sts, BtOpGetLast, "商品化用入荷予定ファイル")
''                Call Input_UnLock
''                Exit Function
''        End Select
        
If Row = 2 Then
    Debug.Print "UPD IN= " & PLN_S_YOTEI(Row, colS_JIKAN)
End If
        
        
        
        If Trim(PLN_S_YOTEI(Row, colKEY_No)) = "" Then
            Upd_Com = BtOpInsert
        Else
            Call UniCode_Conv(K3_PLN_S_YOTEI.KEY_NO, PLN_S_YOTEI(Row, colKEY_No))
            
            sts = BTRV(BtOpGetEqual, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K3_PLN_S_YOTEI, Len(K3_PLN_S_YOTEI), 3)
            Select Case sts
                Case BtNoErr
                    If PLN_S_YOTEI(Row, colSHORI) Then
                        Upd_Com = BtOpDelete
                    Else
                        Upd_Com = BtOpUpdate
                    End If
                
                
                Case BtErrKeyNotFound
                    If PLN_S_YOTEI(Row, colSHORI) Then
                        Skip_Flg = True
                    Else
                        Upd_Com = BtOpInsert
                    End If
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "商品化予定ファイル")
                    Call Input_UnLock
                    Exit Function
            End Select
        End If
        
        If Not Skip_Flg Then
            If Upd_Com = BtOpDelete Then
            Else
                If Upd_Com = BtOpInsert Then
                    Call PLN_S_YOTEI_CLR
                    
                            
                    Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_DT, Format(PLN_S_YOTEI(Row, colYOTEI_DT), "YYYYMMDD"))
                    Call UniCode_Conv(PLN_S_YOTEI_R.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
                    Call UniCode_Conv(PLN_S_YOTEI_R.NAIGAI, "1")
                    Call UniCode_Conv(PLN_S_YOTEI_R.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
                    
                    '------------------------------ 品目マスタより
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                            Call UniCode_Conv(PLN_S_YOTEI_R.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                        
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "1")
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目ファイル")
                            Call Input_UnLock
                            Exit Function
                    End Select
                    '------------------------------ 在庫情報より
                    If Zaiko_Syukei_Proc(SUMI_QTY, _
                                            MI_QTY, _
                                            Right(PLN_S_YOTEI(Row, colJGYOBU), 1), _
                                            "1", _
                                            PLN_S_YOTEI(Row, colHIN_GAI)) = SYS_ERR Then
                        Call Input_UnLock
                        Exit Function
                    End If
                    
                    SUMI_QTY = SUMI_QTY - SAMPLE_QTY
                    If SUMI_QTY < 0 Then
                        SUMI_QTY = 0
                    End If
                    Call UniCode_Conv(PLN_S_YOTEI_R.Z_QTY_MI, Format(MI_QTY, "00000000"))
                    Call UniCode_Conv(PLN_S_YOTEI_R.Z_QTY_S, Format(SUMI_QTY, "00000000"))
                    '------------------------------ 月平均出荷数より
                    Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
                    Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, "1")
                    Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
                    
                    sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                    Select Case sts
                        Case BtNoErr
                            AVE_QTY = Val(StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, vbUnicode))
                            If AVE_QTY = 0 Then
                                Call UniCode_Conv(PLN_S_YOTEI_R.JIZEN, "00000000")
                            Else
                                Call UniCode_Conv(PLN_S_YOTEI_R.JIZEN, Format(CLng(SUMI_QTY / AVE_QTY * 100), "00000000"))
                            End If
                            Call UniCode_Conv(PLN_S_YOTEI_R.JIZEN_NEEDS_QTY, Format(CLng(StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, vbUnicode)) - SUMI_QTY, "00000000"))
                        
                        Case BtErrKeyNotFound
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "商品化予定ファイル")
                            Call Input_UnLock
                            Exit Function
                    End Select
                    
                    
                    '------------------------------ 商品化用入荷予定ファイル
                    Call UniCode_Conv(K1_PLN_Y_NYUKA.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
                    Call UniCode_Conv(K1_PLN_Y_NYUKA.NAIGAI, "1")
                    Call UniCode_Conv(K1_PLN_Y_NYUKA.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
                    Call UniCode_Conv(K1_PLN_Y_NYUKA.N_YOTEI_DT, "")
                    Call UniCode_Conv(K1_PLN_Y_NYUKA.SEQ_NO, "")
                    
                    sts = BTRV(BtOpGetGreater, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K1_PLN_Y_NYUKA, Len(K1_PLN_Y_NYUKA), 1)
                    Select Case sts
                        Case BtNoErr
                        
                            If StrConv(PLN_Y_NYUKA_R.JGYOBU, vbUnicode) = Right(PLN_S_YOTEI(Row, colJGYOBU), 1) And _
                                StrConv(PLN_Y_NYUKA_R.NAIGAI, vbUnicode) = "1" And _
                                Trim(StrConv(PLN_Y_NYUKA_R.HIN_GAI, vbUnicode)) = Trim(PLN_S_YOTEI(Row, colHIN_GAI)) Then
                        
                                Call UniCode_Conv(PLN_S_YOTEI_R.NYUKA_YOTEI_DT, StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode))
                                Call UniCode_Conv(PLN_S_YOTEI_R.NYUKA_YOTEI_QTY, StrConv(PLN_Y_NYUKA_R.N_YOTEI_QTY, vbUnicode))
                        
                        
                            End If
                        
                        Case BtErrEOF
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "商品化用入荷予定ファイル")
                            Call Input_UnLock
                            Exit Function
                    End Select
                    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    見積工数
                    For i = 0 To UBound(MAIN_KOUTEI)
                        MAIN_KOUTEI(i) = 0
                    Next i
                    '①
                    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)) Then
                        
                        wkTANI = Val(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode))
                    Else
                        wkTANI = 0
                    End If
                    If IsNumeric(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)) Then
                        If IsDate(Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 1, 4) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 5, 2) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 7, 4)) Then
                            wkQTY = Val(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode))
                        Else
                            wkQTY = 1
                        End If
                    Else
                        wkQTY = 1
                    End If
                    MAIN_KOUTEI(0) = wkTANI * wkQTY
                    '②
                    '-------------------　構成情報テーブル展開
                    Erase KOUSEI
                    i = -1
        
                    KOUSEI_FLG = False
                    If GetIni(App.EXEName, Right(PLN_S_YOTEI(Row, colJGYOBU), 1), App.EXEName, c) Then
                        SHIMUKE_CODE = ""
                    Else
                        SHIMUKE_CODE = Trim(c)
                    End If
                    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_CODE)
                    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
                    Call UniCode_Conv(K0_P_COMPO.NAIGAI, "1")
                    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
                    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
                                        
                    com = BtOpGetGreater
                                        
                    Do
                        DoEvents
                    
                        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr
                    
                                    
                                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_CODE Or _
                                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Right(PLN_S_YOTEI(Row, colJGYOBU), 1) Or _
                                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> "1" Or _
                                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(PLN_S_YOTEI(Row, colHIN_GAI)) Then
                                
                                    Exit Do
                            
                                End If
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, BtOpGetNext, "構成マスタ")
                                Exit Function
                        End Select
                    
                        If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_KOSOU Then
                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, KOSOU_KBN)
                        End If
                        If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, GAISO_KBN)
                        End If
                    
                        i = i + 1
                        KOUSEI_FLG = True
                                
                        ReDim Preserve KOUSEI(0 To i)
                        '事業部
                        KOUSEI(i).KO_JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                        '国内外
                        KOUSEI(i).KO_NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                        '種別
                        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
                        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                        Select Case sts
                            Case BtNoErr
                                KOUSEI(i).KO_SYUBETSU = Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
                            Case BtErrKeyNotFound
                                KOUSEI(i).KO_SYUBETSU = ""
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                                Exit Function
                        End Select
                        '品番
                        KOUSEI(i).KO_HIN_GAI = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                         
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            Case BtErrKeyNotFound
                                    
                                Call UniCode_Conv(ITEMREC.SEI_KBN, "")
                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
                                Call UniCode_Conv(ITEMREC.S_KOUSU, "")
                                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")
                            
                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        End Select
                        '員数
                        If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
                            KOUSEI(i).KO_QTY = CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                        Else
                            KOUSEI(i).KO_QTY = 1#
                        End If
                        '仕入単価
                        If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                            KOUSEI(i).G_ST_SHITAN = CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode))
                        Else
                            KOUSEI(i).G_ST_SHITAN = 0#
                        End If
                        '売上単価
                        Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
                        
                            Case "1"
                               KOUSEI(i).G_ST_URITAN = 0#
                            Case "2"
                               KOUSEI(i).G_ST_URITAN = 0#
                            Case Else
                                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                                    KOUSEI(i).G_ST_URITAN = CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))
                                Else
                                    KOUSEI(i).G_ST_URITAN = 0#
                                End If
                        End Select
                        '仕入金額計
                        KOUSEI(i).G_ST_SHIKIN = 0#
                        For j = 0 To UBound(SHIZAI_T)
                            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(j) Then
                                
                                
                                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                                    
                                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                                        
                                        If CDbl(KOUSEI(i).KO_QTY) = 0 Then
                                            KOUSEI(i).G_ST_SHIKIN = 0#
                                        Else
                                            KOUSEI(i).G_ST_SHIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_SHITAN)) / CDbl(KOUSEI(i).KO_QTY), 2)
                                        End If
                                    Else
                                        KOUSEI(i).G_ST_SHIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY)) * CDbl(KOUSEI(i).G_ST_SHITAN), 2)
                                    End If
                                End If
                                Exit For
                            End If
                        
                        Next j
                        '売上金額計
                        KOUSEI(i).G_ST_URIKIN = 0
                        KOUSEI(i).G_ST_URIKIN_KUSATU = 0
                    
                        For j = 0 To UBound(SHIZAI_T)
                               
                            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(j) Then
                                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                                        If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then
                                            KOUSEI(i).G_ST_URIKIN = 0#
                                        Else
                                            KOUSEI(i).G_ST_URIKIN = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                                        End If
                                        KOUSEI(i).G_ST_URIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_URITAN)) * CDbl(KOUSEI(i).G_ST_URIKIN), 2)
                                    Else
                                        KOUSEI(i).G_ST_URIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY) * CDbl(KOUSEI(i).G_ST_URITAN)), 2)
                                    End If
                            
                                    
                                Else
                                   
                                    If KUSATU_F Then
                                
                                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                                        
                                        
                                            If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then
                                                KOUSEI(i).G_ST_URIKIN_KUSATU = 0
                                            Else
                                                KOUSEI(i).G_ST_URIKIN_KUSATU = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                                            End If
                                            KOUSEI(i).G_ST_URIKIN_KUSATU = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_URITAN)) * CDbl(KOUSEI(i).G_ST_URIKIN_KUSATU), 2)
                                        
                                        Else
                                            KOUSEI(i).G_ST_URIKIN_KUSATU = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY)) * CDbl(KOUSEI(i).G_ST_URITAN), 2)
                                        End If
                                    
                                    
                                    End If
                                End If
                            End If
                        Next j
                            
                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                            KOUSEI(i).S_KOUSU = 0
                            KOUSEI(i).SEI_SYU_KON = 0
                        Else
                            '作業時間
                            If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                                KOUSEI(i).S_KOUSU = CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode))
                            Else
                                KOUSEI(i).S_KOUSU = 0#
                            End If
                            '集合梱包
                            If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
                                KOUSEI(i).SEI_SYU_KON = CDbl(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode))
                            Else
                                KOUSEI(i).SEI_SYU_KON = 0#
                            End If
                        End If
                    Loop
    
    
                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                        
                            For j = 0 To UBound(SHIZAI_T)
                                If KOUSEI(i).KO_SYUBETSU = SHIZAI_T(j) Then
                                    wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i).S_KOUSU) * CDbl(KOUSEI(i).KO_QTY), 0))
                                    Exit For
                                End If
                        
                            Next j
                        
                        Next i
                    End If
                        
                    wkTANI = wkInt
                    wkQTY = 1
                    MAIN_KOUTEI(1) = wkTANI * wkQTY
    
                    '③
                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                        
                            For j = 0 To UBound(DOUKON_T)
                                If KOUSEI(i).KO_SYUBETSU = DOUKON_T(j) Then
                                    
                                    If IsNumeric(KOUSEI(i).KO_QTY) Then
                                        wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i).KO_QTY), 0))
                                    End If
                                    
                                    
                                    
                                    Exit For
                                End If
                        
                            Next j
                        
                        Next i
                    End If
                    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)) Then
                        wkTANI = CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode))
                    Else
                        wkTANI = 0#
                    End If
                    wkQTY = wkInt
                    MAIN_KOUTEI(2) = wkTANI * wkQTY
                    '④
                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                        
                            For j = 0 To UBound(KAKOU_T)
                                If KOUSEI(i).KO_SYUBETSU = KAKOU_T(j) Then
                                    If IsNumeric(KOUSEI(i).S_KOUSU) Then
                                        wkInt = wkInt + CInt(KOUSEI(i).S_KOUSU)
                                    End If
                                    Exit For
                                End If
                        
                            Next j
                        
                        Next i
                    End If
                    wkTANI = wkInt
                    wkQTY = 1
                    MAIN_KOUTEI(3) = wkTANI * wkQTY
                    '⑤
                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                            
                            
                            For j = 0 To UBound(SHIZAI_T)
                            
                                If KOUSEI(i).KO_SYUBETSU = SHIZAI_T(j) Then
                                    If IsNumeric(KOUSEI(i).SEI_SYU_KON) Then
                                        wkInt = wkInt + CInt(KOUSEI(i).SEI_SYU_KON)
                                    End If
                                End If
                            
                            Next j
                            
                        Next i
                    End If
                    wkTANI = wkInt
                    wkQTY = 1
                    MAIN_KOUTEI(4) = wkTANI * wkQTY
    
    
                    '計
                    wkInt = 0
                    For i = 0 To UBound(MAIN_KOUTEI)
                    
                        wkInt = wkInt + MAIN_KOUTEI(i)
                    Next i
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_KOUSU, Format(ToHalfAdjust(CCur(wkInt) / 60, 1), "000000.0"))
                    
                    
                    Call UniCode_Conv(PLN_S_YOTEI_R.INS_TANTO, App.EXEName)
                
                
                
                    Call UniCode_Conv(PLN_S_YOTEI_R.KEY_NO, KEY_NO)
                
                End If
                '商品化予定日
                If Trim(PLN_S_YOTEI(Row, colYOTEI_DT)) <> "" Then
                    Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_DT, Format(PLN_S_YOTEI(Row, colYOTEI_DT), "YYYYMMDD"))
                Else
                    Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_DT, "")
                End If
                
                '>>>>>>>>>>>>>>>>>>>>>商品化予定工数時間    商品化予定数または見積工数が変更された場合のみ工数を再計算 2011.12.19
                'If IsNumeric(PLN_S_YOTEI(Row, colYOTEI_QTY)) And StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode) Then          '2013.05.16
                If IsNumeric(PLN_S_YOTEI(Row, colYOTEI_QTY)) And IsNumeric(PLN_S_YOTEI(Row, colS_KOUSU)) Then               '2013.05.16
                    If CLng(PLN_S_YOTEI(Row, colYOTEI_QTY)) = CLng(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)) And _
                        CDbl(PLN_S_YOTEI(Row, colS_KOUSU)) = CDbl(StrConv(PLN_S_YOTEI_R.S_KOUSU, vbUnicode)) Then
                    Else
                        Call UniCode_Conv(PLN_S_YOTEI_R.S_JIKAN, Format(ToHalfAdjust(CCur(CDbl(PLN_S_YOTEI(Row, colS_KOUSU))) * CCur(PLN_S_YOTEI(Row, colYOTEI_QTY)), 1), "000000.0"))
                    End If
                Else
                    Call UniCode_Conv(PLN_S_YOTEI_R.S_JIKAN, "000000.0")
                End If
                
                '商品化予定数
                If Not IsNumeric(PLN_S_YOTEI(Row, colYOTEI_QTY)) Then
                    PLN_S_YOTEI(Row, colYOTEI_QTY) = 0
                    Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_QTY, "00000000")
                Else
                    Call UniCode_Conv(PLN_S_YOTEI_R.YOTEI_QTY, Format(CLng(PLN_S_YOTEI(Row, colYOTEI_QTY)), "00000000"))
                End If
                '商品化予定工数
                Call UniCode_Conv(PLN_S_YOTEI_R.S_KOUSU, Format(Val(PLN_S_YOTEI(Row, colS_KOUSU)), "000000.0"))
                '商品化予定工数時間
'2011.11.30                Call UniCode_Conv(PLN_S_YOTEI_R.S_JIKAN, Format(ToHalfAdjust(CCur(Val(PLN_S_YOTEI(Row, colS_KOUSU))) * CCur(PLN_S_YOTEI(Row, colYOTEI_QTY)), 1), "000000.0"))
'2011.12.19                Call UniCode_Conv(PLN_S_YOTEI_R.S_JIKAN, PLN_S_YOTEI(Row, colS_JIKAN))      '2011.11.30
                
                '部品入荷予定日/予定数/KEY_NO
                
If "AMC92F-8T0" = Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) Then
Debug.Print
End If
                
                
                If IsDate(PLN_S_YOTEI(Row, colNYUKA_YOTEI_DT)) Then
                    Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, Format(PLN_S_YOTEI(Row, colNYUKA_YOTEI_DT), "YYYYMMDD"))
                Else
''                    Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, "99999999")
                    Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, "")
                End If
                    
                If IsNumeric(PLN_S_YOTEI(Row, colNYUKA_YOTEI_QTY)) Then
                    
                    
                    Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_QTY, Format(PLN_S_YOTEI(Row, colNYUKA_YOTEI_QTY), "00000000"))

                Else

                    Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_QTY, Format(0, "00000000"))

''2011.10.04
''                    Call UniCode_Conv(PLN_S_YOTEI_R.Y_NYUKA_KEY_NO, PLN_S_YOTEI(Row, colY_NYUKA_KEY_NO))
''
''                    If Trim(PLN_S_YOTEI(Row, colY_NYUKA_KEY_NO)) = "" Then
''                        Upd_NYUKA_Com = BtOpInsert
''
''                    Else
''
''                        Call UniCode_Conv(K4_PLN_Y_NYUKA.KEY_NO, PLN_S_YOTEI(Row, colY_NYUKA_KEY_NO))
''                        sts = BTRV(BtOpGetEqual, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
''                        Select Case sts
''                            Case BtNoErr
''                                Upd_NYUKA_Com = BtOpUpdate
''                            Case BtErrKeyNotFound
''                                Upd_NYUKA_Com = BtOpInsert
''                                Call PLN_Y_NYUKA_CLR
''                            Case Else
''                                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
''                                Exit Function
''                        End Select
''                    End If
''
''                    If Upd_NYUKA_Com = BtOpInsert Then
''
''                        NYUKA_FLG = True
''
''                        Call UniCode_Conv(PLN_Y_NYUKA_R.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
''                        Call UniCode_Conv(PLN_Y_NYUKA_R.NAIGAI, "1")
''                        Call UniCode_Conv(PLN_Y_NYUKA_R.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
''
''                        Call UniCode_Conv(PLN_Y_NYUKA_R.SEQ_NO, "000")
''                        Call UniCode_Conv(PLN_Y_NYUKA_R.DATA_KB, "IN")
''                        Call UniCode_Conv(PLN_Y_NYUKA_R.KEY_NO, NYUKA_KEY_NO)
''
''                        Call UniCode_Conv(PLN_S_YOTEI_R.Y_NYUKA_KEY_NO, NYUKA_KEY_NO)
''
''
''                        '------------------------------ 品目マスタより
''                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
''                        Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
''                        Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
''
''                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
''                        Select Case sts
''                            Case BtNoErr
''                                Call UniCode_Conv(PLN_Y_NYUKA_R.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
''                            Case BtErrKeyNotFound
''
''
''                            Case Else
''                                Call File_Error(sts, BtOpGetEqual, "品目ファイル")
''                                Call Input_UnLock
''                                Exit Function
''                        End Select
''                    End If
''
''
''
''                    Call UniCode_Conv(PLN_Y_NYUKA_R.N_YOTEI_DT, Format(PLN_S_YOTEI(Row, colNYUKA_YOTEI_DT), "YYYYMMDD"))
''                    Call UniCode_Conv(PLN_Y_NYUKA_R.N_YOTEI_QTY, Format(CLng(PLN_S_YOTEI(Row, colNYUKA_YOTEI_QTY)), "00000000"))
''
''
''                    If StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode) <> Format(PLN_S_YOTEI(Row, colNYUKA_YOTEI_DT), "YYYYMMDD") Or _
''                        CLng(StrConv(PLN_Y_NYUKA_R.N_YOTEI_QTY, vbUnicode)) <> CLng(PLN_S_YOTEI(Row, colNYUKA_YOTEI_QTY)) Then
''
''                        If Upd_NYUKA_Com = BtOpUpdate Then
''
''                            NYUKA_FLG = True
''
''                            Call UniCode_Conv(PLN_Y_NYUKA_R.UPD_TANTO, App.EXEName)
''                            Call UniCode_Conv(PLN_Y_NYUKA_R.UPD_TANTO, Format(Now, "YYYYMMDDHHMMSS"))
''                        End If
''
''
''                    End If
''
''                    Do
''                        sts = BTRV(Upd_NYUKA_Com, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
''                        Select Case sts
''                            Case BtNoErr
''                                Exit Do
''                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
''
''                               Beep
''                                ans = MsgBox("商品化予定ファイル」他端末でデータ使用中です。<PLN_S_YOTEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
''                                If ans = vbCancel Then
''                                    Call Input_UnLock
''                                    Exit Function
''                                End If
''
''                            Case Else
''                                Call Input_UnLock
''                                Call File_Error(sts, Upd_Com, "商品化予定ファイル")
''                                Exit Function
''                        End Select
''
''                    Loop
''
''
''                Else
''                    If Trim(StrConv(PLN_S_YOTEI_R.Y_NYUKA_KEY_NO, vbUnicode)) <> "" Then
''                        NYUKA_FLG = True
''                    End If
''                    Call UniCode_Conv(PLN_S_YOTEI_R.Y_NYUKA_KEY_NO, "")
''                    Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, "")
''                    Call UniCode_Conv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_QTY, "00000000")
''                End If
            
            
            
            
                End If
If Trim(PLN_S_YOTEI(Row, colHIN_GAI)) = "ARR31-635" Then
Debug.Print
End If
                '品目マスタ「商品化ｼｽﾃﾑ」用標準工数更新
                Call UniCode_Conv(K0_ITEM.JGYOBU, Right(PLN_S_YOTEI(Row, colJGYOBU), 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                        '>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.05.07
                        'If Not IsNumeric(StrConv(ITEMREC.PLN_KOUSU, vbUnicode)) Then
                        '    Call UniCode_Conv(ITEMREC.PLN_KOUSU, "00000000.00")
                        'End If
                        If Not IsNumeric(StrConv(ITEMREC.PLN_SAGYOU_KOUSU, vbUnicode)) Then
                            Call UniCode_Conv(ITEMREC.PLN_SAGYOU_KOUSU, "00000000.00")
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.05.07
                    
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.05.07
                        'If Val(StrConv(ITEMREC.PLN_KOUSU, vbUnicode)) <> Val(PLN_S_YOTEI(Row, colS_KOUSU)) Then
                        '    Call UniCode_Conv(ITEMREC.PLN_KOUSU, Format(Val(PLN_S_YOTEI(Row, colS_KOUSU)), "00000000.00"))
                        If Val(StrConv(ITEMREC.PLN_SAGYOU_KOUSU, vbUnicode)) <> Val(PLN_S_YOTEI(Row, colS_KOUSU)) Then
                            Call UniCode_Conv(ITEMREC.PLN_SAGYOU_KOUSU, Format(Val(PLN_S_YOTEI(Row, colS_KOUSU)), "00000000.00"))
                        '>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.05.07
                        
                            Call UniCode_Conv(ITEMREC.UPD_TANTO, "PLN04")
                            Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                        
                        
                        
                            Do
                                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE

                                        Beep
                                        ans = MsgBox("品目マスタ」他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                        If ans = vbCancel Then
                                           Call Input_UnLock
                                            Exit Function
                                        End If

                                    Case Else
                                        Call Input_UnLock
                                        Call File_Error(sts, BtOpUpdate, "品目マスタ")
                                        Exit Function
                                End Select

                            Loop
                        
                        
                        End If
                    
                    
                    Case BtErrKeyNotFound
                    
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.05.18
                        Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                        Call UniCode_Conv(ITEMREC.ST_RETU, "")
                        Call UniCode_Conv(ITEMREC.ST_REN, "")
                        Call UniCode_Conv(ITEMREC.ST_DAN, "")
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.05.18
                    
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目ファイル")
                        Call Input_UnLock
                        Exit Function
                End Select
            
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.05.18
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.05.18
                If Upd_Com = BtOpUpdate Then
                    Call UniCode_Conv(PLN_S_YOTEI_R.UPD_TANTO, App.EXEName)
                    Call UniCode_Conv(PLN_S_YOTEI_R.UPD_TANTO, Format(Now, "YYYYMMDDHHMMSS"))
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
            

            Set TDBGrid1.Array = PLN_S_YOTEI
            TDBGrid1.ReBind

            TDBGrid1.Update
            TDBGrid1.Bookmark = Row
        
        End If


If Row = 2 Then
    Debug.Print "UPD OUT= " & PLN_S_YOTEI(Row, colS_JIKAN)
End If


    Next Row
    
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.08
    'Set TDBGrid1.Array = PLN_S_YOTEI
    'TDBGrid1.ReBind
    '
    'TDBGrid1.Update
    'TDBGrid1.MoveFirst
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.08



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定ファイル処理　処理終了！！", Me.hwnd, 0)




    Call Input_UnLock



    Update_Proc = False



End Function

Private Function List_Disp_Proc(NYUKA_FLG As Boolean) As Integer
'----------------------------------------------------------------------------
'                   「商品化予定ファイル」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim com             As Integer

Dim Skip_Flg        As Integer
Dim Row             As Long
Dim i               As Integer

Dim SHIMUKE_CODE    As String * 2       '2011.11.11
Dim c               As String * 128     '2011.11.11


    List_Disp_Proc = True

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定ファイル読込み処理　処理開始！！", Me.hwnd, 0)

                                    'テーブルリセット
    Set PLN_S_YOTEI = Nothing
    Row = Min_Row - 1


    com = BtOpGetFirst



    Do
        DoEvents
        sts = BTRV(com, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K3_PLN_S_YOTEI, Len(K3_PLN_S_YOTEI), 3)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化用入荷予定ファイル")
                Exit Function
        End Select
    
        Skip_Flg = False
        
        
        If IsDate(Text1(ptxYOTEI_DT_S).Text) Then
            If StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) < Format(Text1(ptxYOTEI_DT_S).Text, "YYYYMMDD") Then
                Skip_Flg = True
            End If
        End If
        
        If IsDate(Text1(ptxYOTEI_DT_E).Text) Then
            If StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) > Format(Text1(ptxYOTEI_DT_E).Text, "YYYYMMDD") Then
                Skip_Flg = True
            End If
        End If
        
        
        If Trim(Text1(ptxST_SOKO).Text) = "" Then
        Else
            If Text1(ptxST_SOKO).Text <> StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) Then
                Skip_Flg = True
            End If
        End If
        
        
        
        If Trim(Text1(ptxJIZEN_S).Text) = "" Then   '2011.11.10
        Else
            If Val(StrConv(PLN_S_YOTEI_R.JIZEN, vbUnicode)) < Val(Text1(ptxJIZEN_S).Text) Then
                Skip_Flg = True
            End If
        End If
        
        
        If Trim(Text1(ptxJIZEN_E).Text) = "" Then   '2011.11.10
        Else
            If Val(StrConv(PLN_S_YOTEI_R.JIZEN, vbUnicode)) > Val(Text1(ptxJIZEN_E).Text) Then
                Skip_Flg = True
            End If
        End If
        
        
        If IsDate(Text1(ptxNYUKA_YOTEI_DT_S).Text) Then
            If StrConv(PLN_S_YOTEI_R.NYUKA_YOTEI_DT, vbUnicode) < Format(Text1(ptxNYUKA_YOTEI_DT_S).Text, "YYYYMMDD") Then
                Skip_Flg = True
            End If
        End If
        If IsDate(Text1(ptxNYUKA_YOTEI_DT_E).Text) Then
            If StrConv(PLN_S_YOTEI_R.NYUKA_YOTEI_DT, vbUnicode) > Format(Text1(ptxNYUKA_YOTEI_DT_E).Text, "YYYYMMDD") Then
                Skip_Flg = True
            End If
        End If
        
        '2011.12.19
        If Trim(Text1(ptxHIN_GAI).Text) = "" Then
        Else
            If Trim(Text1(ptxHIN_GAI).Text) <> Left(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), Len(Trim(Text1(ptxHIN_GAI).Text))) Then
                Skip_Flg = True
            End If
        End If
        '2011.12.19
        
        
        If Skip_Flg Then
        Else
            Row = Row + 1
            PLN_S_YOTEI.ReDim Min_Row, Row, Min_Col, Max_Col
            PLN_S_YOTEI(Row, colSHORI) = False
            
            For i = 0 To UBound(JGYOBU_T)
                If StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
                    PLN_S_YOTEI(Row, colJGYOBU) = JGYOBU_T(i).NAME & "          " & JGYOBU_T(i).CODE
                    Exit For
                End If
            Next i

            
            PLN_S_YOTEI(Row, colHIN_GAI) = Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
            PLN_S_YOTEI(Row, colST_TANABAN) = StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(PLN_S_YOTEI_R.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(PLN_S_YOTEI_R.ST_REN, vbUnicode) & "-" & _
                                                StrConv(PLN_S_YOTEI_R.ST_DAN, vbUnicode)



            '2011.11.11 個装資材追加
            PLN_S_YOTEI(Row, colSIZAI) = Trim(StrConv(PLN_S_YOTEI_R.SIZAI, vbUnicode))


'            PLN_S_YOTEI(Row, colJIZEN) = Format(Val(StrConv(PLN_S_YOTEI_R.JIZEN, vbUnicode)), "#0") & "%"
            
            PLN_S_YOTEI(Row, colJIZEN) = Format(Val(StrConv(PLN_S_YOTEI_R.JIZEN, vbUnicode)), "#0")
            
            PLN_S_YOTEI(Row, colJIZEN_NEEDS_QTY) = Format(CLng(StrConv(PLN_S_YOTEI_R.JIZEN_NEEDS_QTY, vbUnicode)), "#,##0")


    
           If IsNumeric(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)) Then
               PLN_S_YOTEI(Row, colYOTEI_QTY) = Trim(Format(CLng(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)), "#,###"))
               PLN_S_YOTEI(Row, colBEF_YOTEI_QTY) = Trim(Format(CLng(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)), "#,###"))
           Else
               PLN_S_YOTEI(Row, colYOTEI_QTY) = ""
               PLN_S_YOTEI(Row, colBEF_YOTEI_QTY) = ""
           End If
        
            If IsNumeric(StrConv(PLN_S_YOTEI_R.Z_QTY_S, vbUnicode)) Then
                PLN_S_YOTEI(Row, colZ_QTY_S) = Format(CLng(StrConv(PLN_S_YOTEI_R.Z_QTY_S, vbUnicode)), "#,##0")
            Else
                PLN_S_YOTEI(Row, colZ_QTY_S) = "0"
            End If
            
            If IsNumeric(StrConv(PLN_S_YOTEI_R.Z_QTY_MI, vbUnicode)) Then
                PLN_S_YOTEI(Row, colZ_QTY_MI) = Format(CLng(StrConv(PLN_S_YOTEI_R.Z_QTY_MI, vbUnicode)), "#,##0")
            Else
                PLN_S_YOTEI(Row, colZ_QTY_MI) = "0"
            End If
        
            If Len(Trim(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode))) = 8 Then
                PLN_S_YOTEI(Row, colYOTEI_DT) = Mid(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode), 7, 2)

            Else
                PLN_S_YOTEI(Row, colYOTEI_DT) = ""
            End If
        
        
        
        
        
''           Call UniCode_Conv(K0_ITEM.JGYOBU, PLN_S_YOTEI(Row, colJGYOBU))
''           Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
''           Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_S_YOTEI(Row, colHIN_GAI))
''           sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
''            Select Case sts
''                Case BtNoErr
''                   If IsNumeric(StrConv(ITEMREC.PLN_KOUSU, vbUnicode)) Then
''                        PLN_S_YOTEI(Row, colS_KOUSU) = Format(Val(StrConv(ITEMREC.PLN_KOUSU, vbUnicode)), "#0.0")
''                    Else
'''''''2012.05.02                        PLN_S_YOTEI(Row, colS_KOUSU) = Format(Val(StrConv(PLN_S_YOTEI_R.S_KOUSU, vbUnicode)), "#0.0")
                        
                               
                        PLN_S_YOTEI(Row, colS_KOUSU) = Format(Val(StrConv(PLN_S_YOTEI_R.SAGYOU_KOUSU, vbUnicode)), "#0.0")       '2012.05.02 見積工数--＞作業工数（商品化工数）
                        
                        
                        PLN_S_YOTEI(Row, colBEF_S_KOUSU) = Format(Val(StrConv(PLN_S_YOTEI_R.S_KOUSU, vbUnicode)), "#0.0")
''                    End If
''                Case BtErrKeyNotFound
''                    PLN_S_YOTEI(Row, colS_KOUSU) = Format(Val(StrConv(PLN_S_YOTEI_R.S_KOUSU, vbUnicode)), "#0.0")
''                Case Else
''                   Call File_Error(sts, BtOpGetEqual, "品目マスタ")
''                    Exit Function
''            End Select
            
            PLN_S_YOTEI(Row, colS_JIKAN) = Format(Val(StrConv(PLN_S_YOTEI_R.S_JIKAN, vbUnicode)), "#0.0")
        
        
If Row = 2 Then
    Debug.Print "DISP IN= " & PLN_S_YOTEI(Row, colS_JIKAN)
End If
        
        
        
            If Len(Trim(StrConv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, vbUnicode))) = 8 Then
                
''                If StrConv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, vbUnicode) = "99999999" Then
''                    PLN_S_YOTEI(Row, colNYUKA_YOTEI_DT) = ""
''                Else
                    PLN_S_YOTEI(Row, colNYUKA_YOTEI_DT) = Mid(StrConv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, vbUnicode), 1, 4) & "/" & _
                                                    Mid(StrConv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, vbUnicode), 5, 2) & "/" & _
                                                    Mid(StrConv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_DT, vbUnicode), 7, 2)
''                End If
            Else
                PLN_S_YOTEI(Row, colNYUKA_YOTEI_DT) = ""
            End If
            
            If IsNumeric(StrConv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_QTY, vbUnicode)) Then
                PLN_S_YOTEI(Row, colNYUKA_YOTEI_QTY) = Trim(Format(CLng(StrConv(PLN_S_YOTEI_R.INP_NYUKA_YOTEI_QTY, vbUnicode)), "#,###"))
            Else
                PLN_S_YOTEI(Row, colNYUKA_YOTEI_QTY) = ""
            End If
        
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2011.11.11 指図票発行日付＆完了（分納）日付の獲得
            If GetIni(App.EXEName, Right(PLN_S_YOTEI(Row, colJGYOBU), 1), App.EXEName, c) Then
                SHIMUKE_CODE = ""
            Else
                SHIMUKE_CODE = Trim(c)
            End If
            
            
            
            
            Call UniCode_Conv(K4_P_SSHIJI_O.SHIMUKE_CODE, SHIMUKE_CODE)
            Call UniCode_Conv(K4_P_SSHIJI_O.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
            Call UniCode_Conv(K4_P_SSHIJI_O.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
            Call UniCode_Conv(K4_P_SSHIJI_O.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K4_P_SSHIJI_O.Print_datetime, StrConv(PLN_S_YOTEI_R.Ins_DateTime, vbUnicode))
            
            com = BtOpGetGreaterEqual
            
            Call UniCode_Conv(PLN_S_YOTEI_R.SASIZU_DateTime, "")
            Call UniCode_Conv(PLN_S_YOTEI_R.S_KAN_DateTime, "")
            
            
            Do
                DoEvents
                sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K4_P_SSHIJI_O, Len(K4_P_SSHIJI_O), 4)
                Select Case sts
                    Case BtNoErr
                            
                        If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_CODE Then
                            Exit Do
                        End If
                            
                        If StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) <> StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) Or _
                            StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) <> StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode) Or _
                            Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) Then
                            Exit Do
                        End If
                            
                            
                        If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                        Else
                            Call UniCode_Conv(PLN_S_YOTEI_R.SASIZU_DateTime, StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode))
                            If Trim(StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode)) = P_KAN_ON Then
                                Call UniCode_Conv(PLN_S_YOTEI_R.S_KAN_DateTime, StrConv(P_SSHIJI_O_REC.KAN_DT, vbUnicode))
                            Else
                                Call UniCode_Conv(K3_P_SUKEIRE.SHIJI_No, StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode))
                                Call UniCode_Conv(K3_P_SUKEIRE.UKEIRE_DT, "")
                                    
                                com = BtOpGetGreaterEqual
                                
                                Do
                                    DoEvents
                                    sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K3_P_SUKEIRE, Len(K3_P_SUKEIRE), 3)
                                    Select Case sts
                                        Case BtNoErr
                                            If StrConv(P_SUKEIRE_REC.SHIJI_No, vbUnicode) <> StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode) Then
                                                Exit Do
                                            End If
                                                        
                                            Call UniCode_Conv(PLN_S_YOTEI_R.S_KAN_DateTime, StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode))
                                
                                        
                                        Case BtErrEOF
                                            Exit Do
                                        Case Else
                                            Call File_Error(sts, com, "商品化指図受入履歴データ")
                                            Exit Function
                                    End Select
                                    com = BtOpGetNext
                                Loop
                                    
                            End If
                        End If
                        
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com, "商品化指図データ（親）")
                        Exit Function
                End Select
            
                com = BtOpGetNext
            
            Loop
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2011.11.11 指図票発行日付＆完了（分納）日付の獲得
            
        
            If Len(Trim(StrConv(PLN_S_YOTEI_R.SASIZU_DateTime, vbUnicode))) >= 8 Then
                PLN_S_YOTEI(Row, colSASIZU_DateTime) = Mid(StrConv(PLN_S_YOTEI_R.SASIZU_DateTime, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(PLN_S_YOTEI_R.SASIZU_DateTime, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(PLN_S_YOTEI_R.SASIZU_DateTime, vbUnicode), 7, 2)
        
            Else
                PLN_S_YOTEI(Row, colSASIZU_DateTime) = ""
            End If
            If Len(Trim(StrConv(PLN_S_YOTEI_R.S_KAN_DateTime, vbUnicode))) >= 8 Then
                PLN_S_YOTEI(Row, colS_KAN_DateTime) = Mid(StrConv(PLN_S_YOTEI_R.S_KAN_DateTime, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(PLN_S_YOTEI_R.S_KAN_DateTime, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(PLN_S_YOTEI_R.S_KAN_DateTime, vbUnicode), 7, 2)
        
            Else
                PLN_S_YOTEI(Row, colS_KAN_DateTime) = ""
            End If
        
                    
            PLN_S_YOTEI(Row, colKEY_No) = StrConv(PLN_S_YOTEI_R.KEY_NO, vbUnicode)
                    
''2011.10.04            PLN_S_YOTEI(Row, colY_NYUKA_KEY_NO) = StrConv(PLN_S_YOTEI_R.Y_NYUKA_KEY_NO, vbUnicode)
        
        
            
            
                        
            PLN_S_YOTEI(Row, colBEF_YOTEI_QTY) = PLN_S_YOTEI(Row, colYOTEI_QTY)         '2011.11.30
            PLN_S_YOTEI(Row, colBEF_S_KOUSU) = PLN_S_YOTEI(Row, colS_KOUSU)             '2011.11.30
        
        
        
        
        
        
        
        
        
        End If

        com = BtOpGetNext

If Row = 2 Then
    Debug.Print "DISP OUT= " & PLN_S_YOTEI(Row, colS_JIKAN)
End If

    Loop


    Set TDBGrid1.Array = PLN_S_YOTEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "件"


    For i = 0 To UBound(Z_Sort_Tbl)
        Z_Sort_Tbl(i) = 0                 'ﾃﾞﾌｫﾙﾄ昇順
    Next i


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化用入荷予定ファイル　[検索]処理終了！！", Me.hwnd, 0)



    If NYUKA_FLG Then
        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
            "商品化予定ファイル処理　処理終了！！「商品化用入荷予定情報が変更されました。商品化用入荷予定メンテナンス画面で確認して下さい。」", Me.hwnd, 0)
    End If



    Call Input_UnLock


    List_Disp_Proc = False
    Exit Function

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    PLN00401.MousePointer = vbHourglass

    Call Ctrl_Lock(PLN00401)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PLN00401)


    PLN00401.MousePointer = vbDefault

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

' ------------------------------------------------------------------------
'       指定した精度の数値に切り上げします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り上げられた数値。
' ------------------------------------------------------------------------
Private Function ToRoundUp(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    
        


    dCoef = (10 ^ iDigits)



    
    
    
    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundUp = (Int(dValue * dCoef) + 1) / dCoef
        Case Is < 0
            ToRoundUp = (Fix(dValue * dCoef) - 1) / dCoef
        Case Else
            ToRoundUp = dValue
    End Select


'    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
'        Case Is > 0
'            ToRoundUp = (Int(dValue * dCoef + 0.9)) / dCoef
'        Case Is < 0
'            ToRoundUp = (Fix(dValue * dCoef - 0.9)) / dCoef
'        Case Else
'            ToRoundUp = dValue
'    End Select



End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に切り捨てします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り捨てられた数値。
' ------------------------------------------------------------------------
Private Function ToRoundDown(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundDown = Int(dValue * dCoef) / dCoef
        Case Is < 0
            ToRoundDown = Fix(dValue * dCoef) / dCoef
        Case Else
            ToRoundDown = dValue
    End Select
End Function





' ------------------------------------------------------------------------
'       指定した精度の数値に四捨五入します。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に四捨五入された数値。
' ------------------------------------------------------------------------
Private Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function



Private Function Err_Check_Proc() As Integer
Dim sts         As Integer
Dim Bookmark    As Variant
    
    
Dim i           As Integer
    
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
    
    
    Err_Check_Proc = True
    
    Set TDBGrid1.Array = PLN_S_YOTEI
    TDBGrid1.Update
    
    
    
    For i = 1 To PLN_S_YOTEI.UpperBound(1)
    
    
If i = 2 Then
    Debug.Print "ERR IN= " & PLN_S_YOTEI(i, colS_JIKAN)
End If
        
        If PLN_S_YOTEI(i, colSHORI) Then
        Else
            'BU
            If Trim(PLN_S_YOTEI(i, colJGYOBU)) = "" Then
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(ＢＵ　必須入力)"
                TDBGrid1.SetFocus
                Exit Function
            End If
                '品番
            Call UniCode_Conv(K0_ITEM.JGYOBU, Right(PLN_S_YOTEI(i, colJGYOBU), 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
            Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_S_YOTEI(i, colHIN_GAI))
                
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    
                    PLN_S_YOTEI(i, colST_TANABAN) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)

                    
                    If Trim(PLN_S_YOTEI(i, colS_KOUSU)) = "" Then
                        
                        '>>>>>>>>>>>>>>>>   2012.05.07
                        'If IsNumeric(StrConv(ITEMREC.PLN_KOUSU, vbUnicode)) Then
                        '    PLN_S_YOTEI(i, colS_KOUSU) = Format(StrConv(ITEMREC.PLN_KOUSU, vbUnicode), "#0.0")
                        If IsNumeric(StrConv(ITEMREC.PLN_SAGYOU_KOUSU, vbUnicode)) Then
                            PLN_S_YOTEI(i, colS_KOUSU) = Format(StrConv(ITEMREC.PLN_SAGYOU_KOUSU, vbUnicode), "#0.0")
                        '>>>>>>>>>>>>>>>>   2012.05.07
                        Else
                            PLN_S_YOTEI(i, colS_KOUSU) = "0.0"
                        End If
                    
                        If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                                StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                            Unload Me
                        End If
                    
                        PLN_S_YOTEI(i, colZ_QTY_S) = Format(SUMI_QTY, "#,##0")
                        PLN_S_YOTEI(i, colZ_QTY_MI) = Format(MI_QTY, "#,##0")
                                                
                    End If
                    
                Case BtErrKeyNotFound
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(品番未登録)"
                    TDBGrid1.SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                    Unload Me
                
            End Select
    
    
            If Trim(PLN_S_YOTEI(i, colYOTEI_QTY)) = "" Then
            Else
                If Not IsNumeric(PLN_S_YOTEI(i, colYOTEI_QTY)) Then
                            
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(商品化予定数)"
                    TDBGrid1.SetFocus
                    Exit Function
                End If
            
            
                If Val(PLN_S_YOTEI(i, colYOTEI_QTY)) < 1 Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(商品化予定数≦０)"
                    TDBGrid1.SetFocus
                    Exit Function
                End If
            End If
                
            If IsNumeric(PLN_S_YOTEI(i, colYOTEI_QTY)) Then
                
                
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   工数時間の自動計算制御2011.11.30
                    
                If Trim(PLN_S_YOTEI(i, colYOTEI_QTY)) = Trim(PLN_S_YOTEI(i, colBEF_YOTEI_QTY)) And _
                    Trim(PLN_S_YOTEI(i, colS_KOUSU)) = Trim(PLN_S_YOTEI(i, colBEF_S_KOUSU)) Then
                                
                
                Else
                
                    If IsNumeric(PLN_S_YOTEI(i, colBEF_YOTEI_QTY)) And IsNumeric(PLN_S_YOTEI(i, colYOTEI_QTY)) Then
                        PLN_S_YOTEI(i, colS_JIKAN) = Format(Round(Val(PLN_S_YOTEI(i, colYOTEI_QTY)) * Val(PLN_S_YOTEI(i, colS_KOUSU)), 2), "#0.0")
                                    
                        PLN_S_YOTEI(i, colBEF_YOTEI_QTY) = PLN_S_YOTEI(i, colYOTEI_QTY)
                        PLN_S_YOTEI(i, colBEF_S_KOUSU) = PLN_S_YOTEI(i, colBEF_S_KOUSU)
                
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   工数時間の自動計算制御2011.11.30
                
                PLN_S_YOTEI(i, colYOTEI_QTY) = Format(CLng(PLN_S_YOTEI(i, colYOTEI_QTY)), "#,###")
''2011.11.30                PLN_S_YOTEI(i, colS_JIKAN) = Format(Round(Val(PLN_S_YOTEI(i, colYOTEI_QTY)) * Val(PLN_S_YOTEI(i, colS_KOUSU)), 2), "#0.0")
            End If

    
    
            If Trim(PLN_S_YOTEI(i, colYOTEI_DT)) = "" Then
''2011.10.31                If IsNumeric(PLN_S_YOTEI(i, colYOTEI_QTY)) Then
''2011.10.31                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。([商品化予定日]商品化予定数入力時、必須入力)"
''2011.10.31                    TDBGrid1.SetFocus
''2011.10.31                    Exit Function
''2011.10.31                End If
            Else
                If Trim(PLN_S_YOTEI(i, colYOTEI_QTY)) = "" Then
                    
''2011.10.31                    If Trim(PLN_S_YOTEI(i, colYOTEI_DT)) <> "" Then
''2011.10.31                        MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。([商品化予定日]商品化予定数未入力時、入力不可)"
''2011.10.31                        TDBGrid1.SetFocus
''2011.10.31                        Exit Function
''2011.10.31                    End If
                    
                Else
                    If Not IsDate(PLN_S_YOTEI(i, colYOTEI_DT)) Then
                        MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(商品化予定日)"
                        TDBGrid1.SetFocus
                        Exit Function
                    End If
            
''2011.10.31                    If PLN_S_YOTEI(i, colYOTEI_DT) < Format(Now, "YYYY/MM/DD") Then
''2011.10.31                        MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(商品化予定日＜本日)"
''2011.10.31                        TDBGrid1.SetFocus
''2011.10.31                        Exit Function
''2011.10.31                    End If
            
                End If
            End If
    
            If Trim(PLN_S_YOTEI(i, colS_KOUSU)) = "" Then
                PLN_S_YOTEI(i, colS_KOUSU) = "0.0"
            End If
    
    
            If Not IsNumeric(PLN_S_YOTEI(i, colS_KOUSU)) Then
            
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(標準工数)"
                TDBGrid1.SetFocus
            Exit Function
            
            End If
    
            If Val(PLN_S_YOTEI(i, colS_KOUSU)) < 0 Then
            
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(標準工数＜０)"
                TDBGrid1.SetFocus
                Exit Function
            
            End If
    
    
    
    
    
    
''2011.11.30            PLN_S_YOTEI(i, colS_JIKAN) = Format(Round(Val(PLN_S_YOTEI(i, colYOTEI_QTY)) * Val(PLN_S_YOTEI(i, colS_KOUSU)), 2), "#0.0")
    
    
            If Trim(PLN_S_YOTEI(i, colNYUKA_YOTEI_DT)) <> "" Then
                If Not IsDate(PLN_S_YOTEI(i, colNYUKA_YOTEI_DT)) Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(入荷予定日)"
                    TDBGrid1.SetFocus
                    Exit Function
                End If
        
''2011.10.31                If PLN_S_YOTEI(i, colNYUKA_YOTEI_DT) < Format(Now, "YYYY/MM/DD") Then
''2011.10.31                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(入荷予定日＜本日)"
''2011.10.31                    TDBGrid1.SetFocus
''2011.10.31                    Exit Function
''2011.10.31                End If
            End If
    
    
            If Trim(PLN_S_YOTEI(i, colNYUKA_YOTEI_QTY)) = "" Then
''2011.10.31                If IsDate(PLN_S_YOTEI(i, colNYUKA_YOTEI_DT)) Then
''2011.10.31                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。([入荷予定数]入荷予定日入力時、必須入力)"
''2011.10.31                    Exit Function
''2011.10.31                End If
            Else
                    
''2011.10.31                If Trim(PLN_S_YOTEI(i, colNYUKA_YOTEI_DT)) = "" Then
''2011.10.31                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。([入荷予定数]入荷予定日未入力時、入力不可)"
''2011.10.31                    Exit Function
''2011.10.31                End If
                    
                If Not IsNumeric(PLN_S_YOTEI(i, colNYUKA_YOTEI_QTY)) Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(入荷予定数)"
                    Exit Function
                End If
            
                If Val(PLN_S_YOTEI(i, colNYUKA_YOTEI_QTY)) < 1 Then
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(入荷予定数≦０)"
                    Exit Function
                End If
            
                PLN_S_YOTEI(i, colNYUKA_YOTEI_QTY) = Format(CLng(PLN_S_YOTEI(i, colNYUKA_YOTEI_QTY)), "#,###")
            
            End If
    
    
    
    
    
        End If
    
If i = 2 Then
    Debug.Print "ERR OUT= " & PLN_S_YOTEI(i, colS_JIKAN)
End If
    
    
    Next i
        
    Set TDBGrid1.Array = PLN_S_YOTEI
        
    
    TDBGrid1.Refresh
    TDBGrid1.Update
    TDBGrid1.SetFocus

    Err_Check_Proc = False

End Function

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
    
    Select Case Index
        Case ptxHIN_GAI
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
    End Select

End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Select Case Index
        Case ptxHIN_GAI
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
    End Select
End Sub
Private Function OutPut_Proc() As Integer
'-------------------------------------------------------------------
'
'   商品化予定データ出力
'
'       2012.04.24
'-------------------------------------------------------------------
Dim Row             As Long

Dim FileNo          As Integer



    OutPut_Proc = True

    Call Input_Lock         '画面項目ロック解除


    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    
    Open (PLN0040CSV) For Output As FileNo



                
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "商品化予定データ出力中", Me.hwnd, 0)
                
                
    
    For Row = 1 To PLN_S_YOTEI.UpperBound(1)
    
        DoEvents
    
        If Row = 1 Then
                
            Write #FileNo, "ＢＵ", "対外品番", "標準棚番", "個装資材", "事前商品化状況％", "事前商品化必要数", _
                            "商品化予定数", "在庫数　済", "在庫数　未", "商品化予定日", "標準工数", "標準時間", _
                             "部品入荷予定日", "部品入荷予定数", "指図票発行日", "完了日"
        
        End If
        'BU
        Write #FileNo, PLN_S_YOTEI(Row, colJGYOBU),
        '対外品番
        Write #FileNo, PLN_S_YOTEI(Row, colHIN_GAI),
        '標準棚番
        Write #FileNo, PLN_S_YOTEI(Row, colST_TANABAN),
        '個装資材
        Write #FileNo, PLN_S_YOTEI(Row, colSIZAI),
        '事前商品化状況％
        Write #FileNo, PLN_S_YOTEI(Row, colJIZEN),
        '事前商品化必要数
        Write #FileNo, PLN_S_YOTEI(Row, colJIZEN_NEEDS_QTY),
        '商品化予定数
        Write #FileNo, PLN_S_YOTEI(Row, colYOTEI_QTY),
        '在庫数　済
        Write #FileNo, PLN_S_YOTEI(Row, colZ_QTY_S),
        '在庫数　未
        Write #FileNo, PLN_S_YOTEI(Row, colZ_QTY_MI),
        '商品化予定日
        Write #FileNo, PLN_S_YOTEI(Row, colYOTEI_DT),
        '標準工数
        Write #FileNo, PLN_S_YOTEI(Row, colS_KOUSU),
        '標準時間
        Write #FileNo, PLN_S_YOTEI(Row, colS_JIKAN),
        '部品入荷予定日
        Write #FileNo, PLN_S_YOTEI(Row, colNYUKA_YOTEI_DT),
        '部品入荷予定数
        Write #FileNo, PLN_S_YOTEI(Row, colNYUKA_YOTEI_QTY),
        '指図票発行日
        Write #FileNo, PLN_S_YOTEI(Row, colSASIZU_DateTime),
        '完了日
        Write #FileNo, PLN_S_YOTEI(Row, colS_KAN_DateTime),
        
        Write #FileNo,
    Next Row
                

    Close #FileNo
    MsgBox "「" & PLN0040CSV & "」は正常に出力されました。"
    
    
    Call Input_UnLock         '画面項目ロック解除


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "商品化予定データ出力終了", Me.hwnd, 0)

    OutPut_Proc = False

    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox PLN0040CSV & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        OutPut_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
    End If

End Function


