VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form SEI00162 
   Caption         =   "[請求システム]見積書作成処理"
   ClientHeight    =   12465
   ClientLeft      =   2025
   ClientTop       =   -3210
   ClientWidth     =   12600
   ControlBox      =   0   'False
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
   ScaleHeight     =   12465
   ScaleWidth      =   12600
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   41
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   40
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   39
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   42
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   38
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   37
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   36
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   35
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1215
   End
   Begin TrueDBGrid80.TDBDropDown TDBDropDown2 
      Height          =   2055
      Left            =   4440
      TabIndex        =   67
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3625
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).ValueItems(0)._DefaultItem=   0
      Columns(0).ValueItems(0).Value=   "aaaa"
      Columns(0).ValueItems(0).Value.vt=   8
      Columns(0).ValueItems(0).DisplayValue=   "aaaa"
      Columns(0).ValueItems(0).DisplayValue.vt=   8
      Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems.Count=   1
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
      DataField       =   "ub_grid2"
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
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   31
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   11040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   30
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   11040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   29
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   10680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   28
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   10680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   27
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   10320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   26
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   25
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   9960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   24
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   23
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   22
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   9600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   21
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   9240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   20
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   9240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   19
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   18
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8880
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   17
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   16
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   15
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   14
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   13
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   12
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   11
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Index           =   10
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
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
      Index           =   34
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
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
      Index           =   33
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
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
      Index           =   32
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Index           =   0
      Left            =   480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7590
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"SEI00162.frx":0000
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   11040
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "閉じる"
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
      Left            =   480
      TabIndex        =   39
      Top             =   0
      Width           =   1215
   End
   Begin TrueDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2055
      Left            =   1680
      TabIndex        =   48
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3625
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).ValueItems(0)._DefaultItem=   0
      Columns(0).ValueItems(0).Value=   "aaaa"
      Columns(0).ValueItems(0).Value.vt=   8
      Columns(0).ValueItems(0).DisplayValue=   "aaaa"
      Columns(0).ValueItems(0).DisplayValue.vt=   8
      Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems.Count=   1
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
      DataField       =   "ub_grid2"
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
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   3975
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   2880
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7011
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   1
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "　　種別"
      Columns(1).DataField=   ""
      Columns(1).DropDown=   "TDBDropDown1"
      Columns(1).DropDown.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   1
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "　事業部"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "TDBDropDown2"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "　 構成品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "       品　名"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "  員数"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "仕入＠"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "販売＠"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "仕入金額計"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "子有無"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=900"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=794"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8193"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2196"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2090"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Button=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1958"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1852"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Button=1"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2831"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2725"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=3757"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=3651"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=8192"
      Splits(0)._ColumnProps(27)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=1508"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=1402"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=1879"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1773"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=2143"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2037"
      Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=2117"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2011"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=8194"
      Splits(0)._ColumnProps(48)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(49)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(50)=   "Column(9).Width=3810"
      Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=3704"
      Splits(0)._ColumnProps(53)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=0,.bgcolor=&H80000016&"
      _StyleDefs(53)  =   ":id=46,.locked=-1"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1,.bgcolor=&H80000005&"
      _StyleDefs(62)  =   ":id=54,.locked=0"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=1,.bgcolor=&H80000005&"
      _StyleDefs(67)  =   ":id=58,.locked=0"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
      _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=1,.bgcolor=&H8000000F&"
      _StyleDefs(72)  =   ":id=62,.locked=-1"
      _StyleDefs(73)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(80)  =   "Named:id=33:Normal"
      _StyleDefs(81)  =   ":id=33,.parent=0"
      _StyleDefs(82)  =   "Named:id=34:Heading"
      _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(84)  =   ":id=34,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=35:Footing"
      _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   "Named:id=36:Selected"
      _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=37:Caption"
      _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(91)  =   "Named:id=38:HighlightRow"
      _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=39:EvenRow"
      _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(95)  =   "Named:id=40:OddRow"
      _StyleDefs(96)  =   ":id=40,.parent=33"
      _StyleDefs(97)  =   "Named:id=41:RecordSelector"
      _StyleDefs(98)  =   ":id=41,.parent=34"
      _StyleDefs(99)  =   "Named:id=42:FilterBar"
      _StyleDefs(100) =   ":id=42,.parent=33"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   975
      Index           =   1
      Left            =   480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9630
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1720
      _Version        =   393217
      BackColor       =   -2147483633
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"SEI00162.frx":00BE
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Index           =   9
      Left            =   10770
      MaxLength       =   8
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "標準棚番"
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
      Left            =   8760
      TabIndex        =   78
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblSHIMUKE 
      Height          =   255
      Left            =   1440
      TabIndex        =   77
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label lblCATEGORY_NAME 
      Height          =   255
      Left            =   2520
      TabIndex        =   76
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label lblGOUKEI_KIN 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10680
      TabIndex        =   35
      Top             =   11760
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "金　額"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   10680
      TabIndex        =   66
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "管理費(5%)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   8400
      TabIndex        =   65
      Top             =   11040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "梱包ASSY"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   8400
      TabIndex        =   64
      Top             =   10680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "副資材"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   8400
      TabIndex        =   63
      Top             =   10320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "梱包材"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   8400
      TabIndex        =   62
      Top             =   9960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "設置工事説明書"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   8400
      TabIndex        =   61
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "品番表示ﾗﾍﾞﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   8400
      TabIndex        =   60
      Top             =   9240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "PE資材"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   8400
      TabIndex        =   59
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "PE加工"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   58
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "PF加工"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   8400
      TabIndex        =   57
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "単位"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   19
      Left            =   9840
      TabIndex        =   56
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "作業工程"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   8400
      TabIndex        =   55
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "商品化工料"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   43
      Left            =   8400
      TabIndex        =   54
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "中西工料"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   44
      Left            =   8400
      TabIndex        =   53
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "合計金額"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   8280
      TabIndex        =   52
      Top             =   11760
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "品名ｶﾃｺﾞﾘｰ"
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
      Index           =   63
      Left            =   240
      TabIndex        =   51
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "見積書備考"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   100
      Left            =   510
      TabIndex        =   49
      Top             =   9450
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "指図票備考"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   81
      Left            =   480
      TabIndex        =   47
      Top             =   7410
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   13440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "担当者"
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
      Left            =   720
      TabIndex        =   46
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "-"
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
      Left            =   11280
      TabIndex        =   45
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "-"
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
      Left            =   10680
      TabIndex        =   44
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "-"
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
      Left            =   10080
      TabIndex        =   43
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "親品番"
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
      Left            =   480
      TabIndex        =   42
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "仕向先"
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
      TabIndex        =   41
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "単価切替日"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   104
      Left            =   9810
      TabIndex        =   50
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "閉じる"
         Index           =   0
         Shortcut        =   {F12}
      End
      Begin VB.Menu SHORI 
         Caption         =   "検索"
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu SHORI 
         Caption         =   "保存"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "単価計算"
         Index           =   3
         Shortcut        =   {F9}
      End
      Begin VB.Menu SHORI 
         Caption         =   "見積書発行"
         Index           =   4
      End
      Begin VB.Menu SHORI 
         Caption         =   "単価登録"
         Index           =   5
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   6
      End
   End
End
Attribute VB_Name = "SEI00162"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'------------------------------------   'テキスト定義

Private Const ptxTanto_Code% = 0            '担当者コード
Private Const ptxTanto_Name% = 1            '担当者名称
Private Const ptxHin_Gai% = 2               '品番
Private Const ptxHin_Name% = 3              '品名

Private Const ptxST_SOKO% = 4               '標準棚番　 倉庫
Private Const ptxST_RETU% = 5               '標準棚番   列
Private Const ptxST_REN% = 6                '標準棚番　 連
Private Const ptxST_DAN% = 7                '標準棚番　 段

Private Const ptxCATEGORY_CODE% = 8         '品名ｶﾃｺﾞﾘｰｺｰﾄﾞ

Private Const ptxTANKA_KIRIKAE_DT% = 9      '単価切替日

Private Const ptxNAKANISHI_TANI% = 10       '中西工料　単位
Private Const ptxNAKANISHI_KIN% = 11        '中西工料　金額

Private Const ptxSHOHIN_TANI% = 12          '商品化工料　単位
Private Const ptxSHOHIN_KIN% = 13           '商品化工料　金額

Private Const ptxPF_KAKOU_TANI% = 14        'PF加工　単位
Private Const ptxPF_KAKOU_KIN% = 15         'PF加工　金額

Private Const ptxPE_KAKOU_TANI% = 16        'PE加工　単位
Private Const ptxPE_KAKOU_KIN% = 17         'PE加工　金額

Private Const ptxPE_SHIZAI_TANI% = 18       'PF資材　単位
Private Const ptxPE_SHIZAI_KIN% = 19        'PF資材　金額

Private Const ptxHINBAN_LABEL_TANI% = 20    '品番表示ﾗﾍﾞﾙ　単位
Private Const ptxHINBAN_LABEL_KIN% = 21    '品番表示ﾗﾍﾞﾙ　金額

Private Const ptxKOUJI_SETSU_TANI% = 22     '設置工事説明書　単位
Private Const ptxKOUJI_SETSU_KIN% = 23     '設置工事説明書　金額

Private Const ptxKONPOU_TANI% = 24          '梱包材　単位
Private Const ptxKONPOU_KIN% = 25           '梱包材　金額

Private Const ptxFUKU_SHIZAI_TANI% = 26     '副資材　単位
Private Const ptxFUKU_SHIZAI_KIN% = 27      '副資材　金額

Private Const ptxKONPOU_ASSY_TANI% = 28     '梱包ASSY　単位
Private Const ptxKONPOU_ASSY_KIN% = 29      '梱包ASSY　金額

Private Const ptxKANRI_TANI% = 30           '管理費　単位
Private Const ptxKANRI_KIN% = 31            '管理費　金額




Private Const ptxS_CLASS_CODE% = 32        '商品化ｸﾗｽ
Private Const ptxF_CLASS_CODE% = 33        '付加ｸﾗｽ
Private Const ptxN_CLASS_CODE% = 34        '内職ｸﾗｽ

Private Const ptxOYA_SYUBETSU% = 35          '親　種別
Private Const ptxOYA_JGYOBU% = 36            '親　事業部
Private Const ptxOYA_S_HIN_GAI% = 37         '親　指図票品番
Private Const ptxOYA_HIN_NAME% = 38          '親　品名
Private Const ptxOYA_QTY% = 39               '親　員数
Private Const ptxOYA_ST_SHITAN% = 40         '親　仕入＠
Private Const ptxOYA_ST_URITAN% = 41         '親　売上＠
Private Const ptxOYA_KINGAKU% = 42           '親　合計金額


'------------------------------------   'コンボ定義
Private Const pcmbSHIMUKE% = 0          '仕向け先
Private Const pcmbCATEGORY_Name% = 1    '品名ｶﾃｺﾞﾘｰ


'------------------------------------   'リッチテキストボックス定義
Private Const prchBIKOU% = 0            '備考
Private Const prchM_BIKOU% = 1          '見積書備考



'------------------------------------   '構成品
Private Const pGrdKOUSEI% = 0


Private Const Min_Row% = 1              '最小行数

Dim Max_Row   As Integer                'グリッド最大表示件数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 9              '最大列数



Private Const ColNO% = 0                '№
Private Const ColKO_SYUBETSU% = 1       '種別
Private Const ColKO_JGYOBU% = 2         '事業部
Private Const ColKO_S_HIN_GAI% = 3      '指図票品番
Private Const ColKO_HIN_NAME% = 4       '品名
Private Const ColKO_QTY% = 5            '員数
Private Const ColG_ST_SHITAN% = 6       '仕入＠
Private Const ColG_ST_URITAN% = 7       '売上＠

Private Const ColG_KINGAKU% = 8         '合計金額
Private Const ColKO_UMU% = 9            '子部品　有無
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        

'-----------------------------------    ドロップダウン
Dim SYUBETSU        As New XArrayDB
Dim JGYOBU          As New XArrayDB



Dim svHin_Gai       As String           '品番
Dim svSHIMUKE_CODE  As String           '仕向け先
Dim svCATEGORY_CODE As String           '品名ｶﾃｺﾞﾘｰｺｰﾄﾞ




Dim EXCEL_TEMPLATE  As String           'EXCELﾃﾝﾌﾟﾚｰﾄ

Dim HIN_INV         As Boolean          '未登録品番の登録可否

'--------------------------------------- EXCEL用定数
Private Const xlCalculationManual% = -4135
Private Const xlLeft% = -4131
Private Const xlCenter% = -4108
Private Const xlBottom% = -4107
Private Const xlNone% = -4142
Private Const xlContinuous% = 1
Private Const xlThin% = 2
Private Const xlAutomatic% = -4105
Private Const xlRight% = -4152
Private Const xlDiagonalDown% = 5
Private Const xlDiagonalUp% = 6
Private Const xlEdgeLeft% = 7
Private Const xlEdgeTop% = 8
Private Const xlEdgeBottom% = 9
Private Const xlEdgeRight% = 10
Private Const xlInsideVertical% = 11
Private Const xlInsideHorizontal% = 12
Private Const xlThick% = 4
Private Const xlCalculationAutomatic% = -4105
Private Const xlPortrait% = 1
'--------------------------------------- EXCEL用定数


Private Const LAST_UPDATE_DAY$ = "[SEI0016] 2016.05.XX XX:X "

Private Sub Command1_Click(Index As Integer)


Dim ans     As Integer
Dim i       As Integer

Dim MESG    As String


    Select Case Index
    
        Case 0      '終了
            Unload Me
    
    End Select






End Sub

Private Sub Form_Activate()

    If Detail_Disp_Proc Then
        Unload Me
    End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer



'    If App.PrevInstance Then
'        Beep
'        MsgBox "同一プログラム実行中です。"
'        End
'    End If


    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]大阪事　見積書作成処理", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
    
    '種別セット
    If SYUBETSU_Set_Proc() Then
        Unload Me
    End If

    '事魚部セット
    If JGYOBU_Set_Proc() Then
        Unload Me
    End If



    SEI00162.Caption = SEI00162.Caption & " " & LAST_UPDATE_DAY

    Call Init_Proc

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
Dim yn  As Integer
                                            
Dim MESG    As String
    
Dim On_Flg  As Boolean
    
    
    MESG = "商品化構成データ／品目単価を保存します。" & vbCrLf
    MESG = MESG & "　　種別／事業部／品番／員数／仕入＠／販売＠" & vbCrLf
    MESG = MESG & "　　指図票備考" & vbCrLf
    MESG = MESG & "よろしいですか？" & vbCrLf
    
    
    yn = MsgBox(MESG, vbYesNo, "確認入力")
    If yn = vbYes Then
        
        
        If Grid_Error_Check_Proc() Then
            
            Set TDBGrid1(pGrdKOUSEI).Array = KO_KOUSEI
            
        
        '    TDBGrid1(pGrdKOUSEI).Refresh
            TDBGrid1(pGrdKOUSEI).Update
        
            TDBGrid1(pGrdKOUSEI).SetFocus
            
            
            Cancel = True
            Exit Sub
        End If
        
        If Update_Proc(On_Flg) Then
            Unload Me
        End If
    
            
        KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Text1(ptxOYA_ST_SHITAN).Text   '親　仕入＠
        KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Text1(ptxOYA_ST_URITAN).Text   '親　売上＠
        KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = Text1(ptxOYA_KINGAKU).Text       '親　合計金額
    
        If On_Flg Then
            KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_UMU) = "  ▽"                          '親　子有無
        Else
            KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_UMU) = ""                              '親　子有無
        End If
    
    
    
    
    
    
    
    
    End If
    
End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    SEI00162.MousePointer = vbHourglass

    Call Ctrl_Lock(SEI00162)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(SEI00162)


    SEI00162.MousePointer = vbDefault

End Sub


Private Sub SHORI_Click(Index As Integer)
    Select Case Index
        Case 0 To 5
            Command1(Index).Value = True

        Case 6      '画面印刷
        
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)

    End Select
                    
    
    


End Sub


Private Function Init_Proc(Optional Start_Pos As Integer = 0) As Integer
'----------------------------------------------------------------------------
'                   画面初期化
'----------------------------------------------------------------------------
Dim i           As Integer

Dim Row         As Integer

Dim c           As String * 128
                                
                                
                                

                                
                                
                                
    Init_Proc = True
                                
                                
    For i = Start_Pos To ptxN_CLASS_CODE
        Text1(i).Text = ""
    Next i
                                
                                
    For i = prchBIKOU To prchM_BIKOU
        RichTextBox1(i).Text = ""
    Next i
                                
                                
                                
    If SYUBETSU_Set_Proc() Then
        Exit Function
    End If
                                
                                
    If JGYOBU_Set_Proc() Then
        Exit Function
    End If
                                
    
    
    
    Init_Proc = True


End Function
Private Function SYUBETSU_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   コードマスタをドロップダウンリストにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer



Dim i           As Integer
    
    SYUBETSU_Set_Proc = True
    
    Set SYUBETSU = Nothing
    
    
    
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    i = 0
    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN06_CD Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "コードマスタ")
                Exit Function
        
        End Select

        
        i = i + 1
        SYUBETSU.ReDim 1, i, 0, 0
        
        
        SYUBETSU(i, 0) = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode)) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
        
        
        com = BtOpGetNext
    
    Loop

    Set TDBDropDown1.Array = SYUBETSU
    TDBDropDown1.ReBind

    SYUBETSU_Set_Proc = False
    



End Function
Private Function JGYOBU_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   事業部をドロップダウンリストにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer



Dim i           As Integer
    
    JGYOBU_Set_Proc = True
    
    Set JGYOBU = Nothing
    
    i = 0
    Do
        If i > UBound(JGYOBU_T) Then
            Exit Do
        End If
        
        i = i + 1
        
        JGYOBU.ReDim 1, i, 0, 0
        JGYOBU(i, 0) = Trim(JGYOBU_T(i - 1).NAME) & "            " & Trim(JGYOBU_T(i - 1).CODE)
    
    Loop
    
    
    

    Set TDBDropDown2.Array = JGYOBU
    TDBDropDown2.ReBind

    JGYOBU_Set_Proc = False
    



End Function



Private Sub TDBGrid1_AfterColUpdate(Index As Integer, ByVal ColIndex As Integer)

Dim sts             As Integer
Dim Bookmark        As Variant
    
    
Dim i               As Integer
    
    
Dim wkDouble        As Double
    
    
Dim wkGoukei        As Double
Dim wkShi_Tan       As Double
Dim wkUri_Tan       As Double
    
    
    
    If TDBGrid1(pGrdKOUSEI).Bookmark = Null Then
        Exit Sub
    End If
    
    If TDBGrid1(pGrdKOUSEI).Bookmark <= 0 Then
        Exit Sub
    End If
    
                    
                    
                    
    Set TDBGrid1(pGrdKOUSEI).Array = KO_KOUSEI
    TDBGrid1(pGrdKOUSEI).Update
                    
                    
                    
    Select Case ColIndex
        
        Case ColKO_JGYOBU, ColKO_S_HIN_GAI
        
            ' 指図票品番の削除
            If Trim(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI)) = "" And _
                Trim(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU)) = "" Then
                
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU) = ""
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU) = ""
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI) = ""
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_NAME) = ""
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = ""
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = ""
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = ""
            
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = ""
                
            
            
            
            Else
                
                
                '品番
                Call UniCode_Conv(K0_ITEM.JGYOBU, Right(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU), 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI) = StrConv(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI), vbUpperCase)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI))
            
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        
                        If HIN_INV Then
                            '未登録品番　可　資材としておく
                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
                            Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                            Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                        Else
                            MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(品番)"
                            Exit Sub
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                        Unload Me
                
                End Select
                '品名
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                '員数
                If KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "" Then
                    KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "1.00"
                End If
                
                '仕入単価
                If Not IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                    KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Format(0, "#0.00")
                Else
                    KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Format(Val(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
                End If
                
                '売上単価
                If Not IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(0, "#0.00")
                Else
                    KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(Val(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                End If
                
                
                '合計金額
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CCur(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2)
                        
            
            End If
                

        Case ColKO_QTY
            If KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "" Then
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "1.00"
            End If

            If Not IsNumeric(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) Then
                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(員数)"
            Else
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = Format(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY), "0.00")
                '合計金額
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CCur(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2)
            End If



        Case ColG_ST_SHITAN

            If Not IsNumeric(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)) Then
                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(仕入単価)"
            Else
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Format(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN), "#0.00")
            End If
            


        Case ColG_ST_URITAN

            If Not IsNumeric(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) Then
                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(売上単価)"
            Else
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN), "#0.00")
    
                '合計金額
                KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CCur(KO_KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2)
            End If


    End Select



    Call Goukei_Proc(wkShi_Tan, wkUri_Tan, wkGoukei)
    Text1(ptxOYA_ST_SHITAN).Text = Format(wkShi_Tan, "#0.00")
    Text1(ptxOYA_ST_URITAN).Text = Format(wkUri_Tan, "#0.00")
    Text1(ptxOYA_KINGAKU).Text = Format(wkGoukei, "#0.00")



    Set TDBGrid1(pGrdKOUSEI).Array = KO_KOUSEI
    

    TDBGrid1(pGrdKOUSEI).Refresh
    TDBGrid1(pGrdKOUSEI).Update

    TDBGrid1(pGrdKOUSEI).SetFocus

End Sub









Private Function P_COMPO_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   構成マスタの読み込み＆表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
    
Dim Row         As Long
    
Dim FAST_FLG    As Boolean
    
Dim wkGoukei    As Double
Dim wkShi_Tan       As Double
Dim wkUri_Tan       As Double
    
    
    P_COMPO_Disp_Proc = True
    Call Input_Lock             '2008.01.15
    
        
    
    
            


    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(SEI00161.Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Right(Text1(ptxOYA_JGYOBU).Text, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxOYA_S_HIN_GAI).Text)
       
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
       
    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        
    Select Case sts
        Case BtNoErr
        
            FAST_FLG = True
        
            '備考
            RichTextBox1(prchBIKOU).Text = RTrim(StrConv(P_COMPO_O_REC.BIKOU, vbUnicode))
        
            '商品化ｸﾗｽ
            Text1(ptxS_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.CLASS_CODE, vbUnicode))
            '付加ｸﾗｽ
            Text1(ptxF_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.F_CLASS_CODE, vbUnicode))
            '内職ｸﾗｽ
            Text1(ptxN_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.N_CLASS_CODE, vbUnicode))

        
        Case BtErrKeyNotFound
            
            FAST_FLG = False
            
            '備考
            RichTextBox1(prchBIKOU).Text = ""
        
            '商品化ｸﾗｽ
            Text1(ptxS_CLASS_CODE).Text = ""
            '付加ｸﾗｽ
            Text1(ptxF_CLASS_CODE).Text = ""
            '内職ｸﾗｽ
            Text1(ptxN_CLASS_CODE).Text = ""
        
        
        Case Else
            
            Set KOUSEI = Nothing
            
            
            Call Input_UnLock           '2008.01.15
            P_COMPO_Disp_Proc = sts
            Exit Function
    End Select




    

    Set KO_KOUSEI = Nothing

    
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(SEI00161.Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Right(Text1(ptxOYA_JGYOBU).Text, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxOYA_S_HIN_GAI).Text)
       
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
       
       
    Row = Min_Row - 1
       
       
    com = BtOpGetGreater
       
        
    
        
    Do
        DoEvents
        
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            
                            
                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(SEI00161.Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Right(Text1(ptxOYA_JGYOBU).Text, 1) Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxOYA_S_HIN_GAI).Text) Then
                
                    Exit Do
            
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock             '2008.01.15
                Call File_Error(sts, BtOpGetNext, "構成マスタ")
                Exit Function
        End Select
        
        
        
        
        Row = Row + 1
                    
        If Grid_Set_Proc(Row) Then
            Exit Function
        End If
        
        com = BtOpGetNext
        
    Loop



    If Row < 49 Then
        For Row = Row + 1 To 50

            KO_KOUSEI.ReDim Min_Row, Row, Min_Col, Max_Col
            KO_KOUSEI(Row, ColNO) = Row
        Next Row
    End If


    Call Goukei_Proc(wkShi_Tan, wkUri_Tan, wkGoukei)
    Text1(ptxOYA_ST_SHITAN).Text = Format(wkShi_Tan, "#0.00")
    Text1(ptxOYA_ST_URITAN).Text = Format(wkUri_Tan, "#0.00")
    Text1(ptxOYA_KINGAKU).Text = Format(wkGoukei, "#0.00")



    Set TDBGrid1(pGrdKOUSEI).Array = KO_KOUSEI
    
    
    TDBGrid1(pGrdKOUSEI).Bookmark = Null
    
    TDBGrid1(pGrdKOUSEI).ReBind
    TDBGrid1(pGrdKOUSEI).Update
    TDBGrid1(pGrdKOUSEI).ScrollBars = dbgAutomatic
    
    If KOUSEI.Count(1) > 0 Then
        TDBGrid1(pGrdKOUSEI).MoveFirst
    End If















    Call Input_UnLock

    
    
    P_COMPO_Disp_Proc = False

End Function
Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'                   構成マスタ==>Gridテーブル
'----------------------------------------------------------------------------

Dim sts As Integer
Dim i   As Integer
Dim j   As Integer
    
Dim com As Integer
Dim Fsw As Integer
    
    Grid_Set_Proc = True

    

    KO_KOUSEI.ReDim Min_Row, Row, Min_Col, Max_Col
    
    'No
    KO_KOUSEI(Row, ColNO) = Row
    
    
    '種別
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
            KO_KOUSEI(Row, ColKO_SYUBETSU) = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode)) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
        
        
        
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Function
    
    End Select
    
    
    '事業部
    For i = 0 To UBound(JGYOBU_T)
    
        If Trim(JGYOBU_T(i).CODE) = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode) Then
            KO_KOUSEI(Row, ColKO_JGYOBU) = Trim(JGYOBU_T(i).NAME) & "            " & Trim(JGYOBU_T(i).CODE)
            Exit For
        End If
    Next i
    
    '指図票品番
    KO_KOUSEI(Row, ColKO_S_HIN_GAI) = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
            KO_KOUSEI(Row, ColKO_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        Case BtErrKeyNotFound
            KO_KOUSEI(Row, ColKO_HIN_NAME) = "未登録品番"
            
        
            Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
            Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select
    '員数
    If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
        KO_KOUSEI(Row, ColKO_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
    Else
        KO_KOUSEI(Row, ColKO_QTY) = "1.00"
    End If
    
    '仕入単価
    If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
        KO_KOUSEI(Row, ColG_ST_SHITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
    Else
        KO_KOUSEI(Row, ColG_ST_SHITAN) = "0.00"
    End If
    
    '売上単価
    If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
        KO_KOUSEI(Row, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
    Else
        KO_KOUSEI(Row, ColG_ST_URITAN) = "0.00"
    End If
    
    
    '合計金額
    KO_KOUSEI(Row, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KO_KOUSEI(Row, ColKO_QTY)) * CCur(KO_KOUSEI(Row, ColG_ST_URITAN))), 2)
    
    Grid_Set_Proc = False
End Function

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




End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に切り捨てします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り捨てられた数値。
' ------------------------------------------------------------------------
Public Function ToRoundDown(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
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
Public Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function

Private Function Detail_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   現在値画面表示
'----------------------------------------------------------------------------

    Detail_Disp_Proc = True
    
    
    
    
    
    
    
    Text1(ptxTanto_Code).Text = SEI00161.Text1(ptxTanto_Code).Text                  '担当者コード
    Text1(ptxTanto_Name).Text = SEI00161.Text1(ptxTanto_Name).Text                  '担当者名称
    Text1(ptxHin_Gai).Text = SEI00161.Text1(ptxHin_Gai).Text                        '品番
    Text1(ptxHin_Name).Text = SEI00161.Text1(ptxHin_Name).Text                      '品名

    Text1(ptxST_SOKO).Text = SEI00161.Text1(ptxST_SOKO).Text                        '標準棚番　 倉庫
    Text1(ptxST_RETU).Text = SEI00161.Text1(ptxST_RETU).Text                        '標準棚番　 列
    Text1(ptxST_REN).Text = SEI00161.Text1(ptxST_REN).Text                          '標準棚番　 連
    Text1(ptxST_DAN).Text = SEI00161.Text1(ptxST_DAN).Text                          '標準棚番　 段

    Text1(ptxCATEGORY_CODE).Text = SEI00161.Text1(ptxCATEGORY_CODE).Text            '品名ｶﾃｺﾞﾘｰｺｰﾄﾞ

    Text1(ptxTANKA_KIRIKAE_DT).Text = SEI00161.Text1(ptxTANKA_KIRIKAE_DT).Text      '単価切替日


    lblCATEGORY_NAME.Caption = SEI00161.Combo1(pcmbCATEGORY_Name).Text              '品名ｶﾃｺﾞﾘｰ
                                                                                    '仕向け先
    lblSHIMUKE.Caption = Mid(SEI00161.Combo1(pcmbSHIMUKE).Text, 1, Len(SEI00161.Combo1(pcmbSHIMUKE).Text) - 4)

    
    
    
    
    
    '-----------------------------------    親情報
                                                                                                    
    Text1(ptxOYA_SYUBETSU).Text = KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU)    '親　種別
    Text1(ptxOYA_JGYOBU).Text = KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU)        '親　事業部
    Text1(ptxOYA_S_HIN_GAI).Text = KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI)  '親　指図票品番
    Text1(ptxOYA_HIN_NAME).Text = KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_NAME)    '親　品名
    Text1(ptxOYA_QTY).Text = KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)              '親　員数
    Text1(ptxOYA_ST_SHITAN).Text = KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)   '親　仕入＠
    Text1(ptxOYA_ST_URITAN).Text = KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)   '親　売上＠
    Text1(ptxOYA_KINGAKU).Text = KOUSEI(SEI00161.TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU)       '親　合計金額
    
    
    
    
    
    
    
    
    
    '-----------------------------------    構成品表示
    If P_COMPO_Disp_Proc() Then
        Exit Function
    End If
    

    '-----------------------------------    作業工程　～　金額表示
    
    
    Text1(ptxNAKANISHI_TANI).Text = SEI00161.Text1(ptxNAKANISHI_TANI).Text          '中西工料　単位
    Text1(ptxNAKANISHI_KIN).Text = SEI00161.Text1(ptxNAKANISHI_KIN).Text            '中西工料　金額
        
    Text1(ptxSHOHIN_TANI).Text = SEI00161.Text1(ptxSHOHIN_TANI).Text                '商品化工料　単位
    Text1(ptxSHOHIN_KIN).Text = SEI00161.Text1(ptxSHOHIN_KIN).Text                  '商品化工料　金額
        
    Text1(ptxPF_KAKOU_TANI).Text = SEI00161.Text1(ptxPF_KAKOU_TANI).Text            'PF加工　単位
    Text1(ptxPF_KAKOU_KIN).Text = SEI00161.Text1(ptxPF_KAKOU_KIN).Text              'PF加工　金額
        
    Text1(ptxPE_KAKOU_TANI).Text = SEI00161.Text1(ptxPE_KAKOU_TANI).Text            'PE加工　単位
    Text1(ptxPE_KAKOU_KIN).Text = SEI00161.Text1(ptxPE_KAKOU_KIN).Text              'PE加工　金額
        
    Text1(ptxPE_SHIZAI_TANI).Text = SEI00161.Text1(ptxPE_SHIZAI_TANI).Text          'PE資材　単位
    Text1(ptxPE_SHIZAI_KIN).Text = SEI00161.Text1(ptxPE_SHIZAI_KIN).Text            'PE資材　金額

    Text1(ptxHINBAN_LABEL_TANI).Text = SEI00161.Text1(ptxHINBAN_LABEL_TANI).Text    '品番表示ﾗﾍﾞﾙ　単位
    Text1(ptxHINBAN_LABEL_KIN).Text = SEI00161.Text1(ptxHINBAN_LABEL_KIN).Text      '品番表示ﾗﾍﾞﾙ　金額

    Text1(ptxKOUJI_SETSU_TANI).Text = SEI00161.Text1(ptxKOUJI_SETSU_TANI).Text      '設置工事説明書　単位
    Text1(ptxKOUJI_SETSU_KIN).Text = SEI00161.Text1(ptxKOUJI_SETSU_KIN).Text        '設置工事説明書　金額

    Text1(ptxKOUJI_SETSU_TANI).Text = SEI00161.Text1(ptxKOUJI_SETSU_TANI).Text      '設置工事説明書　単位
    Text1(ptxKOUJI_SETSU_KIN).Text = SEI00161.Text1(ptxKOUJI_SETSU_KIN).Text        '設置工事説明書　金額

    Text1(ptxKONPOU_TANI).Text = SEI00161.Text1(ptxKONPOU_TANI).Text                '梱包材　単位
    Text1(ptxKONPOU_KIN).Text = SEI00161.Text1(ptxKONPOU_KIN).Text                  '梱包材　金額

    Text1(ptxFUKU_SHIZAI_TANI).Text = SEI00161.Text1(ptxFUKU_SHIZAI_TANI).Text      '副資材　単位
    Text1(ptxFUKU_SHIZAI_KIN).Text = SEI00161.Text1(ptxFUKU_SHIZAI_KIN).Text        '副資材　金額

    Text1(ptxKONPOU_ASSY_TANI).Text = SEI00161.Text1(ptxKONPOU_ASSY_TANI).Text      '梱包ASSY　単位
    Text1(ptxKONPOU_ASSY_KIN).Text = SEI00161.Text1(ptxKONPOU_ASSY_KIN).Text        '梱包ASSY　金額

    Text1(ptxKANRI_TANI).Text = SEI00161.Text1(ptxKANRI_TANI).Text                  '管理費　単位
    Text1(ptxKANRI_KIN).Text = SEI00161.Text1(ptxKANRI_KIN).Text                    '管理費　金額

    
    
    
    lblGOUKEI_KIN.Caption = SEI00161.lblGOUKEI_KIN.Caption                          '合計金額



    Detail_Disp_Proc = False

End Function



Private Function Grid_Error_Check_Proc() As Integer
'----------------------------------------------------------------------------
'                   ｸﾞﾘｯﾄﾞ内容のエラーチェック処理
'----------------------------------------------------------------------------
Dim i               As Long

Dim sts             As Integer
    
Dim j               As Long
    
    
Dim wkGoukei        As Double
Dim wkShi_Tan       As Double
Dim wkUri_Tan       As Double
    
    
    
    Grid_Error_Check_Proc = True
    
    
    
    Set TDBGrid1(pGrdKOUSEI).Array = KO_KOUSEI
    
    
    TDBGrid1(pGrdKOUSEI).Update
    
    
    For i = 1 To KOUSEI.UpperBound(1)
        ' 指図票品番の削除
        If Trim(KO_KOUSEI(i, ColKO_S_HIN_GAI)) = "" Then
            
            KO_KOUSEI(i, ColKO_SYUBETSU) = ""
            KO_KOUSEI(i, ColKO_JGYOBU) = ""
            KO_KOUSEI(i, ColKO_S_HIN_GAI) = ""
            KO_KOUSEI(i, ColKO_HIN_NAME) = ""
            KO_KOUSEI(i, ColKO_QTY) = ""
            KO_KOUSEI(i, ColG_ST_SHITAN) = ""
            KO_KOUSEI(i, ColG_ST_URITAN) = ""
        
            KO_KOUSEI(i, ColG_KINGAKU) = ""
            

        Else
            '品番
            Call UniCode_Conv(K0_ITEM.JGYOBU, Right(KO_KOUSEI(i, ColKO_JGYOBU), 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            
            Call UniCode_Conv(K0_ITEM.HIN_GAI, KO_KOUSEI(i, ColKO_S_HIN_GAI))
    
    
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    KO_KOUSEI(i, ColKO_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    If KO_KOUSEI(i, ColG_ST_SHITAN) = "" Then
                    
                        If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                            KO_KOUSEI(i, ColG_ST_SHITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
                        Else
                            KO_KOUSEI(i, ColG_ST_SHITAN) = Format(CDbl(0), "#0.00")
                        End If
                    End If
                
                    If KO_KOUSEI(i, ColG_ST_URITAN) = "" Then
                    
                        If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                            KO_KOUSEI(i, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                        Else
                            KO_KOUSEI(i, ColG_ST_URITAN) = Format(CDbl(0), "#0.00")
                        End If
                    End If
                
                
                Case BtErrKeyNotFound
                    
                        If HIN_INV Then
                            '未登録品番　可　資材としておく
                            KO_KOUSEI(i, ColKO_HIN_NAME) = "未登録品番"
                        Else
                            MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(品番)"
                            Exit Function
                        End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                    Exit Function
            End Select
                
                '員数
            If IsNumeric(KO_KOUSEI(i, ColKO_QTY)) Then
                KO_KOUSEI(i, ColKO_QTY) = Format(CDbl(KO_KOUSEI(i, ColKO_QTY)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(員数)"
                Exit Function
            End If
                
                
                '仕入＠
            If IsNumeric(KO_KOUSEI(i, ColG_ST_SHITAN)) Then
                KO_KOUSEI(i, ColG_ST_SHITAN) = Format(CDbl(KO_KOUSEI(i, ColG_ST_SHITAN)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(仕入単価)"
                Exit Function
            End If
                '販売＠
            If Trim(KO_KOUSEI(i, ColG_ST_URITAN)) = "" Then
                KO_KOUSEI(i, ColG_ST_URITAN) = "0.00"
            End If
            
            If IsNumeric(KO_KOUSEI(i, ColG_ST_URITAN)) Then
                KO_KOUSEI(i, ColG_ST_URITAN) = Format(CDbl(KO_KOUSEI(i, ColG_ST_URITAN)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(販売単価)"
                Exit Function
    
            End If
    
            '合計金額
            KO_KOUSEI(i, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KO_KOUSEI(i, ColKO_QTY)) * CCur(KO_KOUSEI(i, ColG_ST_URITAN))), 2)
    
    
        End If
    
    
    
    
    
    Next i
    
    
    Call Goukei_Proc(wkShi_Tan, wkUri_Tan, wkGoukei)
    Text1(ptxOYA_ST_SHITAN).Text = Format(wkShi_Tan, "#0.00")
    Text1(ptxOYA_ST_URITAN).Text = Format(wkUri_Tan, "#0.00")
    
    
    Text1(ptxOYA_KINGAKU).Text = Format(wkGoukei, "#0.00")
    
    
    Set TDBGrid1(pGrdKOUSEI).Array = KO_KOUSEI
    

'    TDBGrid1(pGrdKOUSEI).Refresh
    TDBGrid1(pGrdKOUSEI).Update

    TDBGrid1(pGrdKOUSEI).SetFocus
    
    
    

    Grid_Error_Check_Proc = False






End Function
Private Function Update_Proc(Optional On_Flg As Boolean = False) As Integer
'----------------------------------------------------------------------------
'                   構成マスタ出力
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer


Dim i           As Integer
Dim j           As Integer

Dim MESG        As String

Dim D_SEQNO     As Integer

Dim wkGoukei        As Double
Dim wkShi_Tan       As Double
Dim wkUri_Tan       As Double



    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    '---------------------------------------------------    '構成マスタ更新
    '該当データ全件削除
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(SEI00161.Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Right(Text1(ptxOYA_JGYOBU).Text, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxOYA_S_HIN_GAI).Text)
       
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
       
    com = BtOpGetGreater
       
    Do
        
        DoEvents
        
        Do
        
            sts = BTRV(com + BtSNoWait, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                
            Select Case sts
                Case BtNoErr
                
                    If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(SEI00161.Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) <> Right(Text1(ptxOYA_JGYOBU).Text, 1) Or _
                        StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxOYA_S_HIN_GAI).Text) Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "構成マスタ")
                            GoTo Abort_Tran
                        End If
                        sts = BtErrEOF
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "構成マスタ")
                    GoTo Abort_Tran
            End Select
    
        Loop
            
        If sts = BtErrEOF Then
            Exit Do
        End If


        Do
            sts = BTRV(BtOpDelete, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "構成マスタ")
                        End If
                        GoTo Abort_Tran
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpDelete, "構成マスタ")
                    GoTo Abort_Tran
            End Select
        Loop
    
        com = BtOpGetNext
    
    Loop





    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 構成マスタ(ﾍｯﾀﾞｰ)出力
                                                                                '仕向け先ｺｰﾄﾞ
    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, Mid(Right(SEI00161.Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
                                                                                '事業部
    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, Right(Text1(ptxOYA_JGYOBU).Text, 1))
                                                                                '国内外
    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, Text1(ptxOYA_S_HIN_GAI).Text)
    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")

    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, Text1(ptxS_CLASS_CODE).Text)    'ｸﾗｽｺｰﾄﾞ
    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, RichTextBox1(prchBIKOU).Text)        '備考
    
    Call UniCode_Conv(P_COMPO_O_REC.F_CLASS_CODE, Text1(ptxF_CLASS_CODE).Text)  '付加ｺｰﾄﾞ
    
    Call UniCode_Conv(P_COMPO_O_REC.N_CLASS_CODE, Text1(ptxN_CLASS_CODE).Text)  '内職ｺｰﾄﾞ
    
    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")

    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, "SEI16")                         '更新担当者ｺｰﾄﾞ
                                                                                '更新日時
    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


    Do
        
        DoEvents
        
        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "構成マスタ")
                GoTo Abort_Tran
        End Select
    
    Loop

    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 構成マスタ(ﾎﾞﾃﾞｨ)出力
    
    Set TDBGrid1(pGrdKOUSEI).Array = KO_KOUSEI
    
    
    TDBGrid1(pGrdKOUSEI).Update


    D_SEQNO = 0
    On_Flg = False


    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then

    Else


        For i = 1 To KO_KOUSEI.UpperBound(1)
    
    
            If Trim(KO_KOUSEI(i, ColKO_S_HIN_GAI)) = "" Then
            Else
                                                                                            '仕向け先ｺｰﾄﾞ
                Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(SEI00161.Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
                                                                                            '事業部
                Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Right(Text1(ptxOYA_JGYOBU).Text, 1))
                                                                                            '国内外
                Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxOYA_S_HIN_GAI).Text)
            
            
            
                
        
                Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)             'データ区分
                        
                D_SEQNO = D_SEQNO + 10
                        
                Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(D_SEQNO, "000"))  '追番
                                                                                '種別
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, Right(KO_KOUSEI(i, ColKO_SYUBETSU), 2))
            
                                                                                            '子　事業部
                Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, Right(KO_KOUSEI(i, ColKO_JGYOBU), 1))
                Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, NAIGAI_NAI)                      '子　国内外
                Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, KO_KOUSEI(i, ColKO_S_HIN_GAI))  '子　品番
                                                                                            '員数
                Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(KO_KOUSEI(i, ColKO_QTY)), "000.00"))
            
                Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
            
                Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "SEI16")                         '更新担当者ｺｰﾄﾞ
                                                                                            '更新日時
                Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
            
            
                Do
                    
                    DoEvents
                    
                    sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, BtOpInsert, "構成マスタ")
                            GoTo Abort_Tran
                    End Select
                
                Loop
    
                '>>>>>>>>>>>>>  品目単価　更新
                                                                                        
                Call UniCode_Conv(K0_ITEM.JGYOBU, Right(KO_KOUSEI(i, ColKO_JGYOBU), 1)) '事業部
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)                           '国内外
                Call UniCode_Conv(K0_ITEM.HIN_GAI, KO_KOUSEI(i, ColKO_S_HIN_GAI))       '品番
    
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound
                            Beep
                            ans = MsgBox("他端末でデータが変更されました。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            GoTo Abort_Tran
                
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, com + BtSNoWait, "品目マスタ")
                            GoTo Abort_Tran
                   End Select
    
                Loop
    
                If Not IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                End If
                If Not IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                End If
    
    
    
                If CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) <> CDbl(KO_KOUSEI(i, ColG_ST_SHITAN)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN, Format(CDbl(KO_KOUSEI(i, ColG_ST_SHITAN)), "00000000.00"))
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYY/MM/DD"))
                                    
                    Call UniCode_Conv(ITEMREC.UPD_TANTO, "SEI16")
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYYMMDDHHMMSS"))
    
                End If
    
    
                If CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) <> CDbl(KO_KOUSEI(i, ColG_ST_URITAN)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN, Format(CDbl(KO_KOUSEI(i, ColG_ST_URITAN)), "00000000.00"))
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, Format(Now, "YYYY/MM/DD"))
                                    
                    Call UniCode_Conv(ITEMREC.UPD_TANTO, "SEI16")
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYYMMDDHHMMSS"))
    
                End If
    
                Do
                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, com + BtSNoWait, "品目マスタ")
                            GoTo Abort_Tran
                   End Select
    
                Loop
    
    
    
                On_Flg = True
    
    
            End If
        Next i
    End If
    '>>>>>>>>>>>>>  親品番　単価更新
    Call Goukei_Proc(wkShi_Tan, wkUri_Tan, wkGoukei)


    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Text1(ptxOYA_JGYOBU).Text, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxOYA_S_HIN_GAI).Text)
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Beep
                ans = MsgBox("他端末でデータが変更されました。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                GoTo Abort_Tran
    
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, com + BtSNoWait, "品目マスタ")
                GoTo Abort_Tran
       End Select

    Loop

    If Not IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
    End If
    If Not IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
    End If
    
    

    If CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) <> wkShi_Tan Then
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, Format(wkShi_Tan, "00000000.00"))
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYY/MM/DD"))
                        
        Call UniCode_Conv(ITEMREC.UPD_TANTO, "SEI16")
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYYMMDDHHMMSS"))

    End If


    If CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) <> wkUri_Tan Then
        Call UniCode_Conv(ITEMREC.G_ST_URITAN, Format(wkUri_Tan, "00000000.00"))
        Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, Format(Now, "YYYY/MM/DD"))
                        
        Call UniCode_Conv(ITEMREC.UPD_TANTO, "SEI16")
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYYMMDDHHMMSS"))

    End If

    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    
        Select Case sts
            Case BtNoErr
                Exit Do
    
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, com + BtSNoWait, "品目マスタ")
                GoTo Abort_Tran
       End Select

    Loop







End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Call Input_UnLock
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function


Private Sub Goukei_Proc(wkShi_Tan As Double, wkUri_Tan As Double, wkGoukei As Double)
'----------------------------------------------------------------------------
'                   親の金額合計　計算
'----------------------------------------------------------------------------
Dim i       As Integer

    wkShi_Tan = 0
    wkUri_Tan = 0
    wkGoukei = 0


    For i = 1 To KO_KOUSEI.UpperBound(1)
    
    
        If Trim(KO_KOUSEI(i, ColKO_S_HIN_GAI)) <> "" Then
    
            If IsNumeric(KO_KOUSEI(i, ColG_ST_SHITAN)) Then
                wkShi_Tan = wkShi_Tan + CDbl(KO_KOUSEI(i, ColG_ST_SHITAN))
            End If
        End If
    
    
        If Trim(KO_KOUSEI(i, ColKO_S_HIN_GAI)) <> "" Then
    
            If IsNumeric(KO_KOUSEI(i, ColG_ST_URITAN)) Then
                wkUri_Tan = wkUri_Tan + CDbl(KO_KOUSEI(i, ColG_ST_URITAN))
            End If
        End If
    
    
    
        If Trim(KO_KOUSEI(i, ColKO_S_HIN_GAI)) <> "" Then
    
            If IsNumeric(KO_KOUSEI(i, ColG_KINGAKU)) Then
                wkGoukei = wkGoukei + CDbl(KO_KOUSEI(i, ColG_KINGAKU))
            End If
        End If
    
    Next i



End Sub
