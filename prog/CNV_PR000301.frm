VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form CNV_PR000301 
   Caption         =   "éëçﬁç›å…íIâµÇµï\ÉÅÉìÉeÉiÉìÉX CONV_PR00030 2010.12.22"
   ClientHeight    =   10296
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   14988
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10296
   ScaleWidth      =   14988
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.CommandButton Command2 
      Caption         =   "èoå…êîçƒåvéZ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8085
      TabIndex        =   16
      Top             =   1680
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "çáåvílçƒèWåv"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10395
      TabIndex        =   15
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ç›å…ã‡äzçƒèWåv"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   11655
      TabIndex        =   14
      Top             =   1680
      Width           =   1590
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10215
      _ExtentX        =   18013
      _ExtentY        =   2138
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "é˚éxíPà "
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ëOåéç›å…ã‡äz"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ìñåéì¸å…ã‡äz"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "ìñåéèoå…ã‡äz"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "ìñåéç›å…ã‡äz"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "ç∑äzÅÅìñ-ëOåé"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   699
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2096"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1990"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2942"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2836"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2942"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2836"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2942"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2836"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2942"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2836"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2942"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2836"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ÇlÇr ÉSÉVÉbÉN"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ÇlÇr ÉSÉVÉbÉN"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
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
      _StyleDefs(5)   =   ":id=0,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=1200"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ÇlÇr ÉSÉVÉbÉN"
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
      _StyleDefs(26)  =   ":id=43,.fontname=ÇlÇr ÉSÉVÉbÉN"
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
      _StyleDefs(40)  =   ":id=58,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=44"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=45"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=47"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=66,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(45)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(46)  =   ":id=66,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=44"
      _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=32,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(51)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(52)  =   ":id=32,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=44"
      _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=45"
      _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=47"
      _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=44"
      _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=82,.parent=43,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=79,.parent=44"
      _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=80,.parent=45"
      _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=81,.parent=47"
      _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(5).HeadingStyle:id=17,.parent=44"
      _StyleDefs(66)  =   "Splits(0).Columns(5).FooterStyle:id=18,.parent=45"
      _StyleDefs(67)  =   "Splits(0).Columns(5).EditorStyle:id=19,.parent=47"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "èI óπ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10440
      TabIndex        =   12
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9600
      TabIndex        =   11
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7920
      TabIndex        =   9
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   5
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "çX êV"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   9720
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7335
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   14775
      _ExtentX        =   26056
      _ExtentY        =   12933
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ïiî‘"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ïiñº"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ç›å…å≥"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "ëOåééc"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "ì¸å…êî"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "èoå…êî"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "ìñåéç›å…"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "édì¸íPâø"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "ìñåéç›å…ã‡äz"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "édì¸êÊCODE"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "ç≈èIèoâ◊îNåéì˙"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "ç≈èIèoå…êî"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "ëOéÿéc"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "í˜éûç›å…êî"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "ìoò^ì˙ït"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   15
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   699
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=15"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2096"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1990"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5355"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5249"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=2096"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1990"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=0"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=1884"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1778"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=1884"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1778"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=1884"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1778"
      Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=1884"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=1778"
      Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(35)=   "Column(7).Width=2709"
      Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=2604"
      Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(40)=   "Column(8).Width=2709"
      Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2604"
      Splits(0)._ColumnProps(43)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(45)=   "Column(9).Width=1651"
      Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=1545"
      Splits(0)._ColumnProps(48)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(49)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(50)=   "Column(10).Width=3048"
      Splits(0)._ColumnProps(51)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(10)._WidthInPix=2942"
      Splits(0)._ColumnProps(53)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(54)=   "Column(11).Width=2413"
      Splits(0)._ColumnProps(55)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(11)._WidthInPix=2307"
      Splits(0)._ColumnProps(57)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(58)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(59)=   "Column(12).Width=1884"
      Splits(0)._ColumnProps(60)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(12)._WidthInPix=1778"
      Splits(0)._ColumnProps(62)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(63)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(64)=   "Column(13).Width=3810"
      Splits(0)._ColumnProps(65)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(13)._WidthInPix=3704"
      Splits(0)._ColumnProps(67)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(68)=   "Column(14).Width=3810"
      Splits(0)._ColumnProps(69)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(14)._WidthInPix=3704"
      Splits(0)._ColumnProps(71)=   "Column(14).Order=15"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ÇlÇr ÉSÉVÉbÉN"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ÇlÇr ÉSÉVÉbÉN"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
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
      _StyleDefs(5)   =   ":id=0,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=1200"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ÇlÇr ÉSÉVÉbÉN"
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
      _StyleDefs(26)  =   ":id=43,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10,.bgcolor=&HFFFF00&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=58,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(39)  =   ":id=58,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(40)  =   ":id=58,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=44"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=45"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=47"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=16,.parent=43"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=44"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=45"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=47"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=28,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(49)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(50)  =   ":id=28,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=44"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=45"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=47"
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=66,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(55)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(56)  =   ":id=66,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=44"
      _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=32,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(61)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(62)  =   ":id=32,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=44"
      _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=45"
      _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=47"
      _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=82,.parent=43,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=79,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=80,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=81,.parent=47"
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
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=62,.parent=43"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=47"
      _StyleDefs(90)  =   "Splits(0).Columns(11).Style:id=70,.parent=43,.alignment=1"
      _StyleDefs(91)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=44"
      _StyleDefs(92)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=45"
      _StyleDefs(93)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=47"
      _StyleDefs(94)  =   "Splits(0).Columns(12).Style:id=86,.parent=43,.alignment=1"
      _StyleDefs(95)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=44"
      _StyleDefs(96)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=45"
      _StyleDefs(97)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=47"
      _StyleDefs(98)  =   "Splits(0).Columns(13).Style:id=90,.parent=43"
      _StyleDefs(99)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=44"
      _StyleDefs(100) =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=45"
      _StyleDefs(101) =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=47"
      _StyleDefs(102) =   "Splits(0).Columns(14).Style:id=94,.parent=43"
      _StyleDefs(103) =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=44"
      _StyleDefs(104) =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=45"
      _StyleDefs(105) =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=47"
      _StyleDefs(106) =   "Named:id=33:Normal"
      _StyleDefs(107) =   ":id=33,.parent=0"
      _StyleDefs(108) =   "Named:id=34:Heading"
      _StyleDefs(109) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(110) =   ":id=34,.wraptext=-1"
      _StyleDefs(111) =   "Named:id=35:Footing"
      _StyleDefs(112) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(113) =   "Named:id=36:Selected"
      _StyleDefs(114) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(115) =   "Named:id=37:Caption"
      _StyleDefs(116) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(117) =   "Named:id=38:HighlightRow"
      _StyleDefs(118) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(119) =   "Named:id=39:EvenRow"
      _StyleDefs(120) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(121) =   "Named:id=40:OddRow"
      _StyleDefs(122) =   ":id=40,.parent=33"
      _StyleDefs(123) =   "Named:id=41:RecordSelector"
      _StyleDefs(124) =   ":id=41,.parent=34"
      _StyleDefs(125) =   "Named:id=42:FilterBar"
      _StyleDefs(126) =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "CNV_PR000301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

'Glidópä¬ã´---------------------------------

Private Const pSum_GridSTOCK% = 0
Private Const pGridSTOCK% = 1

Private Sum_STOCK       As New XArrayDB

Private Const Sum_Min_Row% = 1              'ç≈è¨çsêî
Private Const Sum_Min_Col% = 0              'ç≈è¨óÒêî
Private Const Sum_Max_Col% = 5              'ç≈ëÂóÒêî

Private Const colSum_G_SYUSHI% = 0          'é˚éxíPà 
Private Const colSum_ZEN_ZAIKO_KIN% = 1     'ëOåéç›å…ã‡äz
Private Const colSum_NYUKO_KIN% = 2         'ìñåéì¸å…ã‡äz
Private Const colSum_SYUKO_KIN% = 3         'ìñåéèoå…ã‡äz
Private Const colSum_ZAIKO_KIN% = 4         'ìñåéç›å…ã‡äz
Private Const colSum_SA_KIN% = 5            'ç∑äz



Private STOCK       As New XArrayDB


Private Const Min_Row% = 1                  'ç≈è¨çsêî
Private Const Min_Col% = 0                  'ç≈è¨óÒêî
Private Const Max_Col% = 14                 'ç≈ëÂóÒêî

Private Const colHIN_GAI% = 0               'éëçﬁïiî‘
Private Const colHIN_NAME% = 1              'ïiñº
Private Const colG_SYUSHI% = 2              'ç›å…å≥Åié˚éxÅj
Private Const colZEN_ZAIKO_QTY% = 3         'ëOåéç›å…
Private Const colNYUKO_QTY% = 4             'ìñåéì¸å…
Private Const colSYUKO_QTY% = 5             'ìñåéèoå…
Private Const colZAIKO_QTY% = 6             'ìñåéç›å…
Private Const colSHI_TANKA% = 7             'édì¸íPâø
Private Const colZAIKO_KIN% = 8             'ìñåéç›å…ã‡äz
Private Const colSHI_CODE% = 9              'ìñåéédì¸êÊ∫∞ƒﬁ

Private Const colLAST_SYUKA_DT% = 10        'ç≈èIèoå…ì˙
Private Const colLAST_SYUKA_QTY% = 11       'ç≈èIèoå…êî

Private Const colMAEGARI_QTY% = 12          'ëOéÿêî

Private Const colMOTO_ZAIKO_QTY% = 13       'å≥ç›å…êî

Private Const colINPUT_DATE% = 14           'ì¸óÕì˙ït


Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ø∞ƒÇÃêßå‰ 0:è∏èá 1:ç~èá
'   íIâµÇµ√ﬁ∞¿ï ŒﬂºﬁºÆ∆›∏ﬁ
Private wP_STOCK_POS    As POSBLK
Private wP_STOCK_REC    As P_STOCK_REC_Tag
Private K0_wP_STOCK     As KEY0_P_STOCK


Private Type SUM_AREA
    SYUSHI      As String
    ZENZAN_KIN  As Long
    NYU_KIN     As Long
    SYU_KIN     As Long
    TOUZAN_KIN  As Long
    SAGAKU_KIN  As Long
End Type

Private G_SYUSHI_TBL    As Variant          ' ëŒè€é˚éx      2007.11.13


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   âÊñ çÄñ⁄ÉçÉbÉNÅiÉCÉxÉìÉgéÊìæïsâ¬Åj
'----------------------------------------------------------------------------

    CNV_PR000301.MousePointer = vbHourglass

    Call Ctrl_Lock(CNV_PR000301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   âÊñ çÄñ⁄ÉçÉbÉNâèúÅiÉCÉxÉìÉgéÊìæâ¬Åj
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(CNV_PR000301)


    CNV_PR000301.MousePointer = vbDefault

End Sub



Private Sub Command1_Click(index As Integer)
Dim ans         As Integer
Dim i           As Integer


Dim sts         As Integer

    Select Case index
        
        Case P_CMD_Upd          'çXêV
        
        
            ans = MsgBox("àÍäáçXêVÇçsÇ¢Ç‹Ç∑Ç©ÅH", vbYesNo, "ämîFì¸óÕ")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
            
            End If
        
        
        
        

        
        
        Case P_CMD_DSP                      'åüçı/ï\é¶
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
        
        
        Case P_CMD_OUT                      '√ﬁ∞¿èoóÕ
        
        Case P_CMD_PRT                      'àÛç¸
 
            
            
        Case P_CMD_End                      'èIóπ
    
            Unload Me
    
    End Select

End Sub

Private Sub Command2_Click(index As Integer)

Dim yn              As Integer
Dim i               As Integer

Dim ZEN_ZAIKO_QTY   As Long

    
    Select Case index
        Case 0
            yn = MsgBox("ìñåéç›å…ã‡äzÇÃçƒèWåvÇçsÇ¢Ç‹Ç∑Ç©ÅH", vbYesNo, "ämîFì¸óÕ")
            If yn = vbNo Then
                Exit Sub
            End If
        
        
            Set TDBGrid1(pGridSTOCK).Array = STOCK
            TDBGrid1(pGridSTOCK).Refresh
            TDBGrid1(pGridSTOCK).Update
        
            For i = 1 To STOCK.UpperBound(1)
            
                If STOCK(i, colSHI_TANKA) = 0 Then
                    STOCK(i, colZAIKO_KIN) = 0
                Else
                    
                    If Not IsNumeric(STOCK(i, colSHI_TANKA)) Or Not IsNumeric(STOCK(i, colZAIKO_QTY)) Then
                        STOCK(i, colZAIKO_KIN) = 0
                    Else
                        STOCK(i, colZAIKO_KIN) = Format(Int(CDbl(CDbl(STOCK(i, colSHI_TANKA)) * CLng(STOCK(i, colZAIKO_QTY)) + 0.5)), "#,##0")
                    End If
                End If
            Next i

    
        Case 1
            yn = MsgBox("ìñåéèoå…êîÇÃçƒèWåvÇçsÇ¢Ç‹Ç∑Ç©ÅH", vbYesNo, "ämîFì¸óÕ")
            If yn = vbNo Then
                Exit Sub
            End If
        
        
            Set TDBGrid1(pGridSTOCK).Array = STOCK
            TDBGrid1(pGridSTOCK).Refresh
            TDBGrid1(pGridSTOCK).Update
        
        
            For i = 1 To STOCK.UpperBound(1)
            
                
                If Trim(STOCK(i, colSHI_CODE)) = "" Then
                    ZEN_ZAIKO_QTY = STOCK(i, colZEN_ZAIKO_QTY)
                
                Else
                    STOCK(i, colSYUKO_QTY) = ZEN_ZAIKO_QTY + STOCK(i, colNYUKO_QTY) - STOCK(i, colZAIKO_QTY)
                End If
            Next i
    
    End Select


    Set TDBGrid1(pGridSTOCK).Array = STOCK
    TDBGrid1(pGridSTOCK).Refresh
    TDBGrid1(pGridSTOCK).Update


End Sub

Private Sub Command3_Click()

Dim yn          As Integer
Dim i           As Integer
Dim j           As Integer

Dim sts         As Integer

Dim SUM_TBL()   As SUM_AREA

Dim G_ZENZAN_KIN  As Long
Dim G_NYU_KIN     As Long
Dim G_SYU_KIN     As Long
Dim G_TOUZAN_KIN  As Long
Dim G_SAGAKU_KIN  As Long


    yn = MsgBox("çáåvílÇÃçƒèWåvÇçsÇ¢Ç‹Ç∑Ç©ÅH", vbYesNo, "ämîFì¸óÕ")
    If yn = vbNo Then
        Exit Sub
    End If


    


    
    


    For i = 1 To STOCK.UpperBound(1)
        
        
        If i = 1 Then
            ReDim Preserve SUM_TBL(0)
            SUM_TBL(0).SYUSHI = STOCK(i, colG_SYUSHI)
            SUM_TBL(0).ZENZAN_KIN = 0
            SUM_TBL(0).NYU_KIN = 0
            SUM_TBL(0).SYU_KIN = 0
            SUM_TBL(0).TOUZAN_KIN = 0
            SUM_TBL(0).SAGAKU_KIN = 0
            j = 0
        Else
        
            For j = 0 To UBound(SUM_TBL)
            
                If SUM_TBL(j).SYUSHI = STOCK(i, colG_SYUSHI) Then
                    Exit For
                End If
            
            Next j
                
            If j > UBound(SUM_TBL) Then
                ReDim Preserve SUM_TBL(j)
                SUM_TBL(j).SYUSHI = STOCK(i, colG_SYUSHI)
                SUM_TBL(j).ZENZAN_KIN = 0
                SUM_TBL(j).NYU_KIN = 0
                SUM_TBL(j).SYU_KIN = 0
                SUM_TBL(j).TOUZAN_KIN = 0
                SUM_TBL(j).SAGAKU_KIN = 0
            End If
        End If
        'ëOåéécçÇ
        If Trim(STOCK(i, colSHI_TANKA)) = "" Then
            Call UniCode_Conv(K0_ITEM.JGYOBU, "S")
            Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
            Call UniCode_Conv(K0_ITEM.HIN_GAI, STOCK(i, colHIN_GAI))
    
    
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) Then
                    Else
                        SUM_TBL(j).ZENZAN_KIN = SUM_TBL(j).ZENZAN_KIN + CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode))
                    End If
                
                Case BtErrKeyNotFound
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ïiñ⁄œΩ¿")
                    Unload Me
            End Select
            
        Else
            'ìñåéì¸å…ã‡äz
            SUM_TBL(j).NYU_KIN = SUM_TBL(j).NYU_KIN + Int(CDbl(CDbl(STOCK(i, colSHI_TANKA)) * CLng(STOCK(i, colNYUKO_QTY)) + 0.5))
            'ìñåéç›å…ã‡äz
            SUM_TBL(j).TOUZAN_KIN = SUM_TBL(j).TOUZAN_KIN + CLng(STOCK(i, colZAIKO_KIN))
            
        End If
    
    Next i


    For i = 0 To UBound(SUM_TBL)


        SUM_TBL(i).SYU_KIN = SUM_TBL(i).ZENZAN_KIN + SUM_TBL(i).NYU_KIN - SUM_TBL(i).TOUZAN_KIN
        SUM_TBL(i).SAGAKU_KIN = SUM_TBL(i).TOUZAN_KIN - SUM_TBL(i).ZENZAN_KIN

    Next i

    Set Sum_STOCK = Nothing

    G_ZENZAN_KIN = 0
    G_NYU_KIN = 0
    G_SYU_KIN = 0
    G_TOUZAN_KIN = 0
    G_SAGAKU_KIN = 0



    For i = 0 To UBound(SUM_TBL)


        Sum_STOCK.ReDim Sum_Min_Row, i + 1, Min_Col, Max_Col

        Sum_STOCK(i + 1, colSum_G_SYUSHI) = SUM_TBL(i).SYUSHI
        Sum_STOCK(i + 1, colSum_ZEN_ZAIKO_KIN) = Format(SUM_TBL(i).ZENZAN_KIN, "#,##0")
        Sum_STOCK(i + 1, colSum_NYUKO_KIN) = Format(SUM_TBL(i).NYU_KIN, "#,##0")
        Sum_STOCK(i + 1, colSum_SYUKO_KIN) = Format(SUM_TBL(i).SYU_KIN, "#,##0")
        Sum_STOCK(i + 1, colSum_ZAIKO_KIN) = Format(SUM_TBL(i).TOUZAN_KIN, "#,##0")
        Sum_STOCK(i + 1, colSum_SA_KIN) = Format(SUM_TBL(i).SAGAKU_KIN, "#,##0")

        G_ZENZAN_KIN = G_ZENZAN_KIN + SUM_TBL(i).ZENZAN_KIN
        G_NYU_KIN = G_NYU_KIN + SUM_TBL(i).NYU_KIN
        G_SYU_KIN = G_SYU_KIN + SUM_TBL(i).SYU_KIN
        G_TOUZAN_KIN = G_TOUZAN_KIN + SUM_TBL(i).TOUZAN_KIN
        G_SAGAKU_KIN = G_SAGAKU_KIN + SUM_TBL(i).SAGAKU_KIN
        


    Next i

    Sum_STOCK.ReDim Sum_Min_Row, i + 1, Min_Col, Max_Col
    Sum_STOCK(i + 1, colSum_G_SYUSHI) = "zzz"
    Sum_STOCK(i + 1, colSum_ZEN_ZAIKO_KIN) = Format(G_ZENZAN_KIN, "#,##0")
    Sum_STOCK(i + 1, colSum_NYUKO_KIN) = Format(G_NYU_KIN, "#,##0")
    Sum_STOCK(i + 1, colSum_SYUKO_KIN) = Format(G_SYU_KIN, "#,##0")
    Sum_STOCK(i + 1, colSum_ZAIKO_KIN) = Format(G_TOUZAN_KIN, "#,##0")
    Sum_STOCK(i + 1, colSum_SA_KIN) = Format(G_SAGAKU_KIN, "#,##0")




    Set TDBGrid1(pSum_GridSTOCK).Array = Sum_STOCK
    TDBGrid1(pSum_GridSTOCK).ReBind
    TDBGrid1(pSum_GridSTOCK).Update
    TDBGrid1(pSum_GridSTOCK).MoveFirst



End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   ÇjÇÖÇô ÇcÇèÇóÇé ëOèàóù
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
        MsgBox "ìØàÍÉvÉçÉOÉâÉÄé¿çsíÜÇ≈Ç∑ÅB"
        End
    End If
                                'ÉçÉOÉtÉ@ÉCÉãñºéÊÇËçûÇ›
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ÉçÉOÉtÉ@ÉCÉãñºÇÃälìæÇ…é∏îsÇµÇ‹ÇµÇΩÅBèàóùÇíÜé~ÇµÇƒâ∫Ç≥Ç¢ÅB"
        End
    End If
    LOG_F = RTrim(c)
                                
                                'ÉçÉOÉtÉ@ÉCÉãñºéÊÇËçûÇ›
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ÉçÉOÉtÉ@ÉCÉãñºÇÃälìæÇ…é∏îsÇµÇ‹ÇµÇΩÅBèàóùÇíÜé~ÇµÇƒâ∫Ç≥Ç¢ÅB"
        End
    End If
    LOG_F = RTrim(c)
                                
                                'ïiñ⁄É}ÉXÉ^ÇnÇoÇdÇm
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                'éëçﬁíIâµÇµÇnÇoÇdÇm
    If P_STOCK_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                'éëçﬁíIâµÇµÇnÇoÇdÇm
    If wP_STOCK_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'éëçﬁíIâµÇµèWåvÇnÇoÇdÇm
    If P_STOCKSUM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
''    GLB_SYUSHI_F = "01"
                                
                                'ëŒè€é˚éxÇÃälìæ 2007.11.13
    If Trim(GLB_SYUSHI_F) = "" Then
    
    Else
    
        If GetIni("PR00030", GLB_SYUSHI_F, "P_SYS", c) Then
            Beep
            MsgBox "ëŒè€é˚éxÇÃälìæÇ…é∏îsÇµÇ‹ÇµÇΩÅBèàóùÇíÜé~ÇµÇƒâ∫Ç≥Ç¢ÅB"
            End
        End If
    
        G_SYUSHI_TBL = Split(Trim(c), ",", -1)
    End If
    
    
    
    
    
    
        
    
    'âÊñ èâä˙ê›íË
    If Init_Proc() Then
        Unload Me
    End If
    

    If List_Disp_Proc() Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            
                                            
                                            
                                            'ïiñ⁄É}ÉXÉ^ÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ïiñ⁄É}ÉXÉ^")
        End If
    End If
                                            
                                            'éëçﬁíIâµÇµÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "éëçﬁíIâµ")
        End If
    End If
                                            'éëçﬁíIâµÇµÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, wP_STOCK_POS, wP_STOCK_REC, Len(wP_STOCK_REC), K0_wP_STOCK, Len(K0_wP_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "wéëçﬁíIâµ")
        End If
    End If
                                            'éëçﬁíIâµèWåvÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "éëçﬁíIâµèWåv")
        End If
    End If
                                            
                                            
    
    Set CNV_PR000301 = Nothing


    End
End Sub





Private Sub TDBGrid1_HeadClick(index As Integer, ByVal ColIndex As Integer)



    Select Case index
        
        
        Case pGridSTOCK
    
    
            If Sort_Tbl(ColIndex) = 0 Then
                Sort_Tbl(ColIndex) = 1
            Else
                If Sort_Tbl(ColIndex) = 1 Then
                    Sort_Tbl(ColIndex) = 0
                End If
            
            End If
            
            If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                            
                STOCK.QuickSort Min_Row, STOCK.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
                
                Set TDBGrid1(index).Array = STOCK
                
                TDBGrid1(index).ReBind
                TDBGrid1(index).Update
                TDBGrid1(index).MoveFirst
        
        
            End If
    
    
    
    End Select




End Sub


Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   ì¸óÕâÊñ ÇÃèâä˙ê›íË
'----------------------------------------------------------------------------
Dim i       As Integer
Dim sts     As Integer


    Init_Proc = True
    
    
    
    
    'ø∞ƒèÓïÒÇÃèâä˙âª
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0               '√ﬁÃ´Ÿƒè∏èá
    Next i
    
    Sort_Tbl(colHIN_NAME) = 9       'ø∞ƒèúäO

    Init_Proc = False

End Function
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           éëçﬁíIâµÇµ√ﬁ∞¿ÇÃï\é¶
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer

Dim Row               As Long


Dim Skip_Flg            As Boolean

Dim i                   As Integer
Dim Mode                As Integer




    List_Disp_Proc = True
    CNV_PR000301.MousePointer = vbHourglass
    
    
    
    
    com = BtOpGetFirst
    
    
    
    
    Do
        sts = BTRV(com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
        Select Case sts
            Case BtNoErr
                If CDbl(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)) = 0 And _
                    CDbl(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) = 0 And _
                    CDbl(StrConv(P_STOCKSUM_REC.SYUKO_KIN, vbUnicode)) = 0 And _
                    CDbl(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)) = 0 Then
                    sts = BTRV(BtOpDelete, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case Else
                            Call File_Error(sts, BtOpDelete, "éëçﬁíIâµèWåv√ﬁ∞¿")
                            Exit Function
                    End Select
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "éëçﬁíIâµèWåv√ﬁ∞¿")
                Exit Function
        End Select
        
        
                
        
        com = BtOpGetNext
    Loop
    
    
    'é˚éxíPà 
    
    Set Sum_STOCK = Nothing
    
    Row = Sum_Min_Row - 1
    
    com = BtOpGetFirst
    Do
        DoEvents
    
        sts = BTRV(com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
            
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "éëçﬁíIâµÇµèWåv√ﬁ∞¿")
                Exit Function
        End Select
    
    
    
    
        Row = Row + 1
        
        Sum_STOCK.ReDim Sum_Min_Row, Row, Sum_Min_Col, Sum_Max_Col
                
                
                
                
        
        'ïiî‘
        Sum_STOCK(Row, colSum_G_SYUSHI) = StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode)
        'ëOåéç›å…ã‡äz
        Sum_STOCK(Row, colSum_ZEN_ZAIKO_KIN) = Format(CLng(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)), "#,##0")
        'ìñåéì¸å…ã‡äz
        Sum_STOCK(Row, colSum_NYUKO_KIN) = Format(CLng(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)), "#,##0")
        'ìñåéèoå…ã‡äz
        Sum_STOCK(Row, colSum_SYUKO_KIN) = Format(CLng(StrConv(P_STOCKSUM_REC.SYUKO_KIN, vbUnicode)), "#,##0")
        'ìñåéç›å…ã‡äz
        Sum_STOCK(Row, colSum_ZAIKO_KIN) = Format(CLng(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)), "#,##0")
        'ìñåéç∑äz
        Sum_STOCK(Row, colSum_SA_KIN) = Format(Sum_STOCK(Row, colSum_ZAIKO_KIN) - Sum_STOCK(Row, colSum_ZEN_ZAIKO_KIN), "#,##0")
        
        
        com = BtOpGetGreater
    Loop
    
    
    
    
    
    Set TDBGrid1(pSum_GridSTOCK).Array = Sum_STOCK
    TDBGrid1(pSum_GridSTOCK).ReBind
    TDBGrid1(pSum_GridSTOCK).Update
    TDBGrid1(pSum_GridSTOCK).MoveFirst
    
    
    
    
    
    
    Set STOCK = Nothing
    
    Row = Min_Row - 1
    
    com = BtOpGetFirst
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
            
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "éëçﬁíIâµÇµ√ﬁ∞¿")
                Exit Function
        End Select
    
    
    
    
        Row = Row + 1
        If Grid_Set_Proc(Row) Then
            Exit Function
        End If
'        If mode = 1 Then
'            Exit Do
'        End If
        com = BtOpGetGreater
    Loop
    
    
    Set TDBGrid1(pGridSTOCK).Array = STOCK
    TDBGrid1(pGridSTOCK).ReBind
    TDBGrid1(pGridSTOCK).Update
    TDBGrid1(pGridSTOCK).MoveFirst
    
    
    CNV_PR000301.MousePointer = vbDefault
    List_Disp_Proc = False
    


End Function

Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           éëçﬁíIâµÇµ√ﬁ∞¿ÇÃì‡óeÇ∏ﬁÿØƒﬁÇ…æØƒÇ∑ÇÈ
'----------------------------------------------------------------------------
Dim sts             As Integer

Dim com             As Integer
Dim Save_Jgyobu     As String
Dim Save_Naigai     As String
Dim Save_Hin_Gai    As String



    Grid_Set_Proc = True
    
    
    
        
    
    
    
    
    STOCK.ReDim Min_Row, Row, Min_Col, Max_Col


    'ïiî‘
    STOCK(Row, colHIN_GAI) = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)
    
    'ïiñº
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_GAI, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "ïiñ⁄É}ÉXÉ^")
            Exit Function
    End Select
    STOCK(Row, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    'ç›å…å≥Åié˚éxÅj
    STOCK(Row, colG_SYUSHI) = StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode)
    
    
    'ìñåéì¸å…êî
    STOCK(Row, colNYUKO_QTY) = Format(CLng(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)), "#,##0")
    'ìñåéèoå…êî
    STOCK(Row, colSYUKO_QTY) = Format(CLng(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode)), "#,##0")
    'ìñåéç›å…êî
    STOCK(Row, colZAIKO_QTY) = Format(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)), "#,##0")
    'édì¸íPâø
    If IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) Then
        STOCK(Row, colSHI_TANKA) = Format(CDbl(StrConv(P_STOCK_REC.TANKA, vbUnicode)), "#,##0.00")
    Else
        STOCK(Row, colSHI_TANKA) = ""
    End If
    'édì¸êÊ
    STOCK(Row, colSHI_CODE) = StrConv(P_STOCK_REC.CODE, vbUnicode)
    'ç›å…ã‡äz
    
'If Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)) = "Y010" Then
'Debug.Print
'End If
    
    
    If IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) Then
        
        
        
        STOCK(Row, colZAIKO_KIN) = Format(ToRoundUp(CCur(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
                                    CDbl(StrConv(P_STOCK_REC.TANKA, vbUnicode))), 0), "#,##0")
    Else
        STOCK(Row, colZAIKO_KIN) = ""
    End If
    
    
    
    'ëOåéç›å…
    STOCK(Row, colZEN_ZAIKO_QTY) = Format(CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)), "#,##0")
    'ç≈èIèoâ◊ì˙
    STOCK(Row, colLAST_SYUKA_DT) = Mid(StrConv(P_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 7, 2)
    'ç≈èIèoå…êî
    If IsNumeric(StrConv(P_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)) Then
        STOCK(Row, colLAST_SYUKA_QTY) = Format(CLng(StrConv(P_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)), "#,##0")
    Else
        STOCK(Row, colLAST_SYUKA_QTY) = ""
    End If
    
    'ëOéÿéc
    If IsNumeric(StrConv(P_STOCK_REC.MAEGARI_QTY, vbUnicode)) Then
        STOCK(Row, colMAEGARI_QTY) = Format(CLng(StrConv(P_STOCK_REC.MAEGARI_QTY, vbUnicode)), "#,##0")
    Else
        STOCK(Row, colMAEGARI_QTY) = ""
    End If
    
    'å≥ç›å…êî
    If IsNumeric(StrConv(P_STOCK_REC.MOTO_ZAIKO_QTY, vbUnicode)) Then
        STOCK(Row, colMOTO_ZAIKO_QTY) = Format(CLng(StrConv(P_STOCK_REC.MOTO_ZAIKO_QTY, vbUnicode)), "#,##0")
    Else
        STOCK(Row, colMOTO_ZAIKO_QTY) = ""
    End If
    
    
    Call UniCode_Conv(K0_wP_STOCK.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_wP_STOCK.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_wP_STOCK.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K0_wP_STOCK.CODE, StrConv(P_STOCK_REC.CODE, vbUnicode))
    Call UniCode_Conv(K0_wP_STOCK.TANKA, StrConv(P_STOCK_REC.TANKA, vbUnicode))
    
    Save_Jgyobu = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
    Save_Naigai = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
    Save_Hin_Gai = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)
    
    com = BtOpGetGreater
    
    
    Do
        DoEvents
        sts = BTRV(com, wP_STOCK_POS, wP_STOCK_REC, Len(wP_STOCK_REC), K0_wP_STOCK, Len(K0_wP_STOCK), 0)
        
        Select Case sts
            Case BtNoErr
                If Save_Jgyobu <> StrConv(wP_STOCK_REC.JGYOBU, vbUnicode) Or _
                    Save_Naigai <> StrConv(wP_STOCK_REC.NAIGAI, vbUnicode) Or _
                    Save_Hin_Gai <> StrConv(wP_STOCK_REC.HIN_GAI, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "wéëçﬁíIâµ√ﬁ∞¿")
                Exit Function
        End Select
    
    
    
        Row = Row + 1
'        If i > 0 Then
            STOCK.ReDim Min_Row, Row, Min_Col, Max_Col
        
'''            STOCK(Row, colHIN_GAI) = ""
''            STOCK(Row, colHIN_NAME) = ""
''            STOCK(Row, colG_SYUSHI) = ""
''''            STOCK(Row, colZEN_ZAIKO_QTY) = ""
            
            
''            STOCK(Row, colLAST_SYUKA_DT) = ""
''            STOCK(Row, colLAST_SYUKA_QTY) = ""

        
'        End If
        
        'ïiî‘
        STOCK(Row, colHIN_GAI) = StrConv(wP_STOCK_REC.HIN_GAI, vbUnicode)
        
        'ïiñº
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(wP_STOCK_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(wP_STOCK_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(wP_STOCK_REC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_GAI, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ïiñ⁄É}ÉXÉ^")
                Exit Function
        End Select
        STOCK(Row, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        'ç›å…å≥Åié˚éxÅj
        STOCK(Row, colG_SYUSHI) = StrConv(wP_STOCK_REC.G_SYUSHI, vbUnicode)
        
        
        'ëOåéç›å…
        STOCK(Row, colZEN_ZAIKO_QTY) = Format(CLng(StrConv(wP_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)), "#,##0")
        'ìñåéì¸å…êî
        STOCK(Row, colNYUKO_QTY) = Format(CLng(StrConv(wP_STOCK_REC.NYUKO_QTY, vbUnicode)), "#,##0")
        'ìñåéèoå…êî
        STOCK(Row, colSYUKO_QTY) = Format(CLng(StrConv(wP_STOCK_REC.SYUKO_QTY, vbUnicode)), "#,##0")
        'ìñåéç›å…êî
        STOCK(Row, colZAIKO_QTY) = Format(CLng(StrConv(wP_STOCK_REC.ZAIKO_QTY, vbUnicode)), "#,##0")
        'édì¸íPâø
        If IsNumeric(StrConv(wP_STOCK_REC.TANKA, vbUnicode)) Then
            STOCK(Row, colSHI_TANKA) = Format(CDbl(StrConv(wP_STOCK_REC.TANKA, vbUnicode)), "#,##0.00")
        Else
            STOCK(Row, colSHI_TANKA) = ""
        End If
        'édì¸êÊ
        STOCK(Row, colSHI_CODE) = StrConv(wP_STOCK_REC.CODE, vbUnicode)
        
        
        'ç›å…ã‡äz
        
        If IsNumeric(StrConv(wP_STOCK_REC.TANKA, vbUnicode)) Then
            
            
        STOCK(Row, colZAIKO_KIN) = Format(ToRoundUp(CCur(CLng(StrConv(wP_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
                                    CDbl(StrConv(wP_STOCK_REC.TANKA, vbUnicode))), 0), "#,##0")
            
'            STOCK(Row, colZAIKO_KIN) = Format(CLng(StrConv(wP_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
'                                        CDbl(StrConv(wP_STOCK_REC.TANKA, vbUnicode)), "#,##0")
        Else
            STOCK(Row, colZAIKO_KIN) = ""
        End If

        'ç≈èIèoâ◊ì˙
        STOCK(Row, colLAST_SYUKA_DT) = Mid(StrConv(wP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(wP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(wP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 7, 2)
        'ç≈èIèoå…êî
        If IsNumeric(StrConv(wP_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)) Then
            STOCK(Row, colLAST_SYUKA_QTY) = Format(CLng(StrConv(wP_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)), "#,##0")
        Else
            STOCK(Row, colLAST_SYUKA_QTY) = ""
        End If



        'ëOéÿéc
        If IsNumeric(StrConv(wP_STOCK_REC.MAEGARI_QTY, vbUnicode)) Then
            STOCK(Row, colMAEGARI_QTY) = Format(CLng(StrConv(wP_STOCK_REC.MAEGARI_QTY, vbUnicode)), "#,##0")
        Else
            STOCK(Row, colMAEGARI_QTY) = ""
        End If
    
        'å≥ç›å…êî
        If IsNumeric(StrConv(wP_STOCK_REC.MOTO_ZAIKO_QTY, vbUnicode)) Then
            STOCK(Row, colMOTO_ZAIKO_QTY) = Format(CLng(StrConv(wP_STOCK_REC.MOTO_ZAIKO_QTY, vbUnicode)), "#,##0")
        Else
            STOCK(Row, colMOTO_ZAIKO_QTY) = ""
        End If
    
        
        'ì¸óÕì˙ït
        STOCK(Row, colINPUT_DATE) = StrConv(wP_STOCK_REC.INPUT_DATE, vbUnicode)
        
        com = BtOpGetNext
    
    
    Loop
        
          
    Call UniCode_Conv(K0_P_STOCK.JGYOBU, Save_Jgyobu)
    Call UniCode_Conv(K0_P_STOCK.NAIGAI, Save_Naigai)
    Call UniCode_Conv(K0_P_STOCK.HIN_GAI, Save_Hin_Gai)
    Call UniCode_Conv(K0_P_STOCK.CODE, "zzzzzz")
    Call UniCode_Conv(K0_P_STOCK.TANKA, "zzzzzzzzzzzz")
          
          
        
        

    Grid_Set_Proc = False

End Function




Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   éëçﬁíIâµÇµ√ﬁ∞¿çÏê¨
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
Dim i           As Integer
    
    
    
    
    Update_Proc = True
    CNV_PR000301.MousePointer = vbHourglass

    com = BtOpGetFirst
    Do
    
        sts = BTRV(com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
            
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "éëçﬁíIâµÇµèWåv√ﬁ∞¿")
                Exit Function
        End Select
    
        sts = BTRV(BtOpDelete, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpDelete, "éëçﬁíIâµÇµèWåv√ﬁ∞¿")
            Exit Function
        End If
    
        com = BtOpGetNext
    
    Loop





    com = BtOpGetFirst
    Do
    
        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
            
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "éëçﬁíIâµÇµ√ﬁ∞¿")
                Exit Function
        End Select
    
        sts = BTRV(BtOpDelete, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpDelete, "éëçﬁíIâµÇµ√ﬁ∞¿")
            Exit Function
        End If
    
        com = BtOpGetNext
    
    Loop
    
    
    
    Set TDBGrid1(pSum_GridSTOCK).Array = Sum_STOCK
    TDBGrid1(pSum_GridSTOCK).Refresh
    TDBGrid1(pSum_GridSTOCK).Update
    
    
    
    
    
    For i = 1 To Sum_STOCK.UpperBound(1)
    
    
        Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, Sum_STOCK(i, colSum_G_SYUSHI))
        

        
        
        
        'ëOåéç›å…ã‡äz
        If IsNumeric(Sum_STOCK(i, colSum_ZEN_ZAIKO_KIN)) Then
            Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, Format(CLng(Sum_STOCK(i, colSum_ZEN_ZAIKO_KIN)), "00000000000"))
        Else
            Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, "00000000000")
        End If
        'ìñåéì¸å…ã‡äz
        If IsNumeric(Sum_STOCK(i, colSum_NYUKO_KIN)) Then
            
            If CLng(Sum_STOCK(i, colSum_NYUKO_KIN)) > 0 Then
                Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, Format(CLng(Sum_STOCK(i, colSum_NYUKO_KIN)), "00000000000"))
            Else
                Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, Format(CLng(Sum_STOCK(i, colSum_NYUKO_KIN)), "0000000000"))
            End If
        Else
            Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, "00000000000")
        End If
        'ìñåéèoå…ã‡äz
        If IsNumeric(Sum_STOCK(i, colSum_SYUKO_KIN)) Then
            
            If CLng(Sum_STOCK(i, colSum_SYUKO_KIN)) < 0 Then
                Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(CLng(Sum_STOCK(i, colSum_SYUKO_KIN)), "0000000000"))
            Else
                Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(CLng(Sum_STOCK(i, colSum_SYUKO_KIN)), "00000000000"))
            End If
        Else
            Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, "00000000000")
        End If
        
        'ìñåéç›å…ã‡äz
        If IsNumeric(Sum_STOCK(i, colSum_ZAIKO_KIN)) Then
            Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, Format(CLng(Sum_STOCK(i, colSum_ZAIKO_KIN)), "00000000000"))
        Else
            Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, "00000000000")
        End If
        Call UniCode_Conv(P_STOCKSUM_REC.FILLER, "")
    
        sts = BTRV(BtOpInsert, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpInsert, "éëçﬁíIâµÇµèWåv√ﬁ∞¿")
            Exit Function
        End If
    
    
    Next i
    
    
    Set TDBGrid1(pGridSTOCK).Array = STOCK
    TDBGrid1(pGridSTOCK).Refresh
    TDBGrid1(pGridSTOCK).Update
    
    
    
    For i = 1 To STOCK.UpperBound(1)
        'éñã∆ïî
        Call UniCode_Conv(P_STOCK_REC.JGYOBU, SHIZAI)
        'çëì‡äO
        Call UniCode_Conv(P_STOCK_REC.NAIGAI, NAIGAI_NAI)
        'ïiî‘
        Call UniCode_Conv(P_STOCK_REC.HIN_GAI, STOCK(i, colHIN_GAI))

Debug.Print STOCK(i, colHIN_GAI)
        
        'édì¸êÊ
        Call UniCode_Conv(P_STOCK_REC.CODE, STOCK(i, colSHI_CODE))
        'édì¸íPâø
        
        
        
        If Trim(STOCK(i, colSHI_CODE)) = "" Then
            Call UniCode_Conv(P_STOCK_REC.TANKA, "")
        Else
            If IsNumeric(STOCK(i, colSHI_TANKA)) Then
                Call UniCode_Conv(P_STOCK_REC.TANKA, Format(CDbl(STOCK(i, colSHI_TANKA)), "00000000.00"))
            Else
                Call UniCode_Conv(P_STOCK_REC.TANKA, "00000000.00")
            End If
        End If
        'ìoò^ì˙ït
        Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, STOCK(i, colINPUT_DATE))
        
        'é˚éxíPà 
        Call UniCode_Conv(P_STOCK_REC.G_SYUSHI, STOCK(i, colG_SYUSHI))
        'ëOåéç›å…
        If IsNumeric(STOCK(i, colZEN_ZAIKO_QTY)) Then
            Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, Format(CDbl(STOCK(i, colZEN_ZAIKO_QTY)), "00000000"))
        Else
            Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")
        End If
        'ì¸å…êî
        If IsNumeric(STOCK(i, colNYUKO_QTY)) Then
            
            If CDbl(STOCK(i, colNYUKO_QTY)) > 0 Then
                Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, Format(CDbl(STOCK(i, colNYUKO_QTY)), "00000000"))
            Else
                Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, Format(CDbl(STOCK(i, colNYUKO_QTY)), "0000000"))
            End If
        Else
            Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, "00000000")
        End If
        
        'èoå…êî
        If IsNumeric(STOCK(i, colSYUKO_QTY)) Then
            
            
            If CDbl(STOCK(i, colSYUKO_QTY)) >= 0 Then
                Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, Format(CDbl(STOCK(i, colSYUKO_QTY)), "00000000"))
            Else
                Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, Format(CDbl(STOCK(i, colSYUKO_QTY)), "0000000"))
            End If
        Else
            Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, "00000000")
        End If
        
        'ç›å…êî
        If IsNumeric(STOCK(i, colZAIKO_QTY)) Then
            Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(CDbl(STOCK(i, colZAIKO_QTY)), "00000000"))
        Else
            Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")
        End If
        
        'ç≈èIèoå…ì˙
        If IsDate(STOCK(i, colLAST_SYUKA_DT)) Then
            Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_DT, Format(CDate(STOCK(i, colLAST_SYUKA_DT)), "YYYYMMDD"))
        Else
            Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_DT, "")
        End If
        'ç≈èIèoâ◊êî
        If IsNumeric(STOCK(i, colLAST_SYUKA_QTY)) Then
            Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_QTY, Format(CDbl(STOCK(i, colLAST_SYUKA_QTY)), "00000000"))
        Else
            Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_QTY, "00000000")
        End If
        
        'ç≈èWåvëOêîó 
        If IsNumeric(STOCK(i, colMOTO_ZAIKO_QTY)) Then
            
            If CDbl(STOCK(i, colMOTO_ZAIKO_QTY)) < 0 Then
                Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, Format(CDbl(STOCK(i, colMOTO_ZAIKO_QTY)), "0000000"))
            Else
                Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, Format(CDbl(STOCK(i, colMOTO_ZAIKO_QTY)), "00000000"))
            End If
        Else
            Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, "00000000")
        End If
    
        'ëOéÿ
        If IsNumeric(STOCK(i, colMAEGARI_QTY)) Then
            Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, Format(CDbl(STOCK(i, colMAEGARI_QTY)), "00000000"))
        Else
            Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, "00000000")
        End If
    
    
    
    
    
    
        Call UniCode_Conv(P_STOCK_REC.FILLER, "")
    
        sts = BTRV(BtOpInsert, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
        
        Select Case sts
        
            Case BtNoErr
            
            Case BtErrDuplicates
            
                
            Case Else
                Call File_Error(sts, BtOpInsert, "éëçﬁíIâµÇµ√ﬁ∞¿")
                
                
                Exit Function
        End Select
    
    
    Next i
    
    
    CNV_PR000301.MousePointer = vbDefault

    Update_Proc = False

End Function
Public Function wP_STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              éëçﬁíIâµÇµ√ﬁ∞¿  ÇnÇoÇdÇm
'*
'*      à¯  êî:Open Mode(BtrieveéQè∆)
'*      ñﬂÇËíl:false ê≥èÌ
'*             true  àŸèÌ
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim Ret             As Long     '2007.11.13


    wP_STOCK_Open = True
                                            'éëçﬁíIâµÉfÅ[É^ÉtÉãÉpÉXéÊçûÇ›
    sts = GetIni("FILE", P_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_STOCK]ì«Ç›çûÇ›ÉGÉâÅ[")
        Exit Function
    End If
    
    
    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - Ret)
    
    
'    FullPath = Trim(c)

    Do
        sts = BTRV(BtOpOpen, wP_STOCK_POS, wP_STOCK_REC, Len(wP_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "wéëçﬁíIâµÇµ√ﬁ∞¿")
                Exit Function
        End Select
    Loop
    
    wP_STOCK_Open = False

End Function


' ------------------------------------------------------------------------
'       éwíËÇµÇΩê∏ìxÇÃêîílÇ…êÿÇËè„Ç∞ÇµÇ‹Ç∑ÅB
'
' @Param    dValue      ä€ÇﬂëŒè€ÇÃî{ê∏ìxïÇìÆè¨êîì_êîÅB
' @Param    iDigits     ñﬂÇËílÇÃóLå¯åÖêîÇÃê∏ìxÅB
' @Return               iDigits Ç…ìôÇµÇ¢ê∏ìxÇÃêîílÇ…êÿÇËè„Ç∞ÇÁÇÍÇΩêîílÅB
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

