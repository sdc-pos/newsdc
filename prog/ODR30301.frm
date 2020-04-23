VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ODR30301 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "半製品管理画面"
   ClientHeight    =   10140
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10140
   ScaleWidth      =   15270
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   8880
      TabIndex        =   11
      Top             =   960
      Width           =   5985
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Index           =   3
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4800
         Width           =   2550
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Index           =   2
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4320
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "更　新"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3960
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4800
         Visible         =   0   'False
         Width           =   1800
      End
      Begin TrueDBGrid80.TDBGrid TDBGrid1 
         Height          =   3915
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   6906
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "子品番"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "子品名"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "員数"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "SDC在庫"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   4
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "在訂"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1508"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8208"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2011"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1879"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8208"
         Splits(0)._ColumnProps(11)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1640"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=8210"
         Splits(0)._ColumnProps(17)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=1852"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1720"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=18"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=1270"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1138"
         Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=17"
         Splits(0)._ColumnProps(28)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(5).Width=4366"
         Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=4233"
         Splits(0)._ColumnProps(33)=   "Column(5)._ColStyle=20"
         Splits(0)._ColumnProps(34)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "半製品情報"
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         DeadAreaBackColor=   -2147483643
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
         _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ Ｐゴシック"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.valignment=2,.bgcolor=&HFF0000&,.bold=0"
         _StyleDefs(7)   =   ":id=1,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
         _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1125,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
         _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
         _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000005&"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFF0000&"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&HFFFFFF&"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.namedParent=37,.bgcolor=&H80FF00&"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.namedParent=39,.bgcolor=&HFFFFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.namedParent=40,.bgcolor=&HFFFFFF&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=17,.parent=87,.alignment=0,.bgcolor=&HFF00&"
         _StyleDefs(37)  =   ":id=17,.locked=-1"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=14,.parent=88"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=15,.parent=89"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=16,.parent=91"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=110,.parent=87,.alignment=0,.bgcolor=&HFF00&"
         _StyleDefs(42)  =   ":id=110,.locked=-1"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=88"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=89"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=91"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=29,.parent=87,.alignment=1,.bgcolor=&HFF00&"
         _StyleDefs(47)  =   ":id=29,.locked=-1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=26,.parent=88"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=27,.parent=89"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=28,.parent=91"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=21,.parent=87,.alignment=1,.bgcolor=&HFF00&"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=18,.parent=88"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=19,.parent=89"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=20,.parent=91"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=25,.parent=87,.alignment=2"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=22,.parent=88"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=23,.parent=89"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=24,.parent=91"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=43,.parent=87"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=30,.parent=88"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=31,.parent=89"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=32,.parent=91"
         _StyleDefs(63)  =   "Named:id=33:Normal"
         _StyleDefs(64)  =   ":id=33,.parent=0"
         _StyleDefs(65)  =   "Named:id=34:Heading"
         _StyleDefs(66)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   ":id=34,.wraptext=-1"
         _StyleDefs(68)  =   "Named:id=35:Footing"
         _StyleDefs(69)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(70)  =   "Named:id=36:Selected"
         _StyleDefs(71)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(72)  =   "Named:id=37:Caption"
         _StyleDefs(73)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(74)  =   "Named:id=38:HighlightRow"
         _StyleDefs(75)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(76)  =   "Named:id=39:EvenRow"
         _StyleDefs(77)  =   ":id=39,.parent=33,.bgcolor=&H80FF00&"
         _StyleDefs(78)  =   "Named:id=40:OddRow"
         _StyleDefs(79)  =   ":id=40,.parent=33,.bgcolor=&HFF0000&"
         _StyleDefs(80)  =   "Named:id=41:RecordSelector"
         _StyleDefs(81)  =   ":id=41,.parent=34"
         _StyleDefs(82)  =   "Named:id=42:FilterBar"
         _StyleDefs(83)  =   ":id=42,.parent=33"
         _StyleDefs(84)  =   "Named:id=13:IO_OK"
         _StyleDefs(85)  =   ":id=13,.parent=42,.bgcolor=&H80000005&"
      End
      Begin VB.Label Lab_Fix 
         Alignment       =   1  '右揃え
         AutoSize        =   -1  'True
         Caption         =   "親品番"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   450
         TabIndex        =   17
         Top             =   4920
         Width           =   720
      End
      Begin VB.Label Lab_Fix 
         Alignment       =   1  '右揃え
         AutoSize        =   -1  'True
         Caption         =   "使用日付"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   210
         TabIndex        =   16
         Top             =   4440
         Width           =   960
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Index           =   0
      Left            =   9000
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   4860
      MaxLength       =   7
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   900
      MaxLength       =   5
      TabIndex        =   1
      Top             =   780
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   13260
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終　了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2205
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更　新"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   165
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1800
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8475
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   14949
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
      Columns(1).Caption=   "使用日"
      Columns(1).DataField=   ""
      Columns(1).DataWidth=   10
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "品番"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "品名"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "部材在庫"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "ＫＥＹ項目"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1032"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=900"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2646"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2514"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=16"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=4366"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4233"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=16"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2408"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8208"
      Splits(0)._ColumnProps(20)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=3519"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3387"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=18"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2778"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=8212"
      Splits(0)._ColumnProps(31)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "半製品情報"
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
      DeadAreaBackColor=   -2147483643
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ Ｐゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.valignment=2,.bgcolor=&HFF0000&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000005&"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFF0000&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFF00&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.namedParent=37,.bgcolor=&H80FF00&"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.namedParent=39,.bgcolor=&H80FF80&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.namedParent=40,.bgcolor=&H80FF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=25,.parent=87,.alignment=2,.bgcolor=&H80000005&"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=22,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=23,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=24,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=87,.alignment=0,.bgcolor=&H80000005&"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=17,.parent=87,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(45)  =   ":id=17,.locked=0"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=14,.parent=88"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=15,.parent=89"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=16,.parent=91"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=110,.parent=87,.alignment=0,.bgcolor=&HFF00&"
      _StyleDefs(50)  =   ":id=110,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=107,.parent=88"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=108,.parent=89"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=109,.parent=91"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=29,.parent=87,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(55)  =   ":id=29,.locked=0"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=26,.parent=88"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=27,.parent=89"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=28,.parent=91"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=138,.parent=87,.locked=-1"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=135,.parent=88"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=136,.parent=89"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=137,.parent=91"
      _StyleDefs(63)  =   "Named:id=33:Normal"
      _StyleDefs(64)  =   ":id=33,.parent=0"
      _StyleDefs(65)  =   "Named:id=34:Heading"
      _StyleDefs(66)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   ":id=34,.wraptext=-1"
      _StyleDefs(68)  =   "Named:id=35:Footing"
      _StyleDefs(69)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   "Named:id=36:Selected"
      _StyleDefs(71)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(72)  =   "Named:id=37:Caption"
      _StyleDefs(73)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(74)  =   "Named:id=38:HighlightRow"
      _StyleDefs(75)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(76)  =   "Named:id=39:EvenRow"
      _StyleDefs(77)  =   ":id=39,.parent=33,.bgcolor=&H80FF00&"
      _StyleDefs(78)  =   "Named:id=40:OddRow"
      _StyleDefs(79)  =   ":id=40,.parent=33,.bgcolor=&HFF0000&"
      _StyleDefs(80)  =   "Named:id=41:RecordSelector"
      _StyleDefs(81)  =   ":id=41,.parent=34"
      _StyleDefs(82)  =   "Named:id=42:FilterBar"
      _StyleDefs(83)  =   ":id=42,.parent=33"
      _StyleDefs(84)  =   "Named:id=13:IO_OK"
      _StyleDefs(85)  =   ":id=13,.parent=42,.bgcolor=&H80000005&"
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "仕向先"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   8220
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "使用月"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Lab_Dsp 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   2235
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "担当者"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   720
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "更新"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   2
      End
   End
End
Attribute VB_Name = "ODR30301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'コンボ用添字
'Private Const pcmbJI = 0            '事業部
Private Const pcmbSM = 0            '仕向け先


'テキスト用添字
Private Const ptxTOP% = 0
Private Const ptxLAST% = 3

Private Const ptxTANTO_CD% = 0
Private Const ptxUSE_YM% = 1

Private Const ptxUSE_YMD = 2
Private Const ptxHIN_GAI = 3


'ラベル用添字
Private Const plabTANTO_NM% = 0

'コマンドボタン用添字
Private Const FuncCOR% = 0       '更新
Private Const FuncEND% = 1       '終了

'ListBox添字
'Private Const plstSRCH% = 0         '


'グリッド更新マーク
Private Const GridOYA% = 0      '親
Private Const GridKO% = 1       '子



Dim Grid_Cor_M      As Integer
Dim Grid_Req_M      As Integer

'グリッド用定義
Private ORDR_OYA   As New XArrayDB
Private ORDR_KO   As New XArrayDB

Private Const OYA_Min_Row% = 1                  '最小行数
Private Const OYA_Max_Row = 9999                '最大行数

Private Const OYA_Min_Col% = 0                  '最小列数
Private Const OYA_Max_Col% = 5                  '最大列数

Private Const Col_OYA_DEL% = 0                  '削除マーク
Private Const Col_OYA_USE_YMD% = 1              '使用日
Private Const Col_OYA_HIN_GAI% = 2              '親品番
Private Const Col_OYA_HIN_NAME% = 3             '親品名
Private Const Col_OYA_SHIJI_QTY% = 4            '部材在庫数
Private Const Col_OYA_KEY% = 5                  'データＫｅｙ情報

Private Const KO_Min_Row% = 1                   '最小行数
Private Const KO_Max_Row = 9999                 '最大行数

Private Const KO_Min_Col% = 0                   '最小列数
Private Const KO_Max_Col% = 5                   '最大列数

Private Const Col_KO_HIN_GAI% = 0               '子品番
Private Const Col_KO_HIN_NAME% = 1              '子品名
Private Const Col_KO_QTY% = 2                   '員数
Private Const Col_KO_USE_QTY% = 3               'ＳＤＣ在庫数
Private Const Col_KO_ZAITEI_F% = 4              '在訂マーク
Private Const Col_KO_KEY% = 5                   'データＫｅｙ情報


Private Sort_Tbl(Col_OYA_DEL To Col_OYA_KEY) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順    2016.01.14


Dim row         As Long                     '対象　行

Dim Cor_Row     As Long                     'カレント行

'Private Const LAST_UPDATE_DAY$ = "([ODR3030] 2016.01.14 16:45)"
Private Const LAST_UPDATE_DAY$ = "([ODR3030] 2016.01.15 09:00)"


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

Private Function Grid_Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   グリッド入力内容エラーチェック
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim yn          As Integer
Dim i           As Integer

    Grid_Err_Chk = True
    
    Set TDBGrid1(GridOYA).Array = ORDR_OYA
                                     
    If ORDR_OYA.Count(1) < 1 Then
        Grid_Err_Chk = False
        Exit Function
    End If
    
    
    
        
    
    
    For i = 1 To ORDR_OYA.Count(1)
        
        
        If ORDR_OYA(i, Col_OYA_DEL) Or _
            Not IsNumeric(ORDR_OYA(i, Col_OYA_KEY)) Then
        Else
        
        
            If Not IsDate(ORDR_OYA(i, Col_OYA_USE_YMD)) Then
                                
                                
                                
                MsgBox "使用日入力エラー"
                                
                                
                Set TDBGrid1(GridOYA).Array = ORDR_OYA
                
                
                TDBGrid1(GridOYA).Bookmark = Null
                
                TDBGrid1(GridOYA).ReBind
                TDBGrid1(GridOYA).Update
                                
                                
                Exit Function
        
            Else
                ORDR_OYA(i, Col_OYA_USE_YMD) = Format(ORDR_OYA(i, Col_OYA_USE_YMD), "YYYY/MM/DD")
        
            End If
            
            
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, GW_JIGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, GW_NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, ORDR_OYA(i, Col_OYA_HIN_GAI))
        
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                
            Select Case sts
                Case BtNoErr
                
                Case BtErrKeyNotFound
                
                
                    MsgBox "品番未登録です。更新できません。"       '2016.01.14
                
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "☆品番未登録☆")
                
                    ORDR_OYA(i, Col_OYA_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                
                    TDBGrid1(GridOYA).Bookmark = Null
                    
                    TDBGrid1(GridOYA).ReBind
                    TDBGrid1(GridOYA).Update
                
                
                
                
                    Exit Function
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            ORDR_OYA(i, Col_OYA_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    
                    
'2016.01.14            Call UniCode_Conv(K0_ODR_HANSEIHIN.USE_YM, Left(Format(Text1(ptxUSE_YM) & "/01", "YYYYMMDD"), 6))
            Call UniCode_Conv(K0_ODR_HANSEIHIN.INPUT_NO, ORDR_OYA(i, Col_OYA_KEY))
            Call UniCode_Conv(K0_ODR_HANSEIHIN.SEQNO, "000")
                                    
            sts = BTRV(BtOpGetEqual, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
            Select Case sts
                Case BtNoErr
                
                    If Trim(StrConv(ODR_HANSEIHIN_O_REC.HIN_GAI, vbUnicode)) <> Trim(ORDR_OYA(i, Col_OYA_HIN_GAI)) Then
                    
                        MsgBox "品番変更不可！！削除後、再登録してください。"
                        
                        Set TDBGrid1(GridOYA).Array = ORDR_OYA
                        
                        
                        TDBGrid1(GridOYA).Bookmark = Null
                        
                        TDBGrid1(GridOYA).ReBind
                        TDBGrid1(GridOYA).Update
                        
                        
                        Exit Function
                    
                    
                    
                    End If
                
                
                Case BtErrKeyNotFound
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "半製品管理")
                    Exit Function
            End Select
                    
                    
        
        
            If Not IsNumeric(ORDR_OYA(i, Col_OYA_SHIJI_QTY)) Then
                MsgBox "数量入力エラー"
                
                
                Set TDBGrid1(GridOYA).Array = ORDR_OYA
                
                
                TDBGrid1(GridOYA).Bookmark = Null
                
                TDBGrid1(GridOYA).ReBind
                TDBGrid1(GridOYA).Update
                
                
                
                Exit Function
        
        
        
        
        
            End If
        End If
    
    Next i
    

    Grid_Err_Chk = False

End Function
Private Function ERR_CHK(Index As Integer) As Integer
'----------------------------------------------------------------------------
'                   テキスト入力内容エラーチェック
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String


    ERR_CHK = True
    
                        '入力文字数チェック
'    If LenB(StrConv(Text1(Index), vbFromUnicode)) > Text1(Index).MaxLength Then
'        MsgBox "入力した項目は（桁あふれエラー）です。", vbExclamation
'        Exit Function
'    End If
    
    Select Case Index
        Case ptxTANTO_CD
            Lab_Dsp(plabTANTO_NM) = ""
            If Trim(Text1(Index)) = "" Then
                MsgBox "担当者を指定して下さい。", vbExclamation
                Exit Function
            End If
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(Index))
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound       'レコード無し
                    MsgBox "担当者　未登録！", vbExclamation
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "TANTO")
                    Exit Function
            End Select
            
            Lab_Dsp(plabTANTO_NM) = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
            GW_TANTO = Trim(Text1(Index))
            
            GW_SIMUKE = Left(Right(Combo1(pcmbSM).Text, 4), 2)
            'GW_JIGYOBU = Right(Combo1(pcmbJI).Text, 1)
            GW_JIGYOBU = Mid(Right(Combo1(pcmbSM).Text, 4), 3, 1)
            GW_NAIGAI = Right(Combo1(pcmbSM).Text, 1)
            
            
        Case ptxUSE_YM
            If Trim(Text1(Index)) = "" Then
                MsgBox "使用年月を指定して下さい。", vbExclamation
                Exit Function
            End If
            
            If Not IsDate(Text1(ptxUSE_YM) & "/01") Then
                MsgBox "使用月エラー！", vbExclamation
                Exit Function
            End If
            
            
    End Select
    
    
    ERR_CHK = False
End Function

Private Function Data_Disp() As Integer
'----------------------------------------------------------------------------
'                   親品番情報の表示
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer
    
Dim X_i         As Long
    
Dim W_Key       As String
Dim W_STR       As String

Dim cnt         As Integer

Dim i           As Integer

    Data_Disp = True
    
    row = OYA_Min_Row - 1
    Call Input_Lock                             '画面項目ロック
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　入力情報　検索中！　＜Data_Disp＞", Me.hwnd, 0)
    
    
    
    'ｿｰﾄ情報の初期化            '2016.01.14
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    Sort_Tbl(Col_OYA_DEL) = 9       'ｿｰﾄ除外
'    Sort_Tbl(Col_OYA_HIN_NAME) = 9  'ｿｰﾄ除外
    
    
    Set ORDR_OYA = Nothing
    
    
    
'2016.01.14    Call UniCode_Conv(K0_ODR_HANSEIHIN.USE_YM, Left(Format(Text1(ptxUSE_YM).Text, "YYYYMMDD"), 6))
    Call UniCode_Conv(K0_ODR_HANSEIHIN.INPUT_NO, "")
    Call UniCode_Conv(K0_ODR_HANSEIHIN.SEQNO, "")
    
    com = BtOpGetGreaterEqual
    Do
        
        sts = BTRV(com, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF       'レコード無し
                Exit Do
            Case Else
                Call File_Error(sts, com, "ODR_ORDER")
                GoTo Err_Exit
        End Select
        
        
'        If StrConv(ODR_HANSEIHIN_O_REC.USE_YM, vbUnicode) <> Left(Format(Text1(ptxUSE_YM).Text, "YYYYMMDD"), 6) Then
'            Exit Do    '2016.01.14
'        End If         '2016.01.14
        
'        Else            '2016.01.14
            If StrConv(ODR_HANSEIHIN_O_REC.SEQNO, vbUnicode) <> "000" Then
            Else
                row = row + 1
                        
                If OYA_Grid_Set_Proc() Then
                    GoTo Err_Exit
                End If
            
            End If
            
'        End If          '2016.01.14
        
        com = BtOpGetNext
        
    Loop
    
    
    
    
    Set TDBGrid1(GridOYA).Array = ORDR_OYA
    
    
    TDBGrid1(GridOYA).Bookmark = Null
    
    TDBGrid1(GridOYA).ReBind
    TDBGrid1(GridOYA).Update
    TDBGrid1(GridOYA).MoveFirst
    TDBGrid1(GridOYA).ScrollBars = dbgAutomatic
    
    
'    If ORDR_OYA.Count(1) > 0 Then
'        If KO_Grid_Set_Proc(TDBGrid1(GridOYA).Bookmark) Then
'            GoTo Err_Exit
'        End If
'
'    End If
    
    
    If row <> (OYA_Min_Row - 1) Then                                                                            '2016.01.14
        ORDR_OYA.QuickSort OYA_Min_Row, ORDR_OYA.UpperBound(1), Col_OYA_USE_YMD, XORDER_ASCEND, XTYPE_STRING    '2016.01.14
        Set TDBGrid1(GridOYA).Array = ORDR_OYA                                                                  '2016.01.14
        TDBGrid1(GridOYA).ReBind                                                                                '2016.01.14
        TDBGrid1(GridOYA).Update                                                                                '2016.01.14
        TDBGrid1(GridOYA).MoveFirst                                                                             '2016.01.14
    End If                                                                                                      '2016.01.14
    
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "現在の親部品情報　表示中　→　追加登録・修正入力して下さい。", Me.hwnd, 0)
    DoEvents
    
    Data_Disp = False
    
Err_Exit:
    Call Input_UnLock                             '画面項目ロック
    
End Function

Private Function OYA_Grid_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   グリッド表示（親）
'               Row   行数
'               mode　FALSE:ﾁｪｯｸOFF  TRUE:ﾁｪｯｸON
'----------------------------------------------------------------------------
Dim sts         As Integer

    OYA_Grid_Set_Proc = True

    ORDR_OYA.ReDim OYA_Min_Row, row, OYA_Min_Col, OYA_Max_Col

    
    '使用日付
    ORDR_OYA(row, Col_OYA_USE_YMD) = Mid(StrConv(ODR_HANSEIHIN_O_REC.USE_YMD, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(ODR_HANSEIHIN_O_REC.USE_YMD, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(ODR_HANSEIHIN_O_REC.USE_YMD, vbUnicode), 7, 2)
    '品番
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ODR_HANSEIHIN_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ODR_HANSEIHIN_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ODR_HANSEIHIN_O_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
            Call UniCode_Conv(ITEMREC.HIN_NAME, "品番未登録！！")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "ITEM")
            Exit Function
    End Select
    ORDR_OYA(row, Col_OYA_HIN_GAI) = Trim(StrConv(ODR_HANSEIHIN_O_REC.HIN_GAI, vbUnicode))
    ORDR_OYA(row, Col_OYA_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    '指示数
    ORDR_OYA(row, Col_OYA_SHIJI_QTY) = Format(CLng(StrConv(ODR_HANSEIHIN_O_REC.SHIJI_QTY, vbUnicode)), "#,##0")
    '
    ORDR_OYA(row, Col_OYA_KEY) = StrConv(ODR_HANSEIHIN_O_REC.INPUT_NO, vbUnicode)



    OYA_Grid_Set_Proc = False

End Function

Private Function KO_Grid_Set_Proc(i As Long) As Integer
'----------------------------------------------------------------------------
'                   グリッド表示（子）
'               Row   行数
'               mode　FALSE:ﾁｪｯｸOFF  TRUE:ﾁｪｯｸON
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer



    KO_Grid_Set_Proc = True


    Set ORDR_KO = Nothing




'2016.01.14    Call UniCode_Conv(K0_ODR_HANSEIHIN.USE_YM, Left(Format(Text1(ptxUSE_YM).Text, "YYYYMMDD"), 6))
    
    Call UniCode_Conv(K0_ODR_HANSEIHIN.INPUT_NO, Format(CInt(ORDR_OYA(TDBGrid1(GridOYA).Bookmark, Col_OYA_KEY)), "0000"))
    Call UniCode_Conv(K0_ODR_HANSEIHIN.SEQNO, "000")
    
    com = BtOpGetGreater
    
    row = 0
    Do
        
        DoEvents
        
        sts = BTRV(com, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_K_REC, Len(ODR_HANSEIHIN_K_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF       'レコード無し
                Exit Do
            Case Else
                Call File_Error(sts, com, "ODR_ORDER")
                Exit Function
        End Select
        
        
'>>>>>>>>>> 2016.01.14
'        If StrConv(ODR_HANSEIHIN_K_REC.USE_YM, vbUnicode) <> Left(Format(Text1(ptxUSE_YM).Text, "YYYYMMDD"), 6) Then
'            Exit Do
'        End If
'>>>>>>>>>> 2016.01.14
        
        If StrConv(ODR_HANSEIHIN_K_REC.INPUT_NO, vbUnicode) <> Format(CInt(ORDR_OYA(TDBGrid1(GridOYA).Bookmark, Col_OYA_KEY)), "0000") Then
            Exit Do
        End If


        If StrConv(ODR_HANSEIHIN_K_REC.SEQNO, vbUnicode) = "000" Then
            Exit Do
        End If

        row = row + 1
        ORDR_KO.ReDim KO_Min_Row, row, KO_Min_Col, KO_Max_Col
        '品番
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ODR_HANSEIHIN_K_REC.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ODR_HANSEIHIN_K_REC.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ODR_HANSEIHIN_K_REC.KO_HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            
                Call UniCode_Conv(ITEMREC.HIN_NAME, "品番未登録！！")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ITEM")
                Exit Function
        End Select
        ORDR_KO(row, Col_KO_HIN_GAI) = Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_HIN_GAI, vbUnicode))
        ORDR_KO(row, Col_KO_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        ORDR_KO(row, Col_KO_QTY) = Format(CDbl(StrConv(ODR_HANSEIHIN_K_REC.KO_QTY, vbUnicode)), "#0.00")
        
        '必要数
        If IsNumeric(ORDR_OYA(i, Col_OYA_SHIJI_QTY)) Then
'            ORDR_KO(row, Col_KO_USE_QTY) = Format(CDbl((ORDR_KO(row, Col_KO_QTY)) * CDbl(ORDR_OYA(i, Col_OYA_SHIJI_QTY))) * -1, "#0")      '2016.01.14
            ORDR_KO(row, Col_KO_USE_QTY) = Format(CDbl((ORDR_KO(row, Col_KO_QTY)) * CDbl(ORDR_OYA(i, Col_OYA_SHIJI_QTY))), "#0")        '2016.01.14
        Else
            ORDR_KO(row, Col_KO_USE_QTY) = 0
        End If
    
    
        '在訂ﾌﾗｸﾞ
        
        
        Select Case StrConv(ODR_HANSEIHIN_K_REC.ZAITEI_F, vbUnicode)
        
            Case "0"
                ORDR_KO(row, Col_KO_ZAITEI_F) = False
            Case "1"
                ORDR_KO(row, Col_KO_ZAITEI_F) = True
        
        End Select
    
        ORDR_KO(row, Col_KO_KEY) = StrConv(ODR_HANSEIHIN_K_REC.SEQNO, vbUnicode)
            
    
    
    Loop
    
    Set TDBGrid1(GridKO).Array = ORDR_KO
    
    
    TDBGrid1(GridKO).Bookmark = Null

    TDBGrid1(GridKO).ReBind
    TDBGrid1(GridKO).Update
    TDBGrid1(GridKO).MoveFirst
    TDBGrid1(GridKO).ScrollBars = dbgAutomatic
'    TDBGrid1(GridKO).Bookmark = 1
    
    Text1(ptxUSE_YMD).Text = ORDR_OYA(i, Col_OYA_USE_YMD)
    Text1(ptxHIN_GAI).Text = ORDR_OYA(i, Col_OYA_HIN_GAI)
    
    
    KO_Grid_Set_Proc = False

End Function



Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    ODR30301.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR30301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR30301)


    ODR30301.MousePointer = vbDefault

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
        Case pcmbSM                 '仕向け先
            GW_SIMUKE = Left(Right(Combo1(pcmbSM).Text, 4), 2)
            'GW_JIGYOBU = Right(Combo1(pcmbJI).Text, 1)
            GW_JIGYOBU = Mid(Right(Combo1(pcmbSM).Text, 4), 3, 1)
            GW_NAIGAI = Right(Combo1(pcmbSM).Text, 1)

        Case Else
    
    End Select
    
    Call Tab_Ctrl(Shift)        '移動
    

End Sub

Private Sub Command1_Click(Index As Integer)
Dim sts     As Integer
Dim yn      As Integer
Dim X_i     As Integer
Dim W_After     As String

Dim W_PC        As String
Dim W_DT        As String
Dim c           As String
Dim W_Path      As String
Dim W_CNT       As Long

    Select Case Index
    
        Case FuncCOR
            If Grid_Cor_M <> True Then
                Exit Sub
            End If
            
            TDBGrid1(GridOYA).Update
            Set ORDR_OYA = TDBGrid1(GridOYA).Array
    
            '入力エラーﾁｪｯｸ
            
            For X_i = ptxTOP To ptxLAST
            
                If ERR_CHK(X_i) Then
                    Exit Sub
                End If
            
            Next X_i
            
            
            
            If Grid_Err_Chk() Then
                
                Command1(FuncCOR).SetFocus      '2016.01.14
                
                Exit Sub
            End If
    
    
            
            yn = MsgBox("更新しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                
                
                
                '更新処理
                If Update_Proc() Then
                    MsgBox "更新失敗しました。", vbExclamation
                    
                    Exit Sub
                End If
                Grid_Cor_M = False
                Grid_Req_M = True
            
            
                If Data_Disp() Then
                    MsgBox "指定条件の注文情報で、表示失敗！", vbExclamation
                End If
            
            End If
            
            DoEvents
            
            If ODR30301.MousePointer <> vbDefault Then
                Call Input_UnLock
            End If
            
            Text1(ptxTOP).SetFocus
            Call Text1_GotFocus(ptxTOP)
            
            
            Exit Sub
            
            
        Case FuncEND
            
            Unload Me
    
    
        Case 2
    
            If KO_Update_Proc() Then
                Unload Me
            End If
    End Select

End Sub

Private Sub Form_Load()
'Dim cc As tagINITCOMMONCONTROLSEX

Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer

Dim sBuffer As String * 255
Dim com     As String

Dim W_DATE  As String



    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "半製品管理画面", Me.hwnd, 0)
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)

    
    '画面初期処理
    
    ODR30301.Caption = ODR30301.Caption & LAST_UPDATE_DAY       '2016.01.14
    
    Show
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。", vbExclamation
        End
    End If
    
    
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
    
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '半製品管理データＯＰＥＮ
    If ODR_HANSEIHIN_Open(BtOpenNomal) Then
        Unload Me
    End If
    
'テキストを設定する
    Text1(ptxUSE_YM) = Left(Format(Date, "yyyy/mm/dd"), 7)
    
    
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSM, P_KBN04_CD, 0) Then
        Unload Me
    End If
    Combo1(pcmbSM).ListIndex = 0
'事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止します。"
        End
    End If
    
    If SET_JGYOBU_T Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止します。"
        End
    End If
    
    GW_SIMUKE = Left(Right(Combo1(pcmbSM).Text, 4), 2)
    'GW_JIGYOBU = Right(Combo1(pcmbJI).Text, 1)
    GW_JIGYOBU = Mid(Right(Combo1(pcmbSM).Text, 4), 3, 1)
    GW_NAIGAI = Right(Combo1(pcmbSM).Text, 1)
    
    'GW_SIMUKE = "01"
    'GW_JIGYOBU = "B"
    'GW_NAIGAI = "1"
    
    GW_HINGAI = ""
    GW_TOUGETU = Left(Format(Date, "yyyymmdd"), 6)
    'Combo1(pcmbSM).SetFocus
       
    Text1(ptxTOP).SetFocus
       
    Grid_Cor_M = False
    Grid_Req_M = False
    row = OYA_Min_Row - 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim yn      As Integer

    If UnloadMode = 1 Then Exit Sub
    
    If Grid_Cor_M = True Then
        yn = MsgBox("更新されていません！！" & Chr(13) & Chr(10) & _
                    "　終了しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
    Else
        yn = MsgBox("終了しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
        'yn = vbYes
    End If
    
    If yn = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts     As Integer






    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "TANTO")
        End If
    End If
    
    
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_COMPO")
        End If
    End If
    
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_CODE")
        End If
    End If
    
    
    
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ITEM")
        End If
    End If

    sts = BTRV(BtOpClose, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ITEM")
        End If
    End If


    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_KANRI")
        End If
    End If



    End
End Sub

Private Sub SHORI_Click(Index As Integer)
Dim yn      As Integer


    Select Case Index
    
        Case 0      '更新
            Call Command1_Click(FuncCOR)
            
            
        Case 1      '画面印刷
            yn = MsgBox("画面印刷しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
        
            
        
        Case 2      '終了
            Call Command1_Click(FuncEND)
    
    End Select


End Sub



Private Sub TDBGrid1_AfterColUpdate(Index As Integer, ByVal ColIndex As Integer)
Dim W_STR       As String
    
Dim W_Before    As String
Dim W_After     As String

'    If TDBGrid1(Index).Bookmark <= 0 Then Exit Sub
    
'    Cor_Row = TDBGrid1(Index).Bookmark
    
    'W_Before = Trim(ORDR_GRID(Cor_Row, ColIndex))
'    W_After = Trim(TDBGrid1(Index).Text)
    
    
'    TDBGrid1(Index).Update
    
'    Select Case Index
'        Case GridOYA
'           Set ORDR_OYA = TDBGrid1(Index).Array
'
'
'        Case GridKO
'           Set ORDR_KO = TDBGrid1(Index).Array
'
'    End Select
    
    'If W_Before <> W_After Then
    '    Grid_Cor_M = True
    'End If
    
    'If Grid_Err_Chk(ColIndex, W_Before, W_After) Then
        
    '    Exit Sub
    'End If

End Sub

Private Sub TDBGrid1_BeforeInsert(Index As Integer, Cancel As Integer)
    
    ORDR_OYA.ReDim OYA_Min_Row, ORDR_OYA.Count(1), OYA_Min_Col, OYA_Max_Col

End Sub






Private Sub TDBGrid1_Change(Index As Integer)
    Grid_Cor_M = True
End Sub

Private Sub TDBGrid1_Click(Index As Integer)
'DblClickに移行 2016.01.14
'    Select Case Index
'
'        Case GridOYA
'
'            If ORDR_OYA.Count(1) <= 0 Then
'                Exit Sub
'            End If
'
'            If TDBGrid1(Index).Bookmark > 0 Then
'
'                If KO_Grid_Set_Proc(TDBGrid1(Index).Bookmark) Then
'                    Unload Me
'                End If
'
'            End If
'
'    End Select
'
End Sub

Private Sub TDBGrid1_DblClick(Index As Integer)
    Select Case Index
    
        Case GridOYA
    
            If ORDR_OYA.Count(1) <= 0 Then
                Exit Sub
            End If
            
            If TDBGrid1(Index).Bookmark > 0 Then
    
                If KO_Grid_Set_Proc(TDBGrid1(Index).Bookmark) Then
                    Unload Me
                End If
    
            End If
    
    End Select

End Sub

Private Sub TDBGrid1_HeadClick(Index As Integer, ByVal ColIndex As Integer)
    
    If Index <> 0 Then
        Exit Sub
    End If
    
    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        ORDR_OYA.QuickSort OYA_Min_Row, ORDR_OYA.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1(GridOYA).Array = ORDR_OYA
        
        TDBGrid1(GridOYA).ReBind
        TDBGrid1(GridOYA).Update
        TDBGrid1(GridOYA).MoveFirst


    End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index))
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index))
    End If

End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim sts As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Text1(Index).Locked = True Then      'ロック中項目なら処理しない
        Call Tab_Ctrl(Shift)    '移動
        Exit Sub
    End If
                        '入力文字数チェック
    If ERR_CHK(Index) Then
        Call Text1_GotFocus(Index)
        Text1(Index).SetFocus
        Exit Sub
    End If
    
'    If Index = ptxUSE_YM Then      '2016.01.14
    If Index = ptxTANTO_CD Then     '2016.01.14
        If Data_Disp Then
            MsgBox "指定条件の注文情報で、表示失敗！", vbExclamation
            Call Text1_GotFocus(ptxTOP%)
            Text1(ptxTOP%).SetFocus
            Exit Sub
        End If
        
        
        TDBGrid1(GridOYA).SetFocus
        
        Exit Sub
    End If
    
    Call Tab_Ctrl(Shift)    '移動
    
End Sub


Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ作成
'----------------------------------------------------------------------------
Dim i           As Integer
Dim sts         As Integer

Dim com         As Integer


Dim INS_KEY     As Integer

    Update_Proc = True
                                     
    Set TDBGrid1(GridOYA).Array = ORDR_OYA
    TDBGrid1(GridOYA).Refresh
    
    TDBGrid1(GridOYA).Update
                                     
    If ORDR_OYA.Count(1) < 1 Then
        Update_Proc = False
        Exit Function
    End If
                                     'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
                                    
                                    
    For i = 1 To ORDR_OYA.Count(1)
                                    
'2016.01.14        Call UniCode_Conv(K0_ODR_HANSEIHIN.USE_YM, Left(Format(Text1(ptxUSE_YM) & "/01", "YYYYMMDD"), 6))
        Call UniCode_Conv(K0_ODR_HANSEIHIN.INPUT_NO, ORDR_OYA(i, Col_OYA_KEY))
        Call UniCode_Conv(K0_ODR_HANSEIHIN.SEQNO, "000")
                                    
        sts = BTRV(BtOpGetEqual, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
            Case BtErrKeyNotFound
                com = BtOpInsert
            Case Else
                Call File_Error(sts, BtOpGetEqual, "半製品管理")
                Exit Function
        End Select
    
    
    
        If ORDR_OYA(i, Col_OYA_DEL) Then
            If com = BtOpUpdate Then
    
                
                        
                sts = BTRV(BtOpDelete, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, BtOpDelete, "半製品管理")
                        Exit Function
                End Select
    
    
    
                com = BtOpGetGreaterEqual
                Do
                    DoEvents
                
                    sts = BTRV(com, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
                    Select Case sts
                        Case BtNoErr
                            
'2016.01.14                            If StrConv(ODR_HANSEIHIN_O_REC.USE_YM, vbUnicode) <> Left(Format(Text1(ptxUSE_YM) & "/01", "YYYYMMDD"), 6) Then
'2016.01.14                                Exit Do
'2016.01.14                            End If
                            
                            If StrConv(ODR_HANSEIHIN_O_REC.INPUT_NO, vbUnicode) <> ORDR_OYA(i, Col_OYA_KEY) Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call File_Error(sts, com, "半製品管理")
                            Exit Function
                    End Select
                
                    sts = BTRV(BtOpDelete, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
                    Select Case sts
                        Case BtNoErr
                        Case Else
                            Call File_Error(sts, BtOpDelete, "半製品管理")
                            Exit Function
                    End Select
                
                
                    com = BtOpGetNext
                
                Loop
    
    
            End If
    
        Else
    
            Select Case com
            
                Case BtOpInsert
                    '追加
                
'2016.01.14                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.USE_YM, Left(Format(Text1(ptxUSE_YM) & "/01", "YYYYMMDD"), 6))
                    
                    
                    
                    
                    
                    
                    INS_KEY = i
                    
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.INPUT_NO, Format(INS_KEY, "0000"))
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.USE_YMD, Format(ORDR_OYA(i, Col_OYA_USE_YMD), "YYYYMMDD"))
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.SEQNO, "000")
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.JGYOBU, GW_JIGYOBU)
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.NAIGAI, GW_NAIGAI)
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.HIN_GAI, ORDR_OYA(i, Col_OYA_HIN_GAI))
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.SHIJI_QTY, Format(CDbl(ORDR_OYA(i, Col_OYA_SHIJI_QTY)), "00000000.00"))
                
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.UPD_TANTO, Text1(ptxTANTO_CD).Text)
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.UPD_DATE, Format(Now, "YYYYMMDD"))
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.UPD_TIME, Format(Now, "HHMMSS"))
                
                
                    
                    Do
                    
                        sts = BTRV(BtOpInsert, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrDuplicates
                                INS_KEY = INS_KEY + 1
                            
                                Call UniCode_Conv(ODR_HANSEIHIN_O_REC.INPUT_NO, Format(INS_KEY, "0000"))
                            
                            Case Else
                                Call File_Error(sts, BtOpDelete, "半製品管理")
                                Exit Function
                        End Select
                    Loop
                    
                    com = BtOpGetGreater
                
                
                    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, GW_SIMUKE)
                    Call UniCode_Conv(K0_P_COMPO.JGYOBU, GW_JIGYOBU)
                    Call UniCode_Conv(K0_P_COMPO.NAIGAI, GW_NAIGAI)
                    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, ORDR_OYA(i, Col_OYA_HIN_GAI))
                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
                    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
                
                
                
                    Do
                    
                        DoEvents
                    
                    
                    
                        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> GW_SIMUKE Then
                                    Exit Do
                                End If
                            
                                If StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> GW_JIGYOBU Then
                                    Exit Do
                                End If
                            
                                If StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> GW_NAIGAI Then
                                    Exit Do
                                End If
                            
                                If Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(ORDR_OYA(i, Col_OYA_HIN_GAI)) Then
                                    Exit Do
                                End If
                            
                                If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                                    Exit Do
                                End If
                            
                            
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, com, "半製品管理")
                                Exit Function
                        End Select
                    
                    
                    
                    
'2016.01.14                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.USE_YM, Left(Format(Text1(ptxUSE_YM) & "/01", "YYYYMMDD"), 6))
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.INPUT_NO, Format(INS_KEY, "0000"))
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.USE_YMD, Format(ORDR_OYA(i, Col_OYA_USE_YMD), "YYYYMMDD"))
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.SEQNO, StrConv(P_COMPO_K_REC.SEQNO, vbUnicode))
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.KO_JGYOBU, SHIZAI)
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.KO_NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.KO_QTY, Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "00000.00"))
'2016.01.14                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.USE_QTY, Format((CDbl(ORDR_OYA(i, Col_OYA_SHIJI_QTY)) * _
'2016.01.14                                                                        CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))) * -1, "00000000.00"))
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.USE_QTY, Format((CDbl(ORDR_OYA(i, Col_OYA_SHIJI_QTY)) * _
                                                                        CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))), "00000000.00"))
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.ZAITEI_F, "0")
                    
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.UPD_TANTO, Text1(ptxTANTO_CD).Text)
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.UPD_DATE, Format(Now, "YYYYMMDD"))
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.UPD_TIME, Format(Now, "HHMMSS"))
                    
                    
                    
                        sts = BTRV(BtOpInsert, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_K_REC, Len(ODR_HANSEIHIN_K_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
                        Select Case sts
                            Case BtNoErr
                            Case Else
                                Call File_Error(sts, BtOpDelete, "半製品管理")
                                Exit Function
                        End Select
                    
                    
                        com = BtOpGetNext
                    
                    Loop
                
                
                
                
                
                
                
                Case BtOpUpdate
                   '変更
    
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.USE_YMD, Format(ORDR_OYA(i, Col_OYA_USE_YMD), "YYYYMMDD"))
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.SHIJI_QTY, Format(CDbl(ORDR_OYA(i, Col_OYA_SHIJI_QTY)), "00000000.00"))
                
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.UPD_TANTO, Text1(ptxTANTO_CD).Text)
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.UPD_DATE, Format(Now, "YYYYMMDD"))
                    Call UniCode_Conv(ODR_HANSEIHIN_O_REC.UPD_TIME, Format(Now, "HHMMSS"))
                
                
                    sts = BTRV(BtOpUpdate, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
                    Select Case sts
                        Case BtNoErr
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "半製品管理")
                            Exit Function
                    End Select
    
    
'2016.01.14                    Call UniCode_Conv(K0_ODR_HANSEIHIN.USE_YM, Left(Format(Text1(ptxUSE_YM) & "/01", "YYYYMMDD"), 6))
                    Call UniCode_Conv(K0_ODR_HANSEIHIN.INPUT_NO, Format(CInt(ORDR_OYA(i, Col_OYA_KEY)), "0000"))
                    Call UniCode_Conv(K0_ODR_HANSEIHIN.SEQNO, "001")
    
                        
    
    
                    com = BtOpGetGreaterEqual
                    Do
                        DoEvents
                    
                        sts = BTRV(com, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_K_REC, Len(ODR_HANSEIHIN_K_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
                        Select Case sts
                            Case BtNoErr
                                
'2016.01.14                                If StrConv(ODR_HANSEIHIN_K_REC.USE_YM, vbUnicode) <> Left(Format(Text1(ptxUSE_YM) & "/01", "YYYYMMDD"), 6) Then
'2016.01.14                                    Exit Do
'2016.01.14                                End If
                                
                                If StrConv(ODR_HANSEIHIN_K_REC.INPUT_NO, vbUnicode) <> Format(CInt(ORDR_OYA(i, Col_OYA_KEY)), "0000") Then
                                    Exit Do
                                End If
                            
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, com, "半製品管理")
                                Exit Function
                        End Select
                    
                        
'2016.01.14                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.USE_QTY, Format((CDbl(ORDR_OYA(i, Col_OYA_SHIJI_QTY)) * _
'2016.01.14                                                                        CDbl(StrConv(ODR_HANSEIHIN_K_REC.KO_QTY, vbUnicode))) * -1, "00000000.00"))
                        Call UniCode_Conv(ODR_HANSEIHIN_K_REC.USE_QTY, Format((CDbl(ORDR_OYA(i, Col_OYA_SHIJI_QTY)) * _
                                                                        CDbl(StrConv(ODR_HANSEIHIN_K_REC.KO_QTY, vbUnicode))), "00000000.00"))
                        
                        sts = BTRV(BtOpUpdate, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_K_REC, Len(ODR_HANSEIHIN_K_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
                        Select Case sts
                            Case BtNoErr
                            Case Else
                                Call File_Error(sts, BtOpDelete, "半製品管理")
                                Exit Function
                        End Select
                    
                    
                        com = BtOpGetNext
                    
                    Loop
    
            End Select
    
    
        End If
    
    Next i
                                    
                                    
                                    
                                        
                                        
End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function

Private Function KO_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   子部品更新
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

Dim i       As Integer


    
    Set TDBGrid1(GridKO).Array = ORDR_KO
    TDBGrid1(GridKO).Refresh
    
    TDBGrid1(GridKO).Update
    
    
    For i = 1 To ORDR_KO.Count(1)
    
        DoEvents
    
    
'2016.01.14        Call UniCode_Conv(K0_ODR_HANSEIHIN.USE_YM, Left(Format(Text1(ptxUSE_YM) & "/01", "YYYYMMDD"), 6))
        Call UniCode_Conv(K0_ODR_HANSEIHIN.INPUT_NO, Format(CInt(ORDR_OYA(TDBGrid1(GridOYA).Bookmark, Col_OYA_KEY)), "0000"))
        Call UniCode_Conv(K0_ODR_HANSEIHIN.SEQNO, Format(ORDR_KO(i, Col_KO_KEY), "000"))
                                        
        sts = BTRV(BtOpGetEqual, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_K_REC, Len(ODR_HANSEIHIN_K_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
            Case BtErrKeyNotFound
                com = BtOpInsert
            Case Else
                Call File_Error(sts, BtOpGetEqual, "半製品管理")
                Unload Me
        End Select
    
    
        If com = BtOpUpdate Then

            Select Case ORDR_KO(i, Col_KO_ZAITEI_F)
            
                Case False
                
                    Call UniCode_Conv(ODR_HANSEIHIN_K_REC.ZAITEI_F, "0")
                
                
                
                
                
                Case True
            
                    Call UniCode_Conv(ODR_HANSEIHIN_K_REC.ZAITEI_F, "1")
            
            
            End Select

            sts = BTRV(BtOpUpdate, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_K_REC, Len(ODR_HANSEIHIN_K_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                Case BtErrKeyNotFound
                    com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpUpdate, "半製品管理")
                    Unload Me
            End Select
        
        End If
    

    Next i

End Function
