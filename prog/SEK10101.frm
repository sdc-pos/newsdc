VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEK10101 
   BackColor       =   &H00C0C0C0&
   Caption         =   "出荷予定作成処理 [SEK1010] 2013.05.01 15:00"
   ClientHeight    =   9210
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   15810
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
   ScaleHeight     =   9210
   ScaleWidth      =   15810
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   11
      Left            =   7440
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｾｷｽｲ注文ﾃﾞｰﾀ検索(納入日/送り先)"
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   41
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   10
      Left            =   120
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   345
      Index           =   9
      Left            =   5160
      MaxLength       =   2
      TabIndex        =   34
      Top             =   8760
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "積水送り先抽出処理"
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   33
      Top             =   8760
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "出荷予定ﾃﾞｰﾀ送り先一括設定"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   8760
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   8
      Left            =   8400
      MaxLength       =   10
      TabIndex        =   31
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   345
      Index           =   7
      Left            =   14520
      MaxLength       =   8
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｾｷｽｲ注文ﾃﾞｰﾀ検索(作成日時/送り先)"
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   8
      Top             =   960
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   6
      Left            =   6240
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   5
      Left            =   4920
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   4
      Left            =   3360
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   3
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   2
      Left            =   120
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   0
      Left            =   720
      MaxLength       =   1
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
      Height          =   495
      Index           =   1
      Left            =   13800
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "出荷予定ﾃﾞｰﾀを作成"
      Height          =   495
      Index           =   0
      Left            =   10680
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   5655
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   9975
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ﾃﾞｰﾀ作成日"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ﾃﾞｰﾀ作成　　時刻"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "送り先"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "送り先名"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "納入日"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "件数"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "件数    (梱包残)"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "ﾃﾞｰﾀ作成"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=873"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2487"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2355"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2196"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2064"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2778"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=7779"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=7646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2487"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2355"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=1905"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1773"
      Splits(0)._ColumnProps(28)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=1905"
      Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=1773"
      Splits(0)._ColumnProps(33)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(35)=   "Column(8).Width=3413"
      Splits(0)._ColumnProps(36)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(8)._WidthInPix=3281"
      Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      EditDropDown    =   0   'False
      HeadLines       =   2
      FootLines       =   1
      AllowArrows     =   0   'False
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ Ｐゴシック"
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
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF80&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFFF00&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF80&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.bgcolor=&HFFFF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=106,.parent=87"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=110,.parent=87"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=87"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=114,.parent=87,.alignment=3"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=111,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=112,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=113,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=87"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=16,.parent=87,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=13,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=14,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=15,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=20,.parent=87,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=24,.parent=87,.alignment=3"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=21,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=22,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=23,.parent=91"
      _StyleDefs(72)  =   "Named:id=33:Normal"
      _StyleDefs(73)  =   ":id=33,.parent=0"
      _StyleDefs(74)  =   "Named:id=34:Heading"
      _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(76)  =   ":id=34,.wraptext=-1"
      _StyleDefs(77)  =   "Named:id=35:Footing"
      _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(79)  =   "Named:id=36:Selected"
      _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=37:Caption"
      _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(83)  =   "Named:id=38:HighlightRow"
      _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=39:EvenRow"
      _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(87)  =   "Named:id=40:OddRow"
      _StyleDefs(88)  =   ":id=40,.parent=33"
      _StyleDefs(89)  =   "Named:id=41:RecordSelector"
      _StyleDefs(90)  =   ":id=41,.parent=34"
      _StyleDefs(91)  =   "Named:id=42:FilterBar"
      _StyleDefs(92)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "納入日"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   40
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblIN 
      BackStyle       =   0  '透明
      BorderStyle     =   1  '実線
      Height          =   375
      Left            =   11880
      TabIndex        =   39
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "/"
      Height          =   255
      Index           =   15
      Left            =   11640
      TabIndex        =   38
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label lblOUT 
      BackStyle       =   0  '透明
      BorderStyle     =   1  '実線
      Height          =   375
      Left            =   10560
      TabIndex        =   37
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "ヵ月分"
      Height          =   255
      Index           =   14
      Left            =   5760
      TabIndex        =   36
      Top             =   8880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "本日より"
      Height          =   255
      Index           =   13
      Left            =   9120
      TabIndex        =   35
      Top             =   8880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "出荷日"
      Height          =   375
      Index           =   12
      Left            =   7560
      TabIndex        =   30
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblBEF_DATETIME 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   240
      Left            =   1920
      TabIndex        =   29
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label lblBEF_START_DENNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   240
      Left            =   6120
      TabIndex        =   28
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label lblSTART_DENNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   240
      Left            =   8520
      TabIndex        =   27
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label lblEXE 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   240
      Left            =   11160
      TabIndex        =   26
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "EXEﾌｫﾙﾀﾞ："
      Height          =   255
      Index           =   11
      Left            =   9960
      TabIndex        =   25
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "今回伝票№："
      Height          =   255
      Index           =   10
      Left            =   7080
      TabIndex        =   24
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "前回伝票№："
      Height          =   255
      Index           =   9
      Left            =   4680
      TabIndex        =   23
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "前回作成日時："
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   22
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   15600
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   15600
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   15600
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "最大件数"
      Height          =   255
      Index           =   7
      Left            =   13440
      TabIndex        =   21
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "送り先"
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   20
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "～"
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   19
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "ﾃﾞｰﾀ作成時刻"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   18
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "～"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   17
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "ﾃﾞｰﾀ作成日"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "運送会社"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   15
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "便№"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "SEK10101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxBin_No% = 0
Private Const ptxUNSOU_KAISHA% = 1

Private Const ptxS_SND_YMD% = 2
Private Const ptxE_SND_YMD% = 3

Private Const ptxS_SND_HMS% = 4
Private Const ptxE_SND_HMS% = 5

Private Const ptxTOK_CD1% = 6
Private Const ptxTOK_CD2% = 11


Private Const ptxListMax% = 7

Private Const ptxSYU_YMD% = 8

Private Const ptxSEL_TUKI% = 9


Private Const ptxNOU_YMD% = 10



Dim Y_Syuka_TEI     As New XArrayDB

Private Const Min_Row% = 1                  '最小行数

Dim Max_Row    As Integer                   'グリッド最大表示件数


Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 8                  '最大列数

Private Const ColNo% = 0                    '№
Private Const ColSND_YMD% = 1               'ﾃﾞｰﾀ作成日
Private Const ColSND_HMS% = 2               'ﾃﾞｰﾀ作成時刻
Private Const ColTOK_CD% = 3                '送り先
Private Const ColL_TOK_NAME% = 4            '送り先名
Private Const ColNOU_YMD% = 5               '送り先名



Private Const ColTOTAL_CNT% = 6             '件数
Private Const ColZAN_CNT% = 7               '件数(梱包残)

Private Const ColDATA_MAKE_DATETIME% = 8    'ﾃﾞｰﾀ作成


Private exeForder   As String

'   『邸別注文データのKEY定義』

Private Type KEY4_Y_SYU_TEI                 'ＫＥＹ４
    SND_YMD(0 To 7)                 As Byte         'データ作成日
    SND_HMS(0 To 5)                 As Byte         'データ作成時刻
    TOK_CD(0 To 7)                  As Byte         '得意先ｺｰﾄﾞ
    CHO_CD(0 To 7)                  As Byte         '直納先ｺｰﾄﾞ
End Type

Private K4_Y_SYU_TEI                As KEY4_Y_SYU_TEI

'>>>>>>>>>>>>>>>>>> 2012.12.25
Private Type KEY5_Y_SYU_TEI                 'ＫＥＹ5
    NOU_YMD(0 To 7)                 As Byte         '納入日

    SND_YMD(0 To 7)                 As Byte         'データ作成日
    SND_HMS(0 To 5)                 As Byte         'データ作成時刻
    TOK_CD(0 To 7)                  As Byte         '得意先ｺｰﾄﾞ
    CHO_CD(0 To 7)                  As Byte         '直納先ｺｰﾄﾞ

End Type

Private K5_Y_SYU_TEI                As KEY5_Y_SYU_TEI
'>>>>>>>>>>>>>>>>>> 2012.12.25


'
Private Type SEK1010_Y_SYU_TEI_FSpeck
    ks0     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体

    ks4     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.12.25
    ks5     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.12.25
    ks6     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.12.25
    ks7     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.12.25
    ks8     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.12.25


End Type

Private SEK1010_Y_SYU_TEI_Speck    As SEK1010_Y_SYU_TEI_FSpeck



Private Type SEK1010_Y_SYU_TEI_FSpeck2

    ks4     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.12.25
    ks5     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.12.25
    ks6     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.12.25
    ks7     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.12.25
    ks8     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.12.25


End Type
Private SEK1010_Y_SYU_TEI_Speck2    As SEK1010_Y_SYU_TEI_FSpeck2




'   『出荷予定ﾃﾞｰﾀ(H)のKEY定義』

Private Type KEY5_DEL_SYU_H             'ＫＥＹ５
    OKURISAKI_CD(0 To 8)            As Byte     '送り先CD
End Type

Private K5_DEL_SYU_H                As KEY5_DEL_SYU_H


'
Private Type SEK1010_DEL_SYU_H_FSpeck
    ks0     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private SEK1010_DEL_SYU_H_Speck    As SEK1010_DEL_SYU_H_FSpeck




Private Sub Command1_Click(index As Integer)
    
Dim ans As Integer
Dim sts As Long
    
    
    Select Case index
    
        Case 0
    
    
                
    
    
    
    
            
            If Y_Syuka_TEI.Count(1) = 0 Then        '2013.05.01
                Command1(index).SetFocus            '2013.05.01
                Exit Sub                            '2013.05.01
            End If                                  '2013.05.01
    
    
    
            If TDBGrid1.Bookmark <= 0 Then
                 
                MsgBox "ﾃﾞｰﾀ作成する行を選択してください"
                Command1(index).SetFocus
                Exit Sub
            
            End If
    
    
            If Not IsDate(Text1(ptxSYU_YMD).Text) Then
            
                MsgBox "出荷日を正しく入力してください"
                Text1(ptxSYU_YMD).SetFocus
                Exit Sub
    
            End If
    
            Text1(ptxSYU_YMD).Text = Format(Text1(ptxSYU_YMD).Text, "YYYY/MM/DD")
    
    
    
    
            If Trim(Y_Syuka_TEI(TDBGrid1.Bookmark, ColDATA_MAKE_DATETIME)) <> "" Then
                ans = MsgBox(TDBGrid1.Bookmark & "行目が選択されています。" & Chr(13) & Chr(10) & _
                    "「出荷予定ﾃﾞｰﾀ」作成済みです。　処理を継続しますか？", vbYesNo + vbDefaultButton2, "確認処理")
                
                If ans = vbNo Then
                    Command1(0).SetFocus
                    Exit Sub
                End If
            End If
    
    
    
            ans = MsgBox(TDBGrid1.Bookmark & "行目が選択されています。" & Chr(13) & Chr(10) & _
                        "「出荷予定作成処理　実行しますか？", vbYesNo + vbDefaultButton2, "確認処理")
            
            If ans = vbYes Then
                
                
                Call Y_SYU_MAKE_PROC
            
                sts = Shell(RTrim(exeForder) & "F102021.exe", vbNormalFocus)
            
                            
            End If
    

        Case 1


            Unload Me


        Case 2
        
            If Trim(Text1(ptxS_SND_YMD).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxS_SND_YMD).Text) Then
                    MsgBox "入力した項目はエラーです。（ﾃﾞｰﾀ作成日(開始)）"
                    Text1(ptxS_SND_YMD).SetFocus
                    Exit Sub
                End If
            End If
        
            If Trim(Text1(ptxE_SND_YMD).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxE_SND_YMD).Text) Then
                    MsgBox "入力した項目はエラーです。（ﾃﾞｰﾀ作成日(終了)）"
                    Text1(ptxE_SND_YMD).SetFocus
                    Exit Sub
                End If
            End If
        
            If Trim(Text1(ptxS_SND_HMS).Text) = "" Then
            Else
                If Len(Trim(Text1(ptxS_SND_HMS).Text)) <> 8 Then
                    MsgBox "入力した項目はエラーです。（ﾃﾞｰﾀ作成時刻(開始)）"
                    Text1(ptxS_SND_HMS).SetFocus
                    Exit Sub
                End If
            End If
        
            If Trim(Text1(ptxE_SND_HMS).Text) = "" Then
            Else
                If Len(Trim(Text1(ptxE_SND_HMS).Text)) <> 8 Then
                    MsgBox "入力した項目はエラーです。（ﾃﾞｰﾀ作成時刻(終了)）"
                    Text1(ptxE_SND_HMS).SetFocus
                    Exit Sub
                End If
            End If
        
        
            If List_Disp_Proc Then
                Unload Me
            End If

            
            TDBGrid1.SetFocus


        '>>>>>>>>>>>>>> 2012.12.25
        Case 3
        
            If Not IsDate(Text1(ptxNOU_YMD).Text) Then
                MsgBox "入力した項目はエラーです。（納入日）"
                Text1(ptxS_SND_YMD).SetFocus
                Exit Sub
            End If
        
        
        
            If List_Disp_NOU_Proc Then
                Unload Me
            End If

            
            TDBGrid1.SetFocus

        '>>>>>>>>>>>>>> 2012.12.25

    End Select

End Sub

Private Sub Command2_Click(index As Integer)
    
    
Dim yn          As Integer
    
    
Dim com         As Integer
Dim sts         As Integer
    
    
Dim OKURI_TBL   As Variant
    
Dim c           As String
    
    
Dim In_Cnt      As Long
Dim Out_Cnt     As Long
    
    
    Select Case index
    
    
        Case 0
    
            yn = MsgBox("出荷予定ﾃﾞｰﾀ送り先一括設定 実行しますか？", vbYesNo + vbDefaultButton2, "確認入力")
            If yn = vbYes Then
                
                If Y_SYU_H_Open(BtOpenNomal) Then                 '出荷予定データ(ホストイメージ)
                    Exit Sub
                End If
                
                
                
                Call UniCode_Conv(K4_Y_SYU_H.ID_NO, "888")
                
                com = BtOpGetGreater
                
                
                Do
                
                    DoEvents
                    
                    sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                    Select Case sts
                        Case BtNoErr
                
                            If Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 3) <> "888" Then
                                Exit Do
                            End If
                
                        Case BtErrEOF
                            Exit Do
                
                        Case Else
                            Call File_Error(sts, com, "出荷予定ﾃﾞｰﾀ(H)")
                            Unload Me
                    End Select
                
                    If GetIni(App.EXEName, StrConv(Y_SYU_HREC.OKURISAKI_CD, vbUnicode), App.EXEName, c) Then
                        c = " "
                    End If
                    OKURI_TBL = Split(Trim(c), ",", -1)
                
                    
                    If UBound(OKURI_TBL) > 5 Then
                        Call UniCode_Conv(Y_SYU_HREC.TEL_NO, CStr(OKURI_TBL(6)))
                        
                        Call UniCode_Conv(Y_SYU_HREC.JYUSHO, CStr(OKURI_TBL(3)) & CStr(OKURI_TBL(4)))
                    Else
                    
                        If UBound(OKURI_TBL) > 3 Then
                        
                            Call UniCode_Conv(Y_SYU_HREC.JYUSHO, CStr(OKURI_TBL(3)) & CStr(OKURI_TBL(4)))
                        
                        
                        Else
                            If UBound(OKURI_TBL) > 2 Then
                            
                                Call UniCode_Conv(Y_SYU_HREC.JYUSHO, CStr(OKURI_TBL(3)))
                            
                            End If
                        
                        End If
                    End If
                    
                    
                    Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                    Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                 
                    
                    
                    
                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                    Select Case sts
                        Case BtNoErr
                
                
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "出荷予定ﾃﾞｰﾀ(H)")
                            Unload Me
                    End Select
                
                
                    com = BtOpGetNext
                
                
                Loop
                
                
                                                                '出荷予定データ(H)ＣＬＯＳＥ
                sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
                If sts Then
                    Call File_Error(sts, BtOpClose, "出荷予定データ(H)")
                End If
                
                
                MsgBox "出荷予定ﾃﾞｰﾀ送り先一括設定 終了しました。"
            End If
    
        Case 1
            yn = MsgBox("積水送り先抽出処理　実行しますか？", vbYesNo + vbDefaultButton2, "確認入力")
            If yn = vbYes Then
                
                com = BtOpGetLast
            
            
                In_Cnt = 0
                Out_Cnt = 0
            
                Do
                    DoEvents
                
                
                    sts = BTRV(com, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K3_DEL_SYU_H, Len(K3_DEL_SYU_H), 3)
                    Select Case sts
                        Case BtNoErr
                
                            If StrConv(DEL_SYU_HREC.SYUKA_YMD, vbUnicode) < _
                                        Format(DateAdd("m", Val(Text1(ptxSEL_TUKI).Text) * -1, Format(Now, "YYYY/MM/DD")), "YYYYMMDD") Then
                                Exit Do
                            End If
                            
                
                        Case BtErrEOF
                            Exit Do
                
                        Case Else
                            Call File_Error(sts, com, "出荷予定ﾃﾞｰﾀ(H)")
                            Unload Me
                    End Select
                
                
'>>>>>>>>>>>>>>>
'                    Call UniCode_Conv(K0_SEK_OKURISAKI.MUKE_CODE, StrConv(DEL_SYU_HREC.OKURISAKI, vbUnicode))
'
'                    sts = BTRV(BtOpGetEqual, SEK_OKURISAKI_POS, SEK_OKURISAKIREC, Len(SEK_OKURISAKIREC), K0_SEK_OKURISAKI, Len(K0_SEK_OKURISAKI), 0)
'                    Select Case sts
'                        Case BtNoErr
'
'
'                        Case BtErrKeyNotFound
'
'
'                            Call UniCode_Conv(K2_Y_SYU_TEI.KEN_NO, StrConv(DEL_SYU_HREC.SEK_KEN_NO, vbUnicode))
'                            Call UniCode_Conv(K2_Y_SYU_TEI.HIN_NO, StrConv(DEL_SYU_HREC.SEK_HIN_NO, vbUnicode))
'
'                            sts = BTRV(BtOpGetEqual, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
'                            Select Case sts
'                                Case BtNoErr
'
'
'
'
'
'                                    Call UniCode_Conv(SEK_OKURISAKIREC.MUKE_CODE, StrConv(DEL_SYU_HREC.OKURISAKI, vbUnicode))
'                                    Call UniCode_Conv(SEK_OKURISAKIREC.MUKE_NAME, StrConv(DEL_SYU_HREC.MUKE_NAME, vbUnicode))
'                                    Call UniCode_Conv(SEK_OKURISAKIREC.JYUSHO, StrConv(DEL_SYU_HREC.JYUSHO, vbUnicode))
'                                    Call UniCode_Conv(SEK_OKURISAKIREC.TEL_NO, StrConv(DEL_SYU_HREC.TEL_NO, vbUnicode))
'                                    Call UniCode_Conv(SEK_OKURISAKIREC.YUBIN_NO, StrConv(DEL_SYU_HREC.YUBIN_NO, vbUnicode))
'                                    Call UniCode_Conv(SEK_OKURISAKIREC.FILLER, "")
'
'                                    sts = BTRV(BtOpInsert, SEK_OKURISAKI_POS, SEK_OKURISAKIREC, Len(SEK_OKURISAKIREC), K0_SEK_OKURISAKI, Len(K0_SEK_OKURISAKI), 0)
'
'                                    Select Case sts
'                                        Case BtNoErr
'
'                                        Case BtErrDuplicates
'
'                                        Case Else
'
'                                            Call File_Error(sts, BtOpInsert, "積水送り先")
'                                            Unload Me
'
'
'                                    End Select
'
'
'
'                                Case BtErrKeyNotFound
'                                Case Else
'                                    Call File_Error(sts, BtOpGetEqual, "邸別注文ﾃﾞｰﾀ")
'                                    Unload Me
'
'                            End Select
'
'
'
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "積水送り先")
'                            Unload Me
'                    End Select
'>>>>>>>>>>>>>>>
                     In_Cnt = In_Cnt + 1
                     lblIN = In_Cnt
                     
                     
                     If Left(StrConv(DEL_SYU_HREC.MUKE_NAME, vbUnicode), 2) = "積水" Then
                         Call UniCode_Conv(SEK_OKURISAKIREC.OKURISAKI_CD, StrConv(DEL_SYU_HREC.OKURISAKI_CD, vbUnicode))
                         Call UniCode_Conv(SEK_OKURISAKIREC.MUKE_NAME, StrConv(DEL_SYU_HREC.MUKE_NAME, vbUnicode))
                         Call UniCode_Conv(SEK_OKURISAKIREC.JYUSHO, StrConv(DEL_SYU_HREC.JYUSHO, vbUnicode))
                         Call UniCode_Conv(SEK_OKURISAKIREC.TEL_NO, StrConv(DEL_SYU_HREC.TEL_NO, vbUnicode))
                         Call UniCode_Conv(SEK_OKURISAKIREC.YUBIN_NO, StrConv(DEL_SYU_HREC.YUBIN_NO, vbUnicode))
                         Call UniCode_Conv(SEK_OKURISAKIREC.FILLER, "")

                         sts = BTRV(BtOpInsert, SEK_OKURISAKI_POS, SEK_OKURISAKIREC, Len(SEK_OKURISAKIREC), K0_SEK_OKURISAKI, Len(K0_SEK_OKURISAKI), 0)

                         Select Case sts

                             Case BtNoErr


                                 Out_Cnt = Out_Cnt + 1
                                 lblOUT = Out_Cnt

                             Case BtErrDuplicates

                             Case Else

                                 Call File_Error(sts, BtOpInsert, "積水送り先")
                                 Unload Me


                        End Select
                    
                     
                     
                     End If
                
                
                    com = BtOpGetPrev
                
                
                Loop
                MsgBox "積水送り先抽出処理 終了しました。"
            
            End If
    End Select


    Command2(index).SetFocus

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()

Dim c   As String * 128

Dim sts As Integer



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
                                
                                '便№取り込み
    If GetIni(App.EXEName, "BIN_NO", App.EXEName, c) Then
    Else
        Text1(ptxBin_No) = RTrim(c)
    End If
                                '運送会社取り込み
    If GetIni(App.EXEName, "UNSOU_KAISHA", App.EXEName, c) Then
    Else
        Text1(ptxUNSOU_KAISHA) = RTrim(c)
    End If
                                '最大件数取り込み
    If GetIni(App.EXEName, "LISTMAX", App.EXEName, c) Then
    Else
        Text1(ptxListMax) = Val(RTrim(c))
    End If
                                
                                '出荷日
    Text1(ptxSYU_YMD) = Format(Now, "YYYY/MM/DD")
                                
                                
                                
                                
    If Y_SYU_TEI_Open(BtOpenNomal) Then                 '出荷予定データ(邸別)
        Unload Me
    End If
                                
    sts = BTRV(BtOpGetFirst, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K4_Y_SYU_TEI, Len(K4_Y_SYU_TEI), 4)
    Select Case sts
        Case BtNoErr
        Case BtErrIvldKey

            If Y_SYU_TEI_Create_Index() Then
                Unload Me
            End If

        Case Else
            Call File_Error(sts, BtOpGetFirst, "邸別注文ﾃﾞｰﾀ", 0)
            Unload Me
    End Select
                                
    Show
                                
                                
    If DEL_SYU_H_Open(BtOpenNomal) Then                 '出荷予定データ(H)
        Unload Me
    End If
                                
'    sts = BTRV(BtOpGetFirst, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K5_DEL_SYU_H, Len(K5_DEL_SYU_H), 5)
'    Select Case sts
'        Case BtNoErr
'        Case BtErrIvldKey
'
'            If DEL_SYU_H_Create_Index() Then
'                Unload Me
'            End If
'
'        Case Else
'            Call File_Error(sts, BtOpGetFirst, "出荷予定データ(H)", 0)
'            Unload Me
'    End Select
                                
    If SEK_OKURISAKI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                'EXEﾌｫﾙﾀﾞｰ取り込み
    exeForder = ""
    If GetIni(App.EXEName, "EXE", App.EXEName, c) Then
    Else
        exeForder = RTrim(c)
    End If


    Call INI_Disp_Proc



    Text1(ptxS_SND_YMD).Text = Format(Now, "YYYY/MM/DD")
    Text1(ptxE_SND_YMD).Text = Format(Now, "YYYY/MM/DD")



    Text1(ptxSEL_TUKI).Text = 1
    

                                '納入日 2012.12.25
    Text1(ptxNOU_YMD) = Format(Now, "YYYY/MM/DD")


End Sub
Private Sub Form_Unload(CANCEL As Integer)
    
Dim sts     As Integer
    
    
    sts = BTRV(BtOpDropSupIndex, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K4_Y_SYU_TEI, Len(K4_Y_SYU_TEI), 4)
    If sts Then
        Call File_Error(sts, BtOpDropSupIndex, "邸別注文データ")
    End If
     
    '2012.12.25
    sts = BTRV(BtOpDropSupIndex, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K4_Y_SYU_TEI, Len(K4_Y_SYU_TEI), 4)
    If sts Then
        Call File_Error(sts, BtOpDropSupIndex, "邸別注文データ")
    End If
    '2012.12.25
    
    
'   sts = BTRV(BtOpDropSupIndex, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K5_DEL_SYU_H, Len(K5_DEL_SYU_H), 5)
'   If sts Then
'       Call File_Error(sts, BtOpDropSupIndex, "出荷予定ﾃﾞｰﾀ(H)")
'   End If
    
    
    
                                                    '邸別注文データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "邸別注文データ")
    End If
                                                    '出荷予定データ(H)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K0_DEL_SYU_H, Len(K0_DEL_SYU_H), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "出荷予定ﾃﾞｰﾀ(H)")
    End If
        
    
    
    sts = BTRV(BtOpReset, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    
    
    Set SEK10101 = Nothing
        
    End
End Sub

Private Sub Y_SYU_MAKE_PROC()

Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer
        
Dim DENNO_A         As String * 1
Dim DENNO           As Long
Dim SEQNO           As Long
        
        
        
Dim FileName        As String
Dim HS_SMEISAI_OP   As Boolean
Dim Ret             As String
Dim HS_SMEISAINo    As Long
        
Dim c               As String * 128
        
Dim Upd_Now         As String
        
Dim wkTOK_CD        As String * 16
        
    DoEvents
        
        
                                '開始伝票№取り込み
    If GetIni(App.EXEName, "START_DENNO", App.EXEName, c) Then
        Beep
        MsgBox "[開始伝票№(START_DENNO)]の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    
    If Len(RTrim(c)) <> 7 Then
        Beep
        MsgBox "[開始伝票№(START_DENNO)]を正しく設定してください(１桁目：英字 ２～６桁目：数時)。処理を中止して下さい。"
        End
    End If
    
    
    DENNO_A = RTrim(Mid(c, 1, 1))
    DENNO = Val(RTrim(Mid(c, 2, 6)))
        
                                '前回伝票№出力
    If WriteIni(App.EXEName, "BEF_START_DENNO", App.EXEName, DENNO_A & Format(DENNO, "000000")) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & "BEF_START_DENNO")
        Unload Me
    End If
                                '前回実行日時出力
    If WriteIni(App.EXEName, "BEF_DATETIME", App.EXEName, Format(Now, "YYYY/MM/DD HH:MM:SS")) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & "BEF_DATETIME")
        Unload Me
    End If
        
        
        
        
        
    '出荷明細ファイル名取り込み & ＯＰＥＮ
    If GetIni("FILE", "HS_SMEISAI", "SYS", c) Then
        Beep
        MsgBox "出荷明細ファイル・ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Exit Sub
    End If
    FileName = Trim(c)

    HS_SMEISAI_OP = False

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & "_" & "B" & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
    
    
    On Error GoTo Exit_Proc
    
    HS_SMEISAINo = FreeFile
    Open FileName For Output As #HS_SMEISAINo
    On Error GoTo 0
    
    HS_SMEISAI_OP = True
        
        
        
    SEQNO = 1
        
        
        
    Call UniCode_Conv(K4_Y_SYU_TEI.SND_YMD, Format(Y_Syuka_TEI(TDBGrid1.Bookmark, ColSND_YMD), "YYYYMMDD"))
    Call UniCode_Conv(K4_Y_SYU_TEI.SND_HMS, Mid(Y_Syuka_TEI(TDBGrid1.Bookmark, ColSND_HMS), 1, 2) & _
                                            Mid(Y_Syuka_TEI(TDBGrid1.Bookmark, ColSND_HMS), 4, 2) & _
                                            Mid(Y_Syuka_TEI(TDBGrid1.Bookmark, ColSND_HMS), 7, 2))
    wkTOK_CD = Y_Syuka_TEI(TDBGrid1.Bookmark, ColTOK_CD)
    Call UniCode_Conv(K4_Y_SYU_TEI.TOK_CD, Left(wkTOK_CD, 8))
    Call UniCode_Conv(K4_Y_SYU_TEI.CHO_CD, Right(wkTOK_CD, 8))
        
        
    com = BtOpGetGreaterEqual

    Upd_Now = Format(Now, "YYYYMMDDHHMMSS")

    Do
        DoEvents
        sts = BTRV(com, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K4_Y_SYU_TEI, Len(K4_Y_SYU_TEI), 4)
        Select Case sts
            Case BtNoErr
            
                If StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode) <> Format(Y_Syuka_TEI(TDBGrid1.Bookmark, ColSND_YMD), "YYYYMMDD") Then
                    Exit Do
                End If
            
            
                If StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode) <> (Mid(Y_Syuka_TEI(TDBGrid1.Bookmark, ColSND_HMS), 1, 2) & _
                                                                    Mid(Y_Syuka_TEI(TDBGrid1.Bookmark, ColSND_HMS), 4, 2) & _
                                                                    Mid(Y_Syuka_TEI(TDBGrid1.Bookmark, ColSND_HMS), 7, 2)) Then
                    Exit Do
                End If
                    
                    
                If RTrim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode) & StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode)) <> _
                    RTrim(Y_Syuka_TEI(TDBGrid1.Bookmark, ColTOK_CD)) Then
            
                    Exit Do
            
                End If
            
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Do
                End If
            
                com = BtOpGetEqual
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "邸別注文データ")
                Exit Do
        End Select
    
            
            
        Print #HS_SMEISAINo, vbTab;
        '№
        Print #HS_SMEISAINo, Format(SEQNO, "#"); vbTab;
        SEQNO = SEQNO + 1
        '出荷日
        'Print #HS_SMEISAINo, Right(Format(Now, "YYYY/MM/DD"), 5); vbTab;
        Print #HS_SMEISAINo, Right(Format(Text1(ptxSYU_YMD), "YYYY/MM/DD"), 5); vbTab;
        '送り先集約CD
        Print #HS_SMEISAINo, vbTab; vbTab;
        '送り先CD
        Print #HS_SMEISAINo, RTrim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode)) & RTrim(StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode)); vbTab; vbTab;
        '送り先名
        Print #HS_SMEISAINo, RTrim(StrConv(Y_SYU_TEI_REC.L_TOK_NAME, vbUnicode)); vbTab;
        '売伝
        Print #HS_SMEISAINo, vbTab; vbTab;
        '伝票№
        Print #HS_SMEISAINo, DENNO_A & Format(DENNO, "000000"); vbTab; ; vbTab;
        DENNO = DENNO + 1
        '品番
        Print #HS_SMEISAINo, RTrim(StrConv(Y_SYU_TEI_REC.HINB_CD, vbUnicode)); vbTab;
        '数量
        Print #HS_SMEISAINo, Format(Val(StrConv(Y_SYU_TEI_REC.JUC_SUU, vbUnicode)), "#"); vbTab; vbTab;
        '注文№
        Print #HS_SMEISAINo, RTrim(StrConv(Y_SYU_TEI_REC.CHU_CD, vbUnicode)); vbTab;
        '得意先CD
        Print #HS_SMEISAINo, RTrim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode)); vbTab;
        '得意先名
        Print #HS_SMEISAINo, RTrim(StrConv(Y_SYU_TEI_REC.L_TOK_NAME, vbUnicode)); vbTab; vbTab;
        '備考
        Print #HS_SMEISAINo, RTrim(StrConv(Y_SYU_TEI_REC.TEI_NM, vbUnicode)); vbTab; vbTab;
        '運送会社
        Print #HS_SMEISAINo, RTrim(Text1(ptxUNSOU_KAISHA).Text); vbTab;
        '便
        Print #HS_SMEISAINo, RTrim(Text1(ptxBin_No).Text); vbTab;
        
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'        Call UniCode_Conv(K5_DEL_SYU_H.OKURISAKI_CD, RTrim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode)) & RTrim(StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode)))
'
'        sts = BTRV(BtOpGetEqual, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K5_DEL_SYU_H, Len(K5_DEL_SYU_H), 5)
'        Select Case sts
'            Case BtNoErr
'            Case BtErrKeyNotFound
'                Call UniCode_Conv(DEL_SYU_HREC.JYUSHO, "")
'                Call UniCode_Conv(DEL_SYU_HREC.TEL_NO, "")
'                Call UniCode_Conv(DEL_SYU_HREC.YUBIN_NO, "")
'            Case Else
'                Call File_Error(sts, BtOpGetEqual, "出荷予定ﾃﾞｰﾀ(H)")
'                Exit Do
'        End Select
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
        
        Call UniCode_Conv(K0_SEK_OKURISAKI.OKURISAKI_CD, RTrim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode)) & RTrim(StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode)))
        sts = BTRV(BtOpGetEqual, SEK_OKURISAKI_POS, SEK_OKURISAKIREC, Len(SEK_OKURISAKIREC), K0_SEK_OKURISAKI, Len(K0_SEK_OKURISAKI), 0)
        Select Case sts
            Case BtNoErr
            
            
                Call UniCode_Conv(DEL_SYU_HREC.JYUSHO, StrConv(SEK_OKURISAKIREC.JYUSHO, vbUnicode))
                Call UniCode_Conv(DEL_SYU_HREC.TEL_NO, StrConv(SEK_OKURISAKIREC.TEL_NO, vbUnicode))
                Call UniCode_Conv(DEL_SYU_HREC.YUBIN_NO, StrConv(SEK_OKURISAKIREC.YUBIN_NO, vbUnicode))
            
            Case BtErrKeyNotFound
                Call UniCode_Conv(DEL_SYU_HREC.JYUSHO, "")
                Call UniCode_Conv(DEL_SYU_HREC.TEL_NO, "")
                Call UniCode_Conv(DEL_SYU_HREC.YUBIN_NO, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "積水送り先")
                Exit Do
        End Select
                
        
        
        
        
        '住所
        Print #HS_SMEISAINo, RTrim(StrConv(DEL_SYU_HREC.JYUSHO, vbUnicode)), vbTab;
        '郵便番号
        Print #HS_SMEISAINo, RTrim(StrConv(DEL_SYU_HREC.YUBIN_NO, vbUnicode)), vbTab;
        '電話番号
        Print #HS_SMEISAINo, RTrim(StrConv(DEL_SYU_HREC.TEL_NO, vbUnicode)), vbTab;
        
        
        '件名管理№
        Print #HS_SMEISAINo, RTrim(StrConv(Y_SYU_TEI_REC.KEN_NO, vbUnicode)); vbTab;
        '品管理№
        Print #HS_SMEISAINo, RTrim(StrConv(Y_SYU_TEI_REC.HIN_NO, vbUnicode))
    
    
        Call UniCode_Conv(Y_SYU_TEI_REC.DATA_MAKE_DATETIME, Upd_Now)
                
        Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, App.EXEName)
        Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, Format(Now, "YYYYMMDDHHMMSS"))
        
        
        sts = BTRV(BtOpUpdate, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpUpdate, "邸別注文データ")
                Exit Do
        End Select
        
        
        
        com = BtOpGetNext
    
    Loop
        
                                                    
    Close #HS_SMEISAINo
                                                    
                                                    
                                    '伝票№出力
    If WriteIni(App.EXEName, "START_DENNO", App.EXEName, DENNO_A & Format(DENNO, "000000")) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & "START_DENNO")
        Unload Me
    End If
                                                    
                                                    
    Call INI_Disp_Proc
                                                    
                                                    
    Command1(2).Value = True
                                                    
                                                    

    Exit Sub


Exit_Proc:
    
    
MsgBox Err.Number
    
    
    If HS_SMEISAI_OP Then
        Close #HS_SMEISAINo
    End If
    


End Sub

Private Sub Text1_GotFocus(index As Integer)
    
    If Text1(index).TabStop = True Then
        Text1(index) = Trim(Text1(index).Text)
        Text1(index).SelStart = 0
        Text1(index).SelLength = Len(Text1(index).Text)
    End If

End Sub

Private Function List_Disp_Proc() As Integer


Dim svSND_YMD               As String
Dim svSND_HMS               As String
Dim svTOK_CD                As String
Dim svL_TOK_NAME            As String
Dim svDATA_MAKE_DATETIME    As String
Dim svNou_YMD               As String


Dim TOTAL_CNT               As Long
Dim ZAN_CNT                 As Long

Dim Row                     As Long
        
Dim com                     As Integer
Dim sts                     As Integer
Dim ans                     As Integer

Dim wkTOK_CD                As String * 16

Dim SKIP_F                  As Boolean


    List_Disp_Proc = True

    Call Input_Lock


                        'テーブルリセット
    Set Y_Syuka_TEI = Nothing
            
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2013.05.01
    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2013.05.01
    
    
    
    
    Row = Min_Row - 1
    
    svSND_YMD = ""
    
        
    Call UniCode_Conv(K4_Y_SYU_TEI.SND_YMD, "")
    Call UniCode_Conv(K4_Y_SYU_TEI.SND_HMS, "")
    Call UniCode_Conv(K4_Y_SYU_TEI.TOK_CD, "")
    Call UniCode_Conv(K4_Y_SYU_TEI.CHO_CD, "")
        
        
    If Trim(Text1(ptxS_SND_YMD).Text) <> "" Then
        Call UniCode_Conv(K4_Y_SYU_TEI.SND_YMD, Format(Text1(ptxS_SND_YMD).Text, "YYYYMMDD"))
    End If
        
    If Trim(Text1(ptxS_SND_HMS).Text) <> "" Then
        Call UniCode_Conv(K4_Y_SYU_TEI.SND_HMS, Mid(Text1(ptxS_SND_HMS).Text, 1, 2) & _
                                                Mid(Text1(ptxS_SND_HMS).Text, 4, 2) & _
                                                Mid(Text1(ptxS_SND_HMS).Text, 7, 2))
    End If
        
        
        
        
    If Trim(Text1(ptxTOK_CD1).Text) <> "" Then
                
'        wkTOK_CD = Text1(ptxTOK_CD).Text                                   '2013.05.01
                
'        Call UniCode_Conv(K4_Y_SYU_TEI.TOK_CD, Left(wkTOK_CD, 8))          '2013.05.01
'        Call UniCode_Conv(K4_Y_SYU_TEI.CHO_CD, Right(wkTOK_CD, 8))         '2013.05.01
            
        Call UniCode_Conv(K4_Y_SYU_TEI.TOK_CD, Text1(ptxTOK_CD1).Text)      '2013.05.01
        Call UniCode_Conv(K4_Y_SYU_TEI.CHO_CD, Text1(ptxTOK_CD2).Text)      '2013.05.01
            
    
    End If
        
    com = BtOpGetGreaterEqual

    Do
        DoEvents
        sts = BTRV(com, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K4_Y_SYU_TEI, Len(K4_Y_SYU_TEI), 4)
        Select Case sts
            Case BtNoErr
                SKIP_F = False
            
                If Trim(Text1(ptxS_SND_YMD).Text) <> "" Then
                    If StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode) > Format(Text1(ptxE_SND_YMD).Text, "YYYYMMDD") Then
                        Exit Do
                    End If
                End If
            
                If Trim(Text1(ptxE_SND_HMS).Text) <> "" Then
                    If StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode) > Mid(Text1(ptxE_SND_HMS).Text, 1, 2) & _
                                                                    Mid(Text1(ptxE_SND_HMS).Text, 4, 2) & _
                                                                    Mid(Text1(ptxE_SND_HMS).Text, 7, 2) Then
                        Exit Do
                    End If
                End If
            
            
                If Trim(Text1(ptxTOK_CD1).Text) <> "" Then
    '2013.05.01
'                    If RTrim(Text1(ptxTOK_CD).Text) <> RTrim((StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode) & StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode))) Then
                    If RTrim(Text1(ptxTOK_CD1).Text) <> RTrim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode)) Or _
                        RTrim(Text1(ptxTOK_CD2).Text) <> RTrim(StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode)) Then
                        SKIP_F = True       '2013.05.01

'                        Exit Do            '2013.05.01
                    End If
    '2013.05.01
                End If
            
            
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Do
                End If
            
                com = BtOpGetEqual
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "邸別注文データ")
                Exit Do
        End Select

                    
        If Not SKIP_F Then                  '2013.05.01
            If Trim(svSND_YMD) = "" Then
                svSND_YMD = StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode)
                svSND_HMS = StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode)
                svTOK_CD = (StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode) & StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode))
                svL_TOK_NAME = StrConv(Y_SYU_TEI_REC.L_TOK_NAME, vbUnicode)
                svDATA_MAKE_DATETIME = StrConv(Y_SYU_TEI_REC.DATA_MAKE_DATETIME, vbUnicode)
                svNou_YMD = StrConv(Y_SYU_TEI_REC.NOU_YMD, vbUnicode)
            
                TOTAL_CNT = 0
                ZAN_CNT = 0
            End If
                        
                        
            If svSND_YMD <> StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode) Or _
                svSND_HMS <> StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode) Or _
                RTrim(svTOK_CD) <> RTrim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode) & StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode)) Then
                
                Row = Row + 1
                Y_Syuka_TEI.ReDim Min_Row, Row, Min_Col, Max_Col
                
                
                '№
                Y_Syuka_TEI(Row, ColNo) = Row
                'ﾃﾞｰﾀ作成日付
                Y_Syuka_TEI(Row, ColSND_YMD) = Mid(svSND_YMD, 1, 4) & "/" & Mid(svSND_YMD, 5, 2) & "/" & Mid(svSND_YMD, 7, 2)
                'ﾃﾞｰﾀ作成時刻
                Y_Syuka_TEI(Row, ColSND_HMS) = Mid(svSND_HMS, 1, 2) & ":" & Mid(svSND_HMS, 3, 2) & ":" & Mid(svSND_HMS, 5, 2)
                '送り先
                Y_Syuka_TEI(Row, ColTOK_CD) = svTOK_CD
                '送り先名
                Y_Syuka_TEI(Row, ColL_TOK_NAME) = svL_TOK_NAME
                
                '納入日
                If Trim(svNou_YMD) <> "" Then
                    Y_Syuka_TEI(Row, ColNOU_YMD) = Mid(svNou_YMD, 1, 4) & "/" & Mid(svNou_YMD, 5, 2) & "/" & Mid(svNou_YMD, 7, 2)
                Else
                    Y_Syuka_TEI(Row, ColNOU_YMD) = ""
                End If
                
                '件数
                Y_Syuka_TEI(Row, ColTOTAL_CNT) = TOTAL_CNT
                '件数(梱包残)
                Y_Syuka_TEI(Row, ColZAN_CNT) = ZAN_CNT
                'ﾃﾞｰﾀ作成
                Y_Syuka_TEI(Row, ColDATA_MAKE_DATETIME) = svDATA_MAKE_DATETIME
                                                    
                                                    
                svSND_YMD = StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode)
                svSND_HMS = StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode)
                svTOK_CD = (StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode) & StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode))
                svL_TOK_NAME = StrConv(Y_SYU_TEI_REC.L_TOK_NAME, vbUnicode)
                svDATA_MAKE_DATETIME = StrConv(Y_SYU_TEI_REC.DATA_MAKE_DATETIME, vbUnicode)
                svNou_YMD = StrConv(Y_SYU_TEI_REC.NOU_YMD, vbUnicode)
            
                TOTAL_CNT = 0
                ZAN_CNT = 0
                                                    
                                                    
            End If
    
            TOTAL_CNT = TOTAL_CNT + 1
            If Trim(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, vbUnicode)) = "" Then
                ZAN_CNT = ZAN_CNT + 1
            End If
        
        End If                          '2013.05.01

        com = BtOpGetNext

    Loop


    If Trim(svSND_YMD) <> "" Then

        Row = Row + 1
        Y_Syuka_TEI.ReDim Min_Row, Row, Min_Col, Max_Col
        
        
        '№
        Y_Syuka_TEI(Row, ColNo) = Row
        'ﾃﾞｰﾀ作成日付
        Y_Syuka_TEI(Row, ColSND_YMD) = Mid(svSND_YMD, 1, 4) & "/" & Mid(svSND_YMD, 5, 2) & "/" & Mid(svSND_YMD, 7, 2)
        'ﾃﾞｰﾀ作成時刻
        Y_Syuka_TEI(Row, ColSND_HMS) = Mid(svSND_HMS, 1, 2) & ":" & Mid(svSND_HMS, 3, 2) & ":" & Mid(svSND_HMS, 5, 2)
        '送り先
        Y_Syuka_TEI(Row, ColTOK_CD) = svTOK_CD
        '送り先名
        Y_Syuka_TEI(Row, ColL_TOK_NAME) = svL_TOK_NAME
        
        '納入日
        If Trim(svNou_YMD) <> "" Then
            Y_Syuka_TEI(Row, ColNOU_YMD) = Mid(svNou_YMD, 1, 4) & "/" & Mid(svNou_YMD, 5, 2) & "/" & Mid(svNou_YMD, 7, 2)
        Else
            Y_Syuka_TEI(Row, ColNOU_YMD) = ""
        End If
        
        '件数
        Y_Syuka_TEI(Row, ColTOTAL_CNT) = TOTAL_CNT
        '件数(梱包残)
        Y_Syuka_TEI(Row, ColZAN_CNT) = ZAN_CNT
        'ﾃﾞｰﾀ作成
        Y_Syuka_TEI(Row, ColDATA_MAKE_DATETIME) = svDATA_MAKE_DATETIME


                                    'DBテーブルリンク
        Set TDBGrid1.Array = Y_Syuka_TEI
        
        
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst
        TDBGrid1.ScrollBars = dbgAutomatic


    End If

    List_Disp_Proc = False

    Call Input_UnLock


End Function



Private Function List_Disp_NOU_Proc() As Integer
'2012.12.25

Dim svSND_YMD               As String
Dim svSND_HMS               As String
Dim svTOK_CD                As String
Dim svL_TOK_NAME            As String
Dim svDATA_MAKE_DATETIME    As String
Dim svNou_YMD               As String


Dim TOTAL_CNT               As Long
Dim ZAN_CNT                 As Long

Dim Row                     As Long
        
Dim com                     As Integer
Dim sts                     As Integer
Dim ans                     As Integer

Dim wkTOK_CD                As String * 16

Dim SKIP_F                  As Boolean

    List_Disp_NOU_Proc = True

    Call Input_Lock


                        'テーブルリセット
    Set Y_Syuka_TEI = Nothing
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2013.05.01
    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2013.05.01
    
    
    Row = Min_Row - 1
    
    svSND_YMD = ""
    
    Call UniCode_Conv(K5_Y_SYU_TEI.NOU_YMD, Format(Text1(ptxNOU_YMD).Text, "YYYYMMDD"))
    Call UniCode_Conv(K5_Y_SYU_TEI.SND_YMD, "")
    Call UniCode_Conv(K5_Y_SYU_TEI.SND_HMS, "")
    Call UniCode_Conv(K5_Y_SYU_TEI.TOK_CD, "")
    Call UniCode_Conv(K5_Y_SYU_TEI.CHO_CD, "")
        
        
        
        
'    If Trim(Text1(ptxTOK_CD).Text) <> "" Then
'        wkTOK_CD = Text1(ptxTOK_CD).Text                                   '2013.05.01
'        Call UniCode_Conv(K5_Y_SYU_TEI.TOK_CD, Left(wkTOK_CD, 8))          '2013.05.01
'        Call UniCode_Conv(K5_Y_SYU_TEI.CHO_CD, Right(wkTOK_CD, 8))         '2013.05.01

'        Call UniCode_Conv(K5_Y_SYU_TEI.TOK_CD, Text1(ptxTOK_CD1).Text)      '2013.05.01
'        Call UniCode_Conv(K5_Y_SYU_TEI.CHO_CD, Text1(ptxTOK_CD2).Text)      '2013.05.01


'    End If
        
    com = BtOpGetGreaterEqual

    Do
        DoEvents
        sts = BTRV(com, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K5_Y_SYU_TEI, Len(K5_Y_SYU_TEI), 5)
        Select Case sts
            Case BtNoErr
            
            
                If Format(Text1(ptxNOU_YMD).Text, "YYYYMMDD") <> StrConv(Y_SYU_TEI_REC.NOU_YMD, vbUnicode) Then
                    Exit Do
                End If
            
            
            
            
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Do
                End If
            
                com = BtOpGetEqual
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "邸別注文データ")
                Exit Do
        End Select


        SKIP_F = False
        If RTrim(Text1(ptxTOK_CD1).Text) <> "" Then
            '2013.05.01
'            If RTrim(Text1(ptxTOK_CD).Text) <> RTrim((StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode) & StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode))) Then
            If RTrim(Text1(ptxTOK_CD1).Text) <> RTrim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode)) Or _
                RTrim(Text1(ptxTOK_CD2).Text) <> RTrim(StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode)) Then
            '2013.05.01
                   
                SKIP_F = True
                
            End If
        End If
                
        If Not SKIP_F Then
                
            If Trim(svSND_YMD) = "" Then
                svSND_YMD = StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode)
                svSND_HMS = StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode)
                svTOK_CD = (StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode) & StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode))
                svL_TOK_NAME = StrConv(Y_SYU_TEI_REC.L_TOK_NAME, vbUnicode)
                svDATA_MAKE_DATETIME = StrConv(Y_SYU_TEI_REC.DATA_MAKE_DATETIME, vbUnicode)
                svNou_YMD = StrConv(Y_SYU_TEI_REC.NOU_YMD, vbUnicode)
            
                TOTAL_CNT = 0
                ZAN_CNT = 0
            End If
                        
                        
            If svSND_YMD <> StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode) Or _
                svSND_HMS <> StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode) Or _
                RTrim(svTOK_CD) <> RTrim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode) & StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode)) Then
                
                Row = Row + 1
                Y_Syuka_TEI.ReDim Min_Row, Row, Min_Col, Max_Col
                
                
                '№
                Y_Syuka_TEI(Row, ColNo) = Row
                'ﾃﾞｰﾀ作成日付
                Y_Syuka_TEI(Row, ColSND_YMD) = Mid(svSND_YMD, 1, 4) & "/" & Mid(svSND_YMD, 5, 2) & "/" & Mid(svSND_YMD, 7, 2)
                'ﾃﾞｰﾀ作成時刻
                Y_Syuka_TEI(Row, ColSND_HMS) = Mid(svSND_HMS, 1, 2) & ":" & Mid(svSND_HMS, 3, 2) & ":" & Mid(svSND_HMS, 5, 2)
                '送り先
                Y_Syuka_TEI(Row, ColTOK_CD) = svTOK_CD
                '送り先名
                Y_Syuka_TEI(Row, ColL_TOK_NAME) = svL_TOK_NAME
                
                '納入日
                If Trim(svNou_YMD) <> "" Then
                    Y_Syuka_TEI(Row, ColNOU_YMD) = Mid(svNou_YMD, 1, 4) & "/" & Mid(svNou_YMD, 5, 2) & "/" & Mid(svNou_YMD, 7, 2)
                Else
                    Y_Syuka_TEI(Row, ColNOU_YMD) = ""
                End If
                
                '件数
                Y_Syuka_TEI(Row, ColTOTAL_CNT) = TOTAL_CNT
                '件数(梱包残)
                Y_Syuka_TEI(Row, ColZAN_CNT) = ZAN_CNT
                'ﾃﾞｰﾀ作成
                Y_Syuka_TEI(Row, ColDATA_MAKE_DATETIME) = svDATA_MAKE_DATETIME
                                                    
                                                    
                svSND_YMD = StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode)
                svSND_HMS = StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode)
                svTOK_CD = (StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode) & StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode))
                svL_TOK_NAME = StrConv(Y_SYU_TEI_REC.L_TOK_NAME, vbUnicode)
                svDATA_MAKE_DATETIME = StrConv(Y_SYU_TEI_REC.DATA_MAKE_DATETIME, vbUnicode)
                svNou_YMD = StrConv(Y_SYU_TEI_REC.NOU_YMD, vbUnicode)
            
                TOTAL_CNT = 0
                ZAN_CNT = 0
                                                    
                                                    
            End If
                    
        
            TOTAL_CNT = TOTAL_CNT + 1
            If Trim(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, vbUnicode)) = "" Then
                ZAN_CNT = ZAN_CNT + 1
            End If
        End If
        com = BtOpGetNext

    Loop


    If Trim(svSND_YMD) <> "" Then

        Row = Row + 1
        Y_Syuka_TEI.ReDim Min_Row, Row, Min_Col, Max_Col
        
        
        '№
        Y_Syuka_TEI(Row, ColNo) = Row
        'ﾃﾞｰﾀ作成日付
        Y_Syuka_TEI(Row, ColSND_YMD) = Mid(svSND_YMD, 1, 4) & "/" & Mid(svSND_YMD, 5, 2) & "/" & Mid(svSND_YMD, 7, 2)
        'ﾃﾞｰﾀ作成時刻
        Y_Syuka_TEI(Row, ColSND_HMS) = Mid(svSND_HMS, 1, 2) & ":" & Mid(svSND_HMS, 3, 2) & ":" & Mid(svSND_HMS, 5, 2)
        '送り先
        Y_Syuka_TEI(Row, ColTOK_CD) = svTOK_CD
        '送り先名
        Y_Syuka_TEI(Row, ColL_TOK_NAME) = svL_TOK_NAME
        
        '納入日
        If Trim(svNou_YMD) <> "" Then
            Y_Syuka_TEI(Row, ColNOU_YMD) = Mid(svNou_YMD, 1, 4) & "/" & Mid(svNou_YMD, 5, 2) & "/" & Mid(svNou_YMD, 7, 2)
        Else
            Y_Syuka_TEI(Row, ColNOU_YMD) = ""
        End If
        
        '件数
        Y_Syuka_TEI(Row, ColTOTAL_CNT) = TOTAL_CNT
        '件数(梱包残)
        Y_Syuka_TEI(Row, ColZAN_CNT) = ZAN_CNT
        'ﾃﾞｰﾀ作成
        Y_Syuka_TEI(Row, ColDATA_MAKE_DATETIME) = svDATA_MAKE_DATETIME


                                    'DBテーブルリンク
        Set TDBGrid1.Array = Y_Syuka_TEI
        
        
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst
        TDBGrid1.ScrollBars = dbgAutomatic


    End If



    List_Disp_NOU_Proc = False

    Call Input_UnLock


End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    SEK10101.MousePointer = vbHourglass


    TDBGrid1.Enabled = False


    Call Ctrl_Lock(SEK10101)

    DoEvents
End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEK10101)

    TDBGrid1.Enabled = True

    SEK10101.MousePointer = vbDefault

    DoEvents
End Sub

Public Function Y_SYU_TEI_Create_Index() As Integer
'-------------------------------------------------------
'
'   『邸別注文ﾃﾞｰﾀのKEY追加』
'
'-------------------------------------------------------
Dim sts As Integer


    Y_SYU_TEI_Create_Index = True


    SEK1010_Y_SYU_TEI_Speck.ks0.keypos = 1                  ' キーポジション
    SEK1010_Y_SYU_TEI_Speck.ks0.keyleng = 8                 ' キー長
                                                            ' キーフラグ
    SEK1010_Y_SYU_TEI_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    SEK1010_Y_SYU_TEI_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    SEK1010_Y_SYU_TEI_Speck.ks0.reserve = &H0               ' 予約済み

    SEK1010_Y_SYU_TEI_Speck.ks1.keypos = 9                  ' キーポジション
    SEK1010_Y_SYU_TEI_Speck.ks1.keyleng = 6                 ' キー長
                                                    ' キーフラグ
    SEK1010_Y_SYU_TEI_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    SEK1010_Y_SYU_TEI_Speck.ks1.keytype = Chr(BtKtString)   ' キータイプ
    SEK1010_Y_SYU_TEI_Speck.ks1.reserve = &H0               ' 予約済み

    SEK1010_Y_SYU_TEI_Speck.ks2.keypos = 52                 ' キーポジション
    SEK1010_Y_SYU_TEI_Speck.ks2.keyleng = 8                 ' キー長
                                                            ' キーフラグ
    SEK1010_Y_SYU_TEI_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    SEK1010_Y_SYU_TEI_Speck.ks2.keytype = Chr(BtKtString)   ' キータイプ
    SEK1010_Y_SYU_TEI_Speck.ks2.reserve = &H0               ' 予約済み

    SEK1010_Y_SYU_TEI_Speck.ks3.keypos = 60                 ' キーポジション
    SEK1010_Y_SYU_TEI_Speck.ks3.keyleng = 8                 ' キー長
                                                            ' キーフラグ
    SEK1010_Y_SYU_TEI_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    SEK1010_Y_SYU_TEI_Speck.ks3.keytype = Chr(BtKtString)   ' キータイプ
    SEK1010_Y_SYU_TEI_Speck.ks3.reserve = &H0               ' 予約済み

    sts = BTRV(BtOpCreatSupIndex, Y_SYU_TEI_POS, SEK1010_Y_SYU_TEI_Speck, Len(SEK1010_Y_SYU_TEI_Speck), K4_Y_SYU_TEI, Len(K4_Y_SYU_TEI), 4)
    If sts Then
        Call File_Error(sts, BtOpCreatSupIndex, "邸別注文ﾃﾞｰﾀ")
        Exit Function
    End If


'>>>>>>>>>> 2012.12.25
    SEK1010_Y_SYU_TEI_Speck2.ks4.keypos = 174                ' キーポジション
    SEK1010_Y_SYU_TEI_Speck2.ks4.keyleng = 8                 ' キー長
                                                            ' キーフラグ
    SEK1010_Y_SYU_TEI_Speck2.ks4.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    SEK1010_Y_SYU_TEI_Speck2.ks4.keytype = Chr(BtKtString)   ' キータイプ
    SEK1010_Y_SYU_TEI_Speck2.ks4.reserve = &H0               ' 予約済み

    SEK1010_Y_SYU_TEI_Speck2.ks5.keypos = 1                  ' キーポジション
    SEK1010_Y_SYU_TEI_Speck2.ks5.keyleng = 8                 ' キー長
                                                            ' キーフラグ
    SEK1010_Y_SYU_TEI_Speck2.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    SEK1010_Y_SYU_TEI_Speck2.ks5.keytype = Chr(BtKtString)   ' キータイプ
    SEK1010_Y_SYU_TEI_Speck2.ks5.reserve = &H0               ' 予約済み

    SEK1010_Y_SYU_TEI_Speck2.ks6.keypos = 9                  ' キーポジション
    SEK1010_Y_SYU_TEI_Speck2.ks6.keyleng = 6                 ' キー長
                                                    ' キーフラグ
    SEK1010_Y_SYU_TEI_Speck2.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    SEK1010_Y_SYU_TEI_Speck2.ks6.keytype = Chr(BtKtString)   ' キータイプ
    SEK1010_Y_SYU_TEI_Speck2.ks6.reserve = &H0               ' 予約済み

    SEK1010_Y_SYU_TEI_Speck2.ks7.keypos = 52                 ' キーポジション
    SEK1010_Y_SYU_TEI_Speck2.ks7.keyleng = 8                 ' キー長
                                                            ' キーフラグ
    SEK1010_Y_SYU_TEI_Speck2.ks7.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    SEK1010_Y_SYU_TEI_Speck2.ks7.keytype = Chr(BtKtString)   ' キータイプ
    SEK1010_Y_SYU_TEI_Speck2.ks7.reserve = &H0               ' 予約済み

    SEK1010_Y_SYU_TEI_Speck2.ks8.keypos = 60                 ' キーポジション
    SEK1010_Y_SYU_TEI_Speck2.ks8.keyleng = 8                 ' キー長
                                                            ' キーフラグ
    SEK1010_Y_SYU_TEI_Speck2.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg
    SEK1010_Y_SYU_TEI_Speck2.ks8.keytype = Chr(BtKtString)   ' キータイプ
    SEK1010_Y_SYU_TEI_Speck2.ks8.reserve = &H0               ' 予約済み

    sts = BTRV(BtOpCreatSupIndex, Y_SYU_TEI_POS, SEK1010_Y_SYU_TEI_Speck2, Len(SEK1010_Y_SYU_TEI_Speck2), K5_Y_SYU_TEI, Len(K5_Y_SYU_TEI), 5)
    If sts Then
        Call File_Error(sts, BtOpCreatSupIndex, "邸別注文ﾃﾞｰﾀ")
        Exit Function
    End If


'>>>>>>>>>> 2012.12.25





    Y_SYU_TEI_Create_Index = False

End Function


'Public Function DEL_SYU_H_Create_Index() As Integer
''-------------------------------------------------------
''
''   『出荷予定ﾃﾞｰﾀ(H)のKEY追加』
''
''-------------------------------------------------------
'Dim sts As Integer
'
'
'    DEL_SYU_H_Create_Index = True
'
'
'    SEK1010_DEL_SYU_H_Speck.ks0.keypos = 481                ' キーポジション
'    SEK1010_DEL_SYU_H_Speck.ks0.keyleng = 9                 ' キー長
'                                                            ' キーフラグ
'    SEK1010_DEL_SYU_H_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfChg
'    SEK1010_DEL_SYU_H_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
'    SEK1010_DEL_SYU_H_Speck.ks0.reserve = &H0               ' 予約済み
'
'
'
'
'    sts = BTRV(BtOpCreatSupIndex, DEL_SYU_H_POS, SEK1010_DEL_SYU_H_Speck, Len(SEK1010_DEL_SYU_H_Speck), K5_DEL_SYU_H, Len(K5_DEL_SYU_H), 5)
'    If sts Then
'        Call File_Error(sts, BtOpCreatSupIndex, "出荷予定ﾃﾞｰﾀ(H)")
'        Exit Function
'    End If
'
'    DEL_SYU_H_Create_Index = False
'
'End Function



Private Sub INI_Disp_Proc()

Dim c   As String * 128

    '前回作成時間
    If GetIni(App.EXEName, "BEF_DATETIME", App.EXEName, c) Then
        lblBEF_DATETIME = ""
    Else
        lblBEF_DATETIME = RTrim(c)
    End If
    '前回伝票№
    If GetIni(App.EXEName, "BEF_START_DENNO", App.EXEName, c) Then
        lblBEF_START_DENNO = ""
    Else
        lblBEF_START_DENNO = RTrim(c)
    End If
    '今回伝票№
    If GetIni(App.EXEName, "START_DENNO", App.EXEName, c) Then
        lblSTART_DENNO = ""
    Else
        lblSTART_DENNO = RTrim(c)
    End If
    



    'EXEﾌｫﾙﾀﾞ
    If GetIni(App.EXEName, "EXE", App.EXEName, c) Then
        lblEXE = ""
    Else
        lblEXE = RTrim(c)
    End If


End Sub
