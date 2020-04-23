VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PR000501 
   Caption         =   "è§ïiâªé¿ê—èWåvï\"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15150
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
   ScaleHeight     =   10305
   ScaleWidth      =   15150
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   8880
      MaxLength       =   10
      TabIndex        =   76
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2160
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   6135
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   10821
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "édå¸ÇØêÊ"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "∏◊Ω"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Åyì‡ïîÅzåèêî"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Åyì‡ïîÅzêîó "
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "ÅyäOïîÅzåèêî"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "ÅyäOïîÅzêîó "
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "ÅyçáåvÅzåèêî"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "ÅyçáåvÅzêîó "
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "ÅyçáåvÅzíPâø"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "ÅyçáåvÅzã‡äz"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "ÅyéëçﬁÅzíPâø"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "ÅyéëçﬁÅzã‡äz"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "ÅyçHóøÅzíPâø"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "ÅyçHóøÅzã‡äz"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "ÅyëºÅzíPâø"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "ÅyëºÅzã‡äz"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Åyå¥âøÅzå¬ëï"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Åyå¥âøÅzäOëï"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Åyå¥âøÅzçHóø"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   19
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=19"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1588"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1482"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2381"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2275"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2381"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2381"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2275"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2381"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2275"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2381"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2275"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2381"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2275"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2381"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2275"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2381"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2275"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=2381"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=2275"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=2381"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=2275"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(61)=   "Column(12).Width=2381"
      Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=2275"
      Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(66)=   "Column(13).Width=2381"
      Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=2275"
      Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=2"
      Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(71)=   "Column(14).Width=2381"
      Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=2275"
      Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(76)=   "Column(15).Width=2381"
      Splits(0)._ColumnProps(77)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(15)._WidthInPix=2275"
      Splits(0)._ColumnProps(79)=   "Column(15)._ColStyle=2"
      Splits(0)._ColumnProps(80)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(81)=   "Column(16).Width=2381"
      Splits(0)._ColumnProps(82)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(16)._WidthInPix=2275"
      Splits(0)._ColumnProps(84)=   "Column(16)._ColStyle=2"
      Splits(0)._ColumnProps(85)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(86)=   "Column(17).Width=2381"
      Splits(0)._ColumnProps(87)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(88)=   "Column(17)._WidthInPix=2275"
      Splits(0)._ColumnProps(89)=   "Column(17)._ColStyle=2"
      Splits(0)._ColumnProps(90)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(91)=   "Column(18).Width=2381"
      Splits(0)._ColumnProps(92)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(18)._WidthInPix=2275"
      Splits(0)._ColumnProps(94)=   "Column(18)._ColStyle=2"
      Splits(0)._ColumnProps(95)=   "Column(18).Order=19"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ÇlÇr ÉSÉVÉbÉN"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ÇlÇr ÉSÉVÉbÉN"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "ê∂éYèWåvñæç◊"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
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
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=110,.parent=43"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=44"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=45"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=47"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=16,.parent=43,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=44"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=45"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=47"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=28,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(53)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(54)  =   ":id=28,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=44"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=45"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=47"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(59)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(60)  =   ":id=66,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=32,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(65)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(66)  =   ":id=32,.fontname=ÇlÇr ÉSÉVÉbÉN"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
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
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=62,.parent=43,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=47"
      _StyleDefs(90)  =   "Splits(0).Columns(11).Style:id=70,.parent=43,.alignment=1"
      _StyleDefs(91)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=44"
      _StyleDefs(92)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=45"
      _StyleDefs(93)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=47"
      _StyleDefs(94)  =   "Splits(0).Columns(12).Style:id=82,.parent=43,.alignment=1"
      _StyleDefs(95)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=44"
      _StyleDefs(96)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=45"
      _StyleDefs(97)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=47"
      _StyleDefs(98)  =   "Splits(0).Columns(13).Style:id=86,.parent=43,.alignment=1"
      _StyleDefs(99)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=44"
      _StyleDefs(100) =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=45"
      _StyleDefs(101) =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=47"
      _StyleDefs(102) =   "Splits(0).Columns(14).Style:id=90,.parent=43,.alignment=1"
      _StyleDefs(103) =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=44"
      _StyleDefs(104) =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=45"
      _StyleDefs(105) =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=47"
      _StyleDefs(106) =   "Splits(0).Columns(15).Style:id=94,.parent=43,.alignment=1"
      _StyleDefs(107) =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=44"
      _StyleDefs(108) =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=45"
      _StyleDefs(109) =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=47"
      _StyleDefs(110) =   "Splits(0).Columns(16).Style:id=98,.parent=43,.alignment=1"
      _StyleDefs(111) =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=44"
      _StyleDefs(112) =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=45"
      _StyleDefs(113) =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=47"
      _StyleDefs(114) =   "Splits(0).Columns(17).Style:id=102,.parent=43,.alignment=1"
      _StyleDefs(115) =   "Splits(0).Columns(17).HeadingStyle:id=99,.parent=44"
      _StyleDefs(116) =   "Splits(0).Columns(17).FooterStyle:id=100,.parent=45"
      _StyleDefs(117) =   "Splits(0).Columns(17).EditorStyle:id=101,.parent=47"
      _StyleDefs(118) =   "Splits(0).Columns(18).Style:id=106,.parent=43,.alignment=1"
      _StyleDefs(119) =   "Splits(0).Columns(18).HeadingStyle:id=103,.parent=44"
      _StyleDefs(120) =   "Splits(0).Columns(18).FooterStyle:id=104,.parent=45"
      _StyleDefs(121) =   "Splits(0).Columns(18).EditorStyle:id=105,.parent=47"
      _StyleDefs(122) =   "Named:id=33:Normal"
      _StyleDefs(123) =   ":id=33,.parent=0"
      _StyleDefs(124) =   "Named:id=34:Heading"
      _StyleDefs(125) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(126) =   ":id=34,.wraptext=-1"
      _StyleDefs(127) =   "Named:id=35:Footing"
      _StyleDefs(128) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(129) =   "Named:id=36:Selected"
      _StyleDefs(130) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(131) =   "Named:id=37:Caption"
      _StyleDefs(132) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(133) =   "Named:id=38:HighlightRow"
      _StyleDefs(134) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(135) =   "Named:id=39:EvenRow"
      _StyleDefs(136) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(137) =   "Named:id=40:OddRow"
      _StyleDefs(138) =   ":id=40,.parent=33"
      _StyleDefs(139) =   "Named:id=41:RecordSelector"
      _StyleDefs(140) =   ":id=41,.parent=34"
      _StyleDefs(141) =   "Named:id=42:FilterBar"
      _StyleDefs(142) =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "èI óπ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10440
      TabIndex        =   15
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9600
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8760
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "àÛ ç¸"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7920
      TabIndex        =   12
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6600
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "åü çı"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   8
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'âEëµÇ¶
      Caption         =   "Å`"
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   75
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'âEëµÇ¶
      Caption         =   "édå¸ÇØêÊ"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   74
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   11640
      TabIndex        =   73
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   11640
      TabIndex        =   72
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   11640
      TabIndex        =   71
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   11640
      TabIndex        =   70
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "çáåvÅ@Å@Å@á@Å{áA"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   9120
      TabIndex        =   69
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "äOíççHóøÅ@Å@Å@áA"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   9120
      TabIndex        =   68
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "á@åv"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   10800
      TabIndex        =   67
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "äOëï"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   10800
      TabIndex        =   66
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   19
      Left            =   9120
      TabIndex        =   65
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "å¬ëï"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   10800
      TabIndex        =   64
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   11640
      TabIndex        =   63
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "édì¸å¥âø"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   9120
      TabIndex        =   62
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "Å@Å@Å@Å@Å@Å@Å@è¡îÔéëçﬁ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   9120
      TabIndex        =   61
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   7440
      TabIndex        =   60
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   7440
      TabIndex        =   59
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   7440
      TabIndex        =   58
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   7440
      TabIndex        =   57
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   7440
      TabIndex        =   56
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   7440
      TabIndex        =   55
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "âøäiç\ê¨"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   7440
      TabIndex        =   54
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   5760
      TabIndex        =   53
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   5760
      TabIndex        =   52
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   5760
      TabIndex        =   51
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   5760
      TabIndex        =   50
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   5760
      TabIndex        =   49
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   5760
      TabIndex        =   48
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   5760
      TabIndex        =   47
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   5760
      TabIndex        =   46
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   4080
      TabIndex        =   45
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   4080
      TabIndex        =   44
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   4080
      TabIndex        =   43
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4080
      TabIndex        =   42
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   41
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   4080
      TabIndex        =   40
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   4080
      TabIndex        =   39
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   38
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   37
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   36
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   35
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   34
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   33
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   32
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   31
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H80000005&
      BorderStyle     =   1  'é¿ê¸
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   30
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "çáÅ@Å@åv"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   5760
      TabIndex        =   29
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "ÇaäOïîê∂éY"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   28
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "Ç`ì‡ïîê∂éY"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   27
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "áBÇªÇÃëº"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   26
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "áAçHóøÇÃïî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   25
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "á@éëçﬁÇÃïî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1080
      TabIndex        =   24
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "ì‡Å@ñÛ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   720
      TabIndex        =   23
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'âEëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "Åiç\ê¨î‰ó¶Åj"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   22
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "ê∂éYã‡äz"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   21
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'âEëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "Åiç\ê¨î‰ó¶Åj"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   20
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "ê∂éYêîó "
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   19
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "ê∂éYåèêî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   18
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "çÄñ⁄/ê∂éYì‡ñÛ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'âEëµÇ¶
      Caption         =   "ëŒè€îNåéì˙"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   16
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "PR000501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'ÉeÉLÉXÉgópìYéö
Private Const ptxSHIMUKE_CODE% = 0          'édå¸ÇØêÊ
Private Const ptxS_YMD% = 1                 'äJénÅ@ëŒè€îNåéì˙
Private Const ptxE_YMD% = 2                 'èIóπÅ@ëŒè€îNåéì˙
'ÉRÉìÉ{ópìYéö
Private Const pcmbSHIMUKE_CODE% = 0         'édå¸ÇØêÊ



'ï\é¶ópÉâÉxÉã
Private Const plblNAI_CNT% = 0              'ì‡ïîê∂éYÅ@ê∂éYåèêî
Private Const plblNAI_SURYO% = 1            'ì‡ïîê∂éYÅ@ê∂éYêîó 
Private Const plblNAI_SU_RITU% = 2          'ì‡ïîê∂éYÅ@ê∂éYêîó ç\ê¨ó¶
Private Const plblNAI_KIN% = 3              'ì‡ïîê∂éYÅ@ê∂éYã‡äz
Private Const plblNAI_KIN_RITU% = 4         'ì‡ïîê∂éYÅ@ê∂éYã‡äzç\ê¨ó¶

Private Const plblNAI_UCHI_SHIZAI% = 5      'ì‡ïîê∂éY  ì‡ñÛÅ@éëçﬁ
Private Const plblNAI_UCHI_KOURYO% = 6      'ì‡ïîê∂éY  ì‡ñÛÅ@çHóø
Private Const plblNAI_UCHI_ETC% = 7         'ì‡ïîê∂éY  ì‡ñÛÅ@ÇªÇÃëº

Private Const plblGAI_CNT% = 8              'äOïîê∂éYÅ@ê∂éYåèêî
Private Const plblGAI_SURYO% = 9            'äOïîê∂éYÅ@ê∂éYêîó 
Private Const plblGAI_SU_RITU% = 10         'äOïîê∂éYÅ@ê∂éYêîó ç\ê¨ó¶
Private Const plblGAI_KIN% = 11             'äOïîê∂éYÅ@ê∂éYã‡äz
Private Const plblGAI_KIN_RITU% = 12        'äOïîê∂éYÅ@ê∂éYã‡äzç\ê¨ó¶

Private Const plblGAI_UCHI_SHIZAI% = 13     'äOïîê∂éY  ì‡ñÛÅ@éëçﬁ
Private Const plblGAI_UCHI_KOURYO% = 14     'äOïîê∂éY  ì‡ñÛÅ@çHóø
Private Const plblGAI_UCHI_ETC% = 15        'äOïîê∂éY  ì‡ñÛÅ@ÇªÇÃëº

Private Const plblGK_CNT% = 16              'çáåvÅ@ê∂éYåèêî
Private Const plblGK_SURYO% = 17            'çáåvÅ@ê∂éYêîó 
Private Const plblGK_SU_RITU% = 18          'çáåvÅ@ê∂éYêîó ç\ê¨ó¶
Private Const plblGK_KIN% = 19              'çáåvÅ@ê∂éYã‡äz
Private Const plblGK_KIN_RITU% = 20         'çáåvÅ@ê∂éYã‡äzç\ê¨ó¶

Private Const plblGK_UCHI_SHIZAI% = 21      'çáåv  ì‡ñÛÅ@éëçﬁ
Private Const plblGK_UCHI_KOURYO% = 22      'çáåv  ì‡ñÛÅ@çHóø
Private Const plblGK_UCHI_ETC% = 23         'çáåv  ì‡ñÛÅ@ÇªÇÃëº


Private Const plblKAKAKU_RITU% = 24         'âøäiç\ê¨Å@ê∂éYã‡äz
Private Const plblSHIZAI_RITU% = 25         'âøäiç\ê¨Å@éëçﬁ
Private Const plblKOURYO_RITU% = 26         'âøäiç\ê¨Å@çHóø
Private Const plblETC_RITU% = 27            'âøäiç\ê¨Å@ÇªÇÃëº

Private Const plblGENKA_KOSOU% = 28         'édì¸å¥âøÅ@å¬ëï
Private Const plblGENKA_GAISOU% = 29        'édì¸å¥âøÅ@äOëï
Private Const plblGENKA_SHIZAI% = 30        'édì¸å¥âøÅ@è¡îÔéëçﬁåv
Private Const plblGENKA_KOURYO% = 31        'édì¸å¥âøÅ@çHóø
Private Const plblGENKA_GK% = 32            'édì¸å¥âøÅ@çáåv





'Glidópä¬ã´---------------------------------

'édì¸ñæç◊
Private Const pGridDETAIL% = 0


Private SEISAN      As New XArrayDB


Private Const Min_Row% = 1                  'ç≈è¨çsêî
Private Const Min_Col% = 0                  'ç≈è¨óÒêî
Private Const Max_Col% = 18                 'ç≈ëÂóÒêî

Private Const colSHIMUKE_CODE% = 0          'édå¸ÇØêÊ
Private Const colCLASS_CODE% = 1            '∏◊Ω∫∞ƒﬁ

Private Const colNAI_CNT% = 2               'ì‡ïîÅ@åèêî
Private Const colNAI_SURYO% = 3             'ì‡ïîÅ@êîó 

Private Const colGAI_CNT% = 4               'äOïîÅ@åèêî
Private Const colGAI_SURYO% = 5             'äOïîÅ@êîó 

Private Const colGK_CNT% = 6                'çáåvÅ@åèêî
Private Const colGK_SURYO% = 7              'çáåvÅ@êîó 
Private Const colGK_TANKA% = 8              'çáåvÅ@íPâø
Private Const colGK_KIN% = 9                'çáåvÅ@ã‡äz

Private Const colSHIZAI_TANKA% = 10         'éëçﬁÅ@íPâø
Private Const colSHIZAI_KIN% = 11           'éëçﬁÅ@ã‡äz
Private Const colKOURYO_TANKA% = 12         'çHóøÅ@íPâø
Private Const colKOURYO_KIN% = 13           'çHóøÅ@ã‡äz
Private Const colETC_TANKA% = 14            'ÇªÇÃëºÅ@íPâø
Private Const colETC_KIN% = 15              'ÇªÇÃëºÅ@ã‡äz

Private Const colGENKA_KOSOU% = 16          'édì¸å¥âøÅ@å¬ëï
Private Const colGENKA_GAISOU% = 17         'édì¸å¥âøÅ@äOëï
Private Const colGENKA_KOURYO% = 18         'édì¸å¥âøÅ@çHóø




Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ø∞ƒÇÃêßå‰ 0:è∏èá 1:ç~èá
Private Tbl_Set_F   As Boolean


Private Type Sum_Area_tag
    CNT                     As Integer      'åèêî
    SURYO                   As Double       'êîó  9(8)V99
                                               
    KINGAKU                 As Long         'è§ïiâªã‡äz     9(10)
    
    SH_KINGAKU              As Long         'éëçﬁã‡äz       9(10)
    
    KO_KINGAKU              As Long         'çHóøã‡äz       9(10)
    
    ETC_KINGAKU             As Long         'ÇªÇÃëºã‡äz     9(10)
End Type

'Private Const LAST_UPDATE_DAY$ = "[PR00050] 2018.10.30 09:30"
Private Const LAST_UPDATE_DAY$ = "[PR00050] 2018.10.30 10:30"


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   âÊñ çÄñ⁄ÉçÉbÉNÅiÉCÉxÉìÉgéÊìæïsâ¬Åj
'----------------------------------------------------------------------------

    PR000501.MousePointer = vbHourglass

    Call Ctrl_Lock(PR000501)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   âÊñ çÄñ⁄ÉçÉbÉNâèúÅiÉCÉxÉìÉgéÊìæâ¬Åj
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PR000501)


    PR000501.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ì¸óÕçÄñ⁄ÇÃÉGÉâÅ[É`ÉFÉbÉN
'----------------------------------------------------------------------------
    
Dim sts     As Integer
Dim com     As Integer
    
Dim i       As Integer
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        Case ptxSHIMUKE_CODE    'édå¸ÇØêÊ∫∞ƒﬁ
        
           
           Combo1(pcmbSHIMUKE_CODE).ListIndex = -1
           For i = 0 To Combo1(pcmbSHIMUKE_CODE).ListCount - 1
               If Text1(ptxSHIMUKE_CODE).Text = Left(Right(Combo1(pcmbSHIMUKE_CODE).List(i), 4), 2) Then
                   Combo1(pcmbSHIMUKE_CODE).ListIndex = i
                   Exit For
               End If
           
           Next i
        
        
        
        Case ptxS_YMD           'ëŒè€îNåéì˙
        
            
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0000/01/01"
            End If
            
            If Not IsDate(Text1(Mode).Text) Then
                MsgBox "ì¸óÕÇµÇΩçÄñ⁄ÇÕÉGÉâÅ[Ç≈Ç∑ÅB"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")
            
            End If
        
        Case ptxE_YMD           'ëŒè€îNåéì˙
        
            
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "9999/12/31"
            End If
            
            If Not IsDate(Text1(Mode).Text) Then
                MsgBox "ì¸óÕÇµÇΩçÄñ⁄ÇÕÉGÉâÅ[Ç≈Ç∑ÅB"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")
            
            End If
        
        
    End Select
        
        
    Error_Check_Proc = False
    

End Function


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    Select Case Index
        Case pcmbSHIMUKE_CODE       'édå¸ÇØêÊ∫∞ƒﬁ
        
            Text1(ptxSHIMUKE_CODE).Text = Trim(Left(Right(Combo1(pcmbSHIMUKE_CODE).Text, 4), 2))
    End Select
    
    
    Call Tab_Ctrl(Shift)        'à⁄ìÆ

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    Select Case Index
        Case pcmbSHIMUKE_CODE       'édå¸ÇØêÊ∫∞ƒﬁ
        
            Text1(ptxSHIMUKE_CODE).Text = Trim(Left(Right(Combo1(pcmbSHIMUKE_CODE).Text, 4), 2))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim Data_Flg    As Boolean

Dim rpt             As New PR00050F1
Dim f               As New PR000502


Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd          'çXêV
        
        Case P_CMD_DEL          'çÌèú
        
        Case P_CMD_DSP                      'åüçı/ï\é¶
        
            For i = ptxS_YMD To ptxE_YMD
            
                If Error_Check_Proc(i) Then     'ÉGÉâÅ[É`ÉFÉbÉN
                    Exit Sub
                End If
            
            Next i
            
            If SUM_Make_Proc(Data_Flg) Then
                Exit Sub
            End If
            
            
            If Not Data_Flg Then
                MsgBox "ëŒè€√ﬁ∞¿Ç™Ç†ÇËÇ‹ÇπÇÒ"
                Exit Sub
            End If
        
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
            Text1(ptxS_YMD).SetFocus
        
        
        Case P_CMD_OUT                      '√ﬁ∞¿èoóÕ
        
        Case P_CMD_PRT                      'àÛç¸
 
            For i = ptxS_YMD To ptxE_YMD
                                            'ÉGÉâÅ[É`ÉFÉbÉN
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            Next i
            
            If SUM_Make_Proc(Data_Flg) Then
                Exit Sub
            End If
            
            
            If Not Data_Flg Then
                MsgBox "ëŒè€√ﬁ∞¿Ç™Ç†ÇËÇ‹ÇπÇÒ"
                Exit Sub
            End If
                
            ans = MsgBox("àÛç¸ÇµÇ‹Ç∑Ç©ÅH", vbYesNo + vbQuestion, "ämîFì¸óÕ")
            If ans = vbYes Then
                
                Set rpt = New PR00050F1
            
                'ÉåÉ|Å[ÉgÇàÛç¸ÇµÇ‹Ç∑ÅBÅitrueÅFàÛç¸É_ÉCÉAÉçÉOÇ†ÇË falseÅFÇ»ÇµÅj
                rpt.PrintReport False
            
                Set rpt = Nothing
                
                
'                f.RunReport rpt
'                f.Show
            
            End If
            
            Text1(ptxS_YMD).SetFocus
            
            
        Case P_CMD_End                      'èIóπ
    
            Unload Me
    
    End Select

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
                                'ïiñ⁄É}ÉXÉ^ÇnÇoÇdÇm
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'ÉNÉâÉXÉ}ÉXÉ^ÇnÇoÇdÇm
    If P_Class_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'éÛï•êÊÉ}ÉXÉ^ÇnÇoÇdÇm
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'ÉRÅ[ÉhÉ}ÉXÉ^ÇnÇoÇdÇm
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'ä«óùÉ}ÉXÉ^ÇnÇoÇdÇm
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'ê∂éYé¿ê—èWåv√ﬁ∞¿ÇnÇoÇdÇm
    If P_SEISAN_SUM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'è§ïiâªéwé¶(êe)√ﬁ∞¿ÇnÇoÇdÇm
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'è§ïiâªéwé¶(éq)√ﬁ∞¿ÇnÇoÇdÇm
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'è§ïiâªéwé¶éÛì¸óöó√ﬁ∞¿ÇnÇoÇdÇm
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    PR000501.Caption = PR000501.Caption & LAST_UPDATE_DAY
    
    Load PR000502
    
    
    
    'ä«óùÉ}ÉXÉ^ÇÃì«Ç›çûÇ›
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)

    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            If P_KANRI_MAKE_Proc() Then
                Unload Me
            End If
        Case Else
            Call File_Error(sts, BtOpGetEqual, "ä«óùÉ}ÉXÉ^")
            Unload Me
    End Select
    
    '∫∞ƒﬁœΩ¿íËã`
    Call P_CODE_TBL_Proc
    
    'édå¸ÇØêÊÇÃÉZÉbÉg
    If Code_Set_Proc(pcmbSHIMUKE_CODE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    
    
    'âÊñ èâä˙ê›íË
    If Init_Proc() Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            'ÉNÉâÉXÉ}ÉXÉ^ÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ÉNÉâÉXÉ}ÉXÉ^")
        End If
    End If
                                            'éÛï•êÊÉ}ÉXÉ^ÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "éÛï•êÊÉ}ÉXÉ^")
        End If
    End If
                                            'ïiñ⁄É}ÉXÉ^ÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ïiñ⁄É}ÉXÉ^")
        End If
    End If
    
                                            'ä«óùÉ}ÉXÉ^ÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ÉNÉâÉXÉ}ÉXÉ^")
        End If
    End If
                                            'ê∂éYé¿ê—èWåv√ﬁ∞¿ÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
        End If
    End If
                                            'è§ïiâªéwé¶ÅiêeÅj√ﬁ∞¿ÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "è§ïiâªéwé¶ÅiêeÅj√ﬁ∞¿")
        End If
    End If
                                            'è§ïiâªéwé¶ÅiéqÅj√ﬁ∞¿ÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "è§ïiâªéwé¶ÅiêeÅj√ﬁ∞¿")
        End If
    End If
                                            'è§ïiâªéÛì¸óöó√ﬁ∞¿ÇbÇkÇnÇrÇd
    sts = BTRV(BtOpClose, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "è§ïiâªéÛì¸óöó√ﬁ∞¿")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PR000501 = Nothing
    Set PR000502 = Nothing


    End
End Sub





Private Sub TDBGrid1_HeadClick(Index As Integer, ByVal ColIndex As Integer)



    Select Case Index
        
        Case pGridDETAIL        'ê∂éYé¿ê—ñæç◊
    
    
            If Sort_Tbl(ColIndex) = 0 Then
                Sort_Tbl(ColIndex) = 1
            Else
                If Sort_Tbl(ColIndex) = 1 Then
                    Sort_Tbl(ColIndex) = 0
                End If
            
            End If
            
            If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                            
                SEISAN.QuickSort Min_Row, SEISAN.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
                
                Set TDBGrid1(Index).Array = SEISAN
                
                TDBGrid1(Index).ReBind
                TDBGrid1(Index).Update
                TDBGrid1(Index).MoveFirst
        
        
            End If
    
    
    
    End Select




End Sub


Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then    'ÉGÉâÅ[É`ÉFÉbÉN
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        'à⁄ìÆ
End Sub
Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   ì¸óÕâÊñ ÇÃèâä˙ê›íË
'----------------------------------------------------------------------------
Dim i       As Integer
Dim sts     As Integer


    Init_Proc = True
    
    
    
    For i = ptxS_YMD To ptxE_YMD
        Text1(i).Text = ""
    Next i
    'èàóùîNåéì˙ÅÅìñì˙
    Text1(ptxS_YMD).Text = Format(Now, "YYYY/MM/DD")
    Text1(ptxE_YMD).Text = Format(Now, "YYYY/MM/DD")
    
    For i = pcmbSHIMUKE_CODE To pcmbSHIMUKE_CODE
        
        Combo1(i).ListIndex = -1
    
    Next i



    
    'ø∞ƒèÓïÒÇÃèâä˙âª
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0               '√ﬁÃ´Ÿƒè∏èá
    Next i

    Init_Proc = False

End Function



Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           éëçﬁéÛì¸ÉfÅ[É^ÇÃï\é¶
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Row             As Long


Dim wkValue         As Double
Dim i               As Integer





    List_Disp_Proc = True
    PR000501.MousePointer = vbHourglass
    
    
    Call UniCode_Conv(K0_P_SEISAN_SUM.SHIMUKE_CODE, Text1(ptxSHIMUKE_CODE))
    Call UniCode_Conv(K0_P_SEISAN_SUM.CLASS_CODE, P_ClassSum_Key)
    
    sts = BTRV(BtOpGetEqual, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
        
            MsgBox "ëŒè€√ﬁ∞¿Ç™Ç†ÇËÇ‹ÇπÇÒ"
            List_Disp_Proc = False
            Exit Function
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
            Exit Function
    End Select
    
    
    
        
                                            'ì‡ïîê∂éYÅ@ê∂éYåèêî
    lblItem(plblNAI_CNT).Caption = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).CNT, vbUnicode)), "#,##0")
                                            'äOïîê∂éYÅ@ê∂éYåèêî
    lblItem(plblGAI_CNT).Caption = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).CNT, vbUnicode)), "#,##0")
                                            'çáåvÅ@ Å@ ê∂éYåèêî
    lblItem(plblGK_CNT).Caption = Format(CInt(lblItem(plblNAI_CNT).Caption) + CInt(lblItem(plblGAI_CNT).Caption), "#,##0")
                                            
                                            'ì‡ïîê∂éYÅ@ê∂éYêîó 
    lblItem(plblNAI_SURYO).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SURYO, vbUnicode)), "#,##0")
                                            'äOïîê∂éYÅ@ê∂éYêîó 
    lblItem(plblGAI_SURYO).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SURYO, vbUnicode)), "#,##0")
                                            'çáåvÅ@ Å@ ê∂éYêîó 
    lblItem(plblGK_SURYO).Caption = Format(CDbl(lblItem(plblNAI_SURYO).Caption) + CDbl(lblItem(plblGAI_SURYO).Caption), "#,##0")
                                            
                                            'ì‡ïîê∂éY  ç\ê¨î‰ó¶
    wkValue = CDbl(lblItem(plblNAI_SURYO).Caption) / (CDbl(lblItem(plblNAI_SURYO).Caption) + CDbl(lblItem(plblGAI_SURYO).Caption)) * 100
    lblItem(plblNAI_SU_RITU).Caption = Format(wkValue, "#0.00") & "%"
                                            
                                            'äOïîê∂éY  ç\ê¨î‰ó¶
    lblItem(plblGAI_SU_RITU).Caption = Format(100 - wkValue, "#0.00") & "%"
                                            'ç\ê¨  ç\ê¨î‰ó¶
    lblItem(plblGK_SU_RITU).Caption = "100.00%"
    
                                            'ì‡ïîê∂éYÅ@ê∂éYã‡äz
    lblItem(plblNAI_KIN).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KINGAKU, vbUnicode)), "#,##0")
                                            'äOïîê∂éYÅ@ê∂éYã‡äz
    lblItem(plblGAI_KIN).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KINGAKU, vbUnicode)), "#,##0")
                                            'çáåvÅ@ê∂éYã‡äz
    lblItem(plblGK_KIN).Caption = Format(CDbl(lblItem(plblNAI_KIN).Caption) + CDbl(lblItem(plblGAI_KIN).Caption), "#,##0")
        
                                            'ì‡ïîê∂éY  ç\ê¨î‰ó¶
    If CDbl(lblItem(plblNAI_KIN).Caption) + CDbl(lblItem(plblGAI_KIN).Caption) = 0 Then
        wkValue = 0
    Else
        wkValue = CDbl(lblItem(plblNAI_KIN).Caption) / (CDbl(lblItem(plblNAI_KIN).Caption) + CDbl(lblItem(plblGAI_KIN).Caption)) * 100
    End If
    lblItem(plblNAI_KIN_RITU).Caption = Format(wkValue, "#0.00") & "%"
                                            
                                            'äOïîê∂éY  ç\ê¨î‰ó¶
    lblItem(plblGAI_KIN_RITU).Caption = Format(100 - wkValue, "#0.00") & "%"
                                            'ç\ê¨  ç\ê¨î‰ó¶
    lblItem(plblGK_KIN_RITU).Caption = "100.00%"
                                            
                                            'ì‡ïîê∂éYÅ@ì‡ñÛÅ@éëçﬁ
    lblItem(plblNAI_UCHI_SHIZAI).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SH_KINGAKU, vbUnicode)), "#,##0")
                                            'ì‡ïîê∂éYÅ@ì‡ñÛÅ@çHóø
    lblItem(plblNAI_UCHI_KOURYO).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KO_KINGAKU, vbUnicode)), "#,##0")
                                            'ì‡ïîê∂éYÅ@ì‡ñÛÅ@ÇªÇÃëº
    lblItem(plblNAI_UCHI_ETC).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).ETC_KINGAKU, vbUnicode)), "#,##0")
                                            'äOïîê∂éYÅ@ì‡ñÛÅ@éëçﬁ
    lblItem(plblGAI_UCHI_SHIZAI).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SH_KINGAKU, vbUnicode)), "#,##0")
                                            'äOïîê∂éYÅ@ì‡ñÛÅ@çHóø
    lblItem(plblGAI_UCHI_KOURYO).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KO_KINGAKU, vbUnicode)), "#,##0")
                                            'äOïîê∂éYÅ@ì‡ñÛÅ@ÇªÇÃëº
    lblItem(plblGAI_UCHI_ETC).Caption = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).ETC_KINGAKU, vbUnicode)), "#,##0")
                                            'çáåvÅ@ì‡ñÛÅ@éëçﬁ
    lblItem(plblGK_UCHI_SHIZAI).Caption = Format(CDbl(lblItem(plblNAI_UCHI_SHIZAI).Caption) + CDbl(lblItem(plblGAI_UCHI_SHIZAI).Caption), "#,##0")
                                            'çáåvÅ@ì‡ñÛÅ@çHóø
    lblItem(plblGK_UCHI_KOURYO).Caption = Format(CDbl(lblItem(plblNAI_UCHI_KOURYO).Caption) + CDbl(lblItem(plblGAI_UCHI_KOURYO).Caption), "#,##0")
                                            'çáåvÅ@ì‡ñÛÅ@ÇªÇÃëº
    lblItem(plblGK_UCHI_ETC).Caption = Format(CDbl(lblItem(plblNAI_UCHI_ETC).Caption) + CDbl(lblItem(plblGAI_UCHI_ETC).Caption), "#,##0")
        
                                            'âøäiç\ê¨î‰
    lblItem(plblKAKAKU_RITU).Caption = "100.00%"
    If (CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + CDbl(lblItem(plblGK_UCHI_ETC).Caption)) = 0 Then
        lblItem(plblSHIZAI_RITU).Caption = "0.00"
    Else
        lblItem(plblSHIZAI_RITU).Caption = Format(CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) / (CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_ETC).Caption)) * 100, "#0.00") & "%"
    End If
    If (CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + CDbl(lblItem(plblGK_UCHI_ETC).Caption)) = 0 Then
        lblItem(plblKOURYO_RITU).Caption = "0.00"
    Else
        lblItem(plblKOURYO_RITU).Caption = Format(CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) / (CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_ETC).Caption)) * 100, "#0.00") & "%"
    End If
    
    If ((CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + CDbl(lblItem(plblGK_UCHI_ETC).Caption))) = 0 Then
        lblItem(plblETC_RITU).Caption = "0.00"
    Else
    
        lblItem(plblETC_RITU).Caption = Format(CDbl(lblItem(plblGK_UCHI_ETC).Caption) / (CDbl(lblItem(plblGK_UCHI_SHIZAI).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_KOURYO).Caption) + _
                                                                                                CDbl(lblItem(plblGK_UCHI_ETC).Caption)) * 100, "#0.00") & "%"
    End If
                                            'è¡îÔéëçﬁ
    lblItem(plblGENKA_KOSOU).Caption = Format(CLng(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode)), "#,##0")
    lblItem(plblGENKA_GAISOU).Caption = Format(CLng(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode)), "#,##0")
    lblItem(plblGENKA_SHIZAI).Caption = Format(CDbl(lblItem(plblGENKA_KOSOU).Caption) + CDbl(lblItem(plblGENKA_GAISOU).Caption), "#,##0")
    lblItem(plblGENKA_KOURYO).Caption = Format(CLng(StrConv(P_SEISAN_SUM_REC.GK_GENKA, vbUnicode)), "#,##0")
    lblItem(plblGENKA_GK).Caption = Format(CDbl(lblItem(plblGENKA_KOURYO).Caption) + CDbl(lblItem(plblGENKA_SHIZAI).Caption), "#,##0")
        
    
    '-------------------------------------  'é¿ê—ñæç◊ÇÃæØƒ
    Set SEISAN = Nothing
    
    Row = Min_Row - 1
    
    Call UniCode_Conv(K0_P_SEISAN_SUM.SHIMUKE_CODE, Text1(ptxSHIMUKE_CODE))
    Call UniCode_Conv(K0_P_SEISAN_SUM.CLASS_CODE, P_ClassSum_Key)
    
    
    com = BtOpGetGreater
    
    
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
                Exit Function
        End Select
    
        If Trim(StrConv(P_SEISAN_SUM_REC.CLASS_CODE, vbUnicode)) = P_ClassSum_Key Then
        Else
    
            Row = Row + 1
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        
        com = BtOpGetNext
    
    
    
    Loop
    
    
    Set TDBGrid1(pGridDETAIL).Array = SEISAN
    TDBGrid1(pGridDETAIL).ReBind
    TDBGrid1(pGridDETAIL).Update
    TDBGrid1(pGridDETAIL).MoveFirst
    
    
    PR000501.MousePointer = vbDefault
    
    
    List_Disp_Proc = False
    


End Function


Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           ê∂éYé¿ê—ÇÃì‡óeÇ∏ﬁÿØƒﬁÇ…æØƒÇ∑ÇÈ
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer

Dim wkValue     As Double


    Grid_Set_Proc = True
    
    
    


    
    
    
    
    SEISAN.ReDim Min_Row, Row, Min_Col, Max_Col


    'édå¸ÇØêÊ
    SEISAN(Row, colSHIMUKE_CODE) = StrConv(P_SEISAN_SUM_REC.SHIMUKE_CODE, vbUnicode)
    '∏◊Ω
    SEISAN(Row, colCLASS_CODE) = StrConv(P_SEISAN_SUM_REC.CLASS_CODE, vbUnicode)
    'ì‡ïîê∂éY åèêî
    SEISAN(Row, colNAI_CNT) = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).CNT, vbUnicode)), "#0")
    'ì‡ïîê∂éY êîó 
    SEISAN(Row, colNAI_SURYO) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SURYO, vbUnicode)), "#0")
    'äOïîê∂éY åèêî
    SEISAN(Row, colGAI_CNT) = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).CNT, vbUnicode)), "#0")
    'ì‡ïîê∂éY êîó 
    SEISAN(Row, colGAI_SURYO) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SURYO, vbUnicode)), "#0")
    'çáåv åèêî
    SEISAN(Row, colGK_CNT) = Format(CInt(SEISAN(Row, colNAI_CNT)) + CInt(SEISAN(Row, colGAI_CNT)), "#0")
    'çáåv êîó 
    SEISAN(Row, colGK_SURYO) = Format(CDbl(SEISAN(Row, colNAI_SURYO)) + CDbl(SEISAN(Row, colGAI_SURYO)), "#0")
    
    'çáåv íPâø
    SEISAN(Row, colGK_TANKA) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).TANKA, vbUnicode)), "#,##0.00")
    'çáåvÅ@ã‡äz
    SEISAN(Row, colGK_KIN) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KINGAKU, vbUnicode)), "#,##0")
    
    'éëçﬁÅ@íPâø
    SEISAN(Row, colSHIZAI_TANKA) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SH_TANKA, vbUnicode)), "#,##0.00")
    'éëçﬁÅ@ã‡äz
    SEISAN(Row, colSHIZAI_KIN) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SH_KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SH_KINGAKU, vbUnicode)), "#,##0")
    
    'çHóøÅ@íPâø
    SEISAN(Row, colKOURYO_TANKA) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KO_TANKA, vbUnicode)), "#,##0.00")
    'çHóøÅ@ã‡äz
    SEISAN(Row, colKOURYO_KIN) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KO_KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KO_KINGAKU, vbUnicode)), "#,##0")
    'ÇªÇÃëºÅ@íPâø
    SEISAN(Row, colETC_TANKA) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).ETC_TANKA, vbUnicode)), "#,##0.00")     'ÇªÇÃëºÅ@ã‡äz
    
    SEISAN(Row, colETC_KIN) = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).ETC_KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).ETC_KINGAKU, vbUnicode)), "#,##0")
    
    
    'édì¸å¥âøÅ@å¬ëï
    SEISAN(Row, colGENKA_KOSOU) = Format(CLng(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode)), "#,##0")
    'édì¸å¥âøÅ@äOëï
    SEISAN(Row, colGENKA_GAISOU) = Format(CLng(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode)), "#,##0")
    'édì¸å¥âøÅ@çHóø
    SEISAN(Row, colGENKA_KOURYO) = Format(CLng(StrConv(P_SEISAN_SUM_REC.GK_GENKA, vbUnicode)), "#,##0")
    
    
    
    
    
    
    
    
    Grid_Set_Proc = False

End Function


Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ÉRÅ[ÉhÉ}ÉXÉ^ÇÉRÉìÉ{Ç…ÉZÉbÉgÇ∑ÇÈÅB
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
                Call File_Error(sts, com, "ÉRÅ[ÉhÉ}ÉXÉ^")
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

Private Function SUM_Make_Proc(Data_Flg As Boolean) As Integer
'----------------------------------------------------------------------------
'                   ê∂éYé¿ê—èWåv√ﬁ∞¿çÏê¨
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim upd_com         As Integer
    
Dim Shizai_com      As Integer
    
Dim SKIP_Flg        As Boolean
    
Dim wkYMD           As String * 8
    
Dim wk_Val          As Double
Dim wk_Suryo        As Double
    
    
    
Dim i               As Integer
    
    
Dim Sum_Area(0 To 1)    As Sum_Area_tag


Dim KO_GENKA        As Long
Dim GA_GENKA        As Long
Dim GK_GENKA        As Long




    
    
    
    
    
    SUM_Make_Proc = True
    PR000501.MousePointer = vbHourglass

    '-----------------------------------------  èWåv√ﬁ∞¿ëSåèçÌèú


    com = BtOpGetFirst



    Do
    
    
        sts = BTRV(com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
        
    '-----------------------------------------  èWåvèàóùäJén
    
    Data_Flg = False
    Call UniCode_Conv(K1_P_SUKEIRE.SHIMUKE_CODE, Text1(ptxSHIMUKE_CODE).Text)

    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K1_P_SUKEIRE, Len(K1_P_SUKEIRE), 1)
            
        Select Case sts
            Case BtNoErr
                'édå¸ÇØêÊ∫∞ƒﬁ
                If Trim(Text1(ptxSHIMUKE_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxSHIMUKE_CODE).Text) <> Trim(StrConv(P_SUKEIRE_REC.SHIMUKE_CODE, vbUnicode)) Then
                        Exit Do
                    End If
                End If
                'éÛì¸îNåéì˙ÇÃÃﬁ⁄∞∏
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "è§ïiâªéwé¶éÛì¸óöó")
                Exit Function
        End Select
        
        SKIP_Flg = False
        
        
        If StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) < Format(CDate(Text1(ptxS_YMD).Text), "YYYYMMDD") Or _
            StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) > Format(CDate(Text1(ptxE_YMD).Text), "YYYYMMDD") Then
            SKIP_Flg = True
        End If
        
        
        If Not SKIP_Flg Then
        
            'éwé¶√ﬁ∞¿ì«Ç›çûÇ›
            Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, StrConv(P_SUKEIRE_REC.SHIJI_No, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    SKIP_Flg = True
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "éëçﬁíçï∂√ﬁ∞¿")
                    Exit Function
            End Select
                
                
            If Not SKIP_Flg Then
                
                Data_Flg = True
                'ê∂éYé¿ê—èWåv√ﬁ∞¿ì«Ç›çûÇ›
                    
                If Trim(Text1(ptxSHIMUKE_CODE).Text) = "" Then
                    Call UniCode_Conv(K0_P_SEISAN_SUM.SHIMUKE_CODE, "")
                Else
                    Call UniCode_Conv(K0_P_SEISAN_SUM.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                End If
                Call UniCode_Conv(K0_P_SEISAN_SUM.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                
                sts = BTRV(BtOpGetEqual, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
                Select Case sts
                    Case BtNoErr
                        upd_com = BtOpUpdate
                    Case BtErrKeyNotFound
                        upd_com = BtOpInsert
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
                        Exit Function
                End Select
                
                
                If upd_com = BtOpInsert Then
                
                
                    '∏◊ΩœΩ¿ì«Ç›çûÇ›
                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        
                        
                            '>>>>>>>>>> 2018.10.30
                            If Not IsNumeric(StrConv(P_CLASSREC.TANKA, vbUnicode)) Then
                                Call UniCode_Conv(P_CLASSREC.TANKA, "00000000.00")
                            End If
                            
                            
                            If Not IsNumeric(StrConv(P_CLASSREC.KOURYOU, vbUnicode)) Then
                                Call UniCode_Conv(P_CLASSREC.KOURYOU, "00000000.00")
                            End If
                            
                            If Not IsNumeric(StrConv(P_CLASSREC.ETC, vbUnicode)) Then
                                Call UniCode_Conv(P_CLASSREC.ETC, "00000000.00")
                            End If
                            '>>>>>>>>>> 2018.10.30
                        
                        
                        Case BtErrKeyNotFound
                            SKIP_Flg = True
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "∏◊ΩœΩ¿")
                            Exit Function
                    End Select
                
                    If Trim(Text1(ptxSHIMUKE_CODE).Text) = "" Then
                        Call UniCode_Conv(P_SEISAN_SUM_REC.SHIMUKE_CODE, "")
                    Else
                        Call UniCode_Conv(P_SEISAN_SUM_REC.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                    End If
                    
                    Call UniCode_Conv(P_SEISAN_SUM_REC.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                
                    For i = 0 To 1
                
                
                        
                        
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).CNT, "00000")
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).SURYO, "00000000.00")
                
                    
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).TANKA, StrConv(P_CLASSREC.TANKA, vbUnicode))
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).KINGAKU, "0000000000")
    
    
    
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).SH_TANKA, "00000000.00")
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).SH_KINGAKU, "0000000000")
    
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).KO_TANKA, StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).KO_KINGAKU, "0000000000")
    
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).ETC_TANKA, StrConv(P_CLASSREC.ETC, vbUnicode))
                        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).ETC_KINGAKU, "0000000000")
    
    
                
                    Next i
                
                    Call UniCode_Conv(P_SEISAN_SUM_REC.KO_GENKA, "00000000000")
                    Call UniCode_Conv(P_SEISAN_SUM_REC.GA_GENKA, "00000000000")
                    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GENKA, "00000000000")
                
                
                    Call UniCode_Conv(P_SEISAN_SUM_REC.FILLER, "")
                
                End If
                
                
                Select Case StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode)
                    Case P_TORI_SYANAI
                        i = 0
                        
                    Case P_TORI_GENERAL, P_TORI_NAISYOKU, P_TORI_GENKIN, P_TORI_ANOTHER, P_TORI_JIKYU
                        i = 1
                End Select
                'ê∂éYåèêî
                Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).CNT, Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).CNT, vbUnicode)) + 1, "00000"))
                'ê∂éYêîó 
                
                Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).SURYO, Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).SURYO, vbUnicode)) + _
                                                                    CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "00000000.00"))
                If Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode)) <> "" Then
                    'äOíççHóø
                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(P_CLASSREC.KOURYOU, "00000000.00")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "∏◊ΩœΩ¿")
                            Exit Function
                    End Select
                    wk_Val = CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                    wk_Val = wk_Val * CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                        
                    wk_Val = wk_Val + CLng(StrConv(P_SEISAN_SUM_REC.GK_GENKA, vbUnicode))
                        
                    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GENKA, Format(CLng(wk_Val), "0000000000"))
                End If
                'éëçﬁì‡ñÛ
                Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, StrConv(P_SUKEIRE_REC.SHIJI_No, vbUnicode))
                Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
                Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
                        
                Shizai_com = BtOpGetGreater
                    
                        
                Do
                    sts = BTRV(Shizai_com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> StrConv(P_SUKEIRE_REC.SHIJI_No, vbUnicode) Then
                                Exit Do
                            End If
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "è§ïiâªéwé¶(éq)√ﬁ∞¿")
                            Exit Function
                    End Select
                    
                    
                    Select Case StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode)
                        Case P_KOSOU    'å¬ëïéëçﬁ
                            
                            If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then
                            
                                wk_Suryo = CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
                            
                            
                                'ïiñ⁄É}ÉXÉ^ì«Ç›çûÇ›
                                Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                                                                                                    
                            
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound
                                    
                                        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                                        
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "ïiñ⁄œΩ¿")
                                        Exit Function
                                End Select
                            
                            
                                If IsNumeric(ITEMREC.G_ST_SHITAN) Then
                                    wk_Val = CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode))
                                Else
                                    wk_Val = 0
                                End If
                                
                                
                                wk_Val = CLng(wk_Val * (wk_Suryo * CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode))))
                            
                            
                                wk_Val = CLng(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode)) + wk_Val
                            
                                Call UniCode_Conv(P_SEISAN_SUM_REC.KO_GENKA, Format(wk_Val, "0000000000"))
                            
                            
                            End If
                            
                            
                        Case P_GAISOU   'äOëïéëçﬁ
                            
                                If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then
                                                                
                                    wk_Suryo = Int(CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)) / CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)))
                            
                                    'ïiñ⁄É}ÉXÉ^ì«Ç›çûÇ›
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                                                                                                    
                            
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                            '>> 2018.10.30
                                            If Not IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                                        
                                                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                                            
                                            End If
                                            '>> 2018.10.30
                                        
                                        
                                        Case BtErrKeyNotFound
                                            Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "ïiñ⁄œΩ¿")
                                            Exit Function
                                    End Select
                            
                            
                                    wk_Val = CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode))
                                    wk_Val = CDbl(wk_Val * wk_Suryo)
                                
                                
                                    wk_Val = CLng(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode)) + wk_Val
                                
                                    Call UniCode_Conv(P_SEISAN_SUM_REC.GA_GENKA, Format(wk_Val, "0000000000"))
                            
                            
                                End If
                        
                    End Select
                    
                    Shizai_com = BtOpGetNext
                    
                Loop
                    
                sts = BTRV(upd_com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, upd_com, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
                        Exit Function
                End Select
    
            End If
        End If
        
        com = BtOpGetNext
    
    Loop
    '-----------------------------------------  èWåv
    com = BtOpGetFirst



    Do
        
        DoEvents
        
        sts = BTRV(com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
                Exit Function
        End Select
    
        
        For i = 0 To 1
        
        
            '>>>>>>>>>> 2018.10.30
            If Not IsNumeric(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).TANKA, vbUnicode)) Then
                Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).TANKA, "00000000.00")
            End If
            
            
            If Not IsNumeric(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).KO_TANKA, vbUnicode)) Then
                Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).KO_TANKA, "00000000.00")
            End If
            
            If Not IsNumeric(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).ETC_TANKA, vbUnicode)) Then
                Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).ETC_TANKA, "00000000.00")
            End If
            '>>>>>>>>>> 2018.10.30
        
        
        
        
            'ê∂éYã‡äz
            
            wk_Val = CLng(CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).SURYO, vbUnicode)) * CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).TANKA, vbUnicode)))
            Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).KINGAKU, Format(wk_Val, "0000000000"))
            'çHóø
            wk_Val = CLng(CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).SURYO, vbUnicode)) * CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).KO_TANKA, vbUnicode)))
            Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).KO_KINGAKU, Format(wk_Val, "0000000000"))
            'ÇªÇÃëº
            wk_Val = CLng(CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).SURYO, vbUnicode)) * CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).ETC_TANKA, vbUnicode)))
            Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).ETC_KINGAKU, Format(wk_Val, "0000000000"))
        
            'éëçﬁÅiãtéZÅj
            wk_Val = CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).KINGAKU, vbUnicode))
            wk_Val = wk_Val - CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).KO_KINGAKU, vbUnicode)) - CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).ETC_KINGAKU, vbUnicode))
            Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).SH_KINGAKU, Format(wk_Val, "0000000000"))
            
            If CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).CNT, vbUnicode)) = 0 Then
                wk_Val = 0
            Else
                wk_Val = CDbl(wk_Val / CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).CNT, vbUnicode)))
            End If
            Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).SH_TANKA, Format(wk_Val, "00000000.00"))
        
        Next i
    
    
        sts = BTRV(BtOpUpdate, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpUpdate, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
                Exit Function
        End Select
    
    
    
        com = BtOpGetNext
    
    Loop


    '-----------------------------------------  çáåv⁄∫∞ƒﬁçÏê¨
    For i = 0 To 1
        Sum_Area(i).CNT = 0
        Sum_Area(i).SURYO = 0
        Sum_Area(i).KINGAKU = 0
        Sum_Area(i).SH_KINGAKU = 0
        Sum_Area(i).KO_KINGAKU = 0
        Sum_Area(i).ETC_KINGAKU = 0
    Next i
    
    
    KO_GENKA = 0
    GA_GENKA = 0
    GK_GENKA = 0
    
    
    
    
    
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
                Exit Function
        End Select
    
    
            
                
                
                
                        
                
                
                
    
    
    
        
        For i = 0 To 1
            'ê∂éYåèêî
            Sum_Area(i).CNT = Sum_Area(i).CNT + CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).CNT, vbUnicode))
            'ê∂éYêîó 
            Sum_Area(i).SURYO = Sum_Area(i).SURYO + CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).SURYO, vbUnicode))
            'ê∂éYã‡äz
            Sum_Area(i).KINGAKU = Sum_Area(i).KINGAKU + CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).KINGAKU, vbUnicode))
            'éëçﬁ
            Sum_Area(i).SH_KINGAKU = Sum_Area(i).SH_KINGAKU + CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).SH_KINGAKU, vbUnicode))
            'çHóø
            Sum_Area(i).KO_KINGAKU = Sum_Area(i).KO_KINGAKU + CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).KO_KINGAKU, vbUnicode))
            'ÇªÇÃëº
            Sum_Area(i).ETC_KINGAKU = Sum_Area(i).ETC_KINGAKU + CLng(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(i).ETC_KINGAKU, vbUnicode))
            
        
        Next i
    
        KO_GENKA = KO_GENKA + CLng(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode))
        GA_GENKA = GA_GENKA + CLng(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode))
        GK_GENKA = GK_GENKA + CLng(StrConv(P_SEISAN_SUM_REC.GK_GENKA, vbUnicode))
    
    
    
    
        com = BtOpGetNext
    
    Loop
    
    
    If Trim(Text1(ptxSHIMUKE_CODE).Text) = "" Then
        Call UniCode_Conv(P_SEISAN_SUM_REC.SHIMUKE_CODE, "")
    Else
        Call UniCode_Conv(P_SEISAN_SUM_REC.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
    End If
                    
    Call UniCode_Conv(P_SEISAN_SUM_REC.CLASS_CODE, P_ClassSum_Key)
            
    For i = 0 To 1
                
                
        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).CNT, Format(Sum_Area(i).CNT, "00000"))
        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).SURYO, Format(Sum_Area(i).SURYO, "00000000.00"))
        
            
        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).TANKA, "")
        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).KINGAKU, Format(Sum_Area(i).KINGAKU, "0000000000"))



        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).SH_TANKA, "")
        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).SH_KINGAKU, Format(Sum_Area(i).SH_KINGAKU, "0000000000"))

        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).KO_TANKA, "")
        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).KO_KINGAKU, Format(Sum_Area(i).KO_KINGAKU, "0000000000"))

        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).ETC_TANKA, "")
        Call UniCode_Conv(P_SEISAN_SUM_REC.UCHIWAKE(i).ETC_KINGAKU, Format(Sum_Area(i).ETC_KINGAKU, "0000000000"))
    
    
                
    Next i
                
    Call UniCode_Conv(P_SEISAN_SUM_REC.KO_GENKA, Format(KO_GENKA, "0000000000"))
    Call UniCode_Conv(P_SEISAN_SUM_REC.GA_GENKA, Format(GA_GENKA, "0000000000"))
    Call UniCode_Conv(P_SEISAN_SUM_REC.GK_GENKA, Format(GK_GENKA, "0000000000"))
                    
    Call UniCode_Conv(P_SEISAN_SUM_REC.FILLER, "")
                
    sts = BTRV(BtOpInsert, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpInsert, "ê∂éYé¿ê—èWåv√ﬁ∞¿")
            Exit Function
    End Select
    
    
    

    PR000501.MousePointer = vbDefault

   SUM_Make_Proc = False

End Function






