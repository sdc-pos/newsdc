VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1030611 
   BackColor       =   &H00FFFFFF&
   Caption         =   "出荷確認"
   ClientHeight    =   12600
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   17610
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   12600
   ScaleWidth      =   17610
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   14280
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   42
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   8
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   9
      Left            =   4560
      MaxLength       =   2
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   10
      Left            =   5160
      MaxLength       =   2
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   4
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   3
      Left            =   6960
      MaxLength       =   1
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   2
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   1
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   0
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   5
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   4320
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   10
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   360
      Index           =   7
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   360
      Index           =   6
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Width           =   732
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   8
      Top             =   720
      Width           =   852
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   12120
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "最　新"
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   12120
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "印　刷"
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   12120
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   12120
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   10935
      Left            =   0
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1080
      Width           =   17325
      _ExtentX        =   30559
      _ExtentY        =   19288
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "注区"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "出荷先"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "送状№"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "ID№"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "伝票№"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "収支"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "品番（外部）"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "品　名"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "出荷数"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "出庫済数"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "検品"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "伝票日付"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "品番(内部)"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "出庫表"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "取り込み日時"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "検品日"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "検品担当者"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "GLICS連携№"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "事業部"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "完了日時"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   20
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=20"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1349"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3201"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3096"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3969"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3863"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2275"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2170"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1640"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1535"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1455"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1349"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2646"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2540"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2619"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=1879"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1773"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=1879"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=1773"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=926"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=820"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=1"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=2037"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=1931"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(61)=   "Column(12).Width=476"
      Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=370"
      Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=8196"
      Splits(0)._ColumnProps(65)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(66)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(67)=   "Column(13).Width=1402"
      Splits(0)._ColumnProps(68)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(13)._WidthInPix=1296"
      Splits(0)._ColumnProps(70)=   "Column(13)._ColStyle=513"
      Splits(0)._ColumnProps(71)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(72)=   "Column(14).Width=3810"
      Splits(0)._ColumnProps(73)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(14)._WidthInPix=3704"
      Splits(0)._ColumnProps(75)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(76)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(77)=   "Column(15).Width=3810"
      Splits(0)._ColumnProps(78)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(15)._WidthInPix=3704"
      Splits(0)._ColumnProps(80)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(81)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(82)=   "Column(16).Width=3969"
      Splits(0)._ColumnProps(83)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(84)=   "Column(16)._WidthInPix=3863"
      Splits(0)._ColumnProps(85)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(86)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(87)=   "Column(17).Width=3810"
      Splits(0)._ColumnProps(88)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(89)=   "Column(17)._WidthInPix=3704"
      Splits(0)._ColumnProps(90)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(91)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(92)=   "Column(18).Width=3810"
      Splits(0)._ColumnProps(93)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(94)=   "Column(18)._WidthInPix=3704"
      Splits(0)._ColumnProps(95)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(96)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(97)=   "Column(19).Width=3810"
      Splits(0)._ColumnProps(98)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(19)._WidthInPix=3704"
      Splits(0)._ColumnProps(100)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(101)=   "Column(19).Order=20"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.alignment=2,.bold=0,.fontsize=1050"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=87,.parent=2,.namedParent=89"
      _StyleDefs(23)  =   "FilterBarStyle:id=90,.parent=1,.namedParent=92"
      _StyleDefs(24)  =   "Splits(0).Style:id=53,.parent=1,.bgcolor=&HFFFF&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=62,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=54,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=55,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=56,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=58,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=57,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=59,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=60,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=61,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=88,.parent=87"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=91,.parent=90"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=14,.parent=53"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=11,.parent=54"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=12,.parent=55"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=13,.parent=57"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=18,.parent=53"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=15,.parent=54"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=16,.parent=55"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=17,.parent=57"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=114,.parent=53"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=111,.parent=54"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=112,.parent=55"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=113,.parent=57"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=48,.parent=53"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=45,.parent=54"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=46,.parent=55"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=47,.parent=57"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=53"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=54"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=55"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=57"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=102,.parent=53"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=19,.parent=54"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=20,.parent=55"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=57"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=70,.parent=53"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=54"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=55"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=57"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=78,.parent=53"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=54"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=55"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=57"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=82,.parent=53,.alignment=1,.locked=0"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=54,.alignment=2"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=55,.alignment=3"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=57"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=86,.parent=53,.alignment=1,.locked=0"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=54,.alignment=2"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=55,.alignment=3"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=57"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=24,.parent=53,.alignment=2,.locked=0"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=21,.parent=54,.alignment=3"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=22,.parent=55,.alignment=3"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=23,.parent=57"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=28,.parent=53"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=25,.parent=54"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=26,.parent=55"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=27,.parent=57"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=40,.parent=53,.alignment=3,.locked=-1"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=37,.parent=54,.alignment=3"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=38,.parent=55,.alignment=3"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=39,.parent=57"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=44,.parent=53,.alignment=2"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=41,.parent=54"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=42,.parent=55"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=43,.parent=57"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=52,.parent=53"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=49,.parent=54"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=50,.parent=55"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=51,.parent=57"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=96,.parent=53"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=93,.parent=54"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=94,.parent=55"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=95,.parent=57"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=100,.parent=53"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=97,.parent=54"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=98,.parent=55"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=99,.parent=57"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=74,.parent=53"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=71,.parent=54"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=72,.parent=55"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=73,.parent=57"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=106,.parent=53"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=103,.parent=54"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=104,.parent=55"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=105,.parent=57"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=110,.parent=53"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=107,.parent=54"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=108,.parent=55"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=109,.parent=57"
      _StyleDefs(116) =   "Named:id=29:Normal"
      _StyleDefs(117) =   ":id=29,.parent=0"
      _StyleDefs(118) =   "Named:id=30:Heading"
      _StyleDefs(119) =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(120) =   ":id=30,.wraptext=-1"
      _StyleDefs(121) =   "Named:id=31:Footing"
      _StyleDefs(122) =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(123) =   "Named:id=32:Selected"
      _StyleDefs(124) =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(125) =   "Named:id=33:Caption"
      _StyleDefs(126) =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(127) =   "Named:id=34:HighlightRow"
      _StyleDefs(128) =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(129) =   "Named:id=35:EvenRow"
      _StyleDefs(130) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(131) =   "Named:id=36:OddRow"
      _StyleDefs(132) =   ":id=36,.parent=29"
      _StyleDefs(133) =   "Named:id=89:RecordSelector"
      _StyleDefs(134) =   ":id=89,.parent=30"
      _StyleDefs(135) =   "Named:id=92:FilterBar"
      _StyleDefs(136) =   ":id=92,.parent=29"
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   240
      Index           =   14
      Left            =   5520
      TabIndex        =   41
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   240
      Index           =   13
      Left            =   4320
      TabIndex        =   40
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   240
      Index           =   12
      Left            =   4920
      TabIndex        =   39
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(空白：全件/1：国内/2：海外)"
      Height          =   240
      Index           =   11
      Left            =   12840
      TabIndex        =   38
      Top             =   240
      Width           =   3360
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "販売区分"
      Height          =   240
      Index           =   10
      Left            =   11400
      TabIndex        =   37
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(空白：全件/1：国内/3：振替/*:1,3)"
      Height          =   240
      Index           =   9
      Left            =   7320
      TabIndex        =   36
      Top             =   240
      Width           =   4080
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ﾃﾞｰﾀ区分"
      Height          =   240
      Index           =   8
      Left            =   5880
      TabIndex        =   35
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "日～"
      Height          =   240
      Index           =   7
      Left            =   3120
      TabIndex        =   33
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   240
      Index           =   6
      Left            =   2520
      TabIndex        =   32
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   240
      Index           =   5
      Left            =   1920
      TabIndex        =   31
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票日付"
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   30
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   240
      Index           =   3
      Left            =   10920
      TabIndex        =   29
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票枚数　検品済／予定"
      Height          =   240
      Index           =   2
      Left            =   7440
      TabIndex        =   28
      Top             =   840
      Width           =   2640
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "注文区分"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   840
      Width           =   960
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
      TabIndex        =   27
      Top             =   12120
      Width           =   180
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷先"
      Height          =   240
      Index           =   0
      Left            =   2280
      TabIndex        =   25
      Top             =   840
      Width           =   720
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Checked         =   -1  'True
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030611"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Const ptxSyuka_YY% = 0          '出荷日　年
Private Const ptxSyuka_MM% = 1          '出荷日　月
Private Const ptxSyuka_DD% = 2          '出荷日　日

Private Const ptxDATA_KBN% = 3          'ﾃﾞｰﾀ区分
Private Const ptxHAN_KBN% = 4           '販売区分


Private Const ptxMUKE_CODE% = 5         '向け先（コード入力用）
Private Const ptxDEN_MAISU_JI% = 6      '伝票枚数　実績
Private Const ptxDEN_MAISU_YO% = 7      '伝票枚数　予定


Private Const ptxE_Syuka_YY% = 8        '出荷日　年     2012.11.13
Private Const ptxE_Syuka_MM% = 9        '出荷日　月     2012.11.13
Private Const ptxE_Syuka_DD% = 10       '出荷日　日     2012.11.13


                                                        '2012.11.13
Private Const Text_Max% = 10             '画面項目別最大ｲﾝﾃﾞｯｸｽ


Private Const pcmbCYU_KBN% = 0          '注文区分
Private Const pcmbMUKE_CODE% = 1        '向け先


Dim SYUKA As New XArrayDB

Private Const Min_Row% = 1              '最小行数
'Private Const Max_Row& = 2000           '最大行数
Dim Max_Row    As Integer               'グリッド最大表示件数

Dim SYUKA_DATA  As String               '出荷データフルパス


Private Const Min_Col% = 0              '最小列数
'Private Const Max_Col% = 17             '最大列数
'Private Const Max_Col% = 18             '最大列数       17-->18 2011.03.30
Private Const Max_Col% = 19             '最大列数       18-->19 2016.09.29

Private Const ColCYU_KBN% = 0           '注文区分
Private Const ColMUKE_CODE% = 1         '出荷先

Private Const ColOKURI_NO% = 2          '送り状№

Private Const ColID_NO% = 3             'ID№
Private Const ColDEN_NO% = 4            '伝票№
Private Const ColSYUKO_SYUSI& = 5       '出庫収支
Private Const ColHIN_GAI% = 6           '品番（外部）
Private Const ColHIN_NAME% = 7          '品名
Private Const ColYOTEI_QTY% = 8         '出荷予定数
Private Const ColFIX_QTY% = 9           '出荷実績
Private Const ColKENPIN_MARK% = 10       '検品
Private Const ColDEN_DT% = 11            '伝票日付
Private Const ColSort_Mark% = 12         'ＳＯＲＴマーク
Private Const ColPrint% = 13            '出庫表印刷マーク
Private Const ColIns_Date% = 14         '取込み日時

Private Const ColKENPIN_Date% = 15      '検品日
Private Const ColKENPIN_TANTO% = 16     '検品担当者

Private Const ColLK_SEQ_NO% = 17        'ﾘﾝｸ№

Private Const ColJGYOBU% = 18           '事業部

Private Const ColKAN_YMD% = 19          '完了日時 2011.03.30



Private Const Sort_MISYUKO$ = "0"       '未出庫
Private Const Sort_SYUKOSUMI$ = "1"     '出庫済
Private Const Sort_KENPIN$ = "2"        '検品済

Private Const KENPIN_ON$ = "○"         '検品済
Private Const KENPIN_OFF$ = "×"        '未検品


Private Inspe_F As Integer              '検品方法


'2011.08.03
Private Sort_Tbl(ColCYU_KBN To ColKAN_YMD) _
                            As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
Dim HEAD_CLICK              As Integer
'2011.08.03

Dim Inspe_Choku_F           As Integer                  '2016.09.29


'Private Const Last_Update_Day$ = "[F103061] 2018.01.12 14:30"
Private Const Last_Update_Day$ = "[F103061] 2018.01.12 15:15"



Private Sub Combo_Click(Index As Integer)
    Select Case Index
        Case pcmbCYU_KBN
            
            
            Text(ptxMUKE_CODE).SetFocus
    End Select

End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        Case pcmbCYU_KBN
            Text(ptxMUKE_CODE).SetFocus
        Case pcmbMUKE_CODE
            Text(ptxMUKE_CODE).Text = Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8)
            If List_Disp_Proc Then
                Unload Me
            End If
    End Select

End Sub


Private Sub Command_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
            
'>>>>>>>>>>>>>>>>   2018.01.12
        Case 3
            Call Input_Lock
            
                        
            
            
            Call Form_HCopy_Win7(Picture1, vbPRPSA4, vbPRORLandscape)

        
            Call Input_UnLock
'>>>>>>>>>>>>>>>>   2018.01.12
        
        Case 7                              '再表示
            Text(ptxMUKE_CODE).Text = Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8)
            If List_Disp_Proc Then
                Unload Me
            End If
        Case 8                              'データ出力
        
            Beep
            ans = MsgBox("「出荷予定」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If ans = vbYes Then
                If OUTPUT_Proc() Then
                    Unload Me
                End If
            End If
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub
Private Sub Form_DblClick()
'    PrintForm                  '2018.01.12
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
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

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
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
                                '出荷データファイル名取り込み
    If GetIni("FILE", "SYUKA_DATA", "SYS", c) Then
        Beep
        MsgBox "出荷データファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    SYUKA_DATA = Trim(c)
                                

    '---------------------------------------------  2011.08.06 SYS.INI-->F103061.INI
                    '最大表示件数の獲得
    If GetIni(App.EXEName, "LISTMAX", App.EXEName, c) Then
        Beep
        MsgBox "最大表示件数の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    Max_Row = CInt(RTrim(c))
                                
                                
                    '検品方法の獲得
    
    Inspe_F = 0
    
    If GetIni(App.EXEName, "Inspection", App.EXEName, c) Then
    Else
        If IsNumeric(Trim(c)) Then
            Inspe_F = CInt(Trim(c))
        End If
    End If
    '---------------------------------------------  2011.08.06 SYS.INI-->F103061.INI
                                
                                
                                
                                
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    '全ＢＵを選択可能とする2006.08.29
''    If UBound(JGYOBU_T) > 0 Then
        ReDim Preserve JGYOBU_T(UBound(JGYOBU_T) + 1)
        JGYOBU_T(UBound(JGYOBU_T)).CODE = "*"
        JGYOBU_T(UBound(JGYOBU_T)).NAME = "全BU"
        JGYOBU_T(UBound(JGYOBU_T)).COLOR = 12
''    End If


    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030611.Caption = "出荷確認（" + RTrim(JGYOBU_T(i).NAME) + ") " & Last_Update_Day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If


                                '直送検品用ﾌｧｲﾙ　ＯＰＥＮ   2016.09.29
    Inspe_Choku_F = 1
    If HTIdDelv_Open(BtOpenNomal) Then
        Inspe_Choku_F = 0
    Else
        If HTDelvNo_Open(BtOpenNomal) Then
            Unload Me
        End If
        If HTDrctId_Open(BtOpenNomal) Then
            Unload Me
        End If
    End If


    If Inspe_Choku_F = 0 Then
        TDBGrid1.Columns(ColOKURI_NO).Visible = False
    End If

    '2011.08.03
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0                 'ﾃﾞﾌｫﾙﾄ昇順
    Next i
    '2011.08.03


'出荷日付
    Text(ptxSyuka_YY).Text = Left(Format(Now, "YYYYMMDD"), 4)
    Text(ptxSyuka_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)
    Text(ptxSyuka_DD).Text = Right(Format(Now, "YYYYMMDD"), 2)

    Text(ptxE_Syuka_YY).Text = Left(Format(Now, "YYYYMMDD"), 4)     '2012.11.13
    Text(ptxE_Syuka_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)   '2012.11.13
    Text(ptxE_Syuka_DD).Text = Right(Format(Now, "YYYYMMDD"), 2)    '2012.11.13


'ﾃﾞｰﾀ区分
    Text(ptxDATA_KBN).Text = "1"
'販売区分
    Text(ptxHAN_KBN).Text = "1"

'向け先設定
    If MTS_Set_Proc() Then
        Unload Me
    End If

'ｺﾝﾎﾞ初期設定
    Combo(pcmbCYU_KBN).AddItem "全て" & "   " & " "
    
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_1 & "   " & CYU_KBN_TUK
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_2 & "   " & CYU_KBN_SPO
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_3 & "   " & CYU_KBN_HJU
'    Combo(pcmbCYU_KBN).AddItem CYU_KBN_4
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_E & "   " & CYU_KBN_BOU
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_T & "   " & CYU_KBN_KIN
    Combo(pcmbCYU_KBN).ListIndex = 0

    Text(ptxSyuka_YY).SetFocus



End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1030611.Caption = "出荷確認（" + RTrim(JGYOBU_T(Index).NAME) + ") " & Last_Update_Day
    SubMenu(Index).Checked = True
    If Last_JGYOBU <> JGYOBU_T(Index).CODE Then
        Last_JGYOBU = JGYOBU_T(Index).CODE
        LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
        LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

    End If

End Sub

Private Function MTS_Set_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim Edit        As String

    MTS_Set_Proc = True
    
    Call Input_Lock
    
    
    Combo(pcmbMUKE_CODE).Clear
    
    Combo(pcmbMUKE_CODE).AddItem "全て　　　" & "   " & Space(16)
        
    
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "向け先マスタ")
                Exit Function
        End Select
        
        Edit = StrConv(MTSREC.MUKE_NAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        
        
        Combo(pcmbMUKE_CODE).AddItem Edit
    
        com = BtOpGetNext
    
    Loop

    If Combo(pcmbMUKE_CODE).ListCount <= 0 Then
    Else
        Combo(pcmbMUKE_CODE).ListIndex = 0
    End If

    Call Input_UnLock

    MTS_Set_Proc = False
End Function


Private Function List_Disp_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim DEN_MAISU   As Long
Dim KAN_MAISU   As Long
    
Dim Skip_Flg    As Boolean
    
    
    List_Disp_Proc = True
                                    
'    Call Input_Lock
                                    
    Me.MousePointer = vbArrowHourglass
                                    
                                    'テーブルリセット
    Set SYUKA = Nothing
                                    '出荷予定読み込み開始
    

'2011.08.04 読込みKEY変更
'    If Last_JGYOBU = "*" Then
'        Call UniCode_Conv(K2_Y_SYU.JGYOBU, "") '事業部
'    Else
'        Call UniCode_Conv(K2_Y_SYU.JGYOBU, Last_JGYOBU) '事業部
'    End If
'                                                    '注文区分
'    Call UniCode_Conv(K2_Y_SYU.KEY_CYU_KBN, "")
'                                                    '向け先
'    Call UniCode_Conv(K2_Y_SYU.KEY_MUKE_CODE, "")
'    Call UniCode_Conv(K2_Y_SYU.KEY_SS_CODE, "")

    Call UniCode_Conv(K9_Y_SYU.KEY_SYUKA_YMD, Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text)
'2011.08.04 読込みKEY変更
    
    
    Row = Min_Row - 1
        
    DEN_MAISU = 0
    KAN_MAISU = 0
    
    
    
    com = BtOpGetGreaterEqual
    
''com = BtOpGetFirst
    Do
        
        DoEvents
        
        
'2011.08.04 読込みKEY変更
'        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K9_Y_SYU, Len(K9_Y_SYU), 9)
'2011.08.04 読込みKEY変更
    
    
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定")
                List_Disp_Proc = SYS_ERR
                Exit Function
        End Select
                                
        Skip_Flg = False
                                
        
'---------------------------------  2012.11.13  範囲指定に変更
'        '出荷日 KEYﾌﾞﾚｰｸ
'        If Len(Trim((Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text))) = 0 Then
'        Else
'            If (Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text) <> StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then
'                '2011.08.04 読込みKEY変更
''                Skip_Flg = True
'                '2011.08.04 読込みKEY変更
'                Exit Do
'            End If
'        End If


        If Len(Trim((Text(ptxE_Syuka_YY).Text & Text(ptxE_Syuka_MM).Text & Text(ptxE_Syuka_DD).Text))) = 0 Then
        Else
            If (Text(ptxE_Syuka_YY).Text & Text(ptxE_Syuka_MM).Text & Text(ptxE_Syuka_DD).Text) < StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then
                Exit Do
            End If
        End If

'---------------------------------  2012.11.13  範囲指定に変更
        
        '事業部 KEYﾌﾞﾚｰｸ
        If Last_JGYOBU = "*" Then
        Else
            If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
'2011.08.04 読込みKEY変更
'                Exit Do
                Skip_Flg = True
'2011.08.04 読込みKEY変更
            End If
        End If
                                
                                
        '注文区分 KEYﾌﾞﾚｰｸ
        If Len(Trim(Right(Combo(pcmbCYU_KBN).Text, 1))) <> 0 Then
            If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCYU_KBN).Text, 1) Then
                Skip_Flg = True
            End If
        End If
        '向け先 KEYﾌﾞﾚｰｸ
        If Len(Trim(Right(Combo(pcmbMUKE_CODE).Text, 16))) <> 0 Then
            If StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) <> Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8) Or _
                StrConv(Y_SYUREC.SS_CODE, vbUnicode) <> Right(Combo(pcmbMUKE_CODE).Text, 8) Then
                Skip_Flg = True
            End If
        End If
        'ﾃﾞｰﾀ区分
        If Trim(Text(ptxDATA_KBN).Text) = "" Then
        Else
            
            If Trim(Text(ptxDATA_KBN).Text) = "*" Then
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" Or StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "3" Then
                Else
                    Skip_Flg = True
                End If
            Else
                If Text(ptxDATA_KBN).Text <> StrConv(Y_SYUREC.DATA_KBN, vbUnicode) Then
                    Skip_Flg = True
                End If
            End If
        End If
        '販売区分
        If Trim(Text(ptxHAN_KBN).Text) = "" Then
        Else
            If Text(ptxHAN_KBN).Text <> StrConv(Y_SYUREC.HAN_KBN, vbUnicode) Then
                Skip_Flg = True
            End If
        End If
                
                
        If Not Skip_Flg Then
            DEN_MAISU = DEN_MAISU + 1
            
                                        '検品完了
            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                KAN_MAISU = KAN_MAISU + 1
            End If
            
            Row = Row + 1
            If Row > Max_Row Then
                Beep
                MsgBox "最大表示行数を超えました。"
                Exit Do
            End If
                    
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
        
    Loop
    
                                
                                'DBテーブルリンク
    If DEN_MAISU < 1 Then
    Else
        SYUKA.QuickSort Min_Row, (SYUKA.UpperBound(1)), ColSort_Mark, XORDER_ASCEND, XTYPE_STRING, _
                                                        ColDEN_NO, XORDER_ASCEND, XTYPE_STRING
    End If
    
    Set TDBGrid1.Array = SYUKA
    
    TDBGrid1.style.Locked = True
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
    Text(ptxDEN_MAISU_JI).Text = Format(KAN_MAISU, "#,##0")
                                
    Text(ptxDEN_MAISU_YO).Text = Format(DEN_MAISU, "#,##0")
    
'    Call Input_UnLock
    
    Me.MousePointer = vbDefault
    
    
    Combo(pcmbMUKE_CODE).SetFocus
    
    List_Disp_Proc = False

    
End Function

Private Function OUTPUT_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

    
Dim ret         As Integer
    

Dim FileNo      As Integer
Dim FileName    As String
    
Dim Skip_Flg    As Boolean
    
    
    OUTPUT_Proc = True
                                    
'    Call Input_Lock

    FileNo = FreeFile
    
    FileName = SYUKA_DATA
    ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), ret) & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - ret)
    
    On Error GoTo Error_Proc
    
    Open (FileName) For Output As FileNo

'    Write #FileNo, "注文区分", "出荷先", "ＩＤ№", "伝票№", "品番（外部）", "品番（内部）", "品名", "出荷予定数", "済み数", "検品", "伝票日付", Format(Now, "yyyy/mm/dd HH:mm:ss") & " 現在"
    Write #FileNo, , , , , , , , , , , , Format(Now, "yyyy/mm/dd HH:mm:ss") & " 現在"
    Write #FileNo, "注文区分", "出荷先", "ＩＤ№", "伝票№", "品番（外部）", "品番（内部）", "品名", "出荷予定数", "済み数", "検品", "伝票日付", "上位ﾘﾝｸ用向け先", "完了日時"

                                    '出荷予定読み込み開始
'    Call UniCode_Conv(K2_Y_SYU.JGYOBU, Last_JGYOBU) '事業部
'
'                                                    '注文区分
'    Call UniCode_Conv(K2_Y_SYU.KEY_CYU_KBN, "")
'                                                    '向け先
'    Call UniCode_Conv(K2_Y_SYU.KEY_MUKE_CODE, "")
'    Call UniCode_Conv(K2_Y_SYU.KEY_SS_CODE, "")
                                                                                                    '2015.07.17
    Call UniCode_Conv(K9_Y_SYU.KEY_SYUKA_YMD, Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text)
    com = BtOpGetGreaterEqual
    Do
'        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)            '2015.07.17
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K9_Y_SYU, Len(K9_Y_SYU), 9)             '2015.07.17
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定")
                OUTPUT_Proc = SYS_ERR
                Exit Function
        End Select
        
'-----------------------------------    2015.07.17      セレクト変更
'                                '事業部 KEYﾌﾞﾚｰｸ
'        If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
'            Exit Do
'        End If
'
'        Skip_Flg = False
'        If Len(Trim(Right(Combo(pcmbCYU_KBN).Text, 1))) <> 0 Then
'            If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCYU_KBN).Text, 1) Then
'                Skip_Flg = True
'            End If
'        End If
'                            '向け先 KEYﾌﾞﾚｰｸ
'
'
'        If Len(Trim(Right(Combo(pcmbMUKE_CODE).Text, 16))) <> 0 Then
'            If StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) <> Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8) Or _
'                StrConv(Y_SYUREC.SS_CODE, vbUnicode) <> Right(Combo(pcmbMUKE_CODE).Text, 8) Then
'                Skip_Flg = True
'            End If
'        End If
'
'        If Len(Trim((Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text))) = 0 Then
'        Else
''2012.11.13            If (Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text) <> StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then
''2012.11.13
'            If (Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text) < StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then
'                Skip_Flg = True
'            End If
'        End If
'
'        '2012.11.13
'        If Len(Trim((Text(ptxE_Syuka_YY).Text & Text(ptxE_Syuka_MM).Text & Text(ptxE_Syuka_DD).Text))) = 0 Then
'        Else
'            If (Text(ptxE_Syuka_YY).Text & Text(ptxE_Syuka_MM).Text & Text(ptxE_Syuka_DD).Text) < StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then
'                Skip_Flg = True
'            End If
'        End If
'        '2012.11.13
'
'
        Skip_Flg = False

        If Len(Trim((Text(ptxE_Syuka_YY).Text & Text(ptxE_Syuka_MM).Text & Text(ptxE_Syuka_DD).Text))) = 0 Then
        Else
            If (Text(ptxE_Syuka_YY).Text & Text(ptxE_Syuka_MM).Text & Text(ptxE_Syuka_DD).Text) < StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then
                Exit Do
            End If
        End If



        '事業部 KEYﾌﾞﾚｰｸ
        If Last_JGYOBU = "*" Then
        Else
            If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Skip_Flg = True
            End If
        End If
                                
                                
        '注文区分 KEYﾌﾞﾚｰｸ
        If Len(Trim(Right(Combo(pcmbCYU_KBN).Text, 1))) <> 0 Then
            If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCYU_KBN).Text, 1) Then
                Skip_Flg = True
            End If
        End If
        '向け先 KEYﾌﾞﾚｰｸ
        If Len(Trim(Right(Combo(pcmbMUKE_CODE).Text, 16))) <> 0 Then
            If StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) <> Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8) Or _
                StrConv(Y_SYUREC.SS_CODE, vbUnicode) <> Right(Combo(pcmbMUKE_CODE).Text, 8) Then
                Skip_Flg = True
            End If
        End If
        'ﾃﾞｰﾀ区分
        If Trim(Text(ptxDATA_KBN).Text) = "" Then
        Else
            
            If Trim(Text(ptxDATA_KBN).Text) = "*" Then
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" Or StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "3" Then
                Else
                    Skip_Flg = True
                End If
            Else
                If Text(ptxDATA_KBN).Text <> StrConv(Y_SYUREC.DATA_KBN, vbUnicode) Then
                    Skip_Flg = True
                End If
            End If
        End If
        '販売区分
        If Trim(Text(ptxHAN_KBN).Text) = "" Then
        Else
            If Text(ptxHAN_KBN).Text <> StrConv(Y_SYUREC.HAN_KBN, vbUnicode) Then
                Skip_Flg = True
            End If
        End If



'-----------------------------------    2015.07.17      セレクト変更
        If Not Skip_Flg Then
            Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                Case CYU_KBN_TUK
                    Write #FileNo, CYU_KBN_1,
                Case CYU_KBN_SPO
                    Write #FileNo, CYU_KBN_2,
                Case CYU_KBN_HJU
                    Write #FileNo, CYU_KBN_3,
                Case CYU_KBN_TOK
                    Write #FileNo, CYU_KBN_4,
                Case CYU_KBN_BOU
                    Write #FileNo, CYU_KBN_E,
                Case CYU_KBN_KIN
                    Write #FileNo, CYU_KBN_T,
                Case Else
                    Write #FileNo, ,
            End Select
            
            
            Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
            Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
            Select Case sts
                Case BtNoErr
                    Write #FileNo, StrConv(MTSREC.MUKE_CODE, vbUnicode) & " " & StrConv(MTSREC.MUKE_DNAME, vbUnicode),
                Case BtErrKeyNotFound
                    Write #FileNo, StrConv(MTSREC.MUKE_CODE, vbUnicode),
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
                    Exit Function
            End Select
            
            
            
            Write #FileNo, StrConv(Y_SYUREC.ID_NO, vbUnicode),
            Write #FileNo, StrConv(Y_SYUREC.DEN_NO, vbUnicode),
            Write #FileNo, StrConv(Y_SYUREC.HIN_NO, vbUnicode),
    '2004        Write #FileNo, StrConv(Y_SYUREC.HIN_NAI, vbUnicode),
                                    '品目マスタ読込み
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Write #FileNo, StrConv(ITEMREC.HIN_NAI, vbUnicode),
                    Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
                Case BtErrKeyNotFound
                    Write #FileNo, ,
                    Write #FileNo, ,
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
                                                                        '出荷予定数
            Write #FileNo, Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#,##0"),
                                                                        '出荷実績数
            Write #FileNo, Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "#,##0"),
                                                                        '検品マーク
            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) = 0 Then
                                    '未検品
                Write #FileNo, KENPIN_OFF,
            Else
                                    '検品済
                Write #FileNo, KENPIN_ON,
            End If
                
            Write #FileNo, Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" _
                            & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" _
                            & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 7, 2),
        
        
            Write #FileNo, Trim(StrConv(Y_SYUREC.LK_MUKE_CODE, vbUnicode));
            
            
            If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = KAN_KBN_FIN Then
                Write #FileNo, Mid(StrConv(Y_SYUREC.KAN_YMD, vbUnicode), 1, 4) & "/" _
                                        & Mid(StrConv(Y_SYUREC.KAN_YMD, vbUnicode), 5, 2) & "/" _
                                        & Mid(StrConv(Y_SYUREC.KAN_YMD, vbUnicode), 7, 2) & " " _
                                        & Mid(StrConv(Y_SYUREC.KAN_HMS, vbUnicode), 1, 2) & ":" _
                                        & Mid(StrConv(Y_SYUREC.KAN_HMS, vbUnicode), 3, 2) & ":" _
                                        & Mid(StrConv(Y_SYUREC.KAN_HMS, vbUnicode), 5, 2),
            Else
                Write #FileNo, "",
            End If
            Write #FileNo,

        
        End If
        com = BtOpGetNext
        
        DoEvents
    Loop

    Close #FileNo
    
'    Call Input_UnLock         '画面項目ロック解除
    
    Beep
    MsgBox "「" & FileName & "」は正常に出力されました。"

    Combo(pcmbMUKE_CODE).SetFocus
    
    OUTPUT_Proc = False
    
    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName & "が使用中です。"
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If


End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1030611.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030611)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030611)


    F1030611.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer

Dim sts As Integer

    
Dim ID_Cnt      As Integer          '2016.09.29
Dim OKURI_NO    As String * 20      '2016.09.29
Dim com         As Integer          '2016.09.29


    
    Grid_Set_Proc = True

    

    SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
    
    Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
        Case CYU_KBN_TUK
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_1
        Case CYU_KBN_SPO
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_2
        Case CYU_KBN_HJU
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_3
        Case CYU_KBN_TOK
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_4
        Case CYU_KBN_BOU
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_E
        Case CYU_KBN_KIN
            SYUKA(Row, ColCYU_KBN) = CYU_KBN_T
        Case Else
            Debug.Print
    End Select
    
    
    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColMUKE_CODE) = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.MUKE_DNAME, vbUnicode)
        Case BtErrKeyNotFound
            SYUKA(Row, ColMUKE_CODE) = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
            Exit Function
    End Select
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    送状№の獲得    2016.09.29
    If Inspe_Choku_F = 1 Then
    
        Call UniCode_Conv(K0_HTIdDelv.IDNO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
        Call UniCode_Conv(K0_HTIdDelv.DelvNo, "")
    
        com = BtOpGetGreater
        ID_Cnt = 0
        Do
            DoEvents
            sts = BTRV(com, HTIdDelv_POS, HTIdDelvREC, Len(HTIdDelvREC), K0_HTIdDelv, Len(K0_HTIdDelv), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(HTIdDelvREC.IDNO, vbUnicode) <> StrConv(Y_SYUREC.ID_NO, vbUnicode) Then
                        Exit Do
                    End If
                        
                    If ID_Cnt = 0 Then
                    
                        OKURI_NO = StrConv(HTIdDelvREC.DelvNo, vbUnicode)
                        
                    End If
                
                
                    ID_Cnt = ID_Cnt + 1
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "HTIdDelv.dat")
                    Exit Function
            End Select
        
            com = BtOpGetNext
        
        Loop
    
    
        If ID_Cnt = 0 Then
            SYUKA(Row, ColOKURI_NO) = ""
        Else
            If ID_Cnt = 1 Then
                SYUKA(Row, ColOKURI_NO) = OKURI_NO
            Else
                SYUKA(Row, ColOKURI_NO) = "*" & OKURI_NO
            End If
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    送状№の獲得    2016.09.29
    
    SYUKA(Row, ColID_NO) = StrConv(Y_SYUREC.ID_NO, vbUnicode)       'ＩＤ№
    SYUKA(Row, ColDEN_NO) = StrConv(Y_SYUREC.DEN_NO, vbUnicode)     '伝票№
    SYUKA(Row, ColSYUKO_SYUSI) = StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)   '出庫収支

    SYUKA(Row, ColHIN_GAI) = StrConv(Y_SYUREC.HIN_NO, vbUnicode)        '品番（外部）
    SYUKA(Row, ColLK_SEQ_NO) = StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)   '上位ﾘﾝｸ用連番
                                                                    '品目マスタ読込み
    
    
'    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
    
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
                                                                    '出荷予定数
    SYUKA(Row, ColYOTEI_QTY) = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#,##0")
                                                                    '出荷実績数
    SYUKA(Row, ColFIX_QTY) = Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "#,##0")
                                                                    '検品マーク
    If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) = 0 Then
                                '未検品
        SYUKA(Row, ColKENPIN_MARK) = KENPIN_OFF
    Else
                                '検品済
        SYUKA(Row, ColKENPIN_MARK) = KENPIN_ON
    End If
            
    SYUKA(Row, ColDEN_DT) = Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" _
                            & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" _
                            & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 7, 2)
    
    If CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) > CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) Then
                                '未出庫　または　出庫中
        SYUKA(Row, ColSort_Mark) = Sort_MISYUKO
    Else
                                '出庫完了　で　未検品
        If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) = 0 Then
            SYUKA(Row, ColSort_Mark) = Sort_SYUKOSUMI
        Else
            SYUKA(Row, ColSort_Mark) = Sort_KENPIN
        End If
    End If
    
    If Len(Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))) = 0 Then
            SYUKA(Row, ColPrint) = ""
    Else
            SYUKA(Row, ColPrint) = "○"
    End If
    If Trim(StrConv(Y_SYUREC.INS_NOW, vbUnicode)) <> "" Then
        SYUKA(Row, ColIns_Date) = Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 1, 4) & "/" _
                                    & Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 5, 2) & "/" _
                                    & Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 7, 2) & " " _
                                    & Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 9, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 11, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.INS_NOW, vbUnicode), 13, 2)

    Else
        SYUKA(Row, ColIns_Date) = ""
    End If
    
    
    If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) <> "" Then
        SYUKA(Row, ColKENPIN_Date) = Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 1, 4) & "/" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 5, 2) & "/" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 7, 2) & " " _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 1, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 3, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 5, 2)

    Else
        SYUKA(Row, ColKENPIN_Date) = ""
    End If
    
    
    If Trim(StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode)) = "POS" Then
        SYUKA(Row, ColKENPIN_TANTO) = "出荷確認画面"
    Else
        Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                Exit Function
        End Select
        
        
        SYUKA(Row, ColKENPIN_TANTO) = StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode) & " " & StrConv(TANTOREC.TANTO_NAME, vbUnicode)
    End If
    
    
    SYUKA(Row, ColJGYOBU) = StrConv(Y_SYUREC.JGYOBU, vbUnicode)
    
    If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = KAN_KBN_FIN Then
        SYUKA(Row, ColKAN_YMD) = Mid(StrConv(Y_SYUREC.KAN_YMD, vbUnicode), 1, 4) & "/" _
                                & Mid(StrConv(Y_SYUREC.KAN_YMD, vbUnicode), 5, 2) & "/" _
                                & Mid(StrConv(Y_SYUREC.KAN_YMD, vbUnicode), 7, 2) & " " _
                                & Mid(StrConv(Y_SYUREC.KAN_HMS, vbUnicode), 1, 2) & ":" _
                                & Mid(StrConv(Y_SYUREC.KAN_HMS, vbUnicode), 3, 2) & ":" _
                                & Mid(StrConv(Y_SYUREC.KAN_HMS, vbUnicode), 5, 2)
    Else
        SYUKA(Row, ColKAN_YMD) = ""
    End If
    

    Grid_Set_Proc = False
End Function

Private Sub TDBGrid1_DblClick()

    If TDBGrid1.Bookmark = -1 Then
    Else
        '2011.08.03
        If HEAD_CLICK Then
            HEAD_CLICK = False
            Exit Sub
        End If
        '2011.08.03
    
        If KENPIN_Update_Proc() Then
            Unload Me
        End If
    End If
    '再表示
'    If List_Disp_Proc Then
'        Unload Me
'    End If


End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
'2011.08.03
'''    TDBGrid1.Bookmark = -1

Dim lngPFstRow  As Long
Dim vntBmk      As Variant
Dim intLeftCol  As Integer
Dim intCol      As Integer
Dim lngCFstRow  As Long
    
    
    
    If SYUKA.Count(1) < 1 Then
        Exit Sub
    End If
    
    HEAD_CLICK = True
        
    
    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
        
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        SYUKA.QuickSort Min_Row, SYUKA.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = SYUKA
        
        
        
        
        
        
        With TDBGrid1
              .SetFocus
              lngPFstRow = TDBGrid1.FirstRow
              vntBmk = .Bookmark
              intLeftCol = .LeftCol
              intCol = .Col
              .ReBind
              .Col = intCol
              .LeftCol = intLeftCol
              .Bookmark = vntBmk
              lngCFstRow = TDBGrid1.FirstRow
              TDBGrid1.Scroll 0, lngPFstRow - lngCFstRow
          End With
        
        TDBGrid1.Update
        TDBGrid1.MoveFirst

    End If
'2011.08.03




End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sts As Integer
Dim i   As Integer

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Select Case Index
        
        Case ptxSyuka_YY
            If Len(Trim(Text(ptxSyuka_YY).Text)) = 0 Then
            Else
            
                If Not IsNumeric(Text(ptxSyuka_YY).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
            End If
        Case ptxSyuka_MM
            If Len(Trim(Text(ptxSyuka_MM).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSyuka_MM).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
                Text(ptxSyuka_MM).Text = Format(CInt(Text(ptxSyuka_MM).Text), "00")
            End If
        Case ptxSyuka_DD
            If Len(Trim(Text(ptxSyuka_DD).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSyuka_DD).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
                Text(ptxSyuka_DD).Text = Format(CInt(Text(ptxSyuka_DD).Text), "00")
            End If
        
        
        
        
        '2012.11.13
        Case ptxE_Syuka_YY
            If Len(Trim(Text(ptxE_Syuka_YY).Text)) = 0 Then
            Else
            
                If Not IsNumeric(Text(ptxE_Syuka_YY).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
            End If
        Case ptxE_Syuka_MM
            If Len(Trim(Text(ptxE_Syuka_MM).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxE_Syuka_MM).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
                Text(ptxE_Syuka_MM).Text = Format(CInt(Text(ptxE_Syuka_MM).Text), "00")
            End If
        Case ptxE_Syuka_DD
            If Len(Trim(Text(ptxE_Syuka_DD).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxE_Syuka_DD).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
                Text(ptxE_Syuka_DD).Text = Format(CInt(Text(ptxE_Syuka_DD).Text), "00")
            End If
        
        
            If (Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text) > (Text(ptxE_Syuka_YY).Text & Text(ptxE_Syuka_MM).Text & Text(ptxE_Syuka_DD).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Exit Sub
            End If
        '2012.11.13
        
        
        Case ptxDATA_KBN
            If Trim(Text(Index).Text) = "" Or Text(Index).Text = "1" Or Text(Index).Text = "3" Or Text(Index).Text = "*" Then
            Else
                Beep
                MsgBox "入力した項目はエラーです。"
                Exit Sub
            End If
        
        Case ptxHAN_KBN
            If Trim(Text(Index).Text) = "" Or Text(Index).Text = "1" Or Text(Index).Text = "2" Then
            Else
                Beep
                MsgBox "入力した項目はエラーです。"
                Exit Sub
            End If
        
        Case ptxMUKE_CODE
            
            Text(Index).Text = StrConv(RTrim(Text(Index).Text), vbUpperCase)
            
            
            Call UniCode_Conv(K2_MTS.MUKE_CODE, Text(Index).Text)
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                        Beep
                        MsgBox "入力した項目はエラーです。(向け先コード)"
                        Exit Sub
                    End If
                                
                Case BtErrKeyNotFound
                                
                    Call UniCode_Conv(K3_MTS.SS_CODE, Text(Index).Text)
                                                        
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                    Select Case sts
                        Case BtNoErr
                                        
                        Case BtErrKeyNotFound
                            Beep
                            MsgBox "入力した項目はエラーです。(向け先コード)"
                            Exit Sub
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                            Unload Me
                    End Select

                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                    Unload Me
            End Select


            For i = 0 To Combo(pcmbMUKE_CODE).ListCount - 1 '向け先
    
                If Right(Combo(pcmbMUKE_CODE).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                    Combo(pcmbMUKE_CODE).ListIndex = i
                    Exit For
                End If
            
    
            Next

            Combo(pcmbMUKE_CODE).SetFocus
    End Select
    
'>>>>>>>>>>>>>>>>>> 2012.11.19
'    For i = Index + 1 To Text_Max
'        If Text(i).Visible And Text(i).Enabled And Text(i).TabStop Then
'            Text(i).SetFocus
'            Exit For
'        End If
'    Next i

    Call Tab_Ctrl(Shift)        '移動
'>>>>>>>>>>>>>>>>>> 2012.11.19


End Sub

Private Function KENPIN_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   検品済更新
'----------------------------------------------------------------------------
Dim sts As Integer
Dim ans As Integer
    
Dim com As Integer  '2016.09.29
    
    
    If TDBGrid1.Bookmark = -1 Then
        Exit Function
    End If
    
    
    KENPIN_Update_Proc = True
                                     'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
                                    '出荷予定の読み込み
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, SYUKA(TDBGrid1.Bookmark, ColJGYOBU))     '事業部
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, SYUKA(TDBGrid1.Bookmark, ColID_NO))   ' ID№
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                MsgBox "他端末で内容が変更されています。最新表示を行ってください。"
                KENPIN_Update_Proc = False
                GoTo Abort_Tran
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    KENPIN_Update_Proc = False
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定")
                GoTo Abort_Tran
        End Select

    Loop
    
    
    If Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode)) <> "" Then
        MsgBox "他端末で処理中です。当画面では処理できません。。"
        KENPIN_Update_Proc = False
        GoTo Abort_Tran
    End If
    
    
                                    
    If Inspe_F = 0 Then
        If StrConv(Y_SYUREC.SURYO, vbUnicode) <> StrConv(Y_SYUREC.JITU_SURYO, vbUnicode) Then
            MsgBox "出庫作業未完了です。検品処理を実行できません。"
            KENPIN_Update_Proc = False
            GoTo Abort_Tran
        
        End If
            
    
    Else
    
''        If Not IsNumeric(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) Then
''        Else
''            If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> 0 Then
''
''                If Not IsNumeric(StrConv(Y_SYUREC.KENPIN_SURYO, vbUnicode)) Then
''                Else
''                    If CLng(StrConv(Y_SYUREC.KENPIN_SURYO, vbUnicode)) <> 0 Then
''
''                        MsgBox "出庫作業中です。検品処理を実行できません。"
''                        KENPIN_Update_Proc = False
''                        GoTo Abort_Tran
''                    End If
''                End If
''            End If
''        End If
    End If
                                    
                                    
    If Trim(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) <> "" Then
        MsgBox "GLICS引渡し済です。当画面では処理できません。"
        KENPIN_Update_Proc = False
        GoTo Abort_Tran
    End If
                                    
    If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
    Else
        If StrConv(Y_SYUREC.G_KENPIN_F, vbUnicode) = "1" Then
        Else
    
            MsgBox "ｽｷｬﾅ検品処理済です。当画面では処理できません。"
            KENPIN_Update_Proc = False
            GoTo Abort_Tran
        End If
    End If
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   直送伝票のチェック  2016.09.29
    If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
        Call UniCode_Conv(K0_HTDrctId.IDNO, SYUKA(TDBGrid1.Bookmark, ColID_NO))   ' ID№
        sts = BTRV(BtOpGetEqual, HTDrctId_POS, HTDrctIdREC, Len(HTDrctIdREC), K0_HTDrctId, Len(K0_HTDrctId), 0)
        Select Case sts
            Case BtNoErr
                MsgBox "直送出荷分です。当画面では処理できません。"
                KENPIN_Update_Proc = False
                GoTo Abort_Tran
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "HTDrctId.dat")
                GoTo Abort_Tran
        End Select
    End If
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   直送伝票のチェック  2016.09.29
    
        
    
    
    
    
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
    
    If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
        '検品済にする
        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
        '2006.07.20 検品担当者出力追加
        Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "POS")
        Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, Format(Now, "HHMMSS"))
        Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "1")
        '予定数--＞実績数（ここには未出庫設定時しか来ない）
        If Inspe_F = 1 Then
            If Not IsNumeric(StrConv(Y_SYUREC.KENPIN_SURYO, vbUnicode)) Then
                If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) = 0 Then
                    Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(Y_SYUREC.SURYO, vbUnicode))
                End If
            
            End If
            
            
            
        
        End If
    Else
        '未検品する
        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
        '2006.07.20 検品担当者出力追加
        Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")
        Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")
        Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "")
    
        If Inspe_F = 1 Then
            
            If IsNumeric(StrConv(Y_SYUREC.KENPIN_SURYO, vbUnicode)) Then
                        
                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(Y_SYUREC.KENPIN_SURYO, vbUnicode))
            End If
            Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, "")
        End If
    
    End If
                                    
                                    '出荷予定書込み
    Do
        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    KENPIN_Update_Proc = False
                    GoTo Abort_Tran
                End If
        
            Case Else
                Call File_Error(sts, BtOpUpdate, "出荷予定")
                GoTo Abort_Tran
        End Select
    Loop
                                        
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 直送検品処理    2016.09.29
    If Inspe_Choku_F = 1 Then
                                        
                                        
        Call UniCode_Conv(K0_HTIdDelv.IDNO, SYUKA(TDBGrid1.Bookmark, ColID_NO))         ' ID№
        Call UniCode_Conv(K0_HTIdDelv.DelvNo, "")                                       ' 送り状№
        
        com = BtOpGetGreater
        
        Do
            DoEvents
            Do
            
                sts = BTRV(com + BtSNoWait, HTIdDelv_POS, HTIdDelvREC, Len(HTIdDelvREC), K0_HTIdDelv, Len(K0_HTIdDelv), 0)
                Select Case sts
                    Case BtNoErr
                        If Trim(StrConv(HTIdDelvREC.IDNO, vbUnicode)) <> Trim(SYUKA(TDBGrid1.Bookmark, ColID_NO)) Then
                            sts = BtErrEOF
                        End If
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    
                        ans = MsgBox("他端末でデータ使用中です。<HTIdDelv.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            KENPIN_Update_Proc = False
                            GoTo Abort_Tran
                        End If
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "HTIdDelv.DAT")
                        GoTo Abort_Tran
                End Select
        
            Loop
                                            
            If sts = BtErrEOF Then
                Exit Do
            End If
                                                
                                            
            Do
            
                sts = BTRV(BtOpDelete, HTIdDelv_POS, HTIdDelvREC, Len(HTIdDelvREC), K0_HTIdDelv, Len(K0_HTIdDelv), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    
                        ans = MsgBox("他端末でデータ使用中です。<HTIdDelv.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            KENPIN_Update_Proc = False
                            GoTo Abort_Tran
                        End If
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpDelete, "HTIdDelv.DAT")
                        GoTo Abort_Tran
                End Select
            Loop
        
            com = BtOpGetNext
        
        Loop
                                        
    End If
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 直送検品処理    2016.09.29
                                        
End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
    
        SYUKA(TDBGrid1.Bookmark, ColKENPIN_MARK) = KENPIN_OFF
        SYUKA(TDBGrid1.Bookmark, ColKENPIN_Date) = ""
        SYUKA(TDBGrid1.Bookmark, ColKENPIN_TANTO) = ""
                                        
    Else
                                        
        SYUKA(TDBGrid1.Bookmark, ColKENPIN_MARK) = KENPIN_ON
        
        SYUKA(TDBGrid1.Bookmark, ColKENPIN_Date) = Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 1, 4) & "/" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 5, 2) & "/" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 7, 2) & " " _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 1, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 3, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 5, 2)
        
        If Trim(StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode)) = "POS" Then
            SYUKA(TDBGrid1.Bookmark, ColKENPIN_TANTO) = "出荷確認画面"
        Else
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode))
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            End Select
            
            
            SYUKA(TDBGrid1.Bookmark, ColKENPIN_TANTO) = StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode) & " " & StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                                            
        End If
                                        
    End If
    
    SYUKA(TDBGrid1.Bookmark, ColFIX_QTY) = Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "#,##0")
    
        
    
    
    Set TDBGrid1.Array = SYUKA
    TDBGrid1.Refresh
    
    TDBGrid1.Update

    
    If IsNumeric(Text(ptxDEN_MAISU_JI).Text) Then
        If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
            Text(ptxDEN_MAISU_JI).Text = Format(CInt(Text(ptxDEN_MAISU_JI).Text) - 1, "#,##0")
        Else
            Text(ptxDEN_MAISU_JI).Text = Format(CInt(Text(ptxDEN_MAISU_JI).Text) + 1, "#,##0")
        End If
    End If
    
    KENPIN_Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function

Private Sub Text_LostFocus(Index As Integer)
    
        
    Select Case Index
        Case ptxMUKE_CODE
            Text(Index).Text = StrConv(RTrim(Text(Index).Text), vbUpperCase)
    End Select

End Sub
