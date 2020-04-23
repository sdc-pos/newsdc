VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PI000901 
   Caption         =   "発注検討処理"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
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
   ScaleHeight     =   10305
   ScaleWidth      =   15240
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   4095
      TabIndex        =   2
      Top             =   120
      Width           =   3690
      Begin VB.OptionButton Option1 
         Caption         =   "受注・未受注分"
         Height          =   255
         Index           =   1
         Left            =   1470
         TabIndex        =   3
         Top             =   240
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         Caption         =   "受注分"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   32
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "表示条件選択"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      TabIndex        =   26
      Top             =   720
      Width           =   15135
      Begin VB.TextBox Text1 
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
         IMEMode         =   3  'ｵﾌ固定
         Index           =   3
         Left            =   2205
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "9999/99/99"
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox Text1 
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
         IMEMode         =   3  'ｵﾌ固定
         Index           =   6
         Left            =   8295
         MaxLength       =   20
         TabIndex        =   9
         Top             =   360
         Width           =   1485
      End
      Begin VB.TextBox Text1 
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
         Index           =   7
         Left            =   9765
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text1 
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
         IMEMode         =   3  'ｵﾌ固定
         Index           =   2
         Left            =   840
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "9999/99/99"
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox Text1 
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
         IMEMode         =   3  'ｵﾌ固定
         Index           =   4
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "1234567890123"
         Top             =   360
         Width           =   1485
      End
      Begin VB.TextBox Text1 
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
         Index           =   5
         Left            =   5670
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   12810
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   11
         Top             =   360
         Width           =   2145
      End
      Begin VB.TextBox Text1 
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
         IMEMode         =   3  'ｵﾌ固定
         Index           =   8
         Left            =   12180
         MaxLength       =   5
         TabIndex        =   10
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         Caption         =   "子品番"
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
         Index           =   2
         Left            =   7560
         TabIndex        =   31
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         Caption         =   "～"
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
         Left            =   1995
         TabIndex        =   30
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         Caption         =   "受注日"
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
         Left            =   210
         TabIndex        =   29
         Top             =   360
         Width           =   645
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
         Index           =   7
         Left            =   3465
         TabIndex        =   28
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         Caption         =   "仕入先"
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
         Left            =   11550
         TabIndex        =   27
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1140
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "注文書"
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
      Left            =   2760
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1920
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1080
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "リスト"
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
      Left            =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7335
      Left            =   105
      TabIndex        =   12
      Top             =   1680
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   12938
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "事"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "国"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "子部品"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "子部品名"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "必要数"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "仕入単価"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "棚番"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "在庫数"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "注文残"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "引当数"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "不足数"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "注文数"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "発注ﾛｯﾄ"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "仕入先CD"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "仕入先名"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "LT"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "納期予定日"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   17
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=17"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=476"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=370"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=476"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=370"
      Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2090"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1984"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=8192"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=3122"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3016"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=8192"
      Splits(0)._ColumnProps(23)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1402"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1296"
      Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=8194"
      Splits(0)._ColumnProps(29)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1773"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1667"
      Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=8194"
      Splits(0)._ColumnProps(35)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=2223"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2117"
      Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=8192"
      Splits(0)._ColumnProps(41)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=1402"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1296"
      Splits(0)._ColumnProps(46)=   "Column(7)._ColStyle=8194"
      Splits(0)._ColumnProps(47)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(49)=   "Column(8).Width=1402"
      Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1296"
      Splits(0)._ColumnProps(52)=   "Column(8)._ColStyle=8194"
      Splits(0)._ColumnProps(53)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(55)=   "Column(9).Width=1402"
      Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1296"
      Splits(0)._ColumnProps(58)=   "Column(9)._ColStyle=8194"
      Splits(0)._ColumnProps(59)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(61)=   "Column(10).Width=1402"
      Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=1296"
      Splits(0)._ColumnProps(64)=   "Column(10)._ColStyle=8194"
      Splits(0)._ColumnProps(65)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(67)=   "Column(11).Width=1402"
      Splits(0)._ColumnProps(68)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(11)._WidthInPix=1296"
      Splits(0)._ColumnProps(70)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(71)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(72)=   "Column(12).Width=1561"
      Splits(0)._ColumnProps(73)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(12)._WidthInPix=1455"
      Splits(0)._ColumnProps(75)=   "Column(12)._ColStyle=8194"
      Splits(0)._ColumnProps(76)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(77)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(78)=   "Column(13).Width=1693"
      Splits(0)._ColumnProps(79)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(13)._WidthInPix=1588"
      Splits(0)._ColumnProps(81)=   "Column(13)._ColStyle=8196"
      Splits(0)._ColumnProps(82)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(83)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(84)=   "Column(14).Width=2090"
      Splits(0)._ColumnProps(85)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(14)._WidthInPix=1984"
      Splits(0)._ColumnProps(87)=   "Column(14)._ColStyle=8196"
      Splits(0)._ColumnProps(88)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(89)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(90)=   "Column(15).Width=1217"
      Splits(0)._ColumnProps(91)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(92)=   "Column(15)._WidthInPix=1111"
      Splits(0)._ColumnProps(93)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(94)=   "Column(15).AllowFocus=0"
      Splits(0)._ColumnProps(95)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(96)=   "Column(16).Width=2461"
      Splits(0)._ColumnProps(97)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(98)=   "Column(16)._WidthInPix=2355"
      Splits(0)._ColumnProps(99)=   "Column(16).Order=17"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.bgcolor=&HFFFF80&,.bold=0,.fontsize=975"
      _StyleDefs(25)  =   ":id=43,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=43,.fontname=ＭＳ ゴシック"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9,.bgcolor=&HD3FEA5&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10,.bgcolor=&HFFFFD2&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=98,.parent=43"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=44"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=45"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=47"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=94,.parent=43"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=91,.parent=44"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=92,.parent=45"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=93,.parent=47"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.locked=-1,.bold=0"
      _StyleDefs(47)  =   ":id=58,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(48)  =   ":id=58,.fontname=ＭＳ ゴシック"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.locked=-1,.bold=0"
      _StyleDefs(53)  =   ":id=62,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(54)  =   ":id=62,.fontname=ＭＳ ゴシック"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=16,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=44"
      _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=45"
      _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=47"
      _StyleDefs(62)  =   "Splits(0).Columns(5).Style:id=28,.parent=43,.alignment=1,.locked=-1,.bold=0"
      _StyleDefs(63)  =   ":id=28,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(64)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(65)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=44"
      _StyleDefs(66)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=45"
      _StyleDefs(67)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=47"
      _StyleDefs(68)  =   "Splits(0).Columns(6).Style:id=66,.parent=43,.alignment=0,.locked=-1,.bold=0"
      _StyleDefs(69)  =   ":id=66,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(70)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=32,.parent=43,.alignment=1,.locked=-1,.bold=0"
      _StyleDefs(75)  =   ":id=32,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(76)  =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(77)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=44"
      _StyleDefs(78)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=45"
      _StyleDefs(79)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=47"
      _StyleDefs(80)  =   "Splits(0).Columns(8).Style:id=70,.parent=43,.alignment=1,.locked=-1,.bold=0"
      _StyleDefs(81)  =   ":id=70,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(82)  =   ":id=70,.fontname=ＭＳ ゴシック"
      _StyleDefs(83)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(9).Style:id=74,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(87)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=47"
      _StyleDefs(90)  =   "Splits(0).Columns(10).Style:id=20,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(91)  =   "Splits(0).Columns(10).HeadingStyle:id=17,.parent=44"
      _StyleDefs(92)  =   "Splits(0).Columns(10).FooterStyle:id=18,.parent=45"
      _StyleDefs(93)  =   "Splits(0).Columns(10).EditorStyle:id=19,.parent=47"
      _StyleDefs(94)  =   "Splits(0).Columns(11).Style:id=24,.parent=43,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(95)  =   "Splits(0).Columns(11).HeadingStyle:id=21,.parent=44"
      _StyleDefs(96)  =   "Splits(0).Columns(11).FooterStyle:id=22,.parent=45"
      _StyleDefs(97)  =   "Splits(0).Columns(11).EditorStyle:id=23,.parent=47"
      _StyleDefs(98)  =   "Splits(0).Columns(12).Style:id=78,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(99)  =   "Splits(0).Columns(12).HeadingStyle:id=75,.parent=44"
      _StyleDefs(100) =   "Splits(0).Columns(12).FooterStyle:id=76,.parent=45"
      _StyleDefs(101) =   "Splits(0).Columns(12).EditorStyle:id=77,.parent=47"
      _StyleDefs(102) =   "Splits(0).Columns(13).Style:id=82,.parent=43,.locked=-1"
      _StyleDefs(103) =   "Splits(0).Columns(13).HeadingStyle:id=79,.parent=44"
      _StyleDefs(104) =   "Splits(0).Columns(13).FooterStyle:id=80,.parent=45"
      _StyleDefs(105) =   "Splits(0).Columns(13).EditorStyle:id=81,.parent=47"
      _StyleDefs(106) =   "Splits(0).Columns(14).Style:id=86,.parent=43,.locked=-1"
      _StyleDefs(107) =   "Splits(0).Columns(14).HeadingStyle:id=83,.parent=44"
      _StyleDefs(108) =   "Splits(0).Columns(14).FooterStyle:id=84,.parent=45"
      _StyleDefs(109) =   "Splits(0).Columns(14).EditorStyle:id=85,.parent=47"
      _StyleDefs(110) =   "Splits(0).Columns(15).Style:id=102,.parent=43"
      _StyleDefs(111) =   "Splits(0).Columns(15).HeadingStyle:id=99,.parent=44"
      _StyleDefs(112) =   "Splits(0).Columns(15).FooterStyle:id=100,.parent=45"
      _StyleDefs(113) =   "Splits(0).Columns(15).EditorStyle:id=101,.parent=47"
      _StyleDefs(114) =   "Splits(0).Columns(16).Style:id=90,.parent=43,.bgcolor=&HFFFFFF&"
      _StyleDefs(115) =   "Splits(0).Columns(16).HeadingStyle:id=87,.parent=44"
      _StyleDefs(116) =   "Splits(0).Columns(16).FooterStyle:id=88,.parent=45"
      _StyleDefs(117) =   "Splits(0).Columns(16).EditorStyle:id=89,.parent=47"
      _StyleDefs(118) =   "Named:id=33:Normal"
      _StyleDefs(119) =   ":id=33,.parent=0"
      _StyleDefs(120) =   "Named:id=34:Heading"
      _StyleDefs(121) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(122) =   ":id=34,.wraptext=-1"
      _StyleDefs(123) =   "Named:id=35:Footing"
      _StyleDefs(124) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(125) =   "Named:id=36:Selected"
      _StyleDefs(126) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(127) =   "Named:id=37:Caption"
      _StyleDefs(128) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(129) =   "Named:id=38:HighlightRow"
      _StyleDefs(130) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(131) =   "Named:id=39:EvenRow"
      _StyleDefs(132) =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&"
      _StyleDefs(133) =   "Named:id=40:OddRow"
      _StyleDefs(134) =   ":id=40,.parent=33"
      _StyleDefs(135) =   "Named:id=41:RecordSelector"
      _StyleDefs(136) =   ":id=41,.parent=34"
      _StyleDefs(137) =   "Named:id=42:FilterBar"
      _StyleDefs(138) =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
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
      Left            =   10440
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Left            =   9600
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Left            =   8760
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   8
      Left            =   7920
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "デ-タ"
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
      Left            =   6615
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5775
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4935
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "検 索"
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
      Left            =   4095
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "担当者"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   25
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "PI000901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    
'テキスト用添字
Private Const ptxTANTO_CODE% = 0            '担当者ｺｰﾄﾞ
Private Const ptxTANTO_NAME% = 1            '担当者名称

Private Const ptxS_ORDER_DT% = 2            '受注日 From
Private Const ptxE_ORDER_DT% = 3            '受注日 To

Private Const ptxO_HIN_GAI% = 4             '親品番
Private Const ptxO_HIN_NAME% = 5            '親品番名称

Private Const ptxK_HIN_GAI% = 6             '子品番
Private Const ptxK_HIN_NAME% = 7            '子品番名称

Private Const ptxORDER_CODE% = 8            '仕入先




'コンボ用添字
Private Const pcmbORDER% = 0                '注文先



'ｺﾏﾝﾄﾞﾎﾞﾀﾝ
Private Const pcomList% = 0                 'ﾘｽﾄ出力
Private Const pcomORDER% = 3                '注文書出力

'ｵﾌﾟｼｮﾝﾎﾞﾀﾝ
Private Const poptORDER% = 0                '受注分
Private Const poptSHIJI% = 1                '指示分



Private Sort_Tbl(colJGYOBU To colY_NOUKI_DT) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
Private Tbl_Set_F   As Boolean
                                            
                                            
Private O_SEL_JGOYBU    As String * 1
Private O_SEL_NAIGAI    As String * 1
                                            
                                            
Private K_SEL_JGOYBU    As String * 1
Private K_SEL_NAIGAI    As String * 1
                                            
Private NOUNYU          As String * 5


Private P_SHKENTO_OSAKA_DATA    As String   '発注検討用ﾃﾞｰﾀ


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PI000901.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000901)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000901)


    PI000901.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        Case ptxTANTO_CODE      '担当者
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(Mode).Text)

            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTANTO_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTANTO_NAME).Text = ""
            
                    MsgBox "入力した項目はエラーです。(担当者エラー)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            
            End Select
        
        Case ptxS_ORDER_DT      '注文日 From
            
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsDate(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。(日付エラー)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If
        
        Case ptxE_ORDER_DT      '注文日 To
            
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsDate(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。(日付エラー)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If
        
        
        Case ptxO_HIN_GAI       '親品番
        
            If Trim(Text1(Mode).Text) = "" Then
                Text1(ptxO_HIN_NAME).Text = ""
            Else
                sts = Item_Read_Proc(Trim(Text1(Mode).Text))
                Select Case sts
                
                    Case BtNoErr
                        O_SEL_JGOYBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                        O_SEL_NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                        Text1(ptxO_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                                                            
                    Case BtErrKeyNotFound
                        
                        Text1(ptxO_HIN_NAME).Text = ""
                        MsgBox "入力した項目はエラーです。(親品番エラー)"
                        Text1(Mode).SetFocus
                        Exit Function
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                
                End Select
            End If
        
        
        
        
        Case ptxK_HIN_GAI       '子品番
        
        
            If Trim(Text1(Mode).Text) = "" Then
                Text1(ptxK_HIN_NAME).Text = ""
            Else
                sts = Item_Read_Proc(Trim(Text1(Mode).Text))
                Select Case sts
                
                    Case BtNoErr
                        K_SEL_JGOYBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                        K_SEL_NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                        Text1(ptxK_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                                                            
                    Case BtErrKeyNotFound
                        
                        Text1(ptxK_HIN_NAME).Text = ""
                        MsgBox "入力した項目はエラーです。(子品番エラー)"
                        Text1(Mode).SetFocus
                        Exit Function
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                
                End Select
            End If
        
        
        
        Case ptxORDER_CODE   '仕入先
            
            If Trim(Text1(Mode).Text) = "" Then
               Combo1(pcmbORDER).ListIndex = -1
            Else
            
               Combo1(pcmbORDER).ListIndex = -1
               For i = 0 To Combo1(pcmbORDER).ListCount - 1
                   If Text1(Mode).Text = Trim(Right(Combo1(pcmbORDER).List(i), 5)) Then
                       Combo1(pcmbORDER).ListIndex = i
                       Exit For
                   End If
               
               Next i
        
               If i > Combo1(pcmbORDER).ListCount - 1 Then
                   MsgBox "入力した項目はエラーです。(仕入先コード)"
                   Text1(Mode).SetFocus
                   Exit Function
               End If
            End If
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  資材注文ﾃﾞｰﾀ更新
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer

Dim ORDERNO         As Integer

Dim i               As Integer
Dim j               As Integer



    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    Set TDBGrid1.Array = SHORDER
    TDBGrid1.Refresh

    TDBGrid1.Update

                                        
    For i = 1 To SHORDER.UpperBound(1)
    
        If IsNumeric(SHORDER(i, colORDER_QTY)) Then
            If CInt(SHORDER(i, colORDER_QTY)) > 0 Then

                '管理ファイルより資材注文番号の獲得
                Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
                
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound
                        
                            If P_KANRI_MAKE_Proc() Then
                                GoTo Abort_Tran
                            End If
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Update_Proc = True
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "管理マスタ")
                            GoTo Abort_Tran
                    
                    End Select
                
                
                Loop
            
                '注文№＋１
                If CLng(StrConv(P_KANRIREC.ORDER_NO, vbUnicode)) = 99999 Then
                    Call UniCode_Conv(P_KANRIREC.ORDER_NO, "00001")
                Else
                    Call UniCode_Conv(P_KANRIREC.ORDER_NO, Format(CLng(StrConv(P_KANRIREC.ORDER_NO, vbUnicode)) + 1, "00000"))
                End If
            
            
                Do
                    
                    DoEvents
                    
                    sts = BTRV(BtOpUpdate, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts Then
                                    Call File_Error(sts, BtOpUnlock, "管理マスタ")
                                End If
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "管理マスタ")
                            GoTo Abort_Tran
                    End Select
                Loop
        
                ORDERNO = CLng(StrConv(P_KANRIREC.ORDER_NO, vbUnicode))
    
    
    
    
                                        
                '注文№
                Call UniCode_Conv(P_SHORDER_REC.ORDER_NO, Format(ORDERNO, "00000"))
                '注文日
                Call UniCode_Conv(P_SHORDER_REC.ORDER_DT, Format(Now, "YYYYMMDD"))
                '発行日時
                Call UniCode_Conv(P_SHORDER_REC.Print_datetime, "")
                '担当者ｺｰﾄﾞ
                Call UniCode_Conv(P_SHORDER_REC.TANTO_CODE, Text1(ptxTANTO_CODE).Text)
                '事業部
                Call UniCode_Conv(P_SHORDER_REC.JGYOBU, SHORDER(i, colJGYOBU))
                '国内外
                Call UniCode_Conv(P_SHORDER_REC.NAIGAI, SHORDER(i, colNAIGAI))
                '品番
                Call UniCode_Conv(P_SHORDER_REC.HIN_GAI, SHORDER(i, colHIN_GAI))
                '注文先
                Call UniCode_Conv(P_SHORDER_REC.ORDER_CODE, SHORDER(i, colORDER_CODE))
                '納入先
                Call UniCode_Conv(P_SHORDER_REC.DELI_CODE, NOUNYU)
                '注文数
                Call UniCode_Conv(P_SHORDER_REC.ORDER_QTY, Format(CDbl(SHORDER(i, colORDER_QTY)), "00000000.00"))
                '予定納期
                Call UniCode_Conv(P_SHORDER_REC.Y_NOUKI_DT, Format(SHORDER(i, colY_NOUKI_DT), "YYYYMMDD"))
                '発注単価
                Call UniCode_Conv(P_SHORDER_REC.TANKA, Format(CDbl(SHORDER(i, colTANKA)), "00000000.00"))
                '発注ﾛｯﾄ
                Call UniCode_Conv(P_SHORDER_REC.LOT, Format(CDbl(SHORDER(i, colLOT)), "00000000"))
                    
                Call UniCode_Conv(P_SHORDER_REC.KAN_F, P_KAN_OFF)                       '完了ﾌﾗｸﾞ
                    
                Call UniCode_Conv(P_SHORDER_REC.KAN_DT, "")                             '完了日
                    
                Call UniCode_Conv(P_SHORDER_REC.BUNNOU_CNT, "00")                       '受入回数
                    
                Call UniCode_Conv(P_SHORDER_REC.UKEIRE_QTY, "00000000.00")              '受入数
                
                Call UniCode_Conv(P_SHORDER_REC.CANCEL_F, P_CANCEL_OFF)                 'ｷｬﾝｾﾙﾌﾗｸﾞ
                    
                Call UniCode_Conv(P_SHORDER_REC.CANCEL_DATETIME, "")                    'ｷｬﾝｾﾙ日時
                
                Call UniCode_Conv(P_SHORDER_REC.PRINT_F, P_PRINT_OFF)                   '印刷ﾌﾗｸﾞ
                
                Call UniCode_Conv(P_SHORDER_REC.WS_NO, WS_NO)                           '入力端末
                
                
                '品目ﾏｽﾀ読込み
                Call UniCode_Conv(K0_ITEM.JGYOBU, SHORDER(i, colJGYOBU))
                Call UniCode_Conv(K0_ITEM.NAIGAI, SHORDER(i, colNAIGAI))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, SHORDER(i, colHIN_GAI))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "品目マスタが他端末で変更されました。更新処理を中止します。"
                        GoTo Abort_Tran
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目ﾏｽﾀ")
                        GoTo Abort_Tran
                End Select
                '仕入区分
                Call UniCode_Conv(P_SHORDER_REC.G_SHIIRE_KBN, StrConv(ITEMREC.G_SHIIRE_KBN, vbUnicode))
                '収支単位
                Call UniCode_Conv(P_SHORDER_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                
                
                '受払先ﾏｽﾀ読込み
                Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, SHORDER(i, colORDER_CODE))
                sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                    
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "受払先マスタが他端末で変更されました。更新処理を中止します。"
                        GoTo Abort_Tran
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "受払先ﾏｽﾀ")
                        GoTo Abort_Tran
                End Select
            
                                                                                            '取引先区分
                Call UniCode_Conv(P_SHORDER_REC.TORI_KBN, StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))
            
                Call UniCode_Conv(P_SHORDER_REC.FILLER, "")
                                                                                            '更新日時
                Call UniCode_Conv(P_SHORDER_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                Do
                    
                    DoEvents
                    
                    sts = BTRV(BtOpInsert, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                If com = BtOpUpdate Then
                                    sts = BTRV(BtOpUnlock, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                                    If sts Then
                                        Call File_Error(sts, BtOpUnlock, "資材注文ﾃﾞｰﾀ")
                                    End If
                                End If
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                            GoTo Abort_Tran
                    End Select
                
                Loop
                
                
                '---------------------------------------------------    '品目マスタ更新
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                
                Do
                
                    sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        
                    Select Case sts
                        Case BtNoErr
                        
                            Exit Do
                        
                        Case BtErrKeyNotFound
                        
                            MsgBox "品目マスタが削除されています。更新を中止します。"
                            GoTo Abort_Tran
                        
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                GoTo Abort_Tran
                            End If
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                            GoTo Abort_Tran
                    End Select
            
                Loop
                
                For j = 0 To 2
                
                    If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(j).CODE, vbUnicode)) = Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) Then
                        Exit For
                    End If
                Next j
                
                
                If j <= 2 Then
                    '前回注文日
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_DT, StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode))
                    '前回注文数
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_QTY, StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
                End If
            
            
                Do
                    
                    DoEvents
                    
                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                If sts Then
                                    Call File_Error(sts, BtOpUnlock, "品目ﾏｽﾀ")
                                End If
                            End If
                            GoTo Abort_Tran
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "品目ﾏｽﾀ")
                            GoTo Abort_Tran
                    End Select
                
                Loop
            End If
        End If
    Next i
        

End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Call Input_UnLock
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbORDER          '注文先
            Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).Text, 5))
    End Select
    
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbORDER          '注文先
            Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).Text, 5))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim rpt         As New PI00090F1
Dim f           As New PI000902

Dim sts         As Integer

    Select Case Index
        
        
        Case pcomList           'ﾘｽﾄ出力
        
                
            
            
            If Tbl_Set_F Then
            
                If Grid_Error_Check_Proc() Then
                    TDBGrid1.SetFocus
                    Exit Sub
                End If
                
                Beep
                ans = MsgBox("リスト出力しますか？", vbYesNo + vbQuestion, "確認入力")
                If ans = vbYes Then
                    If Input_Update_Proc() Then
                        Unload Me
                    End If
                    
                    If List_Print_Proc() Then
                        Unload Me
                    End If
                
                End If
            Else
                MsgBox "印刷対象を選択して下さい！！"
                Exit Sub
            End If
            
            
            Text1(ptxTANTO_CODE).SetFocus
        
        Case pcomORDER          '注文書印刷
        
            If Tbl_Set_F Then
                
                If Grid_Error_Check_Proc() Then
                    TDBGrid1.SetFocus
                    Exit Sub
                End If
                
                Beep
                ans = MsgBox("注文書出力しますか？", vbYesNo + vbQuestion, "確認入力")
                If ans = vbYes Then
                    If Update_Proc() Then
                        Unload Me
                    End If
                    
                    If Print_Proc() Then
                        Unload Me
                    End If
                
                
                    If Data_Make_Proc() Then
                        Unload Me
                    End If
                    
                    
                    
                    
                    If List_Disp_Proc(0) Then
                        Unload Me
                    End If
                
                
                
                
                End If
            Else
                MsgBox "印刷対象を選択して下さい！！"
                Exit Sub
            End If
            
            Text1(ptxTANTO_CODE).SetFocus
        
        
        
        
        Case P_CMD_Upd          '更新
            
        
        Case P_CMD_DEL          '削除
    
        Case P_CMD_DSP          '検索/表示
        
        
            For i = ptxTANTO_CODE To ptxORDER_CODE
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            Next i
        
            Beep
            ans = MsgBox("検索しますか？（入力内容は初期化されます。）", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                
                If Data_Make_Proc() Then
                    Unload Me
                End If
                
                
                
                
                If List_Disp_Proc(0) Then
                    Unload Me
                End If
            
            End If
            
            Text1(ptxTANTO_CODE).SetFocus
        
        
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        
            If Tbl_Set_F Then
                
                
                If Grid_Error_Check_Proc() Then
                    TDBGrid1.SetFocus
                    Exit Sub
                End If
                
                Beep
                ans = MsgBox("データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
                If ans = vbYes Then
                    
                    If Input_Update_Proc() Then
                        Unload Me
                    End If
                    
                    
                    
                    If Data_Output_Proc() Then
                        Unload Me
                    End If
                
                End If
            Else
                MsgBox "出力対象を選択して下さい！！"
                Exit Sub
            End If
            
            Text1(ptxTANTO_CODE).SetFocus
        
        
        
        Case P_CMD_PRT                      '印刷
 
            
            
        Case P_CMD_End                      '終了
    
                        
''            If Tbl_Set_F Then
''                ans = MsgBox("印刷処理が行われていません。入力処理に戻りますか？", vbYesNo + vbQuestion, "確認入力")
''                If ans = vbYes Then
''                Else
''                    Unload Me
''                End If
''            Else
                Unload Me
''            End If
    End Select

End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
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

Dim sBuffer As String

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        WS_NO = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        WS_NO = "???"
    End If


                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                'ﾃﾞｰﾀ出力用ファイル
    If GetIni("FILE", "P_SHKENTO_OSAKA_DATA", "SYS", c) Then
        Beep
        MsgBox "発注検討用ファイル名[P_SHKENTO_OSAKA_DATA]の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    P_SHKENTO_OSAKA_DATA = RTrim(c)
                                
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                
                                
                                '備考１取り込み '2007.07.20
    If GetIni(App.EXEName, "BIKOU_1", "P_SYS", c) Then
        pubBikou_1 = ""
    Else
        pubBikou_1 = Trim(c)
    End If
                                '備考２取り込み '2007.07.20
    If GetIni(App.EXEName, "BIKOU_2", "P_SYS", c) Then
        pubBikou_2 = ""
    Else
        pubBikou_2 = Trim(c)
    End If
                                '備考３取り込み '2007.07.20
    If GetIni(App.EXEName, "BIKOU_3", "P_SYS", c) Then
        pubBikou_3 = ""
    Else
        pubBikou_3 = Trim(c)
    End If
                                
                                
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材注文ﾃﾞｰﾀＯＰＥＮ
    If P_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材注文ﾃﾞｰﾀＯＰＥＮ(別ﾎﾟｲﾝﾀｰ)
    If wP_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '商品化指図ﾃﾞｰﾀ(親)ＯＰＥＮ
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指図ﾃﾞｰﾀ(子)ＯＰＥＮ
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
    
    
    Load PI000901
    
    
    
    '管理マスタの読み込み
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)

    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            If P_KANRI_MAKE_Proc() Then
                Unload Me
            End If
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ")
            Unload Me
    End Select
        
    
    
    
    '注文先
    If Ukeharai_Set_Proc(pcmbORDER) Then
        Unload Me
    End If
    
    
    '画面初期設定
    If Init_Proc() Then
        Unload Me
    End If

    Tbl_Set_F = False

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
    
    
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
                                            '資材注文ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材注文ﾃﾞｰﾀ")
        End If
    End If
                                            '資材注文ﾃﾞｰﾀＣＬＯＳＥ（別ﾎﾟｲﾝﾀｰ）
    sts = BTRV(BtOpClose, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), K2_wP_SHORDER, Len(K2_wP_SHORDER), 2)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材注文ﾃﾞｰﾀ")
        End If
    End If
                                            '在庫ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫ﾃﾞｰﾀ")
        End If
    End If
    
    
                                            '商品化指図ﾃﾞｰﾀ(親)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図ﾃﾞｰﾀ(親)")
        End If
    End If
                                            '商品化指図ﾃﾞｰﾀ(子)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図ﾃﾞｰﾀ(子)")
        End If
    End If
    
                                            '発注検討　大阪向けＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "発注検討　大阪向け")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI000901 = Nothing
    Set PI000902 = Nothing

    End
End Sub
Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)


    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
                    
        Select Case ColIndex
                    
            Case colFUSOKU_QTY
                
                If Sort_Tbl(ColIndex) = 0 Then
                
                    If List_Disp_Proc(1) Then
                        Unload Me
                    End If
                Else
                
                    If List_Disp_Proc(2) Then
                        Unload Me
                    End If
                
                End If
            Case Else
        
                SHORDER.QuickSort Min_Row, SHORDER.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
                
                Set TDBGrid1.Array = SHORDER
                
                TDBGrid1.ReBind
                TDBGrid1.Update
                TDBGrid1.MoveFirst
        End Select

    End If



End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Not Tbl_Set_F Then
        Exit Sub
    End If
        
        
    Select Case LastCol
    
        Case colORDER_QTY
            Set TDBGrid1.Array = SHORDER
            TDBGrid1.Refresh
            TDBGrid1.Update
        
            If Trim(SHORDER(LastRow, LastCol)) = "" Then
                SHORDER(LastRow, colY_NOUKI_DT) = ""
            
                
                Set TDBGrid1.Array = SHORDER
                TDBGrid1.Refresh
        
                TDBGrid1.Update
            
            
            Else
            
                If IsNumeric(SHORDER(LastRow, LastCol)) Then
                    
                    
                    If SHORDER(LastRow, colY_NOUKI_DT) <= 0 Then
                        MsgBox "注文数は１以上を入力して下さい """
                
                        SHORDER(LastRow, LastCol) = ""
                        Set TDBGrid1.Array = SHORDER
                        TDBGrid1.Refresh
                
                        TDBGrid1.Update
                    Else
                    
                    
                        SHORDER(LastRow, colY_NOUKI_DT) = Format(DateAdd("d", CDbl(SHORDER(LastRow, colLT)), Format(Now, "YYYY/MM/DD")), "YYYY/MM/DD")
                        Set TDBGrid1.Array = SHORDER
                        TDBGrid1.Refresh
                
                        TDBGrid1.Update
                    End If
                
                Else
                    MsgBox "注文数は数値を入力して下さい"
            
                    SHORDER(LastRow, LastCol) = ""
                    Set TDBGrid1.Array = SHORDER
                    TDBGrid1.Refresh
            
                    TDBGrid1.Update
            
            
                End If
            End If
        
        Case colY_NOUKI_DT
            
            If Trim(SHORDER(LastRow, LastCol)) = "" Then
            Else
                If Not IsDate(SHORDER(LastRow, LastCol)) Then
                    MsgBox "日付(YYYY/MM/DD)を入力して下さい"
                    SHORDER(LastRow, LastCol) = ""
                    Set TDBGrid1.Array = SHORDER
                    TDBGrid1.Refresh
            
                    TDBGrid1.Update
                End If
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
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動
End Sub
Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   入力画面の初期設定
'----------------------------------------------------------------------------
Dim i           As Integer
Dim sts         As Integer


Dim TANTO_CODE  As String
Dim TANTO_NAME  As String


    Init_Proc = True
    
    TANTO_CODE = Text1(ptxTANTO_CODE).Text
    TANTO_NAME = Text1(ptxTANTO_NAME).Text
    
    
    
    For i = ptxTANTO_CODE To ptxK_HIN_GAI
        Text1(i).Text = ""
    Next i
    
    Text1(ptxTANTO_CODE).Text = TANTO_CODE
    Text1(ptxTANTO_NAME).Text = TANTO_NAME
    
    
    
    '注文日＝当日
''    Text1(ptxS_ORDER_DT).Text = Format(Now, "YYYY/MM/DD")
''    Text1(ptxE_ORDER_DT).Text = Format(Now, "YYYY/MM/DD")


    Option1(poptORDER).Value = True
    Option1(poptSHIJI).Value = False



    For i = pcmbORDER To pcmbORDER
        
        Combo1(i).ListIndex = -1
    
    Next i

    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    Sort_Tbl(colHIN_NAME) = 9       'ｿｰﾄ除外

    Init_Proc = False

End Function
Private Function Ukeharai_Set_Proc(Index As Integer) As Integer
'----------------------------------------------------------------------------
'                   受払先マスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer




Dim i           As Integer
    
    Ukeharai_Set_Proc = True
    
    Combo1(Index).Clear
    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "受払先マスタ")
                Exit Function
        
        End Select

        
        
        Combo1(Index).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop

    Ukeharai_Set_Proc = False
    



End Function
Private Function List_Disp_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'           発注検討ﾃﾞｰﾀの表示
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

Dim Row     As Long

    List_Disp_Proc = True
    Call Input_Lock
    
    
    
    Set SHORDER = Nothing
    Tbl_Set_F = False
    
    
    Select Case Mode
        Case 0, 1
            com = BtOpGetFirst
        Case 2
            com = BtOpGetLast
    
    End Select
    
    Row = Min_Row - 1
       
    Do
    
        DoEvents
    
        Select Case Mode
    
            Case 0
                sts = BTRV(com, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
            
            Case 1, 2
                sts = BTRV(com, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K2_P_SHKENTO_OSAKA, Len(K2_P_SHKENTO_OSAKA), 2)
            
        End Select
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "発注検討ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        
        
        Row = Row + 1
        If Grid_Set_Proc(Row) Then
            Call Input_UnLock
            Exit Function
        End If
        Tbl_Set_F = True
        
        Select Case Mode
    
            Case 0, 1
                com = BtOpGetNext
            
            Case 2
                com = BtOpGetPrev
        End Select
    
    Loop
    
    Set TDBGrid1.Array = SHORDER
            
'    If Row <> (Min_Row - 1) Then
'        SHORDER.QuickSort Min_Row, SHORDER.UpperBound(1), colHIN_GAI, XORDER_ASCEND, XTYPE_STRING
'    End If
            
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    
    
    Call Input_UnLock
    List_Disp_Proc = False
    


End Function

Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           発注検討ﾃﾞｰﾀの内容をｸﾞﾘｯﾄﾞにｾｯﾄする
'----------------------------------------------------------------------------
Dim sts As Integer

    Grid_Set_Proc = True
    
    SHORDER.ReDim Min_Row, Row, Min_Col, Max_Col
    '事業部
    SHORDER(Row, colJGYOBU) = StrConv(P_SHKENTO_OSAKA_REC.JGYOBU, vbUnicode)
    '国内外
    SHORDER(Row, colNAIGAI) = StrConv(P_SHKENTO_OSAKA_REC.NAIGAI, vbUnicode)
    '品番
    SHORDER(Row, colHIN_GAI) = StrConv(P_SHKENTO_OSAKA_REC.HIN_GAI, vbUnicode)
    '品目ﾏｽﾀ読込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, SHORDER(Row, colJGYOBU))
    Call UniCode_Conv(K0_ITEM.NAIGAI, SHORDER(Row, colNAIGAI))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, SHORDER(Row, colHIN_GAI))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
            Exit Function
    End Select
    '品名
    SHORDER(Row, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    '必要量
    SHORDER(Row, colSO_SUU) = Format(CDbl(StrConv(P_SHKENTO_OSAKA_REC.SO_SUU, vbUnicode)), "#,##0")
    '仕入単価
    SHORDER(Row, colTANKA) = Format(CDbl(StrConv(P_SHKENTO_OSAKA_REC.TANKA, vbUnicode)), "#,##0.00")
    '棚番
    If Trim(StrConv(P_SHKENTO_OSAKA_REC.TANKA, vbUnicode)) <> "" Then
        SHORDER(Row, colST_LOCATION) = StrConv(P_SHKENTO_OSAKA_REC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(P_SHKENTO_OSAKA_REC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(P_SHKENTO_OSAKA_REC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(P_SHKENTO_OSAKA_REC.ST_DAN, vbUnicode)
    Else
        SHORDER(Row, colST_LOCATION) = ""
    End If
    '在庫数
    SHORDER(Row, colZAIKO_QTY) = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.ZAIKO_QTY, vbUnicode)), "#,##0")
    '注文残
    SHORDER(Row, colSHIJI_Z_QTY) = Format(CDbl(StrConv(P_SHKENTO_OSAKA_REC.SHIJI_Z_QTY, vbUnicode)), "#,##0")
    '引当残
    SHORDER(Row, colHIKIATE_Z_QTY) = Format(CDbl(StrConv(P_SHKENTO_OSAKA_REC.HIKIATE_Z_QTY, vbUnicode)), "#,##0")
    '不足数
    SHORDER(Row, colFUSOKU_QTY) = Format(CDbl(StrConv(P_SHKENTO_OSAKA_REC.FUSOKU_QTY, vbUnicode)), "#,##0")
    'ﾛｯﾄ数
    SHORDER(Row, colLOT) = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.LOT, vbUnicode)), "#,##0")
    '注文数
    SHORDER(Row, colORDER_QTY) = ""
    '仕入先
    SHORDER(Row, colORDER_CODE) = Trim(StrConv(P_SHKENTO_OSAKA_REC.ORDER_CODE, vbUnicode))
    '仕入先名
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Trim(StrConv(P_SHKENTO_OSAKA_REC.ORDER_CODE, vbUnicode)))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先ﾏｽﾀ")
            Exit Function
    End Select
    SHORDER(Row, colORDER_NAME) = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
    'LT
    SHORDER(Row, colLT) = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.LT, vbUnicode)), "#,##0")
    'LT
    SHORDER(Row, colY_NOUKI_DT) = ""

    Grid_Set_Proc = False


End Function
Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'           注文書印刷処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim Save_Order_Code As String * 5
                
Dim rpt         As New PI00090F1
Dim f           As New PI000902

                
    Print_Proc = True
                
    Call Input_Lock
                
    Call UniCode_Conv(K2_wP_SHORDER.WS_NO, WS_NO)
    Call UniCode_Conv(K2_wP_SHORDER.PRINT_F, P_PRINT_OFF)
    Call UniCode_Conv(K2_wP_SHORDER.ORDER_CODE, "")
    Call UniCode_Conv(K2_wP_SHORDER.ORDER_NO, "")
                
    com = BtOpGetGreaterEqual
                
    Save_Order_Code = ""

                
    Do
        DoEvents
        
        sts = BTRV(com, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), K2_wP_SHORDER, Len(K2_wP_SHORDER), 2)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(wP_SHORDER_REC.WS_NO, vbUnicode) <> WS_NO Or _
                    StrConv(wP_SHORDER_REC.PRINT_F, vbUnicode) <> P_PRINT_OFF Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                Exit Function
        End Select
    
        If Trim(Save_Order_Code) = "" Then
    
            Set rpt = New PI00090F1
        
            'レポートを印刷します。（true：印刷ダイアログあり false：なし）
            rpt.PrintReport False
        
            Set rpt = Nothing
    
    
    
'            f.RunReport rpt
'            f.Show
    
            Save_Order_Code = StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode)
    
    
        End If
    
        If Save_Order_Code <> StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode) Then
    
            Set rpt = New PI00090F1
        
            'レポートを印刷します。（true：印刷ダイアログあり false：なし）
            rpt.PrintReport False
        
            Set rpt = Nothing


'            f.RunReport rpt
'            f.Show
    
            Save_Order_Code = StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode)
    
    
        End If
    
        com = BtOpGetNext
    
    Loop
                

    Call Input_UnLock
    Print_Proc = False

End Function

Private Function Item_Read_Proc(HIN_GAI As String) As Integer
'----------------------------------------------------------------------------
'           品目ﾏｽﾀの読み込み
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer
    
    
    
    For i = 0 To UBound(JGYOBU_T)
    
            
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU_T(i).CODE)
        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)
    
    
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            
        Select Case sts
            Case BtNoErr
                
                
                
                
                Item_Read_Proc = sts
                Exit Function
            Case BtErrKeyNotFound
                Item_Read_Proc = BtErrKeyNotFound
                DoEvents
            Case Else
                Item_Read_Proc = sts
                Exit Function
        End Select
    
    
    
    
    Next i
        
    
    


End Function
Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'           発注検討ﾃﾞｰﾀの作成
'----------------------------------------------------------------------------
Dim sts             As Integer

Dim com             As Integer
Dim com_K           As Integer
Dim UPD_com         As Integer

Dim Skip_F          As Boolean
    
    
Dim Sumi_Zaiko_Qty  As Long
Dim Mi_Zaiko_Qty    As Long
    
    
Dim SHIJI_Z_QTY     As Double

Dim tmpQTY          As Double




    Data_Make_Proc = True
                                
    Call Input_Lock
                                
                                
                                            '発注検討　大阪向けＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "発注検討　大阪向け")
            Call Input_UnLock
            Exit Function
        End If
    End If
                                
                                
                                '発注検討　大阪向けＯＰＥＮ
    If P_SHKENTO_OSAKA_Open(BtOpenNomal, WS_NO) Then
        Call Input_UnLock
        Exit Function
    End If


    '----------------------------   救済の為
    com = BtOpGetFirst
    
    
    Do
        DoEvents
        sts = BTRV(com, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "発注検討ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        sts = BTRV(BtOpDelete, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
        
        If sts <> BtNoErr Then
            Call Input_UnLock
            Call File_Error(sts, BtOpDelete, "発注検討ﾃﾞｰﾀ")
            Exit Function
        End If
        
        com = BtOpGetGreater
    
    Loop



    Call UniCode_Conv(K2_P_SSHIJI_O.ORDER_DT, Format(Text1(ptxS_ORDER_DT).Text, "YYYYMMDD"))
    
    com = BtOpGetGreaterEqual
    
    
    Do
        DoEvents
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K2_P_SSHIJI_O, Len(K2_P_SSHIJI_O), 2)
            
        Select Case sts
            Case BtNoErr
            
            
                If Trim(Text1(ptxE_ORDER_DT).Text) = "" Then
                Else
                    If StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode) > Format(Text1(ptxE_ORDER_DT).Text, "YYYYMMDD") Then
                        Exit Do
                    End If
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "商品化指図ﾃﾞｰﾀ")
                Exit Function
        End Select
    
    
        Skip_F = False
    
        
        '受注分のみ
        If Option1(poptORDER).Value Then
            If Trim(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode)) = "" Then
                Skip_F = True
            End If
        End If
    
        '完了分は対象外
        If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = P_KAN_ON Then
            Skip_F = True
        End If
    
        'ｷｬﾝｾﾙ分は対象外
        If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
            Skip_F = True
        End If
    
        '指図票発行済みは対象外
'2007.06.20        If Trim(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode)) <> "" Then
'2007.06.20            Skip_F = True
'2007.06.20        End If
    
    
        '親品番指定時
        If Trim(Text1(ptxO_HIN_GAI).Text) <> "" Then
            If Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxO_HIN_GAI).Text) Then
                Skip_F = True
            End If
        End If
    
    
        If Skip_F Then
        Else
            
            Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_NO, StrConv(P_SSHIJI_O_REC.SHIJI_NO, vbUnicode))
            Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
            Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
            
            
            com_K = BtOpGetGreaterEqual
            
            Do
            
                DoEvents
                sts = BTRV(com_K, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                    
                Select Case sts
                    Case BtNoErr
                    
                        Skip_F = False
                    
                    
                        If StrConv(P_SSHIJI_O_REC.SHIJI_NO, vbUnicode) <> StrConv(P_SSHIJI_K_REC.SHIJI_NO, vbUnicode) Then
                            Exit Do
                        End If
                    
                        '子品番指定時
                        If Trim(Text1(ptxK_HIN_GAI).Text) <> "" Then
                            If Trim(StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)) <> Trim(Text1(ptxK_HIN_GAI).Text) Then
                                Skip_F = True
                            End If
                        End If
                    
                    
                        '同梱/構成以外はｽｷｯﾌﾟ 2007.04.10
                        If StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                            Skip_F = True
                        End If
                    
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
        
        
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            
                        Select Case sts
                            Case BtNoErr
                            
                                '仕入先指定時
                                If Trim(Text1(ptxORDER_CODE).Text) <> "" Then
                                    If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode)) <> Trim(Text1(ptxORDER_CODE).Text) Then
                                        Skip_F = True
                                    End If
                                End If
                            
                            Case BtErrKeyNotFound
                                
                                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "00000000.00")
                                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, "00000000")
                            
                                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, "000")
                            
                            
                                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                Call UniCode_Conv(ITEMREC.ST_REN, "")
                                Call UniCode_Conv(ITEMREC.ST_DAN, "")
                            
                            
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                Exit Function
                        End Select
                    
                    
                    
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, com, "商品化指図ﾃﾞｰﾀ")
                        Exit Function
                End Select
            
            
                
                If Not Skip_F Then
                
                    Call UniCode_Conv(K0_P_SHKENTO_OSAKA.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_P_SHKENTO_OSAKA.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_P_SHKENTO_OSAKA.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                
                    sts = BTRV(BtOpGetEqual, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
                        
                    Select Case sts
                        Case BtNoErr
                        
                            If StrConv(P_SSHIJI_O_REC.SHIJI_NO, vbUnicode) <> StrConv(P_SSHIJI_K_REC.SHIJI_NO, vbUnicode) Then
                                Exit Do
                            End If
                        
                            
                            
                            UPD_com = BtOpUpdate
                        
                        Case BtErrKeyNotFound
                            
                            UPD_com = BtOpInsert
                        
                        Case Else
                            Call Input_UnLock
                            Call File_Error(sts, BtOpGetEqual, "商品化指図ﾃﾞｰﾀ")
                            Exit Function
                    End Select
                
                                    
                    If UPD_com = BtOpInsert Then
                    
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.SO_SUU, "00000000.00")
                        
                        
                        
    
    
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                        
                            Call UniCode_Conv(P_SHKENTO_OSAKA_REC.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
                        Else
                            Call UniCode_Conv(P_SHKENTO_OSAKA_REC.TANKA, "00000000.00")
                        End If
    
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
        
                        '現在庫集計
                        If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                                Mi_Zaiko_Qty, _
                                                StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode), _
                                                StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode), _
                                                StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)) Then
                            Call Input_UnLock
                            Exit Function
                        End If
                        
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.ZAIKO_QTY, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000000"))
                        
                        '注文残集計
                        If Shiji_Zan_Proc(SHIJI_Z_QTY, _
                                            StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode), _
                                            StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode), _
                                            StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)) Then
                            Call Input_UnLock
                            Exit Function
                        End If
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.SHIJI_Z_QTY, Format(SHIJI_Z_QTY, "00000000.000"))
                        
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.HIKIATE_Z_QTY, "00000000.000")
    
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.FUSOKU_QTY, "00000000.000")
    
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.ORDER_QTY, "00000000.000")
    
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode)) Then
                            Call UniCode_Conv(P_SHKENTO_OSAKA_REC.LOT, StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))
                        Else
                            Call UniCode_Conv(P_SHKENTO_OSAKA_REC.LOT, "00000000")
                        End If
    
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.ORDER_CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
    
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, vbUnicode)) Then
                            Call UniCode_Conv(P_SHKENTO_OSAKA_REC.LT, StrConv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, vbUnicode))
                        Else
                            Call UniCode_Conv(P_SHKENTO_OSAKA_REC.LT, "000")
                        End If
    
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, "")
    
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.REC_NO, "0000")
    
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.FILLER, "")
                    
                    End If
                
                
                
                    
                                        
                    tmpQTY = (CDbl(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode) - CDbl(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))))
                    tmpQTY = CDbl(StrConv(P_SHKENTO_OSAKA_REC.SO_SUU, vbUnicode)) + (CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * tmpQTY)
                    Call UniCode_Conv(P_SHKENTO_OSAKA_REC.SO_SUU, Format(tmpQTY, "00000000.00"))
                
                    sts = BTRV(UPD_com, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
                
                    If sts <> BtNoErr Then
                        Call Input_UnLock
                        Call File_Error(sts, BtOpUpdate, "発注検討ﾃﾞｰﾀ")
                        Exit Function
                    End If
                
                End If
                
                com_K = BtOpGetNext
            
            Loop
        End If
        com = BtOpGetNext
    
    
    Loop


    '-------------------------------    引当済み数の集計

    Call UniCode_Conv(K1_P_SSHIJI_O.KAN_F, P_KAN_OFF)
    Call UniCode_Conv(K1_P_SSHIJI_O.SHIMUKE_CODE, "")
    Call UniCode_Conv(K1_P_SSHIJI_O.JGYOBU, "")
    Call UniCode_Conv(K1_P_SSHIJI_O.NAIGAI, "")
    Call UniCode_Conv(K1_P_SSHIJI_O.KAN_DT, "")
    Call UniCode_Conv(K1_P_SSHIJI_O.SHIJI_NO, "")
    
    com = BtOpGetGreaterEqual


    Do
        DoEvents
        
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K1_P_SSHIJI_O, Len(K1_P_SSHIJI_O), 1)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) <> P_KAN_OFF Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "商品化指図ﾃﾞｰﾀ")
                Exit Function
        End Select
    
'''2007.06.20        If (StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Or _
'''2007.06.20            Trim(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode)) = "") Then
        
        
        '受注分を対象とする
'''2007.06.20        If (StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Or _
'''2007.06.20            Trim(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode)) = "") Then
        
        '指図票発行分を対象とする 2007.06.20
        If (StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Or _
            Trim(StrConv(P_SSHIJI_O_REC.Print_datetime, vbUnicode)) = "") Then
        Else
    
            Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_NO, StrConv(P_SSHIJI_O_REC.SHIJI_NO, vbUnicode))
            Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
            Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
    
            com_K = BtOpGetGreater
    
            Do
                DoEvents
                sts = BTRV(com_K, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                    
                Select Case sts
                    Case BtNoErr
                    
                        If StrConv(P_SSHIJI_K_REC.SHIJI_NO, vbUnicode) <> StrConv(P_SSHIJI_O_REC.SHIJI_NO, vbUnicode) Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, com_K, "商品化指図ﾃﾞｰﾀ")
                        Exit Function
                End Select
            
            
                Call UniCode_Conv(K0_P_SHKENTO_OSAKA.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_SHKENTO_OSAKA.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_SHKENTO_OSAKA.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
            
                sts = BTRV(BtOpGetEqual, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
                    
                Select Case sts
                    Case BtNoErr
                    
                    
                        tmpQTY = CDbl(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)) - CDbl(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))
                        tmpQTY = tmpQTY * CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
                        tmpQTY = CDbl(StrConv(P_SHKENTO_OSAKA_REC.HIKIATE_Z_QTY, vbUnicode)) + tmpQTY
                    
                    
                        Call UniCode_Conv(P_SHKENTO_OSAKA_REC.HIKIATE_Z_QTY, Format(tmpQTY, "00000000.00"))
                                        
                    
                        sts = BTRV(BtOpUpdate, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
                        If sts <> BtNoErr Then
                        
                            Call File_Error(sts, BtOpUpdate, "発注検討ﾃﾞｰﾀ")
                            Exit Function
                        
                        End If
                    Case BtErrKeyNotFound
                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, BtOpGetEqual, "発注検討ﾃﾞｰﾀ")
                        Exit Function
                End Select
            
            
            
                com_K = BtOpGetNext
            
            
            Loop
    
    
        End If
    
        com = BtOpGetNext

    Loop


    '-------------------------------    不足数の集計
    com = BtOpGetFirst
    
    
    Do
        DoEvents
        sts = BTRV(com, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "発注検討ﾃﾞｰﾀ")
                Exit Function
        End Select
    
    
        tmpQTY = CDbl(StrConv(P_SHKENTO_OSAKA_REC.ZAIKO_QTY, vbUnicode))
        tmpQTY = tmpQTY + CDbl(StrConv(P_SHKENTO_OSAKA_REC.SHIJI_Z_QTY, vbUnicode))
    
'''2007.04.23        tmpQTY = tmpQTY - CDbl(StrConv(P_SHKENTO_OSAKA_REC.HIKIATE_Z_QTY, vbUnicode))
        tmpQTY = tmpQTY - CDbl(StrConv(P_SHKENTO_OSAKA_REC.SO_SUU, vbUnicode))
    
    
        If tmpQTY < 0 Then
            Call UniCode_Conv(P_SHKENTO_OSAKA_REC.FUSOKU_QTY, Format(tmpQTY, "0000000.00"))
        Else
            Call UniCode_Conv(P_SHKENTO_OSAKA_REC.FUSOKU_QTY, " " & Format(tmpQTY, "0000000.00"))
        End If
        
        sts = BTRV(BtOpUpdate, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
        
        If sts <> BtNoErr Then
            Call Input_UnLock
            Call File_Error(sts, BtOpUpdate, "発注検討ﾃﾞｰﾀ")
            Exit Function
        End If
        
        com = BtOpGetGreater
    
    Loop

    Call Input_UnLock

    Data_Make_Proc = False
End Function


Private Function Shiji_Zan_Proc(SHIJI_Z_QTY As Double, JGYOBU As String, NAIGAI As String, HIN_GAI As String) As Integer
'----------------------------------------------------------------------------
'           注文残の集計処理
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

Dim Skip_F  As Boolean

Dim tmpQTY  As Double

    Shiji_Zan_Proc = True

    Call UniCode_Conv(K1_P_SHORDER.JGYOBU, JGYOBU)
    Call UniCode_Conv(K1_P_SHORDER.NAIGAI, NAIGAI)
    Call UniCode_Conv(K1_P_SHORDER.HIN_GAI, HIN_GAI)

    Call UniCode_Conv(K1_P_SHORDER.ORDER_DT, "")
    Call UniCode_Conv(K1_P_SHORDER.ORDER_NO, "")

    com = BtOpGetGreater

    SHIJI_Z_QTY = 0


    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K1_P_SHORDER, Len(K1_P_SHORDER), 1)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_SHORDER_REC.JGYOBU, vbUnicode) <> JGYOBU Or _
                    StrConv(P_SHORDER_REC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) <> Trim(HIN_GAI) Then
                    
                    Exit Do
            
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        Skip_F = False
    
        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
            Skip_F = True
        End If
    
        If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then
            Skip_F = True
        End If
    
    
    
    
        If Not Skip_F Then
            tmpQTY = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))
            SHIJI_Z_QTY = SHIJI_Z_QTY + tmpQTY
        End If
    
        com = BtOpGetNext
    
    Loop

    Shiji_Zan_Proc = False

End Function
Private Function List_Print_Proc() As Integer
'----------------------------------------------------------------------------
'           発注検討リスト印刷処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
                
Dim rpt         As New PI00090F2
Dim f           As New PI000902

                
    List_Print_Proc = True
    Call Input_Lock
    
    Set rpt = New PI00090F2
        
    'レポートを印刷します。（true：印刷ダイアログあり false：なし）
    rpt.PrintReport False

    Set rpt = Nothing
    
    Call Input_UnLock
    List_Print_Proc = False

End Function

Private Function Data_Output_Proc() As Integer
'----------------------------------------------------------------------------
'           発注検討ﾃﾞｰﾀをCSV出力する
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim FileNo      As Long
Dim fileName    As String

Dim Fast_Flg    As Boolean


    Data_Output_Proc = True

    Call Input_Lock


    fileName = P_SHKENTO_OSAKA_DATA
    sts = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), sts) & Trim(WS_NO) & Right(Trim(fileName), Len(Trim(fileName)) - sts)

    On Error GoTo Error_Proc
    
    FileNo = FreeFile
    Open (fileName) For Output As FileNo
                                        
    On Error GoTo 0




    Fast_Flg = True

    com = BtOpGetFirst

    Do
        DoEvents
    
    
        sts = BTRV(com, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K1_P_SHKENTO_OSAKA, Len(K1_P_SHKENTO_OSAKA), 1)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "発注検討ﾌｧｲﾙ")
                Exit Function
        End Select
    
    
        If Fast_Flg Then
            Write #FileNo, "品番", "総必要量", "仕入単価", "標準棚番", "現在庫数", "注文残", "引当数", "不足数", "注文数", "発注ﾛｯﾄ", "仕入先", "ﾘｰﾄﾞﾀｲﾑ", "納期予定日"
            Fast_Flg = False
        End If
    
        '品番
        Write #FileNo, Trim(StrConv(P_SHKENTO_OSAKA_REC.HIN_GAI, vbUnicode)),
        '総必要量
        Write #FileNo, Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.SO_SUU, vbUnicode)), "#,##0"),
        '仕入単価
        Write #FileNo, Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.TANKA, vbUnicode)), "#,##0.00"),
        '標準棚番
        If Trim(StrConv(P_SHKENTO_OSAKA_REC.ST_SOKO, vbUnicode)) = "" Then
            Write #FileNo, ,
        Else
            Write #FileNo, StrConv(P_SHKENTO_OSAKA_REC.ST_SOKO, vbUnicode) & "-" & _
                                StrConv(P_SHKENTO_OSAKA_REC.ST_RETU, vbUnicode) & "-" & _
                                StrConv(P_SHKENTO_OSAKA_REC.ST_REN, vbUnicode) & "-" & _
                                StrConv(P_SHKENTO_OSAKA_REC.ST_DAN, vbUnicode),
        End If
        '現在庫
        Write #FileNo, Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.ZAIKO_QTY, vbUnicode)), "#,##0"),
        '注文残
        Write #FileNo, Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.SHIJI_Z_QTY, vbUnicode)), "#,##0"),
        '引当数
        Write #FileNo, Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.HIKIATE_Z_QTY, vbUnicode)), "#,##0"),
        '不足数
        Write #FileNo, Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.FUSOKU_QTY, vbUnicode)), "#,##0"),
        '注文数
        Write #FileNo, Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.ORDER_QTY, vbUnicode)), "#,##0"),
        '発注ﾛｯﾄ数
        Write #FileNo, Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.LOT, vbUnicode)), "#,##0"),
        '仕入先ｺｰﾄﾞ
        Write #FileNo, Trim(StrConv(P_SHKENTO_OSAKA_REC.ORDER_CODE, vbUnicode)),
        'ﾘｰﾄﾞﾀｲﾑ
        Write #FileNo, Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.LT, vbUnicode)), "#,##0"),
        '予定納期
        If Trim(StrConv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, vbUnicode)) = "" Then
            Write #FileNo, ""
        Else
            Write #FileNo, Left(StrConv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, vbUnicode), 4) & "/" & _
                            Mid(StrConv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                            Right(StrConv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, vbUnicode), 2)
        End If
        com = BtOpGetNext
    
    
    Loop


    Close #FileNo
    
    Call Input_UnLock
    
    Beep
    MsgBox "「" & fileName & "」" & "は正常に出力されました。"


    Data_Output_Proc = False

    Exit Function
    
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        Data_Output_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        Data_Output_Proc = True
    End If


End Function

Private Function Input_Update_Proc() As Integer
'----------------------------------------------------------------------------
'           発注検討ﾃﾞｰﾀを入力値で更新する
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer

Dim Skip_F      As Boolean




    Input_Update_Proc = True
    Call Input_Lock

    '-------------------------------    入力値で更新
    Set TDBGrid1.Array = SHORDER
    TDBGrid1.Refresh

    TDBGrid1.Update
    
    
    
    
    For i = 1 To SHORDER.UpperBound(1)
        DoEvents
    
        Skip_F = False
    
        Call UniCode_Conv(K0_P_SHKENTO_OSAKA.JGYOBU, SHORDER(i, colJGYOBU))
        Call UniCode_Conv(K0_P_SHKENTO_OSAKA.NAIGAI, SHORDER(i, colNAIGAI))
        Call UniCode_Conv(K0_P_SHKENTO_OSAKA.HIN_GAI, SHORDER(i, colHIN_GAI))
        sts = BTRV(BtOpGetEqual, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            
                Skip_F = True
            
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "発注検討ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        If Not Skip_F Then
    
            If IsNumeric(SHORDER(i, colORDER_QTY)) Then
                Call UniCode_Conv(P_SHKENTO_OSAKA_REC.ORDER_QTY, Format(CDbl(SHORDER(i, colORDER_QTY)), "00000000.00"))
            Else
                Call UniCode_Conv(P_SHKENTO_OSAKA_REC.ORDER_QTY, "00000000.00")
            End If
            
            
            If IsDate(SHORDER(i, colY_NOUKI_DT)) Then
                Call UniCode_Conv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, Format(SHORDER(i, colY_NOUKI_DT), "YYYYMMDD"))
            Else
                Call UniCode_Conv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, "")
            End If
            
            
            Call UniCode_Conv(P_SHKENTO_OSAKA_REC.REC_NO, Format(i, "0000"))
            
            
            sts = BTRV(BtOpUpdate, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K0_P_SHKENTO_OSAKA, Len(K0_P_SHKENTO_OSAKA), 0)
            
            If sts <> BtNoErr Then
                Call Input_UnLock
                Call File_Error(sts, BtOpUpdate, "発注検討ﾃﾞｰﾀ")
                Exit Function
            End If
    
        End If
    Next i

    Call Input_UnLock


    Input_Update_Proc = False

    
    



End Function

Private Function Grid_Error_Check_Proc() As Integer
'----------------------------------------------------------------------------
'           入力内容のエラーチェック
'----------------------------------------------------------------------------
Dim i   As Integer
Dim yn  As Integer
    
    
    
    Grid_Error_Check_Proc = True
    
    
    Set TDBGrid1.Array = SHORDER
    TDBGrid1.Refresh
    TDBGrid1.Update
    
    
    For i = 1 To SHORDER.UpperBound(1)
        
        If Trim(SHORDER(i, colORDER_QTY)) = "" Then
            SHORDER(i, colY_NOUKI_DT) = ""
        
            Set TDBGrid1.Array = SHORDER
            TDBGrid1.Refresh
            TDBGrid1.Update
        
        End If
    
            
        If Trim(SHORDER(i, colORDER_QTY)) <> "" Then
            If Not IsNumeric(SHORDER(i, colORDER_QTY)) Then
                MsgBox "品番「" & Trim(SHORDER(i, colHIN_GAI)) & "」注文数入力に誤りが有ります。確認して下さい。"
'                MsgBox "注文数入力に誤りが有ります。確認して下さい。"
                
                Set TDBGrid1.Array = SHORDER
                TDBGrid1.Refresh
                TDBGrid1.Update
                
                
                TDBGrid1.SetFocus
                Exit Function
            End If
    
            If CLng(SHORDER(i, colORDER_QTY)) <= 0 Then
                MsgBox "品番「" & Trim(SHORDER(i, colHIN_GAI)) & "」注文数入力に誤りが有ります。確認して下さい。"
                
                Set TDBGrid1.Array = SHORDER
                TDBGrid1.Refresh
                TDBGrid1.Update
                
                
                TDBGrid1.SetFocus
                Exit Function
            End If
    
    
            If Trim(SHORDER(i, colY_NOUKI_DT)) = "" Then
                        
                SHORDER(i, colY_NOUKI_DT) = Format(DateAdd("d", CDbl(SHORDER(i, colLT)), Format(Now, "YYYY/MM/DD")), "YYYY/MM/DD")
                Set TDBGrid1.Array = SHORDER
                TDBGrid1.Refresh
        
                TDBGrid1.Update
            End If
    
    
            If Not IsNumeric(SHORDER(i, colTANKA)) Then
                yn = MsgBox("品番「" & Trim(SHORDER(i, colHIN_GAI)) & "」単価未設定です。継続しますか？", vbYesNo, "確認入力")
                If yn = vbNo Then
    
                    Set TDBGrid1.Array = SHORDER
                    TDBGrid1.Refresh
                    TDBGrid1.Update
                    
                    
                    TDBGrid1.SetFocus
                    Exit Function
                End If
            Else
                If CDbl(SHORDER(i, colTANKA)) = 0 Then
                    yn = MsgBox("品番「" & Trim(SHORDER(i, colHIN_GAI)) & "」単価未設定です。継続しますか？", vbYesNo, "確認入力")
                    If yn = vbNo Then
        
                        Set TDBGrid1.Array = SHORDER
                        TDBGrid1.Refresh
                        TDBGrid1.Update
                        
                        
                        TDBGrid1.SetFocus
                        Exit Function
                    End If
                End If
            End If
        
        
            If Trim(SHORDER(i, colORDER_CODE)) = "" Then
                yn = MsgBox("品番「" & Trim(SHORDER(i, colHIN_GAI)) & "」仕入先未設定です。継続しますか？", vbYesNo, "確認入力")
                If yn = vbNo Then
    
                    Set TDBGrid1.Array = SHORDER
                    TDBGrid1.Refresh
                    TDBGrid1.Update
                    
                    
                    TDBGrid1.SetFocus
                    Exit Function
                End If
        
        
            End If
        End If
    
            
        If Trim(SHORDER(i, colY_NOUKI_DT)) <> "" Then
            If Not IsDate(SHORDER(i, colY_NOUKI_DT)) Then
                
                MsgBox "品番「" & Trim(SHORDER(i, colHIN_GAI)) & "」予定納期の入力に誤りが有ります。確認して下さい。"
                
                Set TDBGrid1.Array = SHORDER
                TDBGrid1.Refresh
                TDBGrid1.Update
                
                
                TDBGrid1.SetFocus
                Exit Function
            
            End If
        
        
        End If
    
    
            
    
    
    Next i

    Grid_Error_Check_Proc = False


End Function
