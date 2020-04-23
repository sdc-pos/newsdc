VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEK00201 
   Caption         =   "積水ハウス邸別出荷ﾃﾞｰﾀ　確認"
   ClientHeight    =   10815
   ClientLeft      =   2025
   ClientTop       =   -3210
   ClientWidth     =   17025
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
   ScaleHeight     =   10815
   ScaleWidth      =   17025
   StartUpPosition =   2  '画面の中央
   Visible         =   0   'False
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
      Index           =   3
      Left            =   1560
      TabIndex        =   23
      Top             =   840
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "抽出条件"
      Height          =   1332
      Left            =   3120
      TabIndex        =   14
      Top             =   360
      Width           =   13680
      Begin VB.TextBox Text1 
         Height          =   372
         Index           =   8
         Left            =   5880
         MaxLength       =   20
         TabIndex        =   7
         Top             =   720
         Width           =   2532
      End
      Begin VB.TextBox Text1 
         Height          =   372
         Index           =   7
         Left            =   2280
         MaxLength       =   13
         TabIndex        =   6
         Top             =   720
         Width           =   1692
      End
      Begin VB.TextBox Text1 
         Height          =   372
         Index           =   6
         Left            =   12000
         MaxLength       =   10
         TabIndex        =   5
         Top             =   720
         Width           =   1452
      End
      Begin VB.TextBox Text1 
         Height          =   372
         Index           =   5
         Left            =   12000
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   1452
      End
      Begin VB.TextBox Text1 
         Height          =   372
         Index           =   4
         Left            =   9600
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   1452
      End
      Begin VB.TextBox Text1 
         Height          =   372
         Index           =   3
         Left            =   5880
         MaxLength       =   20
         TabIndex        =   2
         Top             =   240
         Width           =   2532
      End
      Begin VB.TextBox Text1 
         Height          =   372
         Index           =   2
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1332
      End
      Begin VB.TextBox Text1 
         Height          =   372
         Index           =   1
         Left            =   840
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "梱包ID"
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "邸別ﾗﾍﾞﾙID"
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "品管№"
         Height          =   255
         Index           =   6
         Left            =   11160
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "件管№"
         Height          =   255
         Index           =   5
         Left            =   11160
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "注文№"
         Height          =   255
         Index           =   4
         Left            =   8760
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "品　番"
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "作成日"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "～"
         Height          =   252
         Index           =   1
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   372
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   372
      Index           =   0
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表 示"
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
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8175
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   14420
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "削除"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "処理結果"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "　データ　　　作成日"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "　データ　　　作成時刻"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "連番"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "受注日"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "　　納入　　受入場所"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "　　納入　　　　　受入場所名"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "得意先　　　コード"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "直納先　　　コード"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "得意先品番　　■品番（上）"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "品番　　　　　■品番（下）"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "注文№　　　　　　■指図№（上）"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "出荷順番　　　　　　■指図№（下・左）"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "邸名　　　　　　　■指図№（下・右）"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "受注数量"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "出荷確定日"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "納入日"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "件管№　　　　　　■管理№（上）"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "品管№　　　　　■管理№（下）"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "単品区分"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "邸別ﾗﾍﾞﾙID"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "箱№"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "実出庫数"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "出庫　       担当者"
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "出庫　　　　　日時"
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "梱包　　　担当者"
      Columns(26).DataField=   ""
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "梱包　　　　　日時"
      Columns(27).DataField=   ""
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "照合  　    担当者"
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "照合　　　　　日時"
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "口数"
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "才数"
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "梱包ID"
      Columns(32).DataField=   ""
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "検品  　    担当者"
      Columns(33).DataField=   ""
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "検品　　　　　日時"
      Columns(34).DataField=   ""
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).Caption=   "集合梱包    担当者"
      Columns(35).DataField=   ""
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).Caption=   "集合梱包         日時"
      Columns(36).DataField=   ""
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(37)._VlistStyle=   0
      Columns(37)._MaxComboItems=   5
      Columns(37).Caption=   "追加　　　担当者"
      Columns(37).DataField=   ""
      Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(38)._VlistStyle=   0
      Columns(38)._MaxComboItems=   5
      Columns(38).Caption=   "追加　　　　　日時"
      Columns(38).DataField=   ""
      Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(39)._VlistStyle=   0
      Columns(39)._MaxComboItems=   5
      Columns(39).Caption=   "更新　　　担当者"
      Columns(39).DataField=   ""
      Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(40)._VlistStyle=   0
      Columns(40)._MaxComboItems=   5
      Columns(40).Caption=   "更新　　　　　日時"
      Columns(40).DataField=   ""
      Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(41)._VlistStyle=   0
      Columns(41)._MaxComboItems=   5
      Columns(41).Caption=   "伝票ＩＤ（出荷予定）"
      Columns(41).DataField=   ""
      Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(42)._VlistStyle=   0
      Columns(42)._MaxComboItems=   5
      Columns(42).Caption=   "仕掛中　品番"
      Columns(42).DataField=   ""
      Columns(42)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(43)._VlistStyle=   0
      Columns(43)._MaxComboItems=   5
      Columns(43).Caption=   "仕掛数　バラ"
      Columns(43).DataField=   ""
      Columns(43)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(44)._VlistStyle=   0
      Columns(44)._MaxComboItems=   5
      Columns(44).Caption=   "仕掛数　箱"
      Columns(44).DataField=   ""
      Columns(44)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(45)._VlistStyle=   0
      Columns(45)._MaxComboItems=   5
      Columns(45).Caption=   "品番読込み実績数"
      Columns(45).DataField=   ""
      Columns(45)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   46
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=46"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=556"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2170"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2064"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2381"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2275"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2381"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1376"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1270"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=1588"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1482"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2064"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1958"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2831"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2725"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=1958"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1852"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=1958"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=1852"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=2593"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2487"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=2593"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2487"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=3281"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(53)=   "Column(13).Width=3545"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=3440"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=3334"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=3228"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=2752"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(64)=   "Column(15)._ColStyle=2"
      Splits(0)._ColumnProps(65)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(66)=   "Column(16).Width=2514"
      Splits(0)._ColumnProps(67)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(16)._WidthInPix=2408"
      Splits(0)._ColumnProps(69)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(70)=   "Column(17).Width=2064"
      Splits(0)._ColumnProps(71)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(17)._WidthInPix=1958"
      Splits(0)._ColumnProps(73)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(74)=   "Column(18).Width=3254"
      Splits(0)._ColumnProps(75)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(18)._WidthInPix=3149"
      Splits(0)._ColumnProps(77)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(78)=   "Column(19).Width=2858"
      Splits(0)._ColumnProps(79)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(19)._WidthInPix=2752"
      Splits(0)._ColumnProps(81)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(82)=   "Column(20).Width=1296"
      Splits(0)._ColumnProps(83)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(84)=   "Column(20)._WidthInPix=1191"
      Splits(0)._ColumnProps(85)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(86)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(87)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(88)=   "Column(21)._WidthInPix=2619"
      Splits(0)._ColumnProps(89)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(90)=   "Column(22).Width=1508"
      Splits(0)._ColumnProps(91)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(92)=   "Column(22)._WidthInPix=1402"
      Splits(0)._ColumnProps(93)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(94)=   "Column(23).Width=2752"
      Splits(0)._ColumnProps(95)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(96)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(97)=   "Column(23)._ColStyle=2"
      Splits(0)._ColumnProps(98)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(99)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(100)=   "Column(24).Width=1773"
      Splits(0)._ColumnProps(101)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(102)=   "Column(24)._WidthInPix=1667"
      Splits(0)._ColumnProps(103)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(104)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(105)=   "Column(25).Width=2328"
      Splits(0)._ColumnProps(106)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(25)._WidthInPix=2223"
      Splits(0)._ColumnProps(108)=   "Column(25).Visible=0"
      Splits(0)._ColumnProps(109)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(110)=   "Column(26).Width=1773"
      Splits(0)._ColumnProps(111)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(112)=   "Column(26)._WidthInPix=1667"
      Splits(0)._ColumnProps(113)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(114)=   "Column(27).Width=2328"
      Splits(0)._ColumnProps(115)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(116)=   "Column(27)._WidthInPix=2223"
      Splits(0)._ColumnProps(117)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(118)=   "Column(28).Width=1773"
      Splits(0)._ColumnProps(119)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(120)=   "Column(28)._WidthInPix=1667"
      Splits(0)._ColumnProps(121)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(122)=   "Column(29).Width=2328"
      Splits(0)._ColumnProps(123)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(124)=   "Column(29)._WidthInPix=2223"
      Splits(0)._ColumnProps(125)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(126)=   "Column(30).Width=1244"
      Splits(0)._ColumnProps(127)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(128)=   "Column(30)._WidthInPix=1138"
      Splits(0)._ColumnProps(129)=   "Column(30)._ColStyle=2"
      Splits(0)._ColumnProps(130)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(131)=   "Column(31).Width=1455"
      Splits(0)._ColumnProps(132)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(133)=   "Column(31)._WidthInPix=1349"
      Splits(0)._ColumnProps(134)=   "Column(31)._ColStyle=2"
      Splits(0)._ColumnProps(135)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(136)=   "Column(32).Width=3493"
      Splits(0)._ColumnProps(137)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(138)=   "Column(32)._WidthInPix=3387"
      Splits(0)._ColumnProps(139)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(140)=   "Column(33).Width=1773"
      Splits(0)._ColumnProps(141)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(142)=   "Column(33)._WidthInPix=1667"
      Splits(0)._ColumnProps(143)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(144)=   "Column(34).Width=2328"
      Splits(0)._ColumnProps(145)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(146)=   "Column(34)._WidthInPix=2223"
      Splits(0)._ColumnProps(147)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(148)=   "Column(35).Width=1773"
      Splits(0)._ColumnProps(149)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(150)=   "Column(35)._WidthInPix=1667"
      Splits(0)._ColumnProps(151)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(152)=   "Column(36).Width=2328"
      Splits(0)._ColumnProps(153)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(154)=   "Column(36)._WidthInPix=2223"
      Splits(0)._ColumnProps(155)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(156)=   "Column(37).Width=1773"
      Splits(0)._ColumnProps(157)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(158)=   "Column(37)._WidthInPix=1667"
      Splits(0)._ColumnProps(159)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(160)=   "Column(38).Width=2328"
      Splits(0)._ColumnProps(161)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(162)=   "Column(38)._WidthInPix=2223"
      Splits(0)._ColumnProps(163)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(164)=   "Column(39).Width=1773"
      Splits(0)._ColumnProps(165)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(166)=   "Column(39)._WidthInPix=1667"
      Splits(0)._ColumnProps(167)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(168)=   "Column(40).Width=2328"
      Splits(0)._ColumnProps(169)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(170)=   "Column(40)._WidthInPix=2223"
      Splits(0)._ColumnProps(171)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(172)=   "Column(41).Width=4101"
      Splits(0)._ColumnProps(173)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(174)=   "Column(41)._WidthInPix=3995"
      Splits(0)._ColumnProps(175)=   "Column(41).Order=42"
      Splits(0)._ColumnProps(176)=   "Column(42).Width=3281"
      Splits(0)._ColumnProps(177)=   "Column(42).DividerColor=0"
      Splits(0)._ColumnProps(178)=   "Column(42)._WidthInPix=3175"
      Splits(0)._ColumnProps(179)=   "Column(42).Order=43"
      Splits(0)._ColumnProps(180)=   "Column(43).Width=3281"
      Splits(0)._ColumnProps(181)=   "Column(43).DividerColor=0"
      Splits(0)._ColumnProps(182)=   "Column(43)._WidthInPix=3175"
      Splits(0)._ColumnProps(183)=   "Column(43)._ColStyle=2"
      Splits(0)._ColumnProps(184)=   "Column(43).Order=44"
      Splits(0)._ColumnProps(185)=   "Column(44).Width=3281"
      Splits(0)._ColumnProps(186)=   "Column(44).DividerColor=0"
      Splits(0)._ColumnProps(187)=   "Column(44)._WidthInPix=3175"
      Splits(0)._ColumnProps(188)=   "Column(44)._ColStyle=2"
      Splits(0)._ColumnProps(189)=   "Column(44).Order=45"
      Splits(0)._ColumnProps(190)=   "Column(45).Width=3281"
      Splits(0)._ColumnProps(191)=   "Column(45).DividerColor=0"
      Splits(0)._ColumnProps(192)=   "Column(45)._WidthInPix=3175"
      Splits(0)._ColumnProps(193)=   "Column(45)._ColStyle=2"
      Splits(0)._ColumnProps(194)=   "Column(45).Order=46"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      OLEDropMode     =   1
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
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(27)  =   ":id=14,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(28)  =   ":id=14,.fontname=ＭＳ ゴシック"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=162,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=159,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=160,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=161,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=110,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=107,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=108,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=109,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=3"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=3"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=70,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=74,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=78,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=75,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=76,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=77,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=82,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=79,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=80,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=81,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=86,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=83,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=84,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=85,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=90,.parent=13,.alignment=1"
      _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=87,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=88,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=89,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=94,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=91,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=92,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=93,.parent=17"
      _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=98,.parent=13"
      _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=95,.parent=14"
      _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=96,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=97,.parent=17"
      _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=102,.parent=13"
      _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=99,.parent=14"
      _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=100,.parent=15"
      _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=101,.parent=17"
      _StyleDefs(114) =   "Splits(0).Columns(19).Style:id=106,.parent=13"
      _StyleDefs(115) =   "Splits(0).Columns(19).HeadingStyle:id=103,.parent=14"
      _StyleDefs(116) =   "Splits(0).Columns(19).FooterStyle:id=104,.parent=15"
      _StyleDefs(117) =   "Splits(0).Columns(19).EditorStyle:id=105,.parent=17"
      _StyleDefs(118) =   "Splits(0).Columns(20).Style:id=114,.parent=13"
      _StyleDefs(119) =   "Splits(0).Columns(20).HeadingStyle:id=111,.parent=14"
      _StyleDefs(120) =   "Splits(0).Columns(20).FooterStyle:id=112,.parent=15"
      _StyleDefs(121) =   "Splits(0).Columns(20).EditorStyle:id=113,.parent=17"
      _StyleDefs(122) =   "Splits(0).Columns(21).Style:id=118,.parent=13"
      _StyleDefs(123) =   "Splits(0).Columns(21).HeadingStyle:id=115,.parent=14"
      _StyleDefs(124) =   "Splits(0).Columns(21).FooterStyle:id=116,.parent=15"
      _StyleDefs(125) =   "Splits(0).Columns(21).EditorStyle:id=117,.parent=17"
      _StyleDefs(126) =   "Splits(0).Columns(22).Style:id=122,.parent=13"
      _StyleDefs(127) =   "Splits(0).Columns(22).HeadingStyle:id=119,.parent=14"
      _StyleDefs(128) =   "Splits(0).Columns(22).FooterStyle:id=120,.parent=15"
      _StyleDefs(129) =   "Splits(0).Columns(22).EditorStyle:id=121,.parent=17"
      _StyleDefs(130) =   "Splits(0).Columns(23).Style:id=126,.parent=13,.alignment=1"
      _StyleDefs(131) =   "Splits(0).Columns(23).HeadingStyle:id=123,.parent=14"
      _StyleDefs(132) =   "Splits(0).Columns(23).FooterStyle:id=124,.parent=15"
      _StyleDefs(133) =   "Splits(0).Columns(23).EditorStyle:id=125,.parent=17"
      _StyleDefs(134) =   "Splits(0).Columns(24).Style:id=130,.parent=13"
      _StyleDefs(135) =   "Splits(0).Columns(24).HeadingStyle:id=127,.parent=14"
      _StyleDefs(136) =   "Splits(0).Columns(24).FooterStyle:id=128,.parent=15"
      _StyleDefs(137) =   "Splits(0).Columns(24).EditorStyle:id=129,.parent=17"
      _StyleDefs(138) =   "Splits(0).Columns(25).Style:id=134,.parent=13"
      _StyleDefs(139) =   "Splits(0).Columns(25).HeadingStyle:id=131,.parent=14"
      _StyleDefs(140) =   "Splits(0).Columns(25).FooterStyle:id=132,.parent=15"
      _StyleDefs(141) =   "Splits(0).Columns(25).EditorStyle:id=133,.parent=17"
      _StyleDefs(142) =   "Splits(0).Columns(26).Style:id=138,.parent=13"
      _StyleDefs(143) =   "Splits(0).Columns(26).HeadingStyle:id=135,.parent=14"
      _StyleDefs(144) =   "Splits(0).Columns(26).FooterStyle:id=136,.parent=15"
      _StyleDefs(145) =   "Splits(0).Columns(26).EditorStyle:id=137,.parent=17"
      _StyleDefs(146) =   "Splits(0).Columns(27).Style:id=142,.parent=13"
      _StyleDefs(147) =   "Splits(0).Columns(27).HeadingStyle:id=139,.parent=14"
      _StyleDefs(148) =   "Splits(0).Columns(27).FooterStyle:id=140,.parent=15"
      _StyleDefs(149) =   "Splits(0).Columns(27).EditorStyle:id=141,.parent=17"
      _StyleDefs(150) =   "Splits(0).Columns(28).Style:id=166,.parent=13"
      _StyleDefs(151) =   "Splits(0).Columns(28).HeadingStyle:id=163,.parent=14"
      _StyleDefs(152) =   "Splits(0).Columns(28).FooterStyle:id=164,.parent=15"
      _StyleDefs(153) =   "Splits(0).Columns(28).EditorStyle:id=165,.parent=17"
      _StyleDefs(154) =   "Splits(0).Columns(29).Style:id=170,.parent=13"
      _StyleDefs(155) =   "Splits(0).Columns(29).HeadingStyle:id=167,.parent=14"
      _StyleDefs(156) =   "Splits(0).Columns(29).FooterStyle:id=168,.parent=15"
      _StyleDefs(157) =   "Splits(0).Columns(29).EditorStyle:id=169,.parent=17"
      _StyleDefs(158) =   "Splits(0).Columns(30).Style:id=174,.parent=13,.alignment=1"
      _StyleDefs(159) =   "Splits(0).Columns(30).HeadingStyle:id=171,.parent=14"
      _StyleDefs(160) =   "Splits(0).Columns(30).FooterStyle:id=172,.parent=15"
      _StyleDefs(161) =   "Splits(0).Columns(30).EditorStyle:id=173,.parent=17"
      _StyleDefs(162) =   "Splits(0).Columns(31).Style:id=178,.parent=13,.alignment=1"
      _StyleDefs(163) =   "Splits(0).Columns(31).HeadingStyle:id=175,.parent=14"
      _StyleDefs(164) =   "Splits(0).Columns(31).FooterStyle:id=176,.parent=15"
      _StyleDefs(165) =   "Splits(0).Columns(31).EditorStyle:id=177,.parent=17"
      _StyleDefs(166) =   "Splits(0).Columns(32).Style:id=182,.parent=13"
      _StyleDefs(167) =   "Splits(0).Columns(32).HeadingStyle:id=179,.parent=14"
      _StyleDefs(168) =   "Splits(0).Columns(32).FooterStyle:id=180,.parent=15"
      _StyleDefs(169) =   "Splits(0).Columns(32).EditorStyle:id=181,.parent=17"
      _StyleDefs(170) =   "Splits(0).Columns(33).Style:id=186,.parent=13"
      _StyleDefs(171) =   "Splits(0).Columns(33).HeadingStyle:id=183,.parent=14"
      _StyleDefs(172) =   "Splits(0).Columns(33).FooterStyle:id=184,.parent=15"
      _StyleDefs(173) =   "Splits(0).Columns(33).EditorStyle:id=185,.parent=17"
      _StyleDefs(174) =   "Splits(0).Columns(34).Style:id=190,.parent=13"
      _StyleDefs(175) =   "Splits(0).Columns(34).HeadingStyle:id=187,.parent=14"
      _StyleDefs(176) =   "Splits(0).Columns(34).FooterStyle:id=188,.parent=15"
      _StyleDefs(177) =   "Splits(0).Columns(34).EditorStyle:id=189,.parent=17"
      _StyleDefs(178) =   "Splits(0).Columns(35).Style:id=194,.parent=13"
      _StyleDefs(179) =   "Splits(0).Columns(35).HeadingStyle:id=191,.parent=14"
      _StyleDefs(180) =   "Splits(0).Columns(35).FooterStyle:id=192,.parent=15"
      _StyleDefs(181) =   "Splits(0).Columns(35).EditorStyle:id=193,.parent=17"
      _StyleDefs(182) =   "Splits(0).Columns(36).Style:id=198,.parent=13"
      _StyleDefs(183) =   "Splits(0).Columns(36).HeadingStyle:id=195,.parent=14"
      _StyleDefs(184) =   "Splits(0).Columns(36).FooterStyle:id=196,.parent=15"
      _StyleDefs(185) =   "Splits(0).Columns(36).EditorStyle:id=197,.parent=17"
      _StyleDefs(186) =   "Splits(0).Columns(37).Style:id=146,.parent=13"
      _StyleDefs(187) =   "Splits(0).Columns(37).HeadingStyle:id=143,.parent=14"
      _StyleDefs(188) =   "Splits(0).Columns(37).FooterStyle:id=144,.parent=15"
      _StyleDefs(189) =   "Splits(0).Columns(37).EditorStyle:id=145,.parent=17"
      _StyleDefs(190) =   "Splits(0).Columns(38).Style:id=150,.parent=13"
      _StyleDefs(191) =   "Splits(0).Columns(38).HeadingStyle:id=147,.parent=14"
      _StyleDefs(192) =   "Splits(0).Columns(38).FooterStyle:id=148,.parent=15"
      _StyleDefs(193) =   "Splits(0).Columns(38).EditorStyle:id=149,.parent=17"
      _StyleDefs(194) =   "Splits(0).Columns(39).Style:id=154,.parent=13"
      _StyleDefs(195) =   "Splits(0).Columns(39).HeadingStyle:id=151,.parent=14"
      _StyleDefs(196) =   "Splits(0).Columns(39).FooterStyle:id=152,.parent=15"
      _StyleDefs(197) =   "Splits(0).Columns(39).EditorStyle:id=153,.parent=17"
      _StyleDefs(198) =   "Splits(0).Columns(40).Style:id=158,.parent=13"
      _StyleDefs(199) =   "Splits(0).Columns(40).HeadingStyle:id=155,.parent=14"
      _StyleDefs(200) =   "Splits(0).Columns(40).FooterStyle:id=156,.parent=15"
      _StyleDefs(201) =   "Splits(0).Columns(40).EditorStyle:id=157,.parent=17"
      _StyleDefs(202) =   "Splits(0).Columns(41).Style:id=202,.parent=13"
      _StyleDefs(203) =   "Splits(0).Columns(41).HeadingStyle:id=199,.parent=14"
      _StyleDefs(204) =   "Splits(0).Columns(41).FooterStyle:id=200,.parent=15"
      _StyleDefs(205) =   "Splits(0).Columns(41).EditorStyle:id=201,.parent=17"
      _StyleDefs(206) =   "Splits(0).Columns(42).Style:id=206,.parent=13"
      _StyleDefs(207) =   "Splits(0).Columns(42).HeadingStyle:id=203,.parent=14"
      _StyleDefs(208) =   "Splits(0).Columns(42).FooterStyle:id=204,.parent=15"
      _StyleDefs(209) =   "Splits(0).Columns(42).EditorStyle:id=205,.parent=17"
      _StyleDefs(210) =   "Splits(0).Columns(43).Style:id=210,.parent=13,.alignment=1"
      _StyleDefs(211) =   "Splits(0).Columns(43).HeadingStyle:id=207,.parent=14"
      _StyleDefs(212) =   "Splits(0).Columns(43).FooterStyle:id=208,.parent=15"
      _StyleDefs(213) =   "Splits(0).Columns(43).EditorStyle:id=209,.parent=17"
      _StyleDefs(214) =   "Splits(0).Columns(44).Style:id=214,.parent=13,.alignment=1"
      _StyleDefs(215) =   "Splits(0).Columns(44).HeadingStyle:id=211,.parent=14"
      _StyleDefs(216) =   "Splits(0).Columns(44).FooterStyle:id=212,.parent=15"
      _StyleDefs(217) =   "Splits(0).Columns(44).EditorStyle:id=213,.parent=17"
      _StyleDefs(218) =   "Splits(0).Columns(45).Style:id=218,.parent=13,.alignment=1"
      _StyleDefs(219) =   "Splits(0).Columns(45).HeadingStyle:id=215,.parent=14"
      _StyleDefs(220) =   "Splits(0).Columns(45).FooterStyle:id=216,.parent=15"
      _StyleDefs(221) =   "Splits(0).Columns(45).EditorStyle:id=217,.parent=17"
      _StyleDefs(222) =   "Named:id=33:Normal"
      _StyleDefs(223) =   ":id=33,.parent=0"
      _StyleDefs(224) =   "Named:id=34:Heading"
      _StyleDefs(225) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(226) =   ":id=34,.wraptext=-1"
      _StyleDefs(227) =   "Named:id=35:Footing"
      _StyleDefs(228) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(229) =   "Named:id=36:Selected"
      _StyleDefs(230) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(231) =   "Named:id=37:Caption"
      _StyleDefs(232) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(233) =   "Named:id=38:HighlightRow"
      _StyleDefs(234) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(235) =   "Named:id=39:EvenRow"
      _StyleDefs(236) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(237) =   "Named:id=40:OddRow"
      _StyleDefs(238) =   ":id=40,.parent=33"
      _StyleDefs(239) =   "Named:id=41:RecordSelector"
      _StyleDefs(240) =   ":id=41,.parent=34"
      _StyleDefs(241) =   "Named:id=42:FilterBar"
      _StyleDefs(242) =   ":id=42,.parent=33"
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
      Left            =   1560
      TabIndex        =   10
      Top             =   360
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登 録"
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
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "読込み件数"
      Height          =   255
      Index           =   0
      Left            =   13320
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "登録"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "表示"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   2
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "SEK00201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const ptxIn_Cnt% = 0            '入力カウント
Private Const ptxS_SND_YMD% = 1         '作成日　開始
Private Const ptxE_SND_YMD% = 2         '作成日　終了
Private Const ptxHINB_CD% = 3           '品番
Private Const ptxCHU_CD% = 4            '注文№
Private Const ptxKEN_NO% = 5            '件管№
Private Const ptxHIN_NO% = 6            '品管№

Private Const ptxTEI_LABELID% = 7       '邸別ﾗﾍﾞﾙID
Private Const ptxKONPO_ID% = 8          '梱包ﾗﾍﾞﾙID







Dim Y_Syuka_TEI     As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 45             '最大列数   41-->45 2012.10.24

Private Const colDEL_FLG% = 0           '削除ﾌﾗｸﾞ
Private Const colSHORI% = 1             '処理結果
Private Const colSND_YMD% = 2           'データ作成日
Private Const colSND_HMS% = 3           'データ作成時刻
Private Const colSEQ_NO% = 4            '連番
Private Const colJUC_YMD% = 5           '受注日
Private Const colNOU_CD% = 6            '納入受入場
Private Const colNOU_NM% = 7            '納入受入場名
Private Const colTOK_CD% = 8            '得意先ｺｰﾄﾞ
Private Const colCHO_CD% = 9            '直納先ｺｰﾄﾞ
Private Const colTHINB_CD% = 10         '得意先品番　■品番(上)
Private Const colHINB_CD% = 11          '品番　      ■品番(下)
Private Const colCHU_CD% = 12           '注文№　    ■指図№(上)
Private Const colSYU_JUN% = 13          '出荷順番　  ■指図№(下・左)
Private Const colTEI_NM% = 14           '邸名　      ■指図№(下・右)
Private Const colJUC_SUU% = 15          '受注数量
Private Const colSYU_YMD% = 16          '出荷確定日
Private Const colNOU_YMD% = 17          '納入日
Private Const colKEN_NO% = 18           '件管№　　　■管理№(上)
Private Const colHIN_NO% = 19           '件管№　　　■管理№(下)
Private Const colTANP_KB% = 20          '単品区分

Private Const colTEI_LABELID% = 21      '邸別ﾗﾍﾞﾙID
Private Const colHAKO_NO% = 22          '箱№
Private Const colJITU_SUU% = 23         '実出庫数
Private Const colJITU_TANTO% = 24       '出庫　担当者
Private Const colJITU_DATETIME% = 25    '出庫　日時
Private Const colKONPO_TANTO% = 26      '梱包　担当者
Private Const colKONPO_DATETIME% = 27   '梱包　日時


Private Const colSHOGO_TANTO% = 28      '注文ﾃﾞｰﾀ照合担当
Private Const colSHOGO_DATETIME% = 29   '注文ﾃﾞｰﾀ照合日時

Private Const colKUTI_SU% = 30          '口数
Private Const colSAI_SU% = 31           '才数
    
Private Const colKONPO_ID% = 32         '梱包ID
    
    
Private Const colKENPIN_TANTO% = 33     '検品担当者
Private Const colKENPIN_DATETIME% = 34  '検品日時
    
    
Private Const colSYUGO_KONPO_TANTO% = 35    '集合梱包担当者
Private Const colSYUGO_KONPO_DATETIME% = 36 '集合梱包日時

Private Const colINS_TANTO% = 37        '追加　担当者
Private Const colINS_DateTime% = 38     '追加　日時
Private Const colUPD_TANTO% = 39        '更新　担当者
Private Const colUPD_DATETIME% = 40     '更新　日時

Private Const colID_NO% = 41            'ＩＤ_ＮＯ

Private Const colKEN_HINBAN% = 42       '仕掛中　品番   2012.10.24
Private Const colCNT_BARA_SU% = 43      '検品実績　バラ 2012.10.24
Private Const colCNT_HAKO_SU% = 44      '検品実績　箱   2012.10.24

Private Const colJ_HIN_CHK_CNT% = 45    '検品実績　品番読込み数   2012.10.24




Private CSV_FILE    As String           'CSV出力用ファイル名





Private Sort_Tbl(Min_Col To colID_NO) _
            As Integer                  'ｿｰﾄの制御 0:昇順 1:降順



'Private Const LAST_UPDATE_DAY$ = "[SEK0020] 2012.10.25 16:00"
Private Const LAST_UPDATE_DAY$ = "[SEK0020] 2015.08.07 16:30"












Private Sub Command1_Click(Index As Integer)
    
    Select Case Index
        
        Case 0          '登録
        
        
            If Update_Proc() Then
                Unload Me
            End If
        
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
        
        Case 1          '表示
        
        
            If Not IsDate(Text1(ptxS_SND_YMD).Text) Then
                MsgBox "入力した項目はエラーです（開始日付）"
                Text1(ptxS_SND_YMD).SetFocus
                Exit Sub
            End If
        
            If Not IsDate(Text1(ptxE_SND_YMD).Text) Then
                MsgBox "入力した項目はエラーです（終了日付）"
                Text1(ptxE_SND_YMD).SetFocus
                Exit Sub
            End If
        
            If Text1(ptxS_SND_YMD).Text > Text1(ptxE_SND_YMD).Text Then
                MsgBox "入力した項目はエラーです（日付範囲）"
                Text1(ptxS_SND_YMD).SetFocus
                Exit Sub
            End If
        
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
        
        
        Case 2          '終了
            
            Unload Me
    
    
        Case 3          'ﾃﾞｰﾀ出力
    
            If CSV_DATA_OUT_PROC() Then
            
            End If
    
    End Select
    
    
    
    Command1(Index).SetFocus
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    
    
    If Shift = vbAltMask Then
    
        If Command1(0).Enabled Then
            Command1(0).Enabled = False
            Command1(3).Enabled = False
        Else
            Command1(0).Enabled = True
            Command1(3).Enabled = True
        End If
    
    End If
    
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select

End Sub

Private Sub Form_Load()


Dim c       As String * 128

Dim i       As Integer

    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "積水ハウス邸別注文ﾃﾞｰﾀ　確認", Me.hwnd, 0)
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

    SEK00201.Caption = SEK00201.Caption & " " & LAST_UPDATE_DAY
                                
                                'ＣＳＶﾌｧｲﾙ名取り込み
    If GetIni(App.EXEName, "CSV_FILE", App.EXEName, c) Then
        CSV_FILE = ""
    
    Else
        CSV_FILE = RTrim(c)
        Command1(3).Visible = True
    End If




                                '邸別注文ﾃﾞｰﾀＯＰＥＮ
    If Y_SYU_TEI_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '出荷予定(H)ﾃﾞｰﾀＯＰＥＮ
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目   2013.02.02
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If




    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    Text1(ptxS_SND_YMD).Text = Format(Now, "YYYY/MM/DD")
    Text1(ptxE_SND_YMD).Text = Format(Now, "YYYY/MM/DD")



End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "邸別注文ﾃﾞｰﾀ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set SEK00201 = Nothing



    End

End Sub

Private Sub SHORI_Click(Index As Integer)

    Select Case Index
    
        Case 0
            
            If Command1(0).Enabled Then
            
                Command1(0).Value = True
            End If
        Case 1
            Command1(1).Value = True
        Case 2
            Command1(2).Value = True
    End Select



End Sub



Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   「注文データ」登録処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
    
Dim com             As Integer
Dim Upd_Com         As Integer

Dim Skip_Flg        As Integer

Dim Row             As Long



    If Y_Syuka_TEI.Count(1) = 0 Then
        Exit Function
    End If
    
    Update_Proc = True

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "積水ハウス邸別注文ﾃﾞｰﾀ確認　更新処理開始！！", Me.hwnd, 0)

                                    
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call Input_UnLock
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Update_Proc = SYS_ERR
        Exit Function
    End If
                                    
                                    
                                    
                                    
    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.Refresh
    
    TDBGrid1.Update
                                    
                                    
                                    
                                    
                                    
                                    'テーブルリセット
    For Row = 1 To Y_Syuka_TEI.UpperBound(1)
        
        
        DoEvents
        
        
        Call UniCode_Conv(K0_Y_SYU_TEI.SND_YMD, Y_Syuka_TEI(Row, colSND_YMD))
        Call UniCode_Conv(K0_Y_SYU_TEI.SND_HMS, Y_Syuka_TEI(Row, colSND_HMS))
        Call UniCode_Conv(K0_Y_SYU_TEI.SEQ_NO, Y_Syuka_TEI(Row, colSEQ_NO))
        
        
        Skip_Flg = False
        
        sts = BTRV(BtOpGetEqual, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
        Select Case sts
            Case BtNoErr
            
                Upd_Com = BtOpUpdate
            
                If Y_Syuka_TEI(Row, colDEL_FLG) = "1" Then
            
                    Upd_Com = BtOpDelete
                            
                End If
            
            Case BtErrKeyNotFound
                Upd_Com = BtOpInsert
            
                If Y_Syuka_TEI(Row, colDEL_FLG) = "1" Then
            
                    Skip_Flg = True
                            
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpInsert, "邸別注文データ", 0)
                GoTo Abort_Tran
        End Select
        
        
        
        
        If Not Skip_Flg And Upd_Com <> BtOpDelete Then
        
        
        
            Call UniCode_Conv(Y_SYU_TEI_REC.SND_YMD, Y_Syuka_TEI(Row, colSND_YMD))          'データ作成日
            Call UniCode_Conv(Y_SYU_TEI_REC.SND_HMS, Y_Syuka_TEI(Row, colSND_HMS))          'データ作成時刻
            Call UniCode_Conv(Y_SYU_TEI_REC.SEQ_NO, Y_Syuka_TEI(Row, colSEQ_NO))            '連番
            Call UniCode_Conv(Y_SYU_TEI_REC.JUC_YMD, Y_Syuka_TEI(Row, colJUC_YMD))          '受注日
            Call UniCode_Conv(Y_SYU_TEI_REC.NOU_CD, Y_Syuka_TEI(Row, colNOU_CD))            '納入受入場
            Call UniCode_Conv(Y_SYU_TEI_REC.NOU_NM, Y_Syuka_TEI(Row, colNOU_NM))            '納入受入場名
            Call UniCode_Conv(Y_SYU_TEI_REC.TOK_CD, Y_Syuka_TEI(Row, colTOK_CD))            '得意先ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYU_TEI_REC.CHO_CD, Y_Syuka_TEI(Row, colCHO_CD))            '直納先ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYU_TEI_REC.THINB_CD, Y_Syuka_TEI(Row, colTHINB_CD))        '得意先品番　■品番(上)
            Call UniCode_Conv(Y_SYU_TEI_REC.HINB_CD, Y_Syuka_TEI(Row, colHINB_CD))          '品番　■品番(下)
            Call UniCode_Conv(Y_SYU_TEI_REC.CHU_CD, Y_Syuka_TEI(Row, colCHU_CD))            '注文№　    ■指図№(上)
            Call UniCode_Conv(Y_SYU_TEI_REC.SYU_JUN, Y_Syuka_TEI(Row, colSYU_JUN))          '出荷順番　  ■指図№(下・左)
            Call UniCode_Conv(Y_SYU_TEI_REC.TEI_NM, Y_Syuka_TEI(Row, colTEI_NM))            '邸名　      ■指図№(下・右)
                                                                                            '受注数量
            Call UniCode_Conv(Y_SYU_TEI_REC.JUC_SUU, Format(Val(Y_Syuka_TEI(Row, colJUC_SUU)), "00000000"))
            Call UniCode_Conv(Y_SYU_TEI_REC.SYU_YMD, Y_Syuka_TEI(Row, colSYU_YMD))          '出荷確定日
            Call UniCode_Conv(Y_SYU_TEI_REC.NOU_YMD, Y_Syuka_TEI(Row, colNOU_YMD))          '納入日
            Call UniCode_Conv(Y_SYU_TEI_REC.KEN_NO, Y_Syuka_TEI(Row, colKEN_NO))            '件管№　　　■管理№(上)
            Call UniCode_Conv(Y_SYU_TEI_REC.HIN_NO, Y_Syuka_TEI(Row, colHIN_NO))            '件管№　　　■管理№(下)
            Call UniCode_Conv(Y_SYU_TEI_REC.TANP_KB, Y_Syuka_TEI(Row, colTANP_KB))          '単品区分
            Call UniCode_Conv(Y_SYU_TEI_REC.YOBI1_NM, "")                                   '予備
            
            
            If Upd_Com = BtOpInsert Then
                Call UniCode_Conv(Y_SYU_TEI_REC.GSEQ_NO, "00000")                           '総件数
            End If
            
            
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.TEI_LABELID, Y_Syuka_TEI(Row, colTEI_LABELID))  '邸別ﾗﾍﾞﾙID(注文№■指図№(上)+箱№)
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.HAKO_NO, Y_Syuka_TEI(Row, colHAKO_NO))          '箱№
            
            Call UniCode_Conv(Y_SYU_TEI_REC.JITU_SUU, "")                                   '実出庫数(梱包場への出庫数 現在未使用)
            Call UniCode_Conv(Y_SYU_TEI_REC.JITU_TANTO, "")                                 '出庫　担当者(現在未使用)
            Call UniCode_Conv(Y_SYU_TEI_REC.JITU_DATETIME, "")                              '出庫　日時(現在未使用)
            Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_TANTO, Y_Syuka_TEI(Row, colKONPO_TANTO))  '梱包　担当者
            
                                                                                            '梱包　日時
            If Trim(Y_Syuka_TEI(Row, colKONPO_TANTO)) <> "" Then
                If Trim(Y_Syuka_TEI(Row, colKONPO_DATETIME)) = "" Then
                    Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                Else
                    Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_DATETIME, Y_Syuka_TEI(Row, colKONPO_DATETIME))
                End If
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_DATETIME, "")
            End If
            
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_TANTO, Y_Syuka_TEI(Row, colSHOGO_TANTO))  '照合　担当者
            
                                                                                            '照合  日時
            If Trim(Y_Syuka_TEI(Row, colSHOGO_TANTO)) <> "" Then
                If Trim(Y_Syuka_TEI(Row, colSHOGO_DATETIME)) = "" Then
                    Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                Else
                    Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_DATETIME, Y_Syuka_TEI(Row, colSHOGO_DATETIME))
                End If
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_DATETIME, "")
            End If
            
                                                                                            '口数
            If IsNumeric(Y_Syuka_TEI(Row, colKUTI_SU)) Then
                Call UniCode_Conv(Y_SYU_TEI_REC.KUTI_SU, Format(Y_Syuka_TEI(Row, colKUTI_SU), "0000"))
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.KUTI_SU, "0000")
            End If
                                                                                            '才数
            If IsNumeric(Y_Syuka_TEI(Row, colSAI_SU)) Then
                Call UniCode_Conv(Y_SYU_TEI_REC.SAI_SU, Format(Y_Syuka_TEI(Row, colSAI_SU), "000.00"))
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.SAI_SU, "0000")
            End If
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_ID, Y_Syuka_TEI(Row, colKONPO_ID))        '梱包ID
            
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_TANTO, Y_Syuka_TEI(Row, colKENPIN_TANTO))    '検品　担当者
                                                                                                '検品  日時
            If Trim(Y_Syuka_TEI(Row, colKENPIN_TANTO)) <> "" Then
                If Trim(Y_Syuka_TEI(Row, colKENPIN_DATETIME)) = "" Then
                    Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                Else
                    Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_DATETIME, Y_Syuka_TEI(Row, colKENPIN_DATETIME))
                End If
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_DATETIME, "")
            End If
            
                                                                                                '集合梱包　担当者
            Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, Y_Syuka_TEI(Row, colSYUGO_KONPO_TANTO))
                                                                                                '集合梱包  日時
            If Trim(Y_Syuka_TEI(Row, colSYUGO_KONPO_TANTO)) <> "" Then
                If Trim(Y_Syuka_TEI(Row, colSYUGO_KONPO_DATETIME)) = "" Then
                    Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                Else
                    Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, Y_Syuka_TEI(Row, colSYUGO_KONPO_DATETIME))
                End If
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, "")
            End If
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.DATA_MAKE_DATETIME, "")                         '               2012.10.24
            Call UniCode_Conv(Y_SYU_TEI_REC.GAISO_IRI_QTY, "")                              '               2012.10.24
            Call UniCode_Conv(Y_SYU_TEI_REC.Y_HIN_CHK_CNT, "")                              '               2012.10.24
            Call UniCode_Conv(Y_SYU_TEI_REC.J_HIN_CHK_CNT, "")                              '               2012.10.24
    
    
            
            
                                                        
            Call UniCode_Conv(Y_SYU_TEI_REC.KEN_HINBAN, Y_Syuka_TEI(Row, colKEN_HINBAN))    '仕掛中品番     2012.10.24
                                                                                            '検品実績　バラ 2012.10.24
            Call UniCode_Conv(Y_SYU_TEI_REC.CNT_BARA_SU, Format(Y_Syuka_TEI(Row, colCNT_BARA_SU), "0000000"))
                                                                                            '検品実績　箱   2012.10.24
            Call UniCode_Conv(Y_SYU_TEI_REC.CNT_HAKO_SU, Format(Y_Syuka_TEI(Row, colCNT_HAKO_SU), "0000000"))
                                                                                            '検品実績　箱   2012.10.24
            Call UniCode_Conv(Y_SYU_TEI_REC.J_HIN_CHK_CNT, Format(Y_Syuka_TEI(Row, colJ_HIN_CHK_CNT), "0000000"))
            
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.FILLER, "")                                     'FILLER
            
            
            If Upd_Com = BtOpInsert Then
                                                                                            '追加担当者
                Call UniCode_Conv(Y_SYU_TEI_REC.INS_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                                            '追加日時
                Call UniCode_Conv(Y_SYU_TEI_REC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
            End If
            
            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))   '更新担当者
            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))    '更新日時
                    
                    
        End If
            
            
        Do
            sts = BTRV(Upd_Com, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    
                    Beep
                    ans = MsgBox("「邸別注文データ」他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                
                Case Else
                    Call File_Error(sts, Upd_Com, "邸別注文データ", 0)
                    GoTo Abort_Tran
            End Select
        
        Loop
            
        TDBGrid1.Bookmark = Row
        

    Next Row










                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        Update_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "積水ハウス邸別注文ﾃﾞｰﾀ確認　更新処理終了！！", Me.hwnd, 0)
    
    Call Input_UnLock
    
    Update_Proc = False
    
    Exit Function


Abort_Tran:
    
    Call Input_UnLock
    
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "積水ハウス邸別注文ﾃﾞｰﾀ確認　更新処理異常終了！！", Me.hwnd, 0)



End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「注文データ」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim Skip_F          As Boolean

Dim Row             As Long
Dim Y_SYU_H_CNT     As Long


Dim NullWork        As String

    List_Disp_Proc = True










hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "積水ハウス邸別注文ﾃﾞｰﾀ確認　表示処理開始！！", Me.hwnd, 0)
    DoEvents

                                    
    Call Input_Lock
                                    
                                    
                                    
                                    'テーブルリセット
    Set Y_Syuka_TEI = Nothing
    Row = Min_Row - 1
    Text1(ptxIn_Cnt).Text = Row
    
    
    
    Y_SYU_H_CNT = 0
    
    
    Call UniCode_Conv(K0_Y_SYU_TEI.SND_YMD, Format(Text1(ptxS_SND_YMD).Text, "YYYYMMDD"))
    Call UniCode_Conv(K0_Y_SYU_TEI.SND_HMS, "")
    Call UniCode_Conv(K0_Y_SYU_TEI.SEQ_NO, "")
        
    com = BtOpGetGreaterEqual
        
        
    Do

        DoEvents

        sts = BTRV(com, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
        Select Case sts
            Case BtNoErr
                            
                If StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode) > Format(Text1(ptxE_SND_YMD).Text, "YYYYMMDD") Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_Lock
                Call File_Error(sts, com, "邸別注文データ")
                Exit Function
        End Select



        Skip_F = False


        If Trim(Text1(ptxHINB_CD).Text) <> "" Then
            
            If Trim(Text1(ptxHINB_CD).Text) <> Mid(StrConv(Y_SYU_TEI_REC.HINB_CD, vbUnicode), 1, Len(Trim(Text1(ptxHINB_CD).Text))) Then

                Skip_F = True

            End If
        End If

        If Trim(Text1(ptxCHU_CD).Text) <> "" Then
            If Trim(Text1(ptxCHU_CD).Text) <> Mid(StrConv(Y_SYU_TEI_REC.CHU_CD, vbUnicode), 1, Len(Trim(Text1(ptxCHU_CD).Text))) Then

                Skip_F = True

            End If
        End If


        If Trim(Text1(ptxKEN_NO).Text) <> "" Then
            If Trim(Text1(ptxKEN_NO).Text) <> Mid(StrConv(Y_SYU_TEI_REC.KEN_NO, vbUnicode), 1, Len(Trim(Text1(ptxKEN_NO).Text))) Then

                Skip_F = True

            End If
        End If


        If Trim(Text1(ptxHIN_NO).Text) <> "" Then
            If Trim(Text1(ptxHIN_NO).Text) <> Mid(StrConv(Y_SYU_TEI_REC.HIN_NO, vbUnicode), 1, Len(Trim(Text1(ptxHIN_NO).Text))) Then

                Skip_F = True

            End If
        End If


        If Trim(Text1(ptxTEI_LABELID).Text) <> "" Then
            If Trim(Text1(ptxTEI_LABELID).Text) <> Mid(StrConv(Y_SYU_TEI_REC.TEI_LABELID, vbUnicode), 1, Len(Trim(Text1(ptxTEI_LABELID).Text))) Then

                Skip_F = True

            End If
        End If


        If Trim(Text1(ptxKONPO_ID).Text) <> "" Then
            If Trim(Text1(ptxKONPO_ID).Text) <> Mid(StrConv(Y_SYU_TEI_REC.KONPO_ID, vbUnicode), 1, Len(Trim(Text1(ptxKONPO_ID).Text))) Then

                Skip_F = True

            End If
        End If




        If Not Skip_F Then


'hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
'    "積水ハウス邸別注文ﾃﾞｰﾀ確認　表示処理開始！！ 処理年月日= " & Trim(StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode)), Me.hwnd, 0)
    DoEvents


            Row = Row + 1
            Y_Syuka_TEI.ReDim Min_Row, Row, Min_Col, Max_Col
            
            Y_Syuka_TEI(Row, colSND_YMD) = Trim(StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode))          'データ作成日
            Y_Syuka_TEI(Row, colSND_HMS) = Trim(StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode))          'データ作成時刻
            Y_Syuka_TEI(Row, colSEQ_NO) = Trim(StrConv(Y_SYU_TEI_REC.SEQ_NO, vbUnicode))            '連番
            Y_Syuka_TEI(Row, colJUC_YMD) = Trim(StrConv(Y_SYU_TEI_REC.JUC_YMD, vbUnicode))          '受注日
            Y_Syuka_TEI(Row, colNOU_CD) = Trim(StrConv(Y_SYU_TEI_REC.NOU_CD, vbUnicode))            '納入受入場
            Y_Syuka_TEI(Row, colNOU_NM) = Trim(StrConv(Y_SYU_TEI_REC.NOU_NM, vbUnicode))            '納入受入場名
            Y_Syuka_TEI(Row, colTOK_CD) = Trim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode))            '得意先ｺｰﾄﾞ
            Y_Syuka_TEI(Row, colCHO_CD) = Trim(StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode))            '直納先ｺｰﾄﾞ
            Y_Syuka_TEI(Row, colTHINB_CD) = Trim(StrConv(Y_SYU_TEI_REC.THINB_CD, vbUnicode))        '得意先品番　■品番(上)
            Y_Syuka_TEI(Row, colHINB_CD) = Trim(StrConv(Y_SYU_TEI_REC.HINB_CD, vbUnicode))          '品番　■品番(下)
            Y_Syuka_TEI(Row, colCHU_CD) = Trim(StrConv(Y_SYU_TEI_REC.CHU_CD, vbUnicode))            '注文№　    ■指図№(上)
            Y_Syuka_TEI(Row, colSYU_JUN) = Trim(StrConv(Y_SYU_TEI_REC.SYU_JUN, vbUnicode))          '出荷順番　  ■指図№(下・左)
            Y_Syuka_TEI(Row, colTEI_NM) = Trim(StrConv(Y_SYU_TEI_REC.TEI_NM, vbUnicode))            '邸名　      ■指図№(下・右)
                                                                                                    '受注数量
            Y_Syuka_TEI(Row, colJUC_SUU) = Format(Val(StrConv(Y_SYU_TEI_REC.JUC_SUU, vbUnicode)), "#0")
            Y_Syuka_TEI(Row, colSYU_YMD) = Trim(StrConv(Y_SYU_TEI_REC.SYU_YMD, vbUnicode))          '出荷確定日
            Y_Syuka_TEI(Row, colNOU_YMD) = Trim(StrConv(Y_SYU_TEI_REC.NOU_YMD, vbUnicode))          '納入日
            Y_Syuka_TEI(Row, colKEN_NO) = Trim(StrConv(Y_SYU_TEI_REC.KEN_NO, vbUnicode))            '件管№　　　■管理№(上)
            Y_Syuka_TEI(Row, colHIN_NO) = Trim(StrConv(Y_SYU_TEI_REC.HIN_NO, vbUnicode))            '件管№　　　■管理№(下)
            Y_Syuka_TEI(Row, colTANP_KB) = Trim(StrConv(Y_SYU_TEI_REC.TANP_KB, vbUnicode))          '単品区分
    
    
                                                                                                    '邸別ﾗﾍﾞﾙID
            Y_Syuka_TEI(Row, colTEI_LABELID) = Trim(StrConv(Y_SYU_TEI_REC.TEI_LABELID, vbUnicode))
            Y_Syuka_TEI(Row, colHAKO_NO) = Trim(StrConv(Y_SYU_TEI_REC.HAKO_NO, vbUnicode))          '箱№
            Y_Syuka_TEI(Row, colJITU_SUU) = Trim(StrConv(Y_SYU_TEI_REC.JITU_SUU, vbUnicode))        '実出庫数
                                                                                            
            Y_Syuka_TEI(Row, colJITU_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.JITU_TANTO, vbUnicode))    '出庫　担当者
                                                                                                    '出庫　日時
            Y_Syuka_TEI(Row, colJITU_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.JITU_DATETIME, vbUnicode))
                                                                                                    
            Y_Syuka_TEI(Row, colKONPO_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.KONPO_TANTO, vbUnicode))  '梱包　担当者
                                                                                                    '梱包  日時
            Y_Syuka_TEI(Row, colKONPO_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.KONPO_DATETIME, vbUnicode))
                                                                                                    
            Y_Syuka_TEI(Row, colSHOGO_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.SHOGO_TANTO, vbUnicode))  '照合　担当者
                                                                                                    '照合  日時
            Y_Syuka_TEI(Row, colSHOGO_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.SHOGO_DATETIME, vbUnicode))
                                                                                            
                                                                                                    '口数
            Y_Syuka_TEI(Row, colKUTI_SU) = Format(Val(StrConv(Y_SYU_TEI_REC.KUTI_SU, vbUnicode)), "#0")
                                                                                                    '才数
            Y_Syuka_TEI(Row, colSAI_SU) = Format(Val(StrConv(Y_SYU_TEI_REC.SAI_SU, vbUnicode)), "#0.00")
                                                                                            
            Y_Syuka_TEI(Row, colKONPO_ID) = Trim(StrConv(Y_SYU_TEI_REC.KONPO_ID, vbUnicode))        '梱包ID
                                                                                            
                                                                                                    '検品　担当者
            Y_Syuka_TEI(Row, colKENPIN_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.KENPIN_TANTO, vbUnicode))
                                                                                                    '検品  日時
            Y_Syuka_TEI(Row, colKENPIN_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.KENPIN_DATETIME, vbUnicode))
                                                                                                    '検品　担当者
            Y_Syuka_TEI(Row, colSYUGO_KONPO_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, vbUnicode))
                                                                                                    '検品  日時
            Y_Syuka_TEI(Row, colSYUGO_KONPO_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, vbUnicode))
                                                                                            
                                                                                            
                                                                                            
                                                                                            
                                                                                            
                                                                                            
                                                                                            
                                                                                            
                                                                                            
                                                                                            
                                                                                            
                                                                                            
                                                                                            
                                                                                            
            Y_Syuka_TEI(Row, colINS_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.INS_TANTO, vbUnicode))      '追加　担当者
                                                                                                    '追加  日時
            Y_Syuka_TEI(Row, colINS_DateTime) = Trim(StrConv(Y_SYU_TEI_REC.Ins_DateTime, vbUnicode))
                                                                                            
            Y_Syuka_TEI(Row, colUPD_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.UPD_TANTO, vbUnicode))      '更新　担当者
                                                                                                    '更新  日時
            Y_Syuka_TEI(Row, colUPD_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.UPD_DATETIME, vbUnicode))
    
    
            
            
'            NullWork = "*"
'            If Mid(StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode), 1, 1) < " " Then
'                NullWork = "1"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "2"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.SEQ_NO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "3"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.JUC_YMD, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "4"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.NOU_CD, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "5"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.NOU_NM, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "6"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "7"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "8"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.THINB_CD, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "9"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.HINB_CD, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "10"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.CHU_CD, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "11"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.SYU_JUN, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "12"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.TEI_NM, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "13"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.JUC_SUU, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "14"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.SYU_YMD, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "15"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.NOU_YMD, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "16"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.KEN_NO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "17"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.HIN_NO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "18"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.TANP_KB, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "19"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.YOBI1_NM, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "20"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.GSEQ_NO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "21"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.TEI_LABELID, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "22"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.HAKO_NO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "23"
'            End If
'
'
'            If Mid(StrConv(Y_SYU_TEI_REC.JITU_SUU, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "24"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.JITU_TANTO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "25"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.JITU_DATETIME, vbUnicode), 1, 1) < " " Then
'               NullWork = NullWork & "26"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.KONPO_TANTO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "27"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.KONPO_DATETIME, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "28"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.SHOGO_TANTO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "29"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.SHOGO_DATETIME, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "30"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_KENKAN, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "31"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_TEI_NAME, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "32"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_TOK_NAME, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "33"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_SOTO_NO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "34"
'            End If
'
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_UCHI_NO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "35"
'            End If
'
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_WIDTH, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "36"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_HEIGHT, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "37"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_CONTENT, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "38"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_KNo, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "39"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_SERIES1, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "40"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_SERIES2, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "41"
'            End If
'
'
'            If Mid(StrConv(Y_SYU_TEI_REC.L_PAGE, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "42"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.KUTI_SU, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "43"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.SAI_SU, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "44"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.KONPO_ID, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "45"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.KENPIN_TANTO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "46"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.KENPIN_DATETIME, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "47"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "48"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "49"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.FILLER, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "50"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.INS_TANTO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "51"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.Ins_DateTime, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "52"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.UPD_TANTO, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "53"
'            End If
'
'            If Mid(StrConv(Y_SYU_TEI_REC.UPD_DATETIME, vbUnicode), 1, 1) < " " Then
'                NullWork = NullWork & "54"
'            End If
    
    
            Call UniCode_Conv(K8_Y_SYU_H.SEK_KEN_NO, StrConv(Y_SYU_TEI_REC.KEN_NO, vbUnicode))
            Call UniCode_Conv(K8_Y_SYU_H.SEK_HIN_NO, StrConv(Y_SYU_TEI_REC.HIN_NO, vbUnicode))
    
    
    
            sts = BTRV(BtOpGetEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K8_Y_SYU_H, Len(K8_Y_SYU_H), 8)
            Select Case sts
                Case BtNoErr
                    Y_SYU_H_CNT = Y_SYU_H_CNT + 1
                Case BtErrKeyNotFound
                    Call UniCode_Conv(Y_SYU_HREC.ID_NO, "")
                Case Else
                    Call Input_Lock
                    Call File_Error(sts, BtOpGetEqual, "出荷予定(H)")
                    Exit Function
            End Select
    
            Y_Syuka_TEI(Row, colID_NO) = Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
            
    If Right(Format(Row, "000000"), 3) = "000" Then
            
        Set TDBGrid1.Array = Y_Syuka_TEI
        TDBGrid1.ReBind
        
        TDBGrid1.Update
            
    End If
            
                                                                        '仕掛中　品番   2012.10.24
            Y_Syuka_TEI(Row, colKEN_HINBAN) = Trim(StrConv(Y_SYU_TEI_REC.KEN_HINBAN, vbUnicode))
                                                                        '検品実績　バラ 2012.10.24
            Y_Syuka_TEI(Row, colCNT_BARA_SU) = Trim(StrConv(Y_SYU_TEI_REC.CNT_BARA_SU, vbUnicode))
                                                                        '検品実績　箱 2012.10.24
            Y_Syuka_TEI(Row, colCNT_HAKO_SU) = Trim(StrConv(Y_SYU_TEI_REC.CNT_HAKO_SU, vbUnicode))
                                                                        '品番読込み実績 2012.10.24
            Y_Syuka_TEI(Row, colJ_HIN_CHK_CNT) = Trim(StrConv(Y_SYU_TEI_REC.J_HIN_CHK_CNT, vbUnicode))
            
            
            
            Text1(ptxIn_Cnt).Text = Format(Y_SYU_H_CNT, "#") & "/" & Format(Row, "#")

        End If
        
        com = BtOpGetNext

    Loop


    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "積水ハウス邸別注文ﾃﾞｰﾀ確認　表示処理終了！！", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_Proc = False
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description

End Function

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
    
Dim lngPFstRow  As Long
Dim vntBmk      As Variant
Dim intLeftCol  As Integer
Dim intCol      As Integer
Dim lngCFstRow  As Long
    
    
    
    
    If Y_Syuka_TEI.Count(1) < 1 Then
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
                    
        Y_Syuka_TEI.QuickSort Min_Row, Y_Syuka_TEI.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = Y_Syuka_TEI
        
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

    Call Tab_Ctrl(Shift)        '移動


End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    SEK00201.MousePointer = vbHourglass

    TDBGrid1.Enabled = False

    Call Ctrl_Lock(SEK00201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEK00201)

    TDBGrid1.Enabled = True

    SEK00201.MousePointer = vbDefault

End Sub


Private Function CSV_DATA_OUT_PROC() As Integer
'----------------------------------------------------------------------------
'                   「注文データ」CSV出力処理
'----------------------------------------------------------------------------
Dim FileNo          As Long
Dim i               As Long

Dim sts             As Integer


    If Y_Syuka_TEI.Count(1) < 1 Then
        Exit Function
    End If



    CSV_DATA_OUT_PROC = True
    
    
    
    
    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    Open (CSV_FILE) For Output As FileNo
    On Error GoTo 0

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "積水ハウス邸別注文ﾃﾞｰﾀ確認　ＣＳＶ出力処理開始！！", Me.hwnd, 0)



    For i = 1 To Y_Syuka_TEI.Count(1)
    
        DoEvents
    
        Text1(ptxIn_Cnt).Text = Format(i, "#0")
        
        DoEvents
    
            
    
        Write #FileNo, "*" & Trim(Y_Syuka_TEI(i, colHINB_CD)) & "*",
    
        
        If Trim(Y_Syuka_TEI(i, colTEI_LABELID)) = "" Then
            Write #FileNo, ,
        Else
            Write #FileNo, "*" & Trim(Y_Syuka_TEI(i, colTEI_LABELID)) & "*",
        End If
                
        If Trim(Y_Syuka_TEI(i, colKONPO_ID)) = "" Then
            Write #FileNo, ,
        Else
            Write #FileNo, "*" & Trim(Y_Syuka_TEI(i, colKONPO_ID)) & "*",
        End If
            
            
        Write #FileNo, Val(Y_Syuka_TEI(i, colJUC_SUU)),
                    
                    
        If IsNumeric(Y_Syuka_TEI(i, colID_NO)) Then
            Write #FileNo, "*" & Left(Y_Syuka_TEI(i, colID_NO), 7) & "*",
        Else                                                '2013.08.19
            Write #FileNo, ,                                '2013.08.19
        End If
            
        Write #FileNo, Val(Y_Syuka_TEI(i, colKUTI_SU)),     '2013.08.19
            
        '>>>>>>>>>>>>>>>    2013.02.02
'        Call UniCode_Conv(K0_ITEM.JGYOBU, SETSUBI)
'        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'        Call UniCode_Conv(K0_ITEM.HIN_GAI, Y_Syuka_TEI(i, colHINB_CD))
'
'        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'        Select Case sts
'            Case BtNoErr
'                Write #FileNo, StrConv(ITEMREC.KONPOU_F, vbUnicode),
'            Case BtErrKeyNotFound
'                Write #FileNo, "未登録",
'            Case Else
'                Call Input_Lock
'                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                Exit Function
'        End Select
        '>>>>>>>>>>>>>>>    2013.02.02
            
        Write #FileNo,
    Next i



    MsgBox ("「" & CSV_FILE & "」出力終了！！")
    


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "積水ハウス邸別注文ﾃﾞｰﾀ確認　ＣＳＶ出力処理終了！！", Me.hwnd, 0)

    Close #FileNo


    CSV_DATA_OUT_PROC = False
    Exit Function
    
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox CSV_FILE & "が使用中です。"
    Else
        MsgBox "Err.Number= " & Err.Number & " " & Err.Description
            
    End If





End Function
