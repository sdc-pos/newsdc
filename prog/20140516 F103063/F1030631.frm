VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1030631 
   BackColor       =   &H00FFFFFF&
   Caption         =   "出荷問い合わせ(問合せ№)"
   ClientHeight    =   8880
   ClientLeft      =   2025
   ClientTop       =   2550
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
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   15240
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   12705
      ScaleHeight     =   315
      ScaleWidth      =   480
      TabIndex        =   38
      Top             =   7920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton Command 
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
      Left            =   10320
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "印 刷"
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "ﾃﾞｰﾀ"
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
      Left            =   7830
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "表 示"
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8040
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
      Index           =   3
      Left            =   2640
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8040
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   105
      TabIndex        =   28
      Top             =   240
      Width           =   14925
      Begin VB.TextBox Text 
         Alignment       =   1  '右揃え
         Height          =   360
         Index           =   8
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   44
         Top             =   1320
         Width           =   1908
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   7
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1320
         Width           =   300
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   1
         Left            =   5670
         MaxLength       =   7
         TabIndex        =   3
         Top             =   240
         Width           =   930
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   1
         Left            =   4515
         TabIndex        =   2
         Top             =   360
         Width           =   225
      End
      Begin VB.CheckBox Check1 
         Caption         =   "検品済"
         Height          =   255
         Index           =   1
         Left            =   5775
         TabIndex        =   14
         Top             =   1440
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "未検品"
         Height          =   255
         Index           =   0
         Left            =   4515
         TabIndex        =   13
         Top             =   1440
         Width           =   1065
      End
      Begin VB.ComboBox Combo 
         Height          =   360
         Index           =   1
         Left            =   12810
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   11
         Top             =   840
         Width           =   1965
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   6
         Left            =   9345
         MaxLength       =   20
         TabIndex        =   10
         Top             =   840
         Width           =   2505
      End
      Begin VB.ComboBox Combo 
         Height          =   360
         Index           =   0
         Left            =   5565
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   9
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   5
         Left            =   4515
         MaxLength       =   8
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   2
         Left            =   1470
         MaxLength       =   4
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   3
         Left            =   2415
         MaxLength       =   2
         TabIndex        =   6
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   4
         Left            =   3060
         MaxLength       =   2
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   4
         Top             =   960
         Width           =   225
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   0
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   2505
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "表示件数"
         Height          =   240
         Index           =   11
         Left            =   8388
         TabIndex        =   43
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "(空白：全て)"
         Height          =   240
         Index           =   10
         Left            =   1890
         TabIndex        =   40
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "便"
         Height          =   240
         Index           =   9
         Left            =   1155
         TabIndex        =   39
         Top             =   1440
         Width           =   240
      End
      Begin VB.Line Line2 
         X1              =   4200
         X2              =   4200
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "問合せ№"
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   37
         Top             =   375
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "伝票№"
         Height          =   240
         Index           =   8
         Left            =   4830
         TabIndex        =   36
         Top             =   360
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   105
         X2              =   14700
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "運送会社"
         Height          =   240
         Index           =   3
         Left            =   11865
         TabIndex        =   35
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "品番"
         Height          =   240
         Index           =   2
         Left            =   8715
         TabIndex        =   34
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "得意先"
         Height          =   240
         Index           =   0
         Left            =   3780
         TabIndex        =   33
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "出 荷 日"
         Height          =   240
         Index           =   4
         Left            =   420
         TabIndex        =   32
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "年"
         Height          =   240
         Index           =   5
         Left            =   2100
         TabIndex        =   31
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "月"
         Height          =   240
         Index           =   6
         Left            =   2820
         TabIndex        =   30
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   240
         Index           =   7
         Left            =   3420
         TabIndex        =   29
         Top             =   960
         Width           =   240
      End
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   5655
      Left            =   105
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2040
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   9975
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "問合せ№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ｷｬﾝｾﾙ"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "口"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "№"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "出荷日"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "便"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "集約CD"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "送り先CD"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "送り先名"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "売伝"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "伝票番号"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "品番／品名"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "数量"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "出庫済"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "検品担当者"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "注文№"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "得意先"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "備考"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "運送会社"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "検品日時"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "ID_NO"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "備　考"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "枝番　～"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "才数"
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "梱包Ｆ"
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "口数（単体）"
      Columns(26).DataField=   ""
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "才数(単体)"
      Columns(27).DataField=   ""
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "重量"
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "住所"
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "郵便番号"
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "電話番号"
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "使用状況"
      Columns(32).DataField=   ""
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "件管№"
      Columns(33).DataField=   ""
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "品管№"
      Columns(34).DataField=   ""
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).DataField=   ""
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).DataField=   ""
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   37
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=37"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2434"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2328"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1217"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1111"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=767"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=661"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=900"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=794"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=1958"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1852"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=1032"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=926"
      Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=2143"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2037"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=1958"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=1852"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(8).Width=2143"
      Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=2037"
      Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(41)=   "Column(9).Width=900"
      Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=794"
      Splits(0)._ColumnProps(44)=   "Column(9)._ColStyle=1"
      Splits(0)._ColumnProps(45)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(46)=   "Column(10).Width=1931"
      Splits(0)._ColumnProps(47)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(10)._WidthInPix=1826"
      Splits(0)._ColumnProps(49)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(50)=   "Column(11).Width=2831"
      Splits(0)._ColumnProps(51)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(11)._WidthInPix=2725"
      Splits(0)._ColumnProps(53)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(54)=   "Column(12).Width=1191"
      Splits(0)._ColumnProps(55)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(12)._WidthInPix=1085"
      Splits(0)._ColumnProps(57)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(58)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(59)=   "Column(13).Width=1296"
      Splits(0)._ColumnProps(60)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(13)._WidthInPix=1191"
      Splits(0)._ColumnProps(62)=   "Column(13)._ColStyle=2"
      Splits(0)._ColumnProps(63)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(64)=   "Column(14).Width=2461"
      Splits(0)._ColumnProps(65)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(14)._WidthInPix=2355"
      Splits(0)._ColumnProps(67)=   "Column(14)._ColStyle=0"
      Splits(0)._ColumnProps(68)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(69)=   "Column(15).Width=1640"
      Splits(0)._ColumnProps(70)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(15)._WidthInPix=1535"
      Splits(0)._ColumnProps(72)=   "Column(15)._ColStyle=0"
      Splits(0)._ColumnProps(73)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(74)=   "Column(16).Width=2408"
      Splits(0)._ColumnProps(75)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(16)._WidthInPix=2302"
      Splits(0)._ColumnProps(77)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(78)=   "Column(17).Width=1931"
      Splits(0)._ColumnProps(79)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(17)._WidthInPix=1826"
      Splits(0)._ColumnProps(81)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(82)=   "Column(18).Width=1746"
      Splits(0)._ColumnProps(83)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(84)=   "Column(18)._WidthInPix=1640"
      Splits(0)._ColumnProps(85)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(86)=   "Column(19).Width=1931"
      Splits(0)._ColumnProps(87)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(88)=   "Column(19)._WidthInPix=1826"
      Splits(0)._ColumnProps(89)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(90)=   "Column(20).Width=2514"
      Splits(0)._ColumnProps(91)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(92)=   "Column(20)._WidthInPix=2408"
      Splits(0)._ColumnProps(93)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(94)=   "Column(21).Width=3810"
      Splits(0)._ColumnProps(95)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(96)=   "Column(21)._WidthInPix=3704"
      Splits(0)._ColumnProps(97)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(98)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(99)=   "Column(22).Width=4366"
      Splits(0)._ColumnProps(100)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(101)=   "Column(22)._WidthInPix=4260"
      Splits(0)._ColumnProps(102)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(103)=   "Column(23).Width=2170"
      Splits(0)._ColumnProps(104)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(105)=   "Column(23)._WidthInPix=2064"
      Splits(0)._ColumnProps(106)=   "Column(23)._ColStyle=1"
      Splits(0)._ColumnProps(107)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(108)=   "Column(24).Width=1217"
      Splits(0)._ColumnProps(109)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(110)=   "Column(24)._WidthInPix=1111"
      Splits(0)._ColumnProps(111)=   "Column(24)._ColStyle=2"
      Splits(0)._ColumnProps(112)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(113)=   "Column(25).Width=1217"
      Splits(0)._ColumnProps(114)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(25)._WidthInPix=1111"
      Splits(0)._ColumnProps(116)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(117)=   "Column(26).Width=2381"
      Splits(0)._ColumnProps(118)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(119)=   "Column(26)._WidthInPix=2275"
      Splits(0)._ColumnProps(120)=   "Column(26)._ColStyle=2"
      Splits(0)._ColumnProps(121)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(122)=   "Column(27).Width=3810"
      Splits(0)._ColumnProps(123)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(124)=   "Column(27)._WidthInPix=3704"
      Splits(0)._ColumnProps(125)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(126)=   "Column(28).Width=1455"
      Splits(0)._ColumnProps(127)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(128)=   "Column(28)._WidthInPix=1349"
      Splits(0)._ColumnProps(129)=   "Column(28)._ColStyle=2"
      Splits(0)._ColumnProps(130)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(131)=   "Column(29).Width=3810"
      Splits(0)._ColumnProps(132)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(133)=   "Column(29)._WidthInPix=3704"
      Splits(0)._ColumnProps(134)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(135)=   "Column(30).Width=3810"
      Splits(0)._ColumnProps(136)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(137)=   "Column(30)._WidthInPix=3704"
      Splits(0)._ColumnProps(138)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(139)=   "Column(31).Width=3810"
      Splits(0)._ColumnProps(140)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(141)=   "Column(31)._WidthInPix=3704"
      Splits(0)._ColumnProps(142)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(143)=   "Column(32).Width=1746"
      Splits(0)._ColumnProps(144)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(145)=   "Column(32)._WidthInPix=1640"
      Splits(0)._ColumnProps(146)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(147)=   "Column(33).Width=2328"
      Splits(0)._ColumnProps(148)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(149)=   "Column(33)._WidthInPix=2223"
      Splits(0)._ColumnProps(150)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(151)=   "Column(34).Width=2037"
      Splits(0)._ColumnProps(152)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(153)=   "Column(34)._WidthInPix=1931"
      Splits(0)._ColumnProps(154)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(155)=   "Column(35).Width=3810"
      Splits(0)._ColumnProps(156)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(157)=   "Column(35)._WidthInPix=3704"
      Splits(0)._ColumnProps(158)=   "Column(35).Visible=0"
      Splits(0)._ColumnProps(159)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(160)=   "Column(36).Width=3810"
      Splits(0)._ColumnProps(161)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(162)=   "Column(36)._WidthInPix=3704"
      Splits(0)._ColumnProps(163)=   "Column(36).Visible=0"
      Splits(0)._ColumnProps(164)=   "Column(36).Order=37"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=60,.parent=9,.bgcolor=&HFF80&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=61,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=88,.parent=87"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=91,.parent=90"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=14,.parent=53"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=11,.parent=54"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=12,.parent=55"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=13,.parent=57"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=110,.parent=53,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=54"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=55"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=57"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=130,.parent=53,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=127,.parent=54"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=128,.parent=55"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=129,.parent=57"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=18,.parent=53,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=15,.parent=54"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=16,.parent=55"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=17,.parent=57"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=48,.parent=53"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=45,.parent=54"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=46,.parent=55"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=47,.parent=57"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=114,.parent=53,.alignment=2"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=111,.parent=54"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=112,.parent=55"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=113,.parent=57"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=122,.parent=53"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=119,.parent=54"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=120,.parent=55"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=121,.parent=57"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=126,.parent=53"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=123,.parent=54"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=124,.parent=55"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=125,.parent=57"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=53"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=54"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=55"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=57"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=102,.parent=53,.alignment=2"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=19,.parent=54"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=20,.parent=55"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=101,.parent=57"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=70,.parent=53"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=54"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=55"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=57"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=78,.parent=53"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=54"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=55"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=57"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=53,.alignment=1,.locked=0"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=54,.alignment=3"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=55,.alignment=3"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=57"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=118,.parent=53,.alignment=1"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=115,.parent=54"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=116,.parent=55"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=117,.parent=57"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=40,.parent=53,.alignment=0"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=37,.parent=54"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=38,.parent=55"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=39,.parent=57"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=24,.parent=53,.alignment=0,.locked=0"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=21,.parent=54,.alignment=3"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=22,.parent=55,.alignment=3"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=23,.parent=57"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=28,.parent=53"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=25,.parent=54"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=26,.parent=55"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=27,.parent=57"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=44,.parent=53"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=41,.parent=54"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=42,.parent=55"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=43,.parent=57"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=74,.parent=53"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=71,.parent=54"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=72,.parent=55"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=73,.parent=57"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=52,.parent=53"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=49,.parent=54"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=50,.parent=55"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=51,.parent=57"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=82,.parent=53"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=79,.parent=54"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=80,.parent=55"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=81,.parent=57"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=100,.parent=53"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=97,.parent=54"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=98,.parent=55"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=99,.parent=57"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=106,.parent=53"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=103,.parent=54"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=104,.parent=55"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=105,.parent=57"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=96,.parent=53,.alignment=2"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=93,.parent=54"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=94,.parent=55"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=95,.parent=57"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=146,.parent=53,.alignment=1"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=143,.parent=54"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=144,.parent=55"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=145,.parent=57"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=134,.parent=53"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=131,.parent=54"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=132,.parent=55"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=133,.parent=57"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=138,.parent=53,.alignment=1"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=135,.parent=54"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=136,.parent=55"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=137,.parent=57"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=142,.parent=53"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=139,.parent=54"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=140,.parent=55"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=141,.parent=57"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=150,.parent=53,.alignment=1"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=147,.parent=54"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=148,.parent=55"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=149,.parent=57"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=154,.parent=53"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=151,.parent=54"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=152,.parent=55"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=153,.parent=57"
      _StyleDefs(156) =   "Splits(0).Columns(30).Style:id=158,.parent=53"
      _StyleDefs(157) =   "Splits(0).Columns(30).HeadingStyle:id=155,.parent=54"
      _StyleDefs(158) =   "Splits(0).Columns(30).FooterStyle:id=156,.parent=55"
      _StyleDefs(159) =   "Splits(0).Columns(30).EditorStyle:id=157,.parent=57"
      _StyleDefs(160) =   "Splits(0).Columns(31).Style:id=162,.parent=53"
      _StyleDefs(161) =   "Splits(0).Columns(31).HeadingStyle:id=159,.parent=54"
      _StyleDefs(162) =   "Splits(0).Columns(31).FooterStyle:id=160,.parent=55"
      _StyleDefs(163) =   "Splits(0).Columns(31).EditorStyle:id=161,.parent=57"
      _StyleDefs(164) =   "Splits(0).Columns(32).Style:id=166,.parent=53"
      _StyleDefs(165) =   "Splits(0).Columns(32).HeadingStyle:id=163,.parent=54"
      _StyleDefs(166) =   "Splits(0).Columns(32).FooterStyle:id=164,.parent=55"
      _StyleDefs(167) =   "Splits(0).Columns(32).EditorStyle:id=165,.parent=57"
      _StyleDefs(168) =   "Splits(0).Columns(33).Style:id=170,.parent=53"
      _StyleDefs(169) =   "Splits(0).Columns(33).HeadingStyle:id=167,.parent=54"
      _StyleDefs(170) =   "Splits(0).Columns(33).FooterStyle:id=168,.parent=55"
      _StyleDefs(171) =   "Splits(0).Columns(33).EditorStyle:id=169,.parent=57"
      _StyleDefs(172) =   "Splits(0).Columns(34).Style:id=174,.parent=53"
      _StyleDefs(173) =   "Splits(0).Columns(34).HeadingStyle:id=171,.parent=54"
      _StyleDefs(174) =   "Splits(0).Columns(34).FooterStyle:id=172,.parent=55"
      _StyleDefs(175) =   "Splits(0).Columns(34).EditorStyle:id=173,.parent=57"
      _StyleDefs(176) =   "Splits(0).Columns(35).Style:id=178,.parent=53"
      _StyleDefs(177) =   "Splits(0).Columns(35).HeadingStyle:id=175,.parent=54"
      _StyleDefs(178) =   "Splits(0).Columns(35).FooterStyle:id=176,.parent=55"
      _StyleDefs(179) =   "Splits(0).Columns(35).EditorStyle:id=177,.parent=57"
      _StyleDefs(180) =   "Splits(0).Columns(36).Style:id=182,.parent=53"
      _StyleDefs(181) =   "Splits(0).Columns(36).HeadingStyle:id=179,.parent=54"
      _StyleDefs(182) =   "Splits(0).Columns(36).FooterStyle:id=180,.parent=55"
      _StyleDefs(183) =   "Splits(0).Columns(36).EditorStyle:id=181,.parent=57"
      _StyleDefs(184) =   "Named:id=29:Normal"
      _StyleDefs(185) =   ":id=29,.parent=0"
      _StyleDefs(186) =   "Named:id=30:Heading"
      _StyleDefs(187) =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(188) =   ":id=30,.wraptext=-1"
      _StyleDefs(189) =   "Named:id=31:Footing"
      _StyleDefs(190) =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(191) =   "Named:id=32:Selected"
      _StyleDefs(192) =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(193) =   "Named:id=33:Caption"
      _StyleDefs(194) =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(195) =   "Named:id=34:HighlightRow"
      _StyleDefs(196) =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(197) =   "Named:id=35:EvenRow"
      _StyleDefs(198) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(199) =   "Named:id=36:OddRow"
      _StyleDefs(200) =   ":id=36,.parent=29"
      _StyleDefs(201) =   "Named:id=89:RecordSelector"
      _StyleDefs(202) =   ":id=89,.parent=30"
      _StyleDefs(203) =   "Named:id=92:FilterBar"
      _StyleDefs(204) =   ":id=92,.parent=29"
   End
   Begin VB.Label lblRow 
      Caption         =   "Label1"
      Height          =   315
      Left            =   11610
      TabIndex        =   42
      Top             =   7980
      Visible         =   0   'False
      Width           =   735
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
      Left            =   210
      TabIndex        =   27
      Top             =   8520
      Width           =   180
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
Attribute VB_Name = "F1030631"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const poptSEL_OKURI_NO% = 0     '問合せ№選択
Private Const poptSEL_DEN_NO% = 1       '伝票№選択
Private Const poptSEL_ETC% = 2          'その他選択



Private Const ptxSEL_OKURI_NO% = 0      '問合せ№
Private Const ptxSEL_DEN_NO% = 1        '伝票№
Private Const ptxSEL_SYUKA_YY% = 2      '出荷日(年)
Private Const ptxSEL_SYUKA_MM% = 3      '出荷日(年)
Private Const ptxSEL_SYUKA_DD% = 4      '出荷日(年)

Private Const ptxSEL_MUKE_CODE% = 5     '出荷先
Private Const ptxSEL_HIN_NO% = 6        '品番

Private Const ptxSEL_INS_BIN% = 7       '便             2007.01.16

Private Const ptxDISP_CNT% = 8          '表示件数       2007.01.16



Private Const pcmbSEL_MUKE_NAME% = 0    '出荷先名
Private Const pcmbSEL_UNSOU_KAISHA% = 1 '運送会社

Private Const pchkMI% = 0               '未処理
Private Const pchkSUMI% = 1             '処理済

Private Const UNSOU_KAISHA_ALL = "全  て"


Dim SYUKA As New XArrayDB

Private Const Min_Row% = 1              '最小行数
Dim Max_Row    As Long                  'グリッド最大表示件数



Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 36                 '最大列数       2011.04.23

Private Const ColOKURI_NO% = 0              '問合せ№
Private Const ColCANCEL_F% = 1              'ｷｬﾝｾﾙﾌﾗｸﾞ(X)

Private Const ColKUTI_SU% = 2              '口数   表示位置移動    2007.07.10

Private Const ColSYUKA_NO% = 3              '№
Private Const ColSYUKA_YMD% = 4             '出荷日
Private Const ColINS_BIN% = 5               '便

Private Const ColCOL_OKURISAKI_CD% = 6      '集約送り先CD   2007.07.09
Private Const ColOKURISAKI_CD% = 7          '送り先CD       2007.07.09



Private Const ColOKURISAKI% = 8             '送り先
Private Const ColURIDEN% = 9                '売伝
Private Const ColDEN_NO% = 10                '伝票番号
Private Const ColHIN_NO% = 11                '品番／品名
Private Const ColSURYO% = 12                 '数量
Private Const ColJ_SURYO% = 13              '実績

Private Const ColKENPIN_TANTO_CODE% = 14    '検品担当者 2007.07.10 検品担当者位置移動

Private Const ColORDER_NO% = 15              '注文№
Private Const ColMUKE_CODE% = 16             '得意先
Private Const ColBIKOU% = 17                '備考
Private Const ColUNSOU_KAISHA% = 18         '運送会社
Private Const ColKENPIN_NOW% = 19           '検品日時
Private Const ColID_NO% = 20                'ID_NO

Private Const ColCANCEL_F_A% = 21           'ｷｬﾝｾﾙﾌﾗｸﾞ
Private Const ColINPUT_BIKOU% = 22          '備考（入力値）



Private Const ColOKURI_NO_SEQ% = 23         '送り状枝番 2010.02.11
Private Const ColSAI_SU% = 24               '才数 2010.02.11
Private Const ColKONPOU_F% = 25             '梱包Ｆ 2010.02.11
Private Const ColKUTI_SU_TAN% = 26          '口数 2010.02.11
Private Const ColSAI_SU_TAN% = 27           '才数 2010.02.11


Private Const ColJURYO% = 28                '重量 2010.02.15
Private Const ColJYUSHO% = 29               '住所 2010.02.15


Private Const ColYUBIN_NO% = 30             '郵便番号 2010.02.15
Private Const ColTEL_NO% = 31               '電話番号 2010.02.15

Private Const ColUSE% = 32                  '使用中 2011.04.23

Private Const ColSEK_KEN_NO% = 33           '件管№　　　■管理№(上)   2011.05.18
Private Const ColSEK_HIN_NO% = 34           '品管№　　　■管理№(下)   2011.05.18


Private Const ColY_SYU_INS% = 35            '件管№　　　■管理№(上)   2011.05.18
Private Const ColY_SYU_H_INS% = 36          '品管№　　　■管理№(下)   2011.05.18



Private GW_PATH     As String               'CSV出力パス

Private DATA_FLG    As Boolean

Private wDEL_SYU_H_POS      As POSBLK
Private K0_wDEL_SYU_H       As KEY0_DEL_SYU_H


'2010.08.31
Private Sort_Tbl(ColOKURI_NO To ColSEK_HIN_NO) _
                            As Integer                  'ｿｰﾄの制御 0:昇順 1:降順


Dim HEAD_CLICK              As Integer

Dim IN_CNT                  As Long
Dim SUMI_CNT                As Long
Dim CANCEL_CNT              As Long



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2014.05.14
Dim FUKUYAMA_Name           As String
Dim FUKUYAMA_Length         As Integer
Dim FUKUYAMA_CODE()         As String



Dim SEIBU_Name              As String
Dim SEIBU_Length            As Integer
Dim SEIBU_CODE()            As String


Dim KURUME_Name             As String
Dim KURUME_Length           As Integer
Dim KURUME_CODE()           As String


Dim SAGAWA_Name             As String
Dim SAGAWA_Length           As Integer
Dim SAGAWA_CODE()           As String

Dim YAMATO_Name             As String
Dim YAMATO_Length           As Integer
Dim YAMATO_CODE()           As String



Dim PANA_Name               As String
Dim PANA_Length             As Integer
Dim PANA_CODE()             As String

Dim SEKISUI_Name            As String
Dim SEKISUI_Length          As Integer
Dim SEKISUI_CODE()          As String


Dim NITTSU_Name             As String
Dim NITTSU_Length           As Integer
Dim NITTSU_CODE()           As String



Dim KEN_CHARTER_CD          As String
Dim KEN_CHARTER_NM          As String

Dim KEN_AKABOU_CD           As String
Dim KEN_AKABOU_NM           As String


Dim KEN_LOGISTIC_CD         As String

Dim svOKURI_NO              As String * 20



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2014.05.14



Private Const LAST_UPDATE_DAY$ = "[F103063] 2013.05.16 10:30"






Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    

    Call Tab_Ctrl(Shift)
End Sub


Private Sub Command_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
        Case 7
            
            If Option1(poptSEL_OKURI_NO).Value Then
                If Trim(Text(ptxSEL_OKURI_NO).Text) = "" Then
                    MsgBox "問合せ№を指定して下さい。"
                    Exit Sub
                End If
            
            
            Else
                If IsNumeric(Trim(Text(ptxSEL_SYUKA_MM).Text)) Then
                    Text(ptxSEL_SYUKA_MM).Text = Format(CInt(Trim(Text(ptxSEL_SYUKA_MM).Text)), "00")
                End If
                If IsNumeric(Trim(Text(ptxSEL_SYUKA_DD).Text)) Then
                    Text(ptxSEL_SYUKA_DD).Text = Format(CInt(Trim(Text(ptxSEL_SYUKA_DD).Text)), "00")
                End If
                Text(ptxSEL_MUKE_CODE).Text = Right(Combo(pcmbSEL_MUKE_NAME), 8)
            
                '2007.01.16
                If Trim(Text(ptxSEL_INS_BIN).Text) <> "" Then
                    If Not IsNumeric(Text(ptxSEL_INS_BIN).Text) Then
                        Text(ptxSEL_INS_BIN).Text = ""
                    End If
                End If
            
            
            End If
            
            IN_CNT = 0
            SUMI_CNT = 0
            CANCEL_CNT = 0
            
            If Option1(poptSEL_OKURI_NO).Value Or Option1(poptSEL_ETC).Value Then
                If List_Disp_Proc Then
                    Unload Me
                End If
        
            Else
        
                If List_Disp_DEN_Proc Then
                    Unload Me
                End If

            End If
        
            Text(ptxDISP_CNT).Text = Format(SUMI_CNT, "#0") & "/" & Format(IN_CNT - CANCEL_CNT, "#0") & "(" & Format(IN_CNT, "#0") & ")"
        
        
        Case 8
        
            If Option1(poptSEL_OKURI_NO).Value Then
                If Trim(Text(ptxSEL_OKURI_NO).Text) = "" Then
                    MsgBox "問合せ№を指定して下さい。"
                    Exit Sub
                End If
            
            
            Else
                If IsNumeric(Trim(Text(ptxSEL_SYUKA_MM).Text)) Then
                    Text(ptxSEL_SYUKA_MM).Text = Format(CInt(Trim(Text(ptxSEL_SYUKA_MM).Text)), "00")
                End If
                If IsNumeric(Trim(Text(ptxSEL_SYUKA_DD).Text)) Then
                    Text(ptxSEL_SYUKA_DD).Text = Format(CInt(Trim(Text(ptxSEL_SYUKA_DD).Text)), "00")
                End If
                Text(ptxSEL_MUKE_CODE).Text = Right(Combo(pcmbSEL_MUKE_NAME), 8)
            
                '2007.01.16
                If Trim(Text(ptxSEL_INS_BIN).Text) <> "" Then
                    If Not IsNumeric(Text(ptxSEL_INS_BIN).Text) Then
                        Text(ptxSEL_INS_BIN).Text = ""
                    End If
                End If
            
            End If
        
        
        
            If SDC_FLD_GET("SYS", App.EXEName, GW_PATH) Then
                Unload Me
            End If
            
            If SDC_FLD_Return Then
            Else
                If Option1(poptSEL_OKURI_NO).Value Or Option1(poptSEL_ETC).Value Then
                    If Data_Out_Proc Then
                        Unload Me
                    End If
            
                Else
            
                    If Data_Out_DEN_Proc Then
                        Unload Me
                    End If
    
                End If
                
                
            End If
        
        Case 9
            
            Call Input_Lock     '2013.03.18
            
                        
            
            
            Call Form_HCopy_Win7(Picture1, vbPRPSA4, vbPRORLandscape)

        
            Call Input_UnLock   '2013.03.18
        
        
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub
Private Sub Form_DblClick()
'''    PrintForm
'''    Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
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
Dim i               As Integer
Dim c               As String * 128
Dim sts             As Integer


Dim UNSOU_KAISHA  As Variant

Dim wkUNSOU_KAISHA  As Variant      '2014.05.15


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
                    
                    
    '2010.01.06
                                '出荷ログファイル名取り込み
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "出荷ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    '2010.01.06
                    
                    
                    
                    '最大表示件数の獲得
    If GetIni(App.EXEName, "LISTMAX", "SYS", c) Then
        Beep
        MsgBox "最大表示件数の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    Max_Row = CLng(RTrim(c))
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If


    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030631.Caption = "出荷問い合わせ(問合せ№)（" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
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
                                
                                
                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＯＰＥＮ
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＯＰＥＮ
    If DEL_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＯＰＥＮ
    If wDEL_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If





    Option1(poptSEL_OKURI_NO).Value = True
    Option1(poptSEL_ETC).Value = False

'出荷日付
    Text(ptxSEL_SYUKA_YY).Text = Left(Format(Now, "YYYYMMDD"), 4)
    Text(ptxSEL_SYUKA_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)
    Text(ptxSEL_SYUKA_DD).Text = Right(Format(Now, "YYYYMMDD"), 2)

'向け先設定
    If MTS_Set_Proc() Then
        Unload Me
    End If


'運送会社
    Combo(pcmbSEL_UNSOU_KAISHA).Clear
    Combo(pcmbSEL_UNSOU_KAISHA).AddItem UNSOU_KAISHA_ALL
                    
                    
                                '運送会社名称獲得
    If GetIni(App.EXEName, "UNSOU_KAISHA", "SYS", c) Then
    Else
        UNSOU_KAISHA = Split(Trim(c), ",", -1)
        For i = 0 To UBound(UNSOU_KAISHA)
            Combo(pcmbSEL_UNSOU_KAISHA).AddItem UNSOU_KAISHA(i)
        Next i
    End If
    Combo(pcmbSEL_UNSOU_KAISHA).ListIndex = 0
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 個別運送会社情報　2014.05.15
'福山
    If GetIni("FUKUYAMA", "length", "UNSOU", c) Then
        FUKUYAMA_Name = ""
        FUKUYAMA_Length = 0
        ReDim Preserve FUKUYAMA_CODE(0 To 0)
        FUKUYAMA_CODE(0) = "*"
    Else
        FUKUYAMA_Length = Val(Trim(c))
        
        If GetIni("FUKUYAMA", "Name", "UNSOU", c) Then
            FUKUYAMA_Name = ""
        Else
            FUKUYAMA_Name = Trim(c)
        End If
        
        If GetIni("FUKUYAMA", "Code", "UNSOU", c) Then
            ReDim Preserve FUKUYAMA_CODE(0 To 0)
            FUKUYAMA_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve FUKUYAMA_CODE(0 To i)
                FUKUYAMA_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        
        End If
    End If



'西武
    If GetIni("SEIBU", "length", "UNSOU", c) Then
        SEIBU_Name = ""
        SEIBU_Length = 0
        ReDim Preserve SEIBU_CODE(0 To 0)
        SEIBU_CODE(0) = "*"
    Else
        SEIBU_Length = Val(Trim(c))
        
        If GetIni("SEIBU", "Name", "UNSOU", c) Then
            SEIBU_Name = ""
        Else
            SEIBU_Name = Trim(c)
        End If
        
        If GetIni("SEIBU", "Code", "UNSOU", c) Then
            ReDim Preserve SEIBU_CODE(0 To 0)
            SEIBU_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve SEIBU_CODE(0 To i)
                SEIBU_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        
        End If
    End If


'久留米
    If GetIni("KURUME", "length", "UNSOU", c) Then
        KURUME_Name = ""
        KURUME_Length = 0
        ReDim Preserve KURUME_CODE(0 To 0)
        KURUME_CODE(0) = "*"
    Else
        KURUME_Length = Val(Trim(c))
        
        If GetIni("KURUME", "Name", "UNSOU", c) Then
            KURUME_Name = ""
        Else
            KURUME_Name = Trim(c)
        End If
        
        If GetIni("KURUME", "Code", "UNSOU", c) Then
            ReDim Preserve KURUME_CODE(0 To 0)
            KURUME_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve KURUME_CODE(0 To i)
                KURUME_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        
        End If
    End If


'佐川
    If GetIni("SAGAWA", "length", "UNSOU", c) Then
        SAGAWA_Name = ""
        SAGAWA_Length = 0
        ReDim Preserve SAGAWA_CODE(0 To 0)
        SAGAWA_CODE(0) = "*"
    Else
        SAGAWA_Length = Val(Trim(c))
        
        If GetIni("SAGAWA", "Name", "UNSOU", c) Then
            SAGAWA_Name = ""
        Else
            SAGAWA_Name = Trim(c)
        End If
        
        If GetIni("SAGAWA", "Code", "UNSOU", c) Then
            ReDim Preserve SEIBU_CODE(0 To 0)
            SAGAWA_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve SAGAWA_CODE(0 To i)
                SAGAWA_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        
        End If
    End If

'ヤマト
    If GetIni("YAMATO", "length", "UNSOU", c) Then
        YAMATO_Name = ""
        YAMATO_Length = 0
        ReDim Preserve YAMATO_CODE(0 To 0)
        YAMATO_CODE(0) = "*"
    Else
        YAMATO_Length = Val(Trim(c))
        
        If GetIni("YAMATO", "Name", "UNSOU", c) Then
            YAMATO_Name = ""
        Else
            YAMATO_Name = Trim(c)
        End If
        
        If GetIni("YAMATO", "Code", "UNSOU", c) Then
            ReDim Preserve YAMATO_CODE(0 To 0)
            YAMATO_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve YAMATO_CODE(0 To i)
                YAMATO_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        
        End If
    End If
'ﾊﾟﾅﾎｰﾑ
    If GetIni("PANAHOME", "length", "UNSOU", c) Then
        PANA_Name = ""
        PANA_Length = 0
        ReDim Preserve PANA_CODE(0 To 0)
        PANA_CODE(0) = "*"
    Else
        PANA_Length = Val(Trim(c))
        
        If GetIni("PANAHOME", "Name", "UNSOU", c) Then
            PANA_Name = ""
        Else
            PANA_Name = Trim(c)
        End If
        
        If GetIni("PANAHOME", "Code", "UNSOU", c) Then
            ReDim Preserve PANA_CODE(0 To 0)
            PANA_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve PANA_CODE(0 To i)
                PANA_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        
        End If
    End If
'積水
    If GetIni("SEKISUI", "length", "UNSOU", c) Then
        SEKISUI_Name = ""
        SEKISUI_Length = 0
        ReDim Preserve SEKISUI_CODE(0 To 0)
        SEKISUI_CODE(0) = "*"
    Else
        SEKISUI_Length = Val(Trim(c))
        
        If GetIni("SEKISUI", "Name", "UNSOU", c) Then
            SEKISUI_Name = ""
        Else
            SEKISUI_Name = Trim(c)
        End If
        
        If GetIni("SEKISUI", "Code", "UNSOU", c) Then
            ReDim Preserve SEKISUI_CODE(0 To 0)
            SEKISUI_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve SEKISUI_CODE(0 To i)
                SEKISUI_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        
        End If
    End If
'日通
    If GetIni("NITTSU", "length", "UNSOU", c) Then
        NITTSU_Name = ""
        NITTSU_Length = 0
        ReDim Preserve NITTSU_CODE(0 To 0)
        NITTSU_CODE(0) = "*"
    Else
        NITTSU_Length = Val(Trim(c))
        
        If GetIni("NITTSU", "Name", "UNSOU", c) Then
            NITTSU_Name = ""
        Else
            NITTSU_Name = Trim(c)
        End If
        
        If GetIni("NITTSU", "Code", "UNSOU", c) Then
            ReDim Preserve NITTSU_CODE(0 To 0)
            SEKISUI_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve NITTSU_CODE(0 To i)
                NITTSU_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        
        End If
    End If





'---------------------------------------------- 'チャーター便ｺｰﾄﾞ
    If GetIni("CHARTER", "Code", "UNSOU", c) Then
        KEN_CHARTER_CD = "*"
    Else
        
        wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
    
        KEN_CHARTER_CD = wkUNSOU_KAISHA(0)
        
        If GetIni("CHARTER", "Name", "UNSOU", c) Then
            KEN_CHARTER_NM = ""
        Else
            KEN_CHARTER_NM = Trim(c)
        End If
    End If
'---------------------------------------------- '赤帽便ｺｰﾄﾞ
    If GetIni("AKABOU", "Code", "UNSOU", c) Then
        KEN_AKABOU_CD = "*"
    Else
        wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
    
        KEN_AKABOU_CD = wkUNSOU_KAISHA(0)
    
    
        If GetIni("AKABOU", "Name", "UNSOU", c) Then
            KEN_AKABOU_NM = ""
        Else
            KEN_AKABOU_NM = Trim(c)
        End If
    
    
    End If

'---------------------------------------------- 'ロジステックｺｰﾄﾞ
    If GetIni(App.EXEName, "KEN_LOGISTIC", "SYS", c) Then
        KEN_LOGISTIC_CD = "*"
    Else
        KEN_LOGISTIC_CD = Trim(c)
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 個別運送会社情報　2014.05.15

    
'表示条件
    Check1(pchkMI).Value = vbChecked
    Check1(pchkSUMI).Value = vbChecked
    
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0                 'ﾃﾞﾌｫﾙﾄ昇順
    Next i
    

    Text(ptxSEL_OKURI_NO).SetFocus
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
                                            '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
        End If
    End If
                                            '削除済み出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K0_DEL_SYU_H, Len(K0_DEL_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
        End If
    End If
                                            '削除済み出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wDEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K0_wDEL_SYU_H, Len(K0_wDEL_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
        End If
    End If
    
    
    
    
    sts = BTRV(BtOpReset, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set F1030631 = Nothing


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
    F1030631.Caption = "出荷問い合わせ(問合せ№)（" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
    SubMenu(Index).Checked = True
    If Last_JGYOBU <> JGYOBU_T(Index).CODE Then
        Last_JGYOBU = JGYOBU_T(Index).CODE
        LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
        LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
    End If

End Sub

Private Function MTS_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   出荷先をコンボにセットする
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer
Dim Edit        As String

    MTS_Set_Proc = True
    
    Call Input_Lock
    
    
    Combo(pcmbSEL_MUKE_NAME).Clear
    
    Combo(pcmbSEL_MUKE_NAME).AddItem "全て　　　" & "   " & Space(8)
        
    
    
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
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode)
        
        
        Combo(pcmbSEL_MUKE_NAME).AddItem Edit
    
        com = BtOpGetNext
    
    Loop

    Combo(pcmbSEL_MUKE_NAME).ListIndex = 0

    Call Input_UnLock

    MTS_Set_Proc = False
End Function


Private Function Data_Out_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ検索
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

    
    
Dim Skip_Flg    As Boolean
    
Dim FileNo      As Long
Dim FSW         As Boolean
    
    Data_Out_Proc = True
                                    
    Call Input_Lock
    
    'テキストファイルＯＰＥＮ
    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    
    Open Trim(GW_PATH) For Output As FileNo
                                    
    FSW = True
'---------------------------------- '当日分出荷予定読み込み開始
    If Option1(poptSEL_OKURI_NO).Value Then
    
        Call UniCode_Conv(K2_Y_SYU_H.OKURI_NO, Text(ptxSEL_OKURI_NO).Text)
    
    Else
    
        Call UniCode_Conv(K3_Y_SYU_H.SYUKA_YMD, Text(ptxSEL_SYUKA_YY).Text & _
                                                Text(ptxSEL_SYUKA_MM).Text & _
                                                Text(ptxSEL_SYUKA_DD).Text)
    End If
    
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        Skip_Flg = False
        If Option1(poptSEL_OKURI_NO).Value Then
                
        
            sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K2_Y_SYU_H, Len(K2_Y_SYU_H), 2)
    
            Select Case sts
                Case BtNoErr
            
                    If Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode)) <> Trim(Text(ptxSEL_OKURI_NO).Text) Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    Data_Out_Proc = SYS_ERR
                    Exit Function
            End Select
    
    
    
        Else
    
            sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K3_Y_SYU_H, Len(K3_Y_SYU_H), 3)
    
            Select Case sts
                Case BtNoErr
                        
                    If Trim(Text(ptxSEL_SYUKA_YY).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_YY).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 1, 4) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_SYUKA_MM).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_MM).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 5, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_SYUKA_DD).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_DD).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 7, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_MUKE_CODE).Text) <> "" Then
                        If Trim(Text(ptxSEL_MUKE_CODE).Text) <> Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_HIN_NO).Text) <> "" Then
                        If Trim(Text(ptxSEL_HIN_NO).Text) <> Trim(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                        
                    If Trim(Combo(pcmbSEL_UNSOU_KAISHA).Text) <> UNSOU_KAISHA_ALL Then
                        If Trim(Combo(pcmbSEL_UNSOU_KAISHA).Text) <> Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                
                
                    If Trim(Text(ptxSEL_INS_BIN).Text) <> "" Then  '2007.01.16
                        If Format(CInt(Text(ptxSEL_INS_BIN).Text), "00") <> StrConv(Y_SYU_HREC.INS_BIN, vbUnicode) Then
                            Skip_Flg = True
                        End If
                    End If
                
                
                
                    If Check1(pchkMI).Value <> vbChecked Then
                        If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) = "" Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Check1(pchkSUMI).Value <> vbChecked Then
                        If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) <> "" Then
                            Skip_Flg = True
                        End If
                    End If
                
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    Data_Out_Proc = SYS_ERR
                    Exit Function
            End Select
    
    
        End If
                
                
        If Not Skip_Flg Then
            
                    
            If Data_Output_Proc(FileNo, FSW) Then
                Exit Function
            End If
            FSW = False
        End If
        
        com = BtOpGetNext
        
    Loop
    
    
'---------------------------------- '過日分出荷予定読み込み開始
    If Option1(poptSEL_OKURI_NO).Value Then
    
        Call UniCode_Conv(K2_DEL_SYU_H.OKURI_NO, Text(ptxSEL_OKURI_NO).Text)
    
    Else
    
        Call UniCode_Conv(K3_DEL_SYU_H.SYUKA_YMD, Text(ptxSEL_SYUKA_YY).Text & _
                                                Text(ptxSEL_SYUKA_MM).Text & _
                                                Text(ptxSEL_SYUKA_DD).Text)
    
    End If
    
    
    
    
    
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        Skip_Flg = False
        If Option1(poptSEL_OKURI_NO).Value Then
                
        
            sts = BTRV(com, DEL_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K2_DEL_SYU_H, Len(K2_DEL_SYU_H), 2)
    
            Select Case sts
                Case BtNoErr
            
                    If Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode) <> Trim(Text(ptxSEL_OKURI_NO).Text)) Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    Data_Out_Proc = SYS_ERR
                    Exit Function
            End Select
    
    
    
        Else
    
            sts = BTRV(com, DEL_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K3_DEL_SYU_H, Len(K3_DEL_SYU_H), 3)
    
            Select Case sts
                Case BtNoErr
                        
                    If Trim(Text(ptxSEL_SYUKA_YY).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_YY).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 1, 4) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_SYUKA_MM).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_MM).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 5, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_SYUKA_DD).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_DD).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 7, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_MUKE_CODE).Text) <> "" Then
                        If Trim(Text(ptxSEL_MUKE_CODE).Text) <> Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_HIN_NO).Text) <> "" Then
                        If Trim(Text(ptxSEL_HIN_NO).Text) <> Trim(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                        
                    If Trim(Combo(pcmbSEL_UNSOU_KAISHA).Text) <> UNSOU_KAISHA_ALL Then
                        If Trim(Combo(pcmbSEL_UNSOU_KAISHA).Text) <> Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_INS_BIN).Text) <> "" Then  '2007.01.16
                        If Format(CInt(Text(ptxSEL_INS_BIN).Text), "00") <> StrConv(Y_SYU_HREC.INS_BIN, vbUnicode) Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Check1(pchkMI).Value <> vbChecked Then
                        If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) = "" Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Check1(pchkSUMI).Value <> vbChecked Then
                        If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) <> "" Then
                            Skip_Flg = True
                        End If
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    Data_Out_Proc = SYS_ERR
                    Exit Function
            End Select
    
    
        End If
                
                
        If Not Skip_Flg Then
            
            If Data_Output_Proc(FileNo, FSW) Then
                Exit Function
            End If
            FSW = False
        End If
        
        com = BtOpGetNext
        
    Loop
    
    
    Close #FileNo
    
    
    Call Input_UnLock
    
    
    MsgBox "データ出力終了！！"
    
    Data_Out_Proc = False
    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox GW_PATH & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        Data_Out_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
    End If

    
End Function
Private Function Data_Out_DEN_Proc() As Integer
'----------------------------------------------------------------------------
'                   伝票№での検索
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer
Dim comw        As Integer

Dim Skip_Flg    As Boolean
    
Dim svID_No     As String * 7
    
Dim FileNo      As Long
Dim FSW         As Boolean
    
Dim j           As Integer
Dim k           As Integer
Dim FOUND_FLG   As Boolean

Dim svID_TBL()  As String * 7
    
    
    
    Data_Out_DEN_Proc = True
                                    
    Call Input_Lock
                                    
                                    
    'テキストファイルＯＰＥＮ
    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    
    Open Trim(GW_PATH) For Output As FileNo
                                    
    FSW = True
'---------------------------------- '当日分出荷予定読み込み開始
    
    Call UniCode_Conv(K0_Y_SYU_H.DEN_NO, Text(ptxSEL_DEN_NO).Text)
    Call UniCode_Conv(K0_Y_SYU_H.SEQ_NO, "")
    sts = BTRV(BtOpGetGreaterEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)

    Select Case sts
        Case BtNoErr
    
            If Left(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode), 7) <> Trim(Text(ptxSEL_DEN_NO).Text) Then
                sts = BtErrEOF
            End If
        
        Case BtErrEOF
        
        Case Else
            Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
            Data_Out_DEN_Proc = SYS_ERR
            Exit Function
    
    End Select
    
    
    If sts = BtNoErr Then
        svID_No = StrConv(Y_SYU_HREC.ID_NO, vbUnicode)
        
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, svID_No)
        
    
        com = BtOpGetGreaterEqual
    
        Do
        
            DoEvents
        
        
            sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
    
            Select Case sts
                Case BtNoErr
            
                    If Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) <> svID_No Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    Data_Out_DEN_Proc = SYS_ERR
                    Exit Function
            End Select
                
                
            Skip_Flg = False
            
            If Trim(Text(ptxSEL_SYUKA_YY).Text) = "" And Trim(Text(ptxSEL_SYUKA_MM).Text) = "" And Trim(Text(ptxSEL_SYUKA_DD).Text) = "" Then
            Else
                If StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode) <> _
                    (Text(ptxSEL_SYUKA_YY).Text & Text(ptxSEL_SYUKA_MM).Text & Text(ptxSEL_SYUKA_DD).Text) Then
                                                
                    Skip_Flg = True
                
                End If
            End If
                    
                    
            If Not Skip_Flg Then
                
                
                
                
                        
                If Data_Output_Proc(FileNo, FSW) Then
                    Exit Function
                End If
                FSW = False
            End If
        
            com = BtOpGetNext
        
        Loop
    
    End If
    
'---------------------------------- '過日分出荷予定読み込み開始
    
    svID_No = ""
    j = 0
    comw = BtOpGetGreater
    
    Call UniCode_Conv(K0_wDEL_SYU_H.DEN_NO, Text(ptxSEL_DEN_NO).Text)
    Call UniCode_Conv(K0_wDEL_SYU_H.SEQ_NO, "")
    
    
    Do
        DoEvents
        sts = BTRV(comw, wDEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K0_wDEL_SYU_H, Len(K0_wDEL_SYU_H), 0)
    
        Select Case sts
            Case BtNoErr
        
                If Left(StrConv(DEL_SYU_HREC.DEN_NO, vbUnicode), 7) <> Trim(Text(ptxSEL_DEN_NO).Text) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "削除済み出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                Data_Out_DEN_Proc = SYS_ERR
                Exit Function
        
        End Select
        
        
        FOUND_FLG = False
        If Trim(svID_No) = "" Then
        Else
            For k = 0 To UBound(svID_TBL)
            
            
                If svID_TBL(j) = Left(StrConv(DEL_SYU_HREC.ID_NO, vbUnicode), 7) Then
                    FOUND_FLG = True
                    Exit For
                End If
            
            Next k
        End If
        
        
        If FOUND_FLG Then
        Else
            If Trim(svID_No) = "" Then
                ReDim svID_TBL(j)
            Else
                j = j + 1
                ReDim Preserve svID_TBL(j)
            End If
            svID_TBL(j) = StrConv(DEL_SYU_HREC.ID_NO, vbUnicode)
            
            
            svID_No = StrConv(DEL_SYU_HREC.ID_NO, vbUnicode)
            
            Call UniCode_Conv(K4_DEL_SYU_H.ID_NO, svID_No)
            
            com = BtOpGetGreaterEqual
        
            Do
            
                DoEvents
            
            
                sts = BTRV(com, DEL_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_DEL_SYU_H, Len(K4_DEL_SYU_H), 4)
        
                Select Case sts
                    Case BtNoErr
                
                        If Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) <> svID_No Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                        Data_Out_DEN_Proc = SYS_ERR
                        Exit Function
                End Select
                    
                Skip_Flg = False
                
                If Trim(Text(ptxSEL_SYUKA_YY).Text) = "" And Trim(Text(ptxSEL_SYUKA_MM).Text) = "" And Trim(Text(ptxSEL_SYUKA_DD).Text) = "" Then
                Else
                    If StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode) <> _
                        (Text(ptxSEL_SYUKA_YY).Text & Text(ptxSEL_SYUKA_MM).Text & Text(ptxSEL_SYUKA_DD).Text) Then
                                                    
                        Skip_Flg = True
                    
                    End If
                End If
                    
                    
                If Not Skip_Flg Then
                    
                    If Data_Output_Proc(FileNo, FSW) Then
                        Exit Function
                    End If
                    FSW = False
                End If
            
                com = BtOpGetNext
            
            Loop
        
        End If
        
        comw = BtOpGetNext
    
    Loop
    Close #FileNo
    
    
    
    Call Input_UnLock
    
    MsgBox "データ出力終了！！"
    
    
    Data_Out_DEN_Proc = False

    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox GW_PATH & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        Data_Out_DEN_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
    End If

    
End Function
    

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1030631.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030631)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030631)


    F1030631.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'                   FILE---.GLID
'----------------------------------------------------------------------------

Dim sts As Integer

Dim Seq_No_From As String
Dim Seq_No_To   As String

    
    Grid_Set_Proc = True


    SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
    
    '問合せ№
    SYUKA(Row, ColOKURI_NO) = Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode))
    
    'ｷｬﾝｾﾙ
    If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) = "1" Then
        SYUKA(Row, ColCANCEL_F) = "×"
    Else
        SYUKA(Row, ColCANCEL_F) = ""
    End If
    
    '出荷№
    SYUKA(Row, ColSYUKA_NO) = Trim(StrConv(Y_SYU_HREC.SYUKA_NO, vbUnicode))
    
    '便 2007.01.16
    If IsNumeric(StrConv(Y_SYU_HREC.INS_BIN, vbUnicode)) Then
        SYUKA(Row, ColINS_BIN) = Format(CInt(StrConv(Y_SYU_HREC.INS_BIN, vbUnicode)), "#0")
    Else
        SYUKA(Row, ColINS_BIN) = ""
    End If
    '出荷日
    SYUKA(Row, ColSYUKA_YMD) = Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 7, 2)
    
    
    '送り先集約CD   2007.07.09
    SYUKA(Row, ColCOL_OKURISAKI_CD) = Trim(StrConv(Y_SYU_HREC.COL_OKURISAKI_CD, vbUnicode))
    '送り先CD   2007.07.09
    SYUKA(Row, ColOKURISAKI_CD) = Trim(StrConv(Y_SYU_HREC.OKURISAKI_CD, vbUnicode))
    
    '送り先
    SYUKA(Row, ColOKURISAKI) = Trim(StrConv(Y_SYU_HREC.OKURISAKI, vbUnicode))
    '売り伝
    If StrConv(Y_SYU_HREC.URIDEN, vbUnicode) = "1" Then
        SYUKA(Row, ColURIDEN) = "有"
    Else
        SYUKA(Row, ColURIDEN) = ""
    End If
    '伝票番号
    
    
    If StrConv(Y_SYU_HREC.KYOSEI_END, vbUnicode) = "1" Then
        SYUKA(Row, ColDEN_NO) = "*" & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode))
    Else
        SYUKA(Row, ColDEN_NO) = Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode))
    End If
    
    '品番＆品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYU_HREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
            Exit Function
    End Select
    SYUKA(Row, ColHIN_NO) = Left(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode), 12) & " " & Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    '数量
    SYUKA(Row, ColSURYO) = Format(CLng(StrConv(Y_SYU_HREC.SURYO, vbUnicode)), "#0")
    '実績
    If IsNumeric(StrConv(Y_SYU_HREC.J_SURYO, vbUnicode)) Then
        SYUKA(Row, ColJ_SURYO) = Format(CLng(StrConv(Y_SYU_HREC.J_SURYO, vbUnicode)), "#0")
    Else
        SYUKA(Row, ColJ_SURYO) = 0
    End If
    '注文№
    SYUKA(Row, ColORDER_NO) = StrConv(Y_SYU_HREC.ODER_NO, vbUnicode)
    '得意先ｺｰﾄﾞ＆名称
    SYUKA(Row, ColMUKE_CODE) = StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode) & " " & Trim(StrConv(Y_SYU_HREC.MUKE_NAME, vbUnicode))
    '備考
    SYUKA(Row, ColBIKOU) = StrConv(Y_SYU_HREC.BIKOU, vbUnicode)
    '運送会社
    SYUKA(Row, ColUNSOU_KAISHA) = StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)
    '口数
    
    If IsNumeric(StrConv(Y_SYU_HREC.KUTI_SU, vbUnicode)) Then
        SYUKA(Row, ColKUTI_SU) = Format(CInt(StrConv(Y_SYU_HREC.KUTI_SU, vbUnicode)), "#")
    Else
        SYUKA(Row, ColKUTI_SU) = ""
    End If
    
    '検品日時
    If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) = "" Then
        SYUKA(Row, ColKENPIN_NOW) = ""
    Else
        SYUKA(Row, ColKENPIN_NOW) = Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 7, 2) & " " & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 9, 2) & ":" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 11, 2) & ":" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 13, 2)


    End If
    '担当者
    If Trim(StrConv(Y_SYU_HREC.KENPIN_TANTO_CODE, vbUnicode)) = "" Then
        SYUKA(Row, ColKENPIN_TANTO_CODE) = ""
    Else
        Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(Y_SYU_HREC.KENPIN_TANTO_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound
                Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "担当者ﾏｽﾀ")
                Exit Function
        End Select
    
        SYUKA(Row, ColKENPIN_TANTO_CODE) = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
    
    End If
    
    'ID_NO
    SYUKA(Row, ColID_NO) = Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
    
    
    SYUKA(Row, ColCANCEL_F_A) = StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode)
    '備考（入力値）
    SYUKA(Row, ColINPUT_BIKOU) = Trim(StrConv(Y_SYU_HREC.INPUT_BIKOU, vbUnicode))
    
    
    '2010.02.13
    
    
    
    SYUKA(Row, ColOKURI_NO_SEQ) = StrConv(Y_SYU_HREC.OKURI_NO_SEQ, vbUnicode) & "～" & StrConv(Y_SYU_HREC.OKURI_NO_SEQ_TO, vbUnicode)
    If IsNumeric(StrConv(Y_SYU_HREC.SAI_SU, vbUnicode)) Then
        SYUKA(Row, ColSAI_SU) = Format(StrConv(Y_SYU_HREC.SAI_SU, vbUnicode), "#0.00")
    Else
        SYUKA(Row, ColSAI_SU) = ""
    End If
    
    If StrConv(ITEMREC.KONPOU_F, vbUnicode) = "1" Then
        SYUKA(Row, ColKONPOU_F) = StrConv(ITEMREC.KONPOU_F, vbUnicode)
    
    Else
        SYUKA(Row, ColKONPOU_F) = ""
    End If
    
    If IsNumeric(StrConv(Y_SYU_HREC.KUTI_SU_TAN, vbUnicode)) Then
        SYUKA(Row, ColKUTI_SU_TAN) = Format(StrConv(Y_SYU_HREC.KUTI_SU_TAN, vbUnicode), "#0")
    Else
        SYUKA(Row, ColKUTI_SU_TAN) = ""
    End If
    
    If IsNumeric(StrConv(Y_SYU_HREC.SAI_SU_TAN, vbUnicode)) Then
        SYUKA(Row, ColSAI_SU_TAN) = Format(StrConv(Y_SYU_HREC.SAI_SU_TAN, vbUnicode), "#0.00")
    Else
        SYUKA(Row, ColSAI_SU_TAN) = ""
    End If
    
    
    '2010.02.13
    
    '2010.02.15
    If IsNumeric(StrConv(Y_SYU_HREC.JURYO, vbUnicode)) Then
        SYUKA(Row, ColJURYO) = Format(StrConv(Y_SYU_HREC.JURYO, vbUnicode), "#0.0")
    Else
        SYUKA(Row, ColJURYO) = ""
    End If
    
    
    SYUKA(Row, ColJYUSHO) = Trim(StrConv(Y_SYU_HREC.JYUSHO, vbUnicode))
    
    '2010.02.15
    
    SYUKA(Row, ColYUBIN_NO) = Trim(StrConv(Y_SYU_HREC.YUBIN_NO, vbUnicode))
    SYUKA(Row, ColTEL_NO) = Trim(StrConv(Y_SYU_HREC.TEL_NO, vbUnicode))
    
    
    
    
    '2011.04.23
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, StrConv(Y_SYU_HREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
                
                
                
    sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定")
            Exit Function
    End Select
    If Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode)) = "" Then
    Else
        SYUKA(Row, ColUSE) = Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode)) & ":" & Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode))
    End If
    '2011.04.23
    
    '2011.05.18
        SYUKA(Row, ColSEK_KEN_NO) = StrConv(Y_SYU_HREC.SEK_KEN_NO, vbUnicode)
        SYUKA(Row, ColSEK_HIN_NO) = StrConv(Y_SYU_HREC.SEK_HIN_NO, vbUnicode)
    '2011.05.18
    
    
    SYUKA(Row, ColY_SYU_INS) = StrConv(Y_SYU_HREC.Ins_DateTime, vbUnicode)
    SYUKA(Row, ColY_SYU_H_INS) = StrConv(Y_SYU_HREC.UPD_DATETIME, vbUnicode)
    
    
    
    
    
    Grid_Set_Proc = False
End Function


Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)




    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.05.15
Dim SvBookmark          As Long
    
Dim OKURI_NO_F          As Boolean
Dim FUKUYAMA_CHK_F      As Boolean
    
    
Dim OKURI_NO            As String * 20
   
Dim UNSOU_KAISHA        As String * 40
    
    If SYUKA.Count(1) < 1 Then
        Exit Sub
    End If
    
    
    If TDBGrid1.Bookmark = -1 Then
    Else
    
        Set TDBGrid1.Array = SYUKA
        TDBGrid1.Update
    
    
        SvBookmark = TDBGrid1.Bookmark
        
        
        
        Select Case ColIndex
            Case ColOKURI_NO    '問合せ
                
                
                OKURI_NO = SYUKA(SvBookmark, ColIndex)
                
                
                If Trim(OKURI_NO) = "" Then
                    SYUKA(SvBookmark, ColUNSOU_KAISHA) = ""
                    Exit Sub
                End If
                
                
                If OKURI_NO_CHECK_PROC(OKURI_NO, OKURI_NO_F, FUKUYAMA_CHK_F) Then
                End If
        
                If OKURI_NO_F Then
                    Call OKURI_NO_SET_PROC(OKURI_NO, UNSOU_KAISHA)
                    SYUKA(SvBookmark, ColUNSOU_KAISHA) = UNSOU_KAISHA
                Else
                    SYUKA(SvBookmark, ColUNSOU_KAISHA) = ""
                End If
        
        End Select
    
RESET_PROC:
    
    
        Set TDBGrid1.Array = SYUKA
    
    
    '    TDBGrid1.Refresh
    '    TDBGrid1.ScrollBars = dbgAutomatic
    '    TDBGrid1.Update
    
    
    
        
        TDBGrid1.Bookmark = SvBookmark
        
        TDBGrid1.SetFocus
    
    End If
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.05.15




End Sub

Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.05.15
    If SYUKA.Count(1) < 1 Then
        Exit Sub
    End If



    Select Case ColIndex
        Case ColOKURI_NO
            If TDBGrid1.Bookmark = -1 Then
            Else
                svOKURI_NO = SYUKA(TDBGrid1.Bookmark, ColIndex)
            End If
            
    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.05.15
End Sub

Private Sub TDBGrid1_DblClick()
Dim SvBookmark  As Long
    
    
    
    
    
    
    If SYUKA.Count(1) < 1 Then
        Exit Sub
    End If
    
    If HEAD_CLICK Then
        HEAD_CLICK = False
        Exit Sub
    End If
    
    If TDBGrid1.Bookmark = -1 Then
    Else

        SvBookmark = TDBGrid1.Bookmark
        If Update_Proc(0) Then
            Unload Me
        End If
    
        TDBGrid1.Bookmark = SvBookmark
        TDBGrid1.SetFocus
    
    
    End If

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)

    
Dim lngPFstRow  As Long
Dim vntBmk      As Variant
Dim intLeftCol  As Integer
Dim intCol      As Integer
Dim lngCFstRow  As Long
    
    
    
        '2010.08.31
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
        
        
        
        
        
        
        
        
        
'        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst

    End If
        
        
        '2010.08.31
    
    



End Sub

Private Sub TDBGrid1_LostFocus()
    
    If Not DATA_FLG Then
        Exit Sub
    End If
    
    
    
    
'    If DBGrid_Chk_Proc() Then  2014.05.15
'        Exit Sub               2014.05.15
'    End If                     2014.05.15
    
    
    
    
    
    If Update_Proc(1) Then
        Unload Me
    End If



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
        
        Case ptxSEL_SYUKA_YY
            If Len(Trim(Text(ptxSEL_SYUKA_YY).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSEL_SYUKA_YY).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
            End If
        Case ptxSEL_SYUKA_MM
            If Len(Trim(Text(ptxSEL_SYUKA_MM).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSEL_SYUKA_MM).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
                Text(ptxSEL_SYUKA_MM).Text = Format(CInt(Text(ptxSEL_SYUKA_MM).Text), "00")
            End If
        Case ptxSEL_SYUKA_DD
            If Len(Trim(Text(ptxSEL_SYUKA_DD).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSEL_SYUKA_DD).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
                Text(ptxSEL_SYUKA_DD).Text = Format(CInt(Text(ptxSEL_SYUKA_DD).Text), "00")
            End If
        
        
        
        Case ptxSEL_MUKE_CODE

            For i = 0 To Combo(pcmbSEL_MUKE_NAME).ListCount - 1 '向け先
    
                If Right(Combo(pcmbSEL_MUKE_NAME).List(i), 8) = Trim(Text(ptxSEL_MUKE_CODE)) Then
                    Combo(pcmbSEL_MUKE_NAME).ListIndex = i
                    Exit For
                End If
            
    
            Next
        
            If i > Combo(pcmbSEL_MUKE_NAME).ListCount - 1 Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Combo(pcmbSEL_MUKE_NAME).ListIndex = -1
                Exit Sub
            End If
            
            Combo(pcmbSEL_MUKE_NAME).SetFocus
    End Select
    
    Call Tab_Ctrl(Shift)

End Sub



Private Function Data_Output_Proc(FileNo As Long, FSW As Boolean) As Integer
'----------------------------------------------------------------------------
'                   ＣＳＶ出力
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer


    
    Data_Output_Proc = True


    If FSW Then
                                'ヘッダ出力 口数の出力位置移動　2007.07.10
        Write #FileNo, "問合せ№", _
                            "ｷｬﾝｾﾙ", _
                            "口", _
                            "№", _
                            "出荷日", _
                            "便", _
                            "送り先集約CD", _
                            "送り先CD", _
                            "送り先名", _
                            "売伝", _
                            "伝票番号", _
                            "品番", "品名", _
                            "数量", _
                            "出庫済み数", _
                            "注文№", _
                            "得意先", _
                            "備考", _
                            "運送会社", _
                            "検品日時", _
                            "検品担当者", _
                            "ID_NO", _
                            "備考", _
                            "枝番", _
                            "才数", _
                            "梱包区分"
    
    End If


    Write #FileNo, Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode)),
    
    
    If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) = "1" Then
        Write #FileNo, "×",
    Else
        Write #FileNo, "",
    End If
    
    
    If IsNumeric(StrConv(Y_SYU_HREC.KUTI_SU, vbUnicode)) Then
        Write #FileNo, Format(CInt(StrConv(Y_SYU_HREC.KUTI_SU, vbUnicode)), "#"),
    Else
        Write #FileNo, ,
    End If
   
    
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.SYUKA_NO, vbUnicode)),
    Write #FileNo, Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 7, 2),
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.INS_BIN, vbUnicode)),
    
    '2007.07.09 送り先集約CD
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.COL_OKURISAKI_CD, vbUnicode)),
    '2007.07.09 送り先CD
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.OKURISAKI_CD, vbUnicode)),
    
    
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.OKURISAKI, vbUnicode)),
    If StrConv(Y_SYU_HREC.URIDEN, vbUnicode) = "1" Then
        Write #FileNo, "有",
    Else
        Write #FileNo, "",
    End If
    
    If StrConv(Y_SYU_HREC.KYOSEI_END, vbUnicode) = "1" Then
        Write #FileNo, "*" & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)),
    Else
        Write #FileNo, Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)),
    End If
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYU_HREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
            Exit Function
    End Select
    Write #FileNo, "*" & Left(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode), 12) & "*";
    Write #FileNo, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)),
    Write #FileNo, Format(CLng(StrConv(Y_SYU_HREC.SURYO, vbUnicode)), "#0"),
    If IsNumeric(StrConv(Y_SYU_HREC.J_SURYO, vbUnicode)) Then
        Write #FileNo, Format(CLng(StrConv(Y_SYU_HREC.J_SURYO, vbUnicode)), "#0"),
    Else
        Write #FileNo, ,
    End If
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.ODER_NO, vbUnicode)),
    Write #FileNo, StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode) & " " & Trim(StrConv(Y_SYU_HREC.MUKE_NAME, vbUnicode)),
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.BIKOU, vbUnicode)),
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)),
'    If IsNumeric(StrConv(Y_SYU_HREC.KUTI_SU, vbUnicode)) Then
'        Write #FileNo, Format(CInt(StrConv(Y_SYU_HREC.KUTI_SU, vbUnicode)), "#"),
'    Else
'        Write #FileNo, ,
'    End If
    
    If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) = "" Then
        Write #FileNo, ,
    Else
        Write #FileNo, Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 7, 2) & " " & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 9, 2) & ":" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 11, 2) & ":" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 13, 2),


    End If
    If Trim(StrConv(Y_SYU_HREC.KENPIN_TANTO_CODE, vbUnicode)) = "" Then
        Write #FileNo, ,
    Else
        Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(Y_SYU_HREC.KENPIN_TANTO_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound
                Call UniCode_Conv(TANTOREC.TANTO_NAME, StrConv(Y_SYU_HREC.KENPIN_TANTO_CODE, vbUnicode))
            Case Else
                Call File_Error(sts, BtOpGetEqual, "担当者ﾏｽﾀ")
                Exit Function
        End Select
    
        Write #FileNo, Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode)),
    
    End If
    Write #FileNo, "*" & Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)) & "*",
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.INPUT_BIKOU, vbUnicode)),
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.OKURI_NO_SEQ, vbUnicode)),
    Write #FileNo, Trim(StrConv(Y_SYU_HREC.SAI_SU, vbUnicode)),
    
    Write #FileNo, Trim(StrConv(ITEMREC.KONPOU_F, vbUnicode)),
    Write #FileNo, (Trim(StrConv(ITEMREC.KUTI_SU, vbUnicode)))
    
    
    
    
    
    
    
    Data_Output_Proc = False
    
    

    
End Function
Private Function List_Disp_DEN_Proc() As Integer
'----------------------------------------------------------------------------
'                   伝票№での検索
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim comw         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim DEN_MAISU   As Long
Dim KAN_MAISU   As Long
    
Dim Skip_Flg    As Boolean
    
Dim svID_No     As String * 7
    
    
Dim j           As Integer
Dim k           As Integer
Dim FOUND_FLG   As Boolean

Dim svID_TBL()  As String * 7
    
    
    
    
    List_Disp_DEN_Proc = True
                                    
'    Call Input_Lock
    F1030631.MousePointer = vbHourglass
                                    
    'ｿｰﾄ情報の初期化    2010.08.31
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
    Next i
    HEAD_CLICK = False
                                    
                                    'テーブルリセット
    Set SYUKA = Nothing
    
    DATA_FLG = False
'---------------------------------- '当日分出荷予定読み込み開始
    
    Call UniCode_Conv(K0_Y_SYU_H.DEN_NO, Text(ptxSEL_DEN_NO).Text)
    Call UniCode_Conv(K0_Y_SYU_H.SEQ_NO, "")
    sts = BTRV(BtOpGetGreaterEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)

    Select Case sts
        Case BtNoErr
    
            If Left(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode), 7) <> Trim(Text(ptxSEL_DEN_NO).Text) Then
                sts = BtErrEOF
            End If
        
        Case BtErrEOF
        
        Case Else
            Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
            List_Disp_DEN_Proc = SYS_ERR
            Exit Function
    
    End Select
    
    
    If sts = BtNoErr Then
        
        svID_No = StrConv(Y_SYU_HREC.ID_NO, vbUnicode)
        
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, svID_No)
        
        Row = Min_Row - 1
    
        com = BtOpGetGreaterEqual
        
        Do
        
            DoEvents
        
        
            sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
    
            Select Case sts
                Case BtNoErr
            
                    If Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) <> svID_No Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    List_Disp_DEN_Proc = SYS_ERR
                    Exit Function
            End Select
                    
                    
            Skip_Flg = False
            
            If Trim(Text(ptxSEL_SYUKA_YY).Text) = "" And Trim(Text(ptxSEL_SYUKA_MM).Text) = "" And Trim(Text(ptxSEL_SYUKA_DD).Text) = "" Then
            Else
                If StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode) <> _
                    (Text(ptxSEL_SYUKA_YY).Text & Text(ptxSEL_SYUKA_MM).Text & Text(ptxSEL_SYUKA_DD).Text) Then
                                                
                    Skip_Flg = True
                
                End If
            End If
                    
                    
            If Not Skip_Flg Then
                    
                    
    
    
                    
                    
                    
                    
                Row = Row + 1
                If Row > Max_Row Then
                    Beep
                    MsgBox "最大表示行数を超えました。"
                    Exit Do
                End If
                DATA_FLG = True
                        
                If Grid_Set_Proc(Row) Then
                    Exit Function
                End If
        
            
            End If
            
            com = BtOpGetNext
        
        Loop
    
    End If
    
'---------------------------------- '過日分出荷予定読み込み開始
    svID_No = ""
    j = 0
    comw = BtOpGetGreater
    
    Call UniCode_Conv(K0_wDEL_SYU_H.DEN_NO, Text(ptxSEL_DEN_NO).Text)
    Call UniCode_Conv(K0_wDEL_SYU_H.SEQ_NO, "")
    
    
    Do
        DoEvents
        sts = BTRV(comw, wDEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K0_wDEL_SYU_H, Len(K0_wDEL_SYU_H), 0)
    
        Select Case sts
            Case BtNoErr
        
                If Left(StrConv(DEL_SYU_HREC.DEN_NO, vbUnicode), 7) <> Trim(Text(ptxSEL_DEN_NO).Text) Then
                    sts = BtErrEOF
                    
                    
                   Exit Do
                End If
            
            
            
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                
                Call File_Error(sts, comw, "削除済み出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                List_Disp_DEN_Proc = SYS_ERR
                Exit Function
        
        End Select
        
        
        Skip_Flg = False
        If Trim(Text(ptxSEL_SYUKA_YY).Text) = "" And Trim(Text(ptxSEL_SYUKA_MM).Text) = "" And Trim(Text(ptxSEL_SYUKA_DD).Text) = "" Then
        Else
            If StrConv(DEL_SYU_HREC.SYUKA_YMD, vbUnicode) <> _
                (Text(ptxSEL_SYUKA_YY).Text & Text(ptxSEL_SYUKA_MM).Text & Text(ptxSEL_SYUKA_DD).Text) Then

                Skip_Flg = True

            End If
        End If
        
        If Not Skip_Flg Then
        
        
            FOUND_FLG = False
            If Trim(svID_No) = "" Then
            Else
                For k = 0 To UBound(svID_TBL)
                
                
                    If svID_TBL(j) = Left(StrConv(DEL_SYU_HREC.ID_NO, vbUnicode), 7) Then
                        FOUND_FLG = True
                        Exit For
                    End If
                
                Next k
            End If
            
            
            If FOUND_FLG Then
            Else
                
                If Trim(svID_No) = "" Then
                    ReDim svID_TBL(j)
                Else
                    j = j + 1
                    ReDim Preserve svID_TBL(j)
                End If
                svID_TBL(j) = StrConv(DEL_SYU_HREC.ID_NO, vbUnicode)
                
                svID_No = StrConv(DEL_SYU_HREC.ID_NO, vbUnicode)
                
                Call UniCode_Conv(K4_DEL_SYU_H.ID_NO, svID_No)
                
            
                com = BtOpGetGreaterEqual
            
                Do
                
                    DoEvents
                
                
                
                    sts = BTRV(com, DEL_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_DEL_SYU_H, Len(K4_DEL_SYU_H), 4)
            
                    Select Case sts
                        Case BtNoErr
                    
                            If Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) <> svID_No Then
                                sts = BtErrEOF
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                            List_Disp_DEN_Proc = SYS_ERR
                            Exit Function
                    End Select
        
                    
                    Skip_Flg = False
                    
                    If Trim(Text(ptxSEL_SYUKA_YY).Text) = "" And Trim(Text(ptxSEL_SYUKA_MM).Text) = "" And Trim(Text(ptxSEL_SYUKA_DD).Text) = "" Then
                    Else
                        If StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode) <> _
                            (Text(ptxSEL_SYUKA_YY).Text & Text(ptxSEL_SYUKA_MM).Text & Text(ptxSEL_SYUKA_DD).Text) Then
            
                            Skip_Flg = True
            
                        End If
                    End If
                        
                    If Not Skip_Flg Then
                        
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
                
            End If
        End If
        comw = BtOpGetNext
    Loop
        
    
    Set TDBGrid1.Array = SYUKA
    
    TDBGrid1.Style.Locked = False
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
    
'    Call Input_UnLock
    F1030631.MousePointer = vbDefault
   
    
    
    List_Disp_DEN_Proc = False

    
End Function


Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ検索
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim DEN_MAISU   As Long
Dim KAN_MAISU   As Long
    
Dim Skip_Flg    As Boolean
    
    
    List_Disp_Proc = True
'''    Call Input_Lock
    F1030631.MousePointer = vbHourglass
                                    
                                    
    'ｿｰﾄ情報の初期化    2010.08.31
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
    Next i
    HEAD_CLICK = False
                                    
    DATA_FLG = False
                                    
                                    'テーブルリセット
    Set SYUKA = Nothing
                                    
'---------------------------------- '当日分出荷予定読み込み開始
    If Option1(poptSEL_OKURI_NO).Value Then
    
        Call UniCode_Conv(K2_Y_SYU_H.OKURI_NO, Text(ptxSEL_OKURI_NO).Text)
    
    Else
    
        Call UniCode_Conv(K3_Y_SYU_H.SYUKA_YMD, Text(ptxSEL_SYUKA_YY).Text & _
                                                Text(ptxSEL_SYUKA_MM).Text & _
                                                Text(ptxSEL_SYUKA_DD).Text)
    End If
    
    
    
    
    
    
    Row = Min_Row - 1
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        Skip_Flg = False
        If Option1(poptSEL_OKURI_NO).Value Then
                
        
            sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K2_Y_SYU_H, Len(K2_Y_SYU_H), 2)
    
            Select Case sts
                Case BtNoErr
            
                    If Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode)) <> Trim(Text(ptxSEL_OKURI_NO).Text) Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    List_Disp_Proc = SYS_ERR
                    Exit Function
            End Select
    
    
    
        Else
    
            sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K3_Y_SYU_H, Len(K3_Y_SYU_H), 3)
    
            Select Case sts
                Case BtNoErr
                        
                    If Trim(Text(ptxSEL_SYUKA_YY).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_YY).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 1, 4) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_SYUKA_MM).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_MM).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 5, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_SYUKA_DD).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_DD).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 7, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_MUKE_CODE).Text) <> "" Then
                        If Trim(Text(ptxSEL_MUKE_CODE).Text) <> Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_HIN_NO).Text) <> "" Then
                        If Trim(Text(ptxSEL_HIN_NO).Text) <> Trim(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                        
                    If Trim(Combo(pcmbSEL_UNSOU_KAISHA).Text) <> UNSOU_KAISHA_ALL Then
                        If Trim(Combo(pcmbSEL_UNSOU_KAISHA).Text) <> Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                    
                    If Trim(Text(ptxSEL_INS_BIN).Text) <> "" Then      '2007.01.16
                        If Format(CInt(Text(ptxSEL_INS_BIN).Text), "00") <> StrConv(Y_SYU_HREC.INS_BIN, vbUnicode) Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Check1(pchkMI).Value <> vbChecked Then
                        If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) = "" Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Check1(pchkSUMI).Value <> vbChecked Then
                        If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) <> "" Then
                            Skip_Flg = True
                        End If
                    End If
                
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    List_Disp_Proc = SYS_ERR
                    Exit Function
            End Select
    
    
        End If
                
                
        If Not Skip_Flg Then
            
     If Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)) = "068742387" Then
Debug.Print
     End If
            IN_CNT = IN_CNT + 1
            If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) = "1" Then
                CANCEL_CNT = CANCEL_CNT + 1
                
            End If
            If Trim(StrConv(Y_SYU_HREC.KENPIN_TANTO_CODE, vbUnicode)) <> "" And StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) <> "1" Then
                SUMI_CNT = SUMI_CNT + 1
' Call LOG_OUT(LOG_F, StrConv(Y_SYU_HREC.ID_NO, vbUnicode) & StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) & StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode))
            
            End If
            
            If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) = "1" Or Not IsNumeric(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) Then
            
            
            
            End If
            
            
            
            Row = Row + 1
            If Row > Max_Row Then
                Beep
                MsgBox "最大表示行数を超えました。"
                Exit Do
            End If
            DATA_FLG = True
                    
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
        
    Loop
    
    
'---------------------------------- '過日分出荷予定読み込み開始
    If Option1(poptSEL_OKURI_NO).Value Then
    
        Call UniCode_Conv(K2_DEL_SYU_H.OKURI_NO, Text(ptxSEL_OKURI_NO).Text)
    
    Else
    
        Call UniCode_Conv(K3_DEL_SYU_H.SYUKA_YMD, Text(ptxSEL_SYUKA_YY).Text & _
                                                Text(ptxSEL_SYUKA_MM).Text & _
                                                Text(ptxSEL_SYUKA_DD).Text)
    
    End If
    
    
    
    
    
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        Skip_Flg = False
        If Option1(poptSEL_OKURI_NO).Value Then
                
        
            sts = BTRV(com, DEL_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K2_DEL_SYU_H, Len(K2_DEL_SYU_H), 2)
    
            Select Case sts
                Case BtNoErr
            
                    If Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode)) <> Trim(Text(ptxSEL_OKURI_NO).Text) Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    List_Disp_Proc = SYS_ERR
                    Exit Function
            End Select
    
    
    
        Else
    
            sts = BTRV(com, DEL_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K3_DEL_SYU_H, Len(K3_DEL_SYU_H), 3)
    
            Select Case sts
                Case BtNoErr
                        
                    If Trim(Text(ptxSEL_SYUKA_YY).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_YY).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 1, 4) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_SYUKA_MM).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_MM).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 5, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_SYUKA_DD).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_DD).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 7, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_MUKE_CODE).Text) <> "" Then
                        If Trim(Text(ptxSEL_MUKE_CODE).Text) <> Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_HIN_NO).Text) <> "" Then
                        If Trim(Text(ptxSEL_HIN_NO).Text) <> Trim(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                        
                    If Trim(Combo(pcmbSEL_UNSOU_KAISHA).Text) <> UNSOU_KAISHA_ALL Then
                        If Trim(Combo(pcmbSEL_UNSOU_KAISHA).Text) <> Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                
                    
                    If Trim(Text(ptxSEL_INS_BIN).Text) <> "" Then      '2007.01.16
                        If Format(CInt(Text(ptxSEL_INS_BIN).Text), "00") <> StrConv(Y_SYU_HREC.INS_BIN, vbUnicode) Then
                            Skip_Flg = True
                        End If
                    End If
                    
                    If Check1(pchkMI).Value <> vbChecked Then
                        If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) = "" Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Check1(pchkSUMI).Value <> vbChecked Then
                        If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) <> "" Then
                            Skip_Flg = True
                        End If
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    List_Disp_Proc = SYS_ERR
                    Exit Function
            End Select
    
    
        End If
                
                
        If Not Skip_Flg Then
        
            
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
    
    
    Set TDBGrid1.Array = SYUKA
    
    TDBGrid1.Style.Locked = False
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
lblRow.Caption = Row
'''    Call Input_UnLock
    F1030631.MousePointer = vbDefault
    
    
    
    List_Disp_Proc = False

    
End Function


Private Function Update_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   出荷予定更新処理
'----------------------------------------------------------------------------
Dim sts As Integer
Dim ans As Integer
Dim i   As Integer
    
    
    If TDBGrid1.Bookmark = -1 Then
        Exit Function
    End If
    
    
    Update_Proc = True
                                     'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
                                    
    Select Case Mode
        
        '1行更新(キャンセル)
        Case 0
                                    
                                            '出荷予定の読み込み
            Call UniCode_Conv(K4_Y_SYU_H.ID_NO, SYUKA(TDBGrid1.Bookmark, ColID_NO))     'ID№
            
            Do
            
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                        MsgBox "他端末で内容が変更されています。最新表示を行ってください。"
                        Update_Proc = False
                        GoTo Abort_Tran
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Update_Proc = False
                            GoTo Abort_Tran
                        End If
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                        GoTo Abort_Tran
                End Select
        
            Loop
                                            
                                            
            'ｷｬﾝｾﾙﾌﾗｸﾞ
            If SYUKA(TDBGrid1.Bookmark, ColCANCEL_F_A) = "1" Then
               SYUKA(TDBGrid1.Bookmark, ColCANCEL_F_A) = ""
               SYUKA(TDBGrid1.Bookmark, ColCANCEL_F) = ""
            Else
               SYUKA(TDBGrid1.Bookmark, ColCANCEL_F_A) = "1"
               SYUKA(TDBGrid1.Bookmark, ColCANCEL_F) = "×"
            End If
            Call UniCode_Conv(Y_SYU_HREC.CANCEL_F, SYUKA(TDBGrid1.Bookmark, ColCANCEL_F_A))
                                            '出荷予定書込み
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Update_Proc = False
                            GoTo Abort_Tran
                        End If
                
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                        GoTo Abort_Tran
                End Select
            Loop
                                        
            '2010.01.06
            If SYUKA_LOG_ON Then
                
                
                Call UniCode_Conv(K0_Y_SYU.JGYOBU, StrConv(Y_SYU_HREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, SYUKA(TDBGrid1.Bookmark, ColID_NO))
                
                
                
                sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        If SYUKA(TDBGrid1.Bookmark, ColCANCEL_F_A) = "1" Then
                            Call SYUKA_LOG_OUT_PROC("CANCEL", "ON ")
                        Else
                            Call SYUKA_LOG_OUT_PROC("CANCEL", "OFF")
                        End If
                    Case BtErrKeyNotFound
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定")
                        GoTo Abort_Tran
                End Select
                
                
                
                
            End If
            '2010.01.06
                                        
                                        
        '全行更新(備考)
        Case 1
                                        
                                        
            Set TDBGrid1.Array = SYUKA
            TDBGrid1.Refresh
            TDBGrid1.Update
                                        
                                        
            For i = 1 To SYUKA.UpperBound(1)
                                            '出荷予定の読み込み
                Call UniCode_Conv(K4_Y_SYU_H.ID_NO, SYUKA(i, ColID_NO))     'ID№
            
                Do
            
                    sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound
'''                            MsgBox "他端末で内容が変更されています。最新表示を行ってください。"
'''                            Update_Proc = False
'''                            GoTo Abort_Tran
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    
                            ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Update_Proc = False
                                GoTo Abort_Tran
                            End If
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                            GoTo Abort_Tran
                    End Select
            
                Loop
                                            
                If sts = BtNoErr Then
                                            
                                            
                    '備考
                    Call UniCode_Conv(Y_SYU_HREC.INPUT_BIKOU, SYUKA(i, ColINPUT_BIKOU))
                                                
                                                
                                                
                                                
                    '問合せ№   2014.05.15
                    Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, SYUKA(i, ColOKURI_NO))
                    '運送会社   2014.05.15
                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SYUKA(i, ColUNSOU_KAISHA))
                    
                                                
                                                
                                                '出荷予定書込み
                    Do
                        sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Update_Proc = False
                                    GoTo Abort_Tran
                                End If
                    
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                                GoTo Abort_Tran
                        End Select
                    Loop
                End If
            Next i
    End Select
End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    
        
    
    
    Set TDBGrid1.Array = SYUKA
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function


Private Function wDEL_SYU_H_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    wDEL_SYU_H_Open = True
                                            '出荷予定データフルパス取込み
    sts = GetIni("FILE", DEL_SYU_H_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [DEL_SYU_H]読み込みエラー ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wDEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
                Exit Function
        End Select
    Loop
    wDEL_SYU_H_Open = False
End Function

'Sub Form_PrintForm(obj_Pic As Object, pr_Size As Integer, pr_Orient As Integer)
''00/02/12「ＨＭ」より引用
''---------------------------------------------------------------------------
''           カレントフォームのハードコピー
''
''［引数］obj_Pic   ：ｲﾒｰｼﾞ取込み用ﾋﾟｸﾁｬｰｵﾌﾞｼﾞｪｸﾄ（FORMの見えない位置に配置）
''　　　　pr_Size   ：印刷用紙サイズ
''　　　　pr_Orient ：印刷用紙方向
''
''《キー操作について》
''　　Ｗｉｎ９５／９８ではキーを「押す」「離す」をまとめて行えるが、
''　　ＷｉｎＮＴでは「押す」「離す」を別々にしないと認識してくれない
''
''《ハードコピー使用上の注意》
''　サブCALL時点でフォーカスを持つFORMが印刷される。
''　一旦クリップボードに取り込んだ画像を、ピクチャボックスに読み込んで印刷する為、
''　画像読み込み用のピクチャボックスコントロールを引数として渡す。
''　ピクチャボックスは、FORM上の見えない位置に配置するか、Visible=Falseにする。
''
''---------------------------------------------------------------------------
'Dim sngPrnRatio As Single
'Dim sngPrnHeight As Single
'Dim sngPrnWidth As Single
'Dim sngPicPosX As Single
'Dim sngPicPosY As Single
'Dim sngPicRatio As Single
'Dim sngPicWidth As Single
'Dim sngPicHeight As Single
'
'Dim c As String
'Dim USE_Printer As String
'Dim Wk_Printer As Printer
'
'Dim Pri_Name As Printer
'
'
'
'
'
'
'
'    For Each Pri_Name In Printers
'        If Pri_Name.DeviceName = Printer.DeviceName Then
'            USE_Printer = Pri_Name.DeviceName
'            Exit For
'        End If
'    Next
'
'
'    For Each Wk_Printer In Printers
'        c = RTrim(Wk_Printer.DeviceName)
'        If Wk_Printer.DeviceName = USE_Printer Then
'            Set Printer = Wk_Printer
'            Exit For
'        End If
'    Next
'
'
'
'
'    With Printer
'        '印刷用紙の設定
'        .PaperSize = pr_Size         '用紙サイズ
'        .Orientation = pr_Orient     '上にして印刷する用紙の辺（長辺，短辺）
'
'    End With
'
'    PrintForm
'
'End Sub

Public Sub OKURI_NO_SET_PROC(OKURI_NO As String, UNSOU_KAISHA As String)
'---------------------------------- 2014.05.15
    
Dim k   As Integer
    
    If Len(Trim(OKURI_NO)) = FUKUYAMA_Length Then
        
        If FUKUYAMA_CODE(0) = "*" Then
'            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA_Name)
    
           UNSOU_KAISHA = FUKUYAMA_Name
    
        Else
            For k = 0 To UBound(FUKUYAMA_CODE)
            
                If Mid(OKURI_NO, 1, Len(FUKUYAMA_CODE(k))) = FUKUYAMA_CODE(k) Then
'                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA_Name)
                    UNSOU_KAISHA = FUKUYAMA_Name
                End If
            Next k

    
        End If
    End If

    If Len(Trim(OKURI_NO)) = SEIBU_Length Then
        If SEIBU_CODE(0) = "*" Then
            'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SEIBU_Name)
            UNSOU_KAISHA = SEIBU_Name
    
        Else
            For k = 0 To UBound(SEIBU_CODE)
            
                If Mid(OKURI_NO, 1, Len(SEIBU_CODE(k))) = SEIBU_CODE(k) Then
'                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SEIBU_Name)
                    UNSOU_KAISHA = SEIBU_Name
                End If
            Next k

    
        End If
    End If

    If Len(Trim(OKURI_NO)) = KURUME_Length Then
    
        If KURUME_CODE(0) = "*" Then
            'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME_Name)
            UNSOU_KAISHA = KURUME_Name
        Else
            For k = 0 To UBound(KURUME_CODE)
            
                If Mid(OKURI_NO, 1, Len(KURUME_CODE(k))) = KURUME_CODE(k) Then
                    'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME_Name)
                    UNSOU_KAISHA = KURUME_Name
                End If
            Next k
        End If
    
    End If

    If Len(Trim(OKURI_NO)) = SAGAWA_Length Then
    
        If SAGAWA_CODE(0) = "*" Then
            'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA_Name)
            UNSOU_KAISHA = SAGAWA_Name
        Else
            For k = 0 To UBound(SAGAWA_CODE)
            
                If Mid(OKURI_NO, 1, Len(SAGAWA_CODE(k))) = SAGAWA_CODE(k) Then
                    'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA_Name)
                    UNSOU_KAISHA = SAGAWA_Name
                End If
            Next k
        End If
    
    End If
    
    If Len(Trim(OKURI_NO)) = YAMATO_Length Then
    
        If YAMATO_CODE(0) = "*" Then
            'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, YAMATO_Name)
            UNSOU_KAISHA = YAMATO_Name
        Else
            For k = 0 To UBound(YAMATO_CODE)
            
                If Mid(OKURI_NO, 1, Len(YAMATO_CODE(k))) = YAMATO_CODE(k) Then
                    'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, YAMATO_Name)
                    UNSOU_KAISHA = YAMATO_Name
                End If
            Next k
        End If
    
    End If
    
    
    
    
    
    'ﾊﾟﾅﾎｰﾑ 2011.06.06
    If Len(Trim(OKURI_NO)) = PANA_Length Then
    
        If PANA_CODE(0) = "*" Then
            'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, PANA_Name)
            UNSOU_KAISHA = PANA_Name
        Else
            For k = 0 To UBound(PANA_CODE)
            
                If Mid(OKURI_NO, 1, Len(PANA_CODE(k))) = PANA_CODE(k) Then
                    'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, PANA_Name)
                    UNSOU_KAISHA = PANA_Name
                End If
            Next k
        End If
    
    End If
    
    '積水 2011.06.06
    If Len(Trim(OKURI_NO)) = SEKISUI_Length Then
    
        If SEKISUI_CODE(0) = "*" Then
            'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SEKISUI_Name)
            UNSOU_KAISHA = SEKISUI_Name
        Else
            For k = 0 To UBound(SEKISUI_CODE)
            
                If Mid(OKURI_NO, 1, Len(PANA_CODE(k))) = SEKISUI_CODE(k) Then
                    'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SEKISUI_Name)
                    UNSOU_KAISHA = SEKISUI_Name
                End If
            Next k
        End If
    
    End If
    '日通 2014.05.15
    If Len(Trim(OKURI_NO)) = NITTSU_Length Then
    
        If NITTSU_CODE(0) = "*" Then
            'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SEKISUI_Name)
            UNSOU_KAISHA = NITTSU_Name
        Else
            For k = 0 To UBound(NITTSU_CODE)
            
                If Mid(OKURI_NO, 1, Len(PANA_CODE(k))) = NITTSU_CODE(k) Then
                    'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SEKISUI_Name)
                    UNSOU_KAISHA = NITTSU_Name
                End If
            Next k
        End If
    
    End If
    
    
    
    
    
    '物流 2010.02.19
    If Trim(OKURI_NO) = Trim(KEN_CHARTER_CD) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "物流")
        'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KEN_CHARTER_NM)
        UNSOU_KAISHA = KEN_CHARTER_NM
    End If
    
    '赤帽 2010.02.19
    If Trim(OKURI_NO) = Trim(KEN_AKABOU_CD) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "赤帽")
        'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KEN_AKABOU_NM)
        UNSOU_KAISHA = KEN_AKABOU_NM
    End If
    
    'ロジステック 2010.02.19
    If Trim(OKURI_NO) = Trim(KEN_LOGISTIC_CD) Then
        'Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "ロジステック")
        UNSOU_KAISHA = "ロジステック"
    End If

End Sub

Public Function OKURI_NO_CHECK_PROC(OKURI_NO As String, OKURI_NO_F As Boolean, FUKUYAMA_CHK_F As Boolean) As Integer
'---------------------------------- 2014.05.15
    
    
    
Dim k   As Integer
    
    OKURI_NO_F = False
    FUKUYAMA_CHK_F = False
    
    
    If Len(Trim(OKURI_NO)) = FUKUYAMA_Length Then
        
        If FUKUYAMA_CODE(0) = "*" Then
            OKURI_NO_F = True
    
    
        Else
            For k = 0 To UBound(FUKUYAMA_CODE)
            
                If Mid(OKURI_NO, 1, Len(FUKUYAMA_CODE(k))) = FUKUYAMA_CODE(k) Then
                    OKURI_NO_F = True
                End If
            Next k

    
        End If
    
    
    
        '2010.04.28 福山チェック追加
        If Not IsNumeric(Left(OKURI_NO, FUKUYAMA_Length - 1)) Then
            FUKUYAMA_CHK_F = True
        Else
            If Left(OKURI_NO, FUKUYAMA_Length - 1) - (ToRoundDown(CCur(Left(OKURI_NO, FUKUYAMA_Length - 1) / 7), 0) * 7) <> Val(Right(OKURI_NO, 1)) Then
                FUKUYAMA_CHK_F = True
            End If
        End If
        '2010.04.28
    
    
    
    
    End If

    If Len(Trim(OKURI_NO)) = SEIBU_Length Then
        If SEIBU_CODE(0) = "*" Then
            OKURI_NO_F = True
    
    
        Else
            For k = 0 To UBound(SEIBU_CODE)
            
                If Mid(OKURI_NO, 1, Len(SEIBU_CODE(k))) = SEIBU_CODE(k) Then
                    OKURI_NO_F = True
                End If
            Next k

    
        End If
    End If

    If Len(Trim(OKURI_NO)) = KURUME_Length Then
    
        If KURUME_CODE(0) = "*" Then
            OKURI_NO_F = True
        Else
            For k = 0 To UBound(KURUME_CODE)
            
                If Mid(OKURI_NO, 1, Len(KURUME_CODE(k))) = KURUME_CODE(k) Then
                    OKURI_NO_F = True
                End If
            Next k
        End If
    
    End If

    If Len(Trim(OKURI_NO)) = SAGAWA_Length Then
    
        If SAGAWA_CODE(0) = "*" Then
            OKURI_NO_F = True
        Else
            For k = 0 To UBound(SAGAWA_CODE)
            
                If Mid(OKURI_NO, 1, Len(SAGAWA_CODE(k))) = SAGAWA_CODE(k) Then
                    OKURI_NO_F = True
                End If
            Next k
        End If
    
    End If
    
    If Len(Trim(OKURI_NO)) = YAMATO_Length Then
    
        If YAMATO_CODE(0) = "*" Then
            OKURI_NO_F = True
        Else
            For k = 0 To UBound(YAMATO_CODE)
            
                If Mid(OKURI_NO, 1, Len(YAMATO_CODE(k))) = YAMATO_CODE(k) Then
                    OKURI_NO_F = True
                End If
            Next k
        End If
    
    End If




    'ﾊﾟﾅﾎｰﾑ 2011.06.06
    If Len(Trim(OKURI_NO)) = PANA_Length Then
    
        If PANA_CODE(0) = "*" Then
            OKURI_NO_F = True
        Else
            For k = 0 To UBound(PANA_CODE)
            
                If Mid(OKURI_NO, 1, Len(PANA_CODE(k))) = PANA_CODE(k) Then
                    OKURI_NO_F = True
                End If
            Next k
        End If
    
    End If
    '積水 2011.06.06
    If Len(Trim(OKURI_NO)) = SEKISUI_Length Then
    
        If SEKISUI_CODE(0) = "*" Then
            OKURI_NO_F = True
        Else
            For k = 0 To UBound(SEKISUI_CODE)
            
                If Mid(OKURI_NO, 1, Len(SEKISUI_CODE(k))) = SEKISUI_CODE(k) Then
                    OKURI_NO_F = True
                End If
            Next k
        End If
    
    End If
    '日通 204.05.15
    If Len(Trim(OKURI_NO)) = NITTSU_Length Then
    
        If NITTSU_CODE(0) = "*" Then
            OKURI_NO_F = True
        Else
            For k = 0 To UBound(SEKISUI_CODE)
            
                If Mid(OKURI_NO, 1, Len(NITTSU_CODE(k))) = NITTSU_CODE(k) Then
                    OKURI_NO_F = True
                End If
            Next k
        End If
    
    End If



    'ﾁｬﾀｰ便
    If Trim(OKURI_NO) = Trim(KEN_CHARTER_CD) Then
        OKURI_NO_F = True
    End If
    
    '赤帽
    If Trim(OKURI_NO) = Trim(KEN_AKABOU_CD) Then
        OKURI_NO_F = True
    End If
    
    'ロジステック
    If Trim(OKURI_NO) = Trim(KEN_LOGISTIC_CD) Then
        OKURI_NO_F = True
    End If



End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に切り捨てします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り捨てられた数値。
'
'
'       2012.03.25  frm より　移管
'
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


Private Function DBGrid_Chk_Proc() As Integer
'2014.05.15
Dim i           As Integer
Dim OKURI_NO    As String * 20
    
Dim OKURI_NO_F      As Boolean
Dim FUKUYAMA_CHK_F  As Boolean
    
    
    If TDBGrid1.Bookmark = -1 Then
        Exit Function
    End If
    
    
    DBGrid_Chk_Proc = True
    
    
    
    Set TDBGrid1.Array = SYUKA
    TDBGrid1.Refresh
    TDBGrid1.Update
                                        
                                        
    For i = 1 To SYUKA.UpperBound(1)

        OKURI_NO = SYUKA(i, ColOKURI_NO)
        
        If Trim(OKURI_NO) = "" Then
        Else
            If Not IsNumeric(OKURI_NO) Then
                MsgBox i & "行目　送り状№エラー"
                        
                                            
                TDBGrid1.Bookmark = i
                TDBGrid1.Col = ColOKURI_NO
                
                
                Set TDBGrid1.Array = SYUKA
        
    
                TDBGrid1.Refresh
                TDBGrid1.Update
                
                TDBGrid1.SetFocus
                        
                        
                    
                        
                Exit Function
            
            End If
        
        
        
            If OKURI_NO_CHECK_PROC(OKURI_NO, OKURI_NO_F, FUKUYAMA_CHK_F) Then
            End If
            If Not OKURI_NO_F Then
                MsgBox i & "行目　送り状№エラー"
                        
                                            
                TDBGrid1.Bookmark = i
                TDBGrid1.Col = ColOKURI_NO
                
                
                Set TDBGrid1.Array = SYUKA
        
    
                TDBGrid1.Refresh
                TDBGrid1.Update
                
                TDBGrid1.SetFocus
                        
                        
                    
                        
                Exit Function
            End If
       
        
                If FUKUYAMA_CHK_F Then
                    MsgBox i & "行目　福山 ﾁｪｯｸﾃﾞｼﾞｯﾄｴﾗｰ"
                
                    TDBGrid1.Bookmark = i
                    TDBGrid1.Col = ColOKURI_NO
                    
                    
                    Set TDBGrid1.Array = SYUKA
            
        
                    TDBGrid1.Refresh
                    TDBGrid1.Update
                    
                    TDBGrid1.SetFocus
                            
                            
                        
                            
                    Exit Function
                End If
        
        
        
        End If

    Next i

    DBGrid_Chk_Proc = False

End Function
