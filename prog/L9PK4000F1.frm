VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form L9PK4000F1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "包　装　計　画"
   ClientHeight    =   10272
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10272
   ScaleWidth      =   15240
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "9"
      Top             =   480
      Width           =   300
   End
   Begin VB.Timer Tim_Disp 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10560
      Top             =   120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "表示対象切替"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   10965
      TabIndex        =   33
      Top             =   105
      Width           =   2895
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "未確定"
         Height          =   360
         Index           =   2
         Left            =   1815
         TabIndex        =   36
         Top             =   240
         Width           =   960
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "確定"
         Height          =   240
         Index           =   1
         Left            =   1035
         TabIndex        =   35
         Top             =   300
         Width           =   765
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "全て"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   34
         Top             =   300
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   8265
      Left            =   60
      TabIndex        =   32
      Top             =   975
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   14584
      _LayoutType     =   4
      _RowHeight      =   21
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   512
      Columns(1)._MaxComboItems=   30
      Columns(1).ValueItems(0)._DefaultItem=   0
      Columns(1).ValueItems(0).Value=   "10"
      Columns(1).ValueItems(0).Value.vt=   8
      Columns(1).ValueItems(0).DisplayValue=   "　"
      Columns(1).ValueItems(0).DisplayValue.vt=   8
      Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(1)._DefaultItem=   0
      Columns(1).ValueItems(1).Value=   "20"
      Columns(1).ValueItems(1).Value.vt=   8
      Columns(1).ValueItems(1).DisplayValue=   "　"
      Columns(1).ValueItems(1).DisplayValue.vt=   8
      Columns(1).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(2)._DefaultItem=   0
      Columns(1).ValueItems(2).Value=   "30"
      Columns(1).ValueItems(2).Value.vt=   8
      Columns(1).ValueItems(2).DisplayValue=   "　"
      Columns(1).ValueItems(2).DisplayValue.vt=   8
      Columns(1).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(3)._DefaultItem=   0
      Columns(1).ValueItems(3).Value=   "40"
      Columns(1).ValueItems(3).Value.vt=   8
      Columns(1).ValueItems(3).DisplayValue=   "　"
      Columns(1).ValueItems(3).DisplayValue.vt=   8
      Columns(1).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(4)._DefaultItem=   0
      Columns(1).ValueItems(4).Value=   "50"
      Columns(1).ValueItems(4).Value.vt=   8
      Columns(1).ValueItems(4).DisplayValue=   "　"
      Columns(1).ValueItems(4).DisplayValue.vt=   8
      Columns(1).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(5)._DefaultItem=   0
      Columns(1).ValueItems(5).Value=   "60"
      Columns(1).ValueItems(5).Value.vt=   8
      Columns(1).ValueItems(5).DisplayValue=   "　"
      Columns(1).ValueItems(5).DisplayValue.vt=   8
      Columns(1).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(6)._DefaultItem=   0
      Columns(1).ValueItems(6).Value=   "70"
      Columns(1).ValueItems(6).Value.vt=   8
      Columns(1).ValueItems(6).DisplayValue=   "　"
      Columns(1).ValueItems(6).DisplayValue.vt=   8
      Columns(1).ValueItems(6)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(7)._DefaultItem=   0
      Columns(1).ValueItems(7).Value=   "80"
      Columns(1).ValueItems(7).Value.vt=   8
      Columns(1).ValueItems(7).DisplayValue=   "　"
      Columns(1).ValueItems(7).DisplayValue.vt=   8
      Columns(1).ValueItems(7)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(8)._DefaultItem=   0
      Columns(1).ValueItems(8).Value=   "90"
      Columns(1).ValueItems(8).Value.vt=   8
      Columns(1).ValueItems(8).DisplayValue=   "　"
      Columns(1).ValueItems(8).DisplayValue.vt=   8
      Columns(1).ValueItems(8)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(9)._DefaultItem=   0
      Columns(1).ValueItems(9).Value=   "100"
      Columns(1).ValueItems(9).Value.vt=   8
      Columns(1).ValueItems(9).DisplayValue=   "　"
      Columns(1).ValueItems(9).DisplayValue.vt=   8
      Columns(1).ValueItems(9)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(10)._DefaultItem=   0
      Columns(1).ValueItems(10).Value=   "110"
      Columns(1).ValueItems(10).Value.vt=   8
      Columns(1).ValueItems(10).DisplayValue=   "　"
      Columns(1).ValueItems(10).DisplayValue.vt=   8
      Columns(1).ValueItems(10)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(11)._DefaultItem=   0
      Columns(1).ValueItems(11).Value=   "120"
      Columns(1).ValueItems(11).Value.vt=   8
      Columns(1).ValueItems(11).DisplayValue=   "　"
      Columns(1).ValueItems(11).DisplayValue.vt=   8
      Columns(1).ValueItems(11)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(12)._DefaultItem=   0
      Columns(1).ValueItems(12).Value=   "130"
      Columns(1).ValueItems(12).Value.vt=   8
      Columns(1).ValueItems(12).DisplayValue=   "　"
      Columns(1).ValueItems(12).DisplayValue.vt=   8
      Columns(1).ValueItems(12)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(13)._DefaultItem=   0
      Columns(1).ValueItems(13).Value=   "140"
      Columns(1).ValueItems(13).Value.vt=   8
      Columns(1).ValueItems(13).DisplayValue=   "　"
      Columns(1).ValueItems(13).DisplayValue.vt=   8
      Columns(1).ValueItems(13)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(14)._DefaultItem=   0
      Columns(1).ValueItems(14).Value=   "150"
      Columns(1).ValueItems(14).Value.vt=   8
      Columns(1).ValueItems(14).DisplayValue=   "　"
      Columns(1).ValueItems(14).DisplayValue.vt=   8
      Columns(1).ValueItems(14)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(15)._DefaultItem=   0
      Columns(1).ValueItems(15).Value=   "160"
      Columns(1).ValueItems(15).Value.vt=   8
      Columns(1).ValueItems(15).DisplayValue=   "　"
      Columns(1).ValueItems(15).DisplayValue.vt=   8
      Columns(1).ValueItems(15)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(16)._DefaultItem=   0
      Columns(1).ValueItems(16).Value=   "170"
      Columns(1).ValueItems(16).Value.vt=   8
      Columns(1).ValueItems(16).DisplayValue=   "　"
      Columns(1).ValueItems(16).DisplayValue.vt=   8
      Columns(1).ValueItems(16)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(17)._DefaultItem=   0
      Columns(1).ValueItems(17).Value=   "180"
      Columns(1).ValueItems(17).Value.vt=   8
      Columns(1).ValueItems(17).DisplayValue=   "　"
      Columns(1).ValueItems(17).DisplayValue.vt=   8
      Columns(1).ValueItems(17)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(18)._DefaultItem=   0
      Columns(1).ValueItems(18).Value=   "190"
      Columns(1).ValueItems(18).Value.vt=   8
      Columns(1).ValueItems(18).DisplayValue=   "　"
      Columns(1).ValueItems(18).DisplayValue.vt=   8
      Columns(1).ValueItems(18)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(19)._DefaultItem=   0
      Columns(1).ValueItems(19).Value=   "200"
      Columns(1).ValueItems(19).Value.vt=   8
      Columns(1).ValueItems(19).DisplayValue=   "　"
      Columns(1).ValueItems(19).DisplayValue.vt=   8
      Columns(1).ValueItems(19)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(20)._DefaultItem=   -1
      Columns(1).ValueItems(20).Value=   "210"
      Columns(1).ValueItems(20).Value.vt=   8
      Columns(1).ValueItems(20).DisplayValue=   "　"
      Columns(1).ValueItems(20).DisplayValue.vt=   8
      Columns(1).ValueItems(20)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(21)._DefaultItem=   0
      Columns(1).ValueItems(21).Value=   "220"
      Columns(1).ValueItems(21).Value.vt=   8
      Columns(1).ValueItems(21).DisplayValue=   "　"
      Columns(1).ValueItems(21).DisplayValue.vt=   8
      Columns(1).ValueItems(21)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(22)._DefaultItem=   0
      Columns(1).ValueItems(22).Value=   "230"
      Columns(1).ValueItems(22).Value.vt=   8
      Columns(1).ValueItems(22).DisplayValue=   "　"
      Columns(1).ValueItems(22).DisplayValue.vt=   8
      Columns(1).ValueItems(22)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(23)._DefaultItem=   0
      Columns(1).ValueItems(23).Value=   "240"
      Columns(1).ValueItems(23).Value.vt=   8
      Columns(1).ValueItems(23).DisplayValue=   "　"
      Columns(1).ValueItems(23).DisplayValue.vt=   8
      Columns(1).ValueItems(23)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(24)._DefaultItem=   0
      Columns(1).ValueItems(24).Value=   "250"
      Columns(1).ValueItems(24).Value.vt=   8
      Columns(1).ValueItems(24).DisplayValue=   "　"
      Columns(1).ValueItems(24).DisplayValue.vt=   8
      Columns(1).ValueItems(24)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(25)._DefaultItem=   0
      Columns(1).ValueItems(25).Value=   "260"
      Columns(1).ValueItems(25).Value.vt=   8
      Columns(1).ValueItems(25).DisplayValue=   "　"
      Columns(1).ValueItems(25).DisplayValue.vt=   8
      Columns(1).ValueItems(25)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(26)._DefaultItem=   0
      Columns(1).ValueItems(26).Value=   "270"
      Columns(1).ValueItems(26).Value.vt=   8
      Columns(1).ValueItems(26).DisplayValue=   "　"
      Columns(1).ValueItems(26).DisplayValue.vt=   8
      Columns(1).ValueItems(26)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(27)._DefaultItem=   0
      Columns(1).ValueItems(27).Value=   "280"
      Columns(1).ValueItems(27).Value.vt=   8
      Columns(1).ValueItems(27).DisplayValue=   "　"
      Columns(1).ValueItems(27).DisplayValue.vt=   8
      Columns(1).ValueItems(27)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(28)._DefaultItem=   0
      Columns(1).ValueItems(28).Value=   "290"
      Columns(1).ValueItems(28).Value.vt=   8
      Columns(1).ValueItems(28).DisplayValue=   "　"
      Columns(1).ValueItems(28).DisplayValue.vt=   8
      Columns(1).ValueItems(28)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(29)._DefaultItem=   0
      Columns(1).ValueItems(29).Value=   "300"
      Columns(1).ValueItems(29).Value.vt=   8
      Columns(1).ValueItems(29).DisplayValue=   "　"
      Columns(1).ValueItems(29).DisplayValue.vt=   8
      Columns(1).ValueItems(29)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(30)._DefaultItem=   0
      Columns(1).ValueItems(30).Value=   "11"
      Columns(1).ValueItems(30).Value.vt=   8
      Columns(1).ValueItems(30).DisplayValue=   "確"
      Columns(1).ValueItems(30).DisplayValue.vt=   8
      Columns(1).ValueItems(30)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(31)._DefaultItem=   0
      Columns(1).ValueItems(31).Value=   "12"
      Columns(1).ValueItems(31).Value.vt=   8
      Columns(1).ValueItems(31).DisplayValue=   "再"
      Columns(1).ValueItems(31).DisplayValue.vt=   8
      Columns(1).ValueItems(31)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(32)._DefaultItem=   0
      Columns(1).ValueItems(32).Value=   "19"
      Columns(1).ValueItems(32).Value.vt=   8
      Columns(1).ValueItems(32).DisplayValue=   "レ"
      Columns(1).ValueItems(32).DisplayValue.vt=   8
      Columns(1).ValueItems(32)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(33)._DefaultItem=   0
      Columns(1).ValueItems(33).Value=   "21"
      Columns(1).ValueItems(33).Value.vt=   8
      Columns(1).ValueItems(33).DisplayValue=   "確"
      Columns(1).ValueItems(33).DisplayValue.vt=   8
      Columns(1).ValueItems(33)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(34)._DefaultItem=   0
      Columns(1).ValueItems(34).Value=   "22"
      Columns(1).ValueItems(34).Value.vt=   8
      Columns(1).ValueItems(34).DisplayValue=   "再"
      Columns(1).ValueItems(34).DisplayValue.vt=   8
      Columns(1).ValueItems(34)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(35)._DefaultItem=   0
      Columns(1).ValueItems(35).Value=   "29"
      Columns(1).ValueItems(35).Value.vt=   8
      Columns(1).ValueItems(35).DisplayValue=   "レ"
      Columns(1).ValueItems(35).DisplayValue.vt=   8
      Columns(1).ValueItems(35)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(36)._DefaultItem=   0
      Columns(1).ValueItems(36).Value=   "31"
      Columns(1).ValueItems(36).Value.vt=   8
      Columns(1).ValueItems(36).DisplayValue=   "確"
      Columns(1).ValueItems(36).DisplayValue.vt=   8
      Columns(1).ValueItems(36)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(37)._DefaultItem=   0
      Columns(1).ValueItems(37).Value=   "32"
      Columns(1).ValueItems(37).Value.vt=   8
      Columns(1).ValueItems(37).DisplayValue=   "再"
      Columns(1).ValueItems(37).DisplayValue.vt=   8
      Columns(1).ValueItems(37)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(38)._DefaultItem=   0
      Columns(1).ValueItems(38).Value=   "39"
      Columns(1).ValueItems(38).Value.vt=   8
      Columns(1).ValueItems(38).DisplayValue=   "レ"
      Columns(1).ValueItems(38).DisplayValue.vt=   8
      Columns(1).ValueItems(38)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(39)._DefaultItem=   0
      Columns(1).ValueItems(39).Value=   "41"
      Columns(1).ValueItems(39).Value.vt=   8
      Columns(1).ValueItems(39).DisplayValue=   "確"
      Columns(1).ValueItems(39).DisplayValue.vt=   8
      Columns(1).ValueItems(39)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(40)._DefaultItem=   0
      Columns(1).ValueItems(40).Value=   "42"
      Columns(1).ValueItems(40).Value.vt=   8
      Columns(1).ValueItems(40).DisplayValue=   "再"
      Columns(1).ValueItems(40).DisplayValue.vt=   8
      Columns(1).ValueItems(40)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(41)._DefaultItem=   0
      Columns(1).ValueItems(41).Value=   "49"
      Columns(1).ValueItems(41).Value.vt=   8
      Columns(1).ValueItems(41).DisplayValue=   "レ"
      Columns(1).ValueItems(41).DisplayValue.vt=   8
      Columns(1).ValueItems(41)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(42)._DefaultItem=   0
      Columns(1).ValueItems(42).Value=   "51"
      Columns(1).ValueItems(42).Value.vt=   8
      Columns(1).ValueItems(42).DisplayValue=   "確"
      Columns(1).ValueItems(42).DisplayValue.vt=   8
      Columns(1).ValueItems(42)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(43)._DefaultItem=   0
      Columns(1).ValueItems(43).Value=   "52"
      Columns(1).ValueItems(43).Value.vt=   8
      Columns(1).ValueItems(43).DisplayValue=   "再"
      Columns(1).ValueItems(43).DisplayValue.vt=   8
      Columns(1).ValueItems(43)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(44)._DefaultItem=   0
      Columns(1).ValueItems(44).Value=   "59"
      Columns(1).ValueItems(44).Value.vt=   8
      Columns(1).ValueItems(44).DisplayValue=   "レ"
      Columns(1).ValueItems(44).DisplayValue.vt=   8
      Columns(1).ValueItems(44)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(45)._DefaultItem=   0
      Columns(1).ValueItems(45).Value=   "61"
      Columns(1).ValueItems(45).Value.vt=   8
      Columns(1).ValueItems(45).DisplayValue=   "確"
      Columns(1).ValueItems(45).DisplayValue.vt=   8
      Columns(1).ValueItems(45)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(46)._DefaultItem=   0
      Columns(1).ValueItems(46).Value=   "62"
      Columns(1).ValueItems(46).Value.vt=   8
      Columns(1).ValueItems(46).DisplayValue=   "再"
      Columns(1).ValueItems(46).DisplayValue.vt=   8
      Columns(1).ValueItems(46)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(47)._DefaultItem=   0
      Columns(1).ValueItems(47).Value=   "69"
      Columns(1).ValueItems(47).Value.vt=   8
      Columns(1).ValueItems(47).DisplayValue=   "レ"
      Columns(1).ValueItems(47).DisplayValue.vt=   8
      Columns(1).ValueItems(47)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(48)._DefaultItem=   0
      Columns(1).ValueItems(48).Value=   "71"
      Columns(1).ValueItems(48).Value.vt=   8
      Columns(1).ValueItems(48).DisplayValue=   "確"
      Columns(1).ValueItems(48).DisplayValue.vt=   8
      Columns(1).ValueItems(48)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(49)._DefaultItem=   0
      Columns(1).ValueItems(49).Value=   "72"
      Columns(1).ValueItems(49).Value.vt=   8
      Columns(1).ValueItems(49).DisplayValue=   "再"
      Columns(1).ValueItems(49).DisplayValue.vt=   8
      Columns(1).ValueItems(49)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(50)._DefaultItem=   0
      Columns(1).ValueItems(50).Value=   "79"
      Columns(1).ValueItems(50).Value.vt=   8
      Columns(1).ValueItems(50).DisplayValue=   "レ"
      Columns(1).ValueItems(50).DisplayValue.vt=   8
      Columns(1).ValueItems(50)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(51)._DefaultItem=   0
      Columns(1).ValueItems(51).Value=   "81"
      Columns(1).ValueItems(51).Value.vt=   8
      Columns(1).ValueItems(51).DisplayValue=   "確"
      Columns(1).ValueItems(51).DisplayValue.vt=   8
      Columns(1).ValueItems(51)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(52)._DefaultItem=   0
      Columns(1).ValueItems(52).Value=   "82"
      Columns(1).ValueItems(52).Value.vt=   8
      Columns(1).ValueItems(52).DisplayValue=   "再"
      Columns(1).ValueItems(52).DisplayValue.vt=   8
      Columns(1).ValueItems(52)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(53)._DefaultItem=   0
      Columns(1).ValueItems(53).Value=   "89"
      Columns(1).ValueItems(53).Value.vt=   8
      Columns(1).ValueItems(53).DisplayValue=   "レ"
      Columns(1).ValueItems(53).DisplayValue.vt=   8
      Columns(1).ValueItems(53)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(54)._DefaultItem=   0
      Columns(1).ValueItems(54).Value=   "91"
      Columns(1).ValueItems(54).Value.vt=   8
      Columns(1).ValueItems(54).DisplayValue=   "確"
      Columns(1).ValueItems(54).DisplayValue.vt=   8
      Columns(1).ValueItems(54)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(55)._DefaultItem=   0
      Columns(1).ValueItems(55).Value=   "92"
      Columns(1).ValueItems(55).Value.vt=   8
      Columns(1).ValueItems(55).DisplayValue=   "再"
      Columns(1).ValueItems(55).DisplayValue.vt=   8
      Columns(1).ValueItems(55)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(56)._DefaultItem=   0
      Columns(1).ValueItems(56).Value=   "99"
      Columns(1).ValueItems(56).Value.vt=   8
      Columns(1).ValueItems(56).DisplayValue=   "レ"
      Columns(1).ValueItems(56).DisplayValue.vt=   8
      Columns(1).ValueItems(56)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(57)._DefaultItem=   0
      Columns(1).ValueItems(57).Value=   "101"
      Columns(1).ValueItems(57).Value.vt=   8
      Columns(1).ValueItems(57).DisplayValue=   "確"
      Columns(1).ValueItems(57).DisplayValue.vt=   8
      Columns(1).ValueItems(57)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(58)._DefaultItem=   0
      Columns(1).ValueItems(58).Value=   "111"
      Columns(1).ValueItems(58).Value.vt=   8
      Columns(1).ValueItems(58).DisplayValue=   "確"
      Columns(1).ValueItems(58).DisplayValue.vt=   8
      Columns(1).ValueItems(58)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(59)._DefaultItem=   0
      Columns(1).ValueItems(59).Value=   "121"
      Columns(1).ValueItems(59).Value.vt=   8
      Columns(1).ValueItems(59).DisplayValue=   "確"
      Columns(1).ValueItems(59).DisplayValue.vt=   8
      Columns(1).ValueItems(59)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(60)._DefaultItem=   0
      Columns(1).ValueItems(60).Value=   "131"
      Columns(1).ValueItems(60).Value.vt=   8
      Columns(1).ValueItems(60).DisplayValue=   "確"
      Columns(1).ValueItems(60).DisplayValue.vt=   8
      Columns(1).ValueItems(60)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(61)._DefaultItem=   0
      Columns(1).ValueItems(61).Value=   "141"
      Columns(1).ValueItems(61).Value.vt=   8
      Columns(1).ValueItems(61).DisplayValue=   "確"
      Columns(1).ValueItems(61).DisplayValue.vt=   8
      Columns(1).ValueItems(61)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(62)._DefaultItem=   0
      Columns(1).ValueItems(62).Value=   "151"
      Columns(1).ValueItems(62).Value.vt=   8
      Columns(1).ValueItems(62).DisplayValue=   "確"
      Columns(1).ValueItems(62).DisplayValue.vt=   8
      Columns(1).ValueItems(62)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(63)._DefaultItem=   0
      Columns(1).ValueItems(63).Value=   "161"
      Columns(1).ValueItems(63).Value.vt=   8
      Columns(1).ValueItems(63).DisplayValue=   "確"
      Columns(1).ValueItems(63).DisplayValue.vt=   8
      Columns(1).ValueItems(63)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(64)._DefaultItem=   0
      Columns(1).ValueItems(64).Value=   "171"
      Columns(1).ValueItems(64).Value.vt=   8
      Columns(1).ValueItems(64).DisplayValue=   "確"
      Columns(1).ValueItems(64).DisplayValue.vt=   8
      Columns(1).ValueItems(64)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(65)._DefaultItem=   0
      Columns(1).ValueItems(65).Value=   "181"
      Columns(1).ValueItems(65).Value.vt=   8
      Columns(1).ValueItems(65).DisplayValue=   "確"
      Columns(1).ValueItems(65).DisplayValue.vt=   8
      Columns(1).ValueItems(65)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(66)._DefaultItem=   0
      Columns(1).ValueItems(66).Value=   "191"
      Columns(1).ValueItems(66).Value.vt=   8
      Columns(1).ValueItems(66).DisplayValue=   "確"
      Columns(1).ValueItems(66).DisplayValue.vt=   8
      Columns(1).ValueItems(66)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(67)._DefaultItem=   0
      Columns(1).ValueItems(67).Value=   "201"
      Columns(1).ValueItems(67).Value.vt=   8
      Columns(1).ValueItems(67).DisplayValue=   "確"
      Columns(1).ValueItems(67).DisplayValue.vt=   8
      Columns(1).ValueItems(67)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(68)._DefaultItem=   0
      Columns(1).ValueItems(68).Value=   "211"
      Columns(1).ValueItems(68).Value.vt=   8
      Columns(1).ValueItems(68).DisplayValue=   "確"
      Columns(1).ValueItems(68).DisplayValue.vt=   8
      Columns(1).ValueItems(68)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(69)._DefaultItem=   0
      Columns(1).ValueItems(69).Value=   "221"
      Columns(1).ValueItems(69).Value.vt=   8
      Columns(1).ValueItems(69).DisplayValue=   "確"
      Columns(1).ValueItems(69).DisplayValue.vt=   8
      Columns(1).ValueItems(69)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(70)._DefaultItem=   0
      Columns(1).ValueItems(70).Value=   "231"
      Columns(1).ValueItems(70).Value.vt=   8
      Columns(1).ValueItems(70).DisplayValue=   "確"
      Columns(1).ValueItems(70).DisplayValue.vt=   8
      Columns(1).ValueItems(70)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(71)._DefaultItem=   0
      Columns(1).ValueItems(71).Value=   "241"
      Columns(1).ValueItems(71).Value.vt=   8
      Columns(1).ValueItems(71).DisplayValue=   "確"
      Columns(1).ValueItems(71).DisplayValue.vt=   8
      Columns(1).ValueItems(71)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(72)._DefaultItem=   0
      Columns(1).ValueItems(72).Value=   "251"
      Columns(1).ValueItems(72).Value.vt=   8
      Columns(1).ValueItems(72).DisplayValue=   "確"
      Columns(1).ValueItems(72).DisplayValue.vt=   8
      Columns(1).ValueItems(72)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(73)._DefaultItem=   0
      Columns(1).ValueItems(73).Value=   "261"
      Columns(1).ValueItems(73).Value.vt=   8
      Columns(1).ValueItems(73).DisplayValue=   "確"
      Columns(1).ValueItems(73).DisplayValue.vt=   8
      Columns(1).ValueItems(73)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(74)._DefaultItem=   0
      Columns(1).ValueItems(74).Value=   "271"
      Columns(1).ValueItems(74).Value.vt=   8
      Columns(1).ValueItems(74).DisplayValue=   "確"
      Columns(1).ValueItems(74).DisplayValue.vt=   8
      Columns(1).ValueItems(74)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(75)._DefaultItem=   0
      Columns(1).ValueItems(75).Value=   "281"
      Columns(1).ValueItems(75).Value.vt=   8
      Columns(1).ValueItems(75).DisplayValue=   "確"
      Columns(1).ValueItems(75).DisplayValue.vt=   8
      Columns(1).ValueItems(75)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(76)._DefaultItem=   0
      Columns(1).ValueItems(76).Value=   "291"
      Columns(1).ValueItems(76).Value.vt=   8
      Columns(1).ValueItems(76).DisplayValue=   "確"
      Columns(1).ValueItems(76).DisplayValue.vt=   8
      Columns(1).ValueItems(76)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(77)._DefaultItem=   0
      Columns(1).ValueItems(77).Value=   "301"
      Columns(1).ValueItems(77).Value.vt=   8
      Columns(1).ValueItems(77).DisplayValue=   "確"
      Columns(1).ValueItems(77).DisplayValue.vt=   8
      Columns(1).ValueItems(77)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(78)._DefaultItem=   0
      Columns(1).ValueItems(78).Value=   "102"
      Columns(1).ValueItems(78).Value.vt=   8
      Columns(1).ValueItems(78).DisplayValue=   "再"
      Columns(1).ValueItems(78).DisplayValue.vt=   8
      Columns(1).ValueItems(78)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(79)._DefaultItem=   0
      Columns(1).ValueItems(79).Value=   "112"
      Columns(1).ValueItems(79).Value.vt=   8
      Columns(1).ValueItems(79).DisplayValue=   "再"
      Columns(1).ValueItems(79).DisplayValue.vt=   8
      Columns(1).ValueItems(79)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(80)._DefaultItem=   0
      Columns(1).ValueItems(80).Value=   "122"
      Columns(1).ValueItems(80).Value.vt=   8
      Columns(1).ValueItems(80).DisplayValue=   "再"
      Columns(1).ValueItems(80).DisplayValue.vt=   8
      Columns(1).ValueItems(80)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(81)._DefaultItem=   0
      Columns(1).ValueItems(81).Value=   "132"
      Columns(1).ValueItems(81).Value.vt=   8
      Columns(1).ValueItems(81).DisplayValue=   "再"
      Columns(1).ValueItems(81).DisplayValue.vt=   8
      Columns(1).ValueItems(81)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(82)._DefaultItem=   0
      Columns(1).ValueItems(82).Value=   "142"
      Columns(1).ValueItems(82).Value.vt=   8
      Columns(1).ValueItems(82).DisplayValue=   "再"
      Columns(1).ValueItems(82).DisplayValue.vt=   8
      Columns(1).ValueItems(82)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(83)._DefaultItem=   0
      Columns(1).ValueItems(83).Value=   "152"
      Columns(1).ValueItems(83).Value.vt=   8
      Columns(1).ValueItems(83).DisplayValue=   "再"
      Columns(1).ValueItems(83).DisplayValue.vt=   8
      Columns(1).ValueItems(83)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(84)._DefaultItem=   0
      Columns(1).ValueItems(84).Value=   "162"
      Columns(1).ValueItems(84).Value.vt=   8
      Columns(1).ValueItems(84).DisplayValue=   "再"
      Columns(1).ValueItems(84).DisplayValue.vt=   8
      Columns(1).ValueItems(84)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(85)._DefaultItem=   0
      Columns(1).ValueItems(85).Value=   "172"
      Columns(1).ValueItems(85).Value.vt=   8
      Columns(1).ValueItems(85).DisplayValue=   "再"
      Columns(1).ValueItems(85).DisplayValue.vt=   8
      Columns(1).ValueItems(85)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(86)._DefaultItem=   0
      Columns(1).ValueItems(86).Value=   "182"
      Columns(1).ValueItems(86).Value.vt=   8
      Columns(1).ValueItems(86).DisplayValue=   "再"
      Columns(1).ValueItems(86).DisplayValue.vt=   8
      Columns(1).ValueItems(86)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(87)._DefaultItem=   0
      Columns(1).ValueItems(87).Value=   "192"
      Columns(1).ValueItems(87).Value.vt=   8
      Columns(1).ValueItems(87).DisplayValue=   "再"
      Columns(1).ValueItems(87).DisplayValue.vt=   8
      Columns(1).ValueItems(87)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(88)._DefaultItem=   0
      Columns(1).ValueItems(88).Value=   "202"
      Columns(1).ValueItems(88).Value.vt=   8
      Columns(1).ValueItems(88).DisplayValue=   "再"
      Columns(1).ValueItems(88).DisplayValue.vt=   8
      Columns(1).ValueItems(88)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(89)._DefaultItem=   0
      Columns(1).ValueItems(89).Value=   "212"
      Columns(1).ValueItems(89).Value.vt=   8
      Columns(1).ValueItems(89).DisplayValue=   "再"
      Columns(1).ValueItems(89).DisplayValue.vt=   8
      Columns(1).ValueItems(89)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(90)._DefaultItem=   0
      Columns(1).ValueItems(90).Value=   "222"
      Columns(1).ValueItems(90).Value.vt=   8
      Columns(1).ValueItems(90).DisplayValue=   "再"
      Columns(1).ValueItems(90).DisplayValue.vt=   8
      Columns(1).ValueItems(90)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(91)._DefaultItem=   0
      Columns(1).ValueItems(91).Value=   "232"
      Columns(1).ValueItems(91).Value.vt=   8
      Columns(1).ValueItems(91).DisplayValue=   "再"
      Columns(1).ValueItems(91).DisplayValue.vt=   8
      Columns(1).ValueItems(91)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(92)._DefaultItem=   0
      Columns(1).ValueItems(92).Value=   "242"
      Columns(1).ValueItems(92).Value.vt=   8
      Columns(1).ValueItems(92).DisplayValue=   "再"
      Columns(1).ValueItems(92).DisplayValue.vt=   8
      Columns(1).ValueItems(92)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(93)._DefaultItem=   0
      Columns(1).ValueItems(93).Value=   "252"
      Columns(1).ValueItems(93).Value.vt=   8
      Columns(1).ValueItems(93).DisplayValue=   "再"
      Columns(1).ValueItems(93).DisplayValue.vt=   8
      Columns(1).ValueItems(93)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(94)._DefaultItem=   0
      Columns(1).ValueItems(94).Value=   "262"
      Columns(1).ValueItems(94).Value.vt=   8
      Columns(1).ValueItems(94).DisplayValue=   "再"
      Columns(1).ValueItems(94).DisplayValue.vt=   8
      Columns(1).ValueItems(94)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(95)._DefaultItem=   0
      Columns(1).ValueItems(95).Value=   "272"
      Columns(1).ValueItems(95).Value.vt=   8
      Columns(1).ValueItems(95).DisplayValue=   "再"
      Columns(1).ValueItems(95).DisplayValue.vt=   8
      Columns(1).ValueItems(95)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(96)._DefaultItem=   0
      Columns(1).ValueItems(96).Value=   "282"
      Columns(1).ValueItems(96).Value.vt=   8
      Columns(1).ValueItems(96).DisplayValue=   "再"
      Columns(1).ValueItems(96).DisplayValue.vt=   8
      Columns(1).ValueItems(96)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(97)._DefaultItem=   0
      Columns(1).ValueItems(97).Value=   "292"
      Columns(1).ValueItems(97).Value.vt=   8
      Columns(1).ValueItems(97).DisplayValue=   "再"
      Columns(1).ValueItems(97).DisplayValue.vt=   8
      Columns(1).ValueItems(97)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(98)._DefaultItem=   0
      Columns(1).ValueItems(98).Value=   "302"
      Columns(1).ValueItems(98).Value.vt=   8
      Columns(1).ValueItems(98).DisplayValue=   "再"
      Columns(1).ValueItems(98).DisplayValue.vt=   8
      Columns(1).ValueItems(98)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(99)._DefaultItem=   0
      Columns(1).ValueItems(99).Value=   "109"
      Columns(1).ValueItems(99).Value.vt=   8
      Columns(1).ValueItems(99).DisplayValue=   "レ"
      Columns(1).ValueItems(99).DisplayValue.vt=   8
      Columns(1).ValueItems(99)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(100)._DefaultItem=   0
      Columns(1).ValueItems(100).Value=   "119"
      Columns(1).ValueItems(100).Value.vt=   8
      Columns(1).ValueItems(100).DisplayValue=   "レ"
      Columns(1).ValueItems(100).DisplayValue.vt=   8
      Columns(1).ValueItems(100)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(101)._DefaultItem=   0
      Columns(1).ValueItems(101).Value=   "129"
      Columns(1).ValueItems(101).Value.vt=   8
      Columns(1).ValueItems(101).DisplayValue=   "レ"
      Columns(1).ValueItems(101).DisplayValue.vt=   8
      Columns(1).ValueItems(101)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(102)._DefaultItem=   0
      Columns(1).ValueItems(102).Value=   "139"
      Columns(1).ValueItems(102).Value.vt=   8
      Columns(1).ValueItems(102).DisplayValue=   "レ"
      Columns(1).ValueItems(102).DisplayValue.vt=   8
      Columns(1).ValueItems(102)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(103)._DefaultItem=   0
      Columns(1).ValueItems(103).Value=   "149"
      Columns(1).ValueItems(103).Value.vt=   8
      Columns(1).ValueItems(103).DisplayValue=   "レ"
      Columns(1).ValueItems(103).DisplayValue.vt=   8
      Columns(1).ValueItems(103)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(104)._DefaultItem=   0
      Columns(1).ValueItems(104).Value=   "159"
      Columns(1).ValueItems(104).Value.vt=   8
      Columns(1).ValueItems(104).DisplayValue=   "レ"
      Columns(1).ValueItems(104).DisplayValue.vt=   8
      Columns(1).ValueItems(104)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(105)._DefaultItem=   0
      Columns(1).ValueItems(105).Value=   "169"
      Columns(1).ValueItems(105).Value.vt=   8
      Columns(1).ValueItems(105).DisplayValue=   "レ"
      Columns(1).ValueItems(105).DisplayValue.vt=   8
      Columns(1).ValueItems(105)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(106)._DefaultItem=   0
      Columns(1).ValueItems(106).Value=   "179"
      Columns(1).ValueItems(106).Value.vt=   8
      Columns(1).ValueItems(106).DisplayValue=   "レ"
      Columns(1).ValueItems(106).DisplayValue.vt=   8
      Columns(1).ValueItems(106)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(107)._DefaultItem=   0
      Columns(1).ValueItems(107).Value=   "189"
      Columns(1).ValueItems(107).Value.vt=   8
      Columns(1).ValueItems(107).DisplayValue=   "レ"
      Columns(1).ValueItems(107).DisplayValue.vt=   8
      Columns(1).ValueItems(107)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(108)._DefaultItem=   0
      Columns(1).ValueItems(108).Value=   "199"
      Columns(1).ValueItems(108).Value.vt=   8
      Columns(1).ValueItems(108).DisplayValue=   "レ"
      Columns(1).ValueItems(108).DisplayValue.vt=   8
      Columns(1).ValueItems(108)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(109)._DefaultItem=   0
      Columns(1).ValueItems(109).Value=   "209"
      Columns(1).ValueItems(109).Value.vt=   8
      Columns(1).ValueItems(109).DisplayValue=   "レ"
      Columns(1).ValueItems(109).DisplayValue.vt=   8
      Columns(1).ValueItems(109)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(110)._DefaultItem=   0
      Columns(1).ValueItems(110).Value=   "219"
      Columns(1).ValueItems(110).Value.vt=   8
      Columns(1).ValueItems(110).DisplayValue=   "レ"
      Columns(1).ValueItems(110).DisplayValue.vt=   8
      Columns(1).ValueItems(110)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(111)._DefaultItem=   0
      Columns(1).ValueItems(111).Value=   "229"
      Columns(1).ValueItems(111).Value.vt=   8
      Columns(1).ValueItems(111).DisplayValue=   "レ"
      Columns(1).ValueItems(111).DisplayValue.vt=   8
      Columns(1).ValueItems(111)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(112)._DefaultItem=   0
      Columns(1).ValueItems(112).Value=   "239"
      Columns(1).ValueItems(112).Value.vt=   8
      Columns(1).ValueItems(112).DisplayValue=   "レ"
      Columns(1).ValueItems(112).DisplayValue.vt=   8
      Columns(1).ValueItems(112)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(113)._DefaultItem=   0
      Columns(1).ValueItems(113).Value=   "249"
      Columns(1).ValueItems(113).Value.vt=   8
      Columns(1).ValueItems(113).DisplayValue=   "レ"
      Columns(1).ValueItems(113).DisplayValue.vt=   8
      Columns(1).ValueItems(113)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(114)._DefaultItem=   0
      Columns(1).ValueItems(114).Value=   "259"
      Columns(1).ValueItems(114).Value.vt=   8
      Columns(1).ValueItems(114).DisplayValue=   "レ"
      Columns(1).ValueItems(114).DisplayValue.vt=   8
      Columns(1).ValueItems(114)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(115)._DefaultItem=   0
      Columns(1).ValueItems(115).Value=   "269"
      Columns(1).ValueItems(115).Value.vt=   8
      Columns(1).ValueItems(115).DisplayValue=   "レ"
      Columns(1).ValueItems(115).DisplayValue.vt=   8
      Columns(1).ValueItems(115)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(116)._DefaultItem=   0
      Columns(1).ValueItems(116).Value=   "279"
      Columns(1).ValueItems(116).Value.vt=   8
      Columns(1).ValueItems(116).DisplayValue=   "レ"
      Columns(1).ValueItems(116).DisplayValue.vt=   8
      Columns(1).ValueItems(116)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(117)._DefaultItem=   0
      Columns(1).ValueItems(117).Value=   "289"
      Columns(1).ValueItems(117).Value.vt=   8
      Columns(1).ValueItems(117).DisplayValue=   "レ"
      Columns(1).ValueItems(117).DisplayValue.vt=   8
      Columns(1).ValueItems(117)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(118)._DefaultItem=   0
      Columns(1).ValueItems(118).Value=   "299"
      Columns(1).ValueItems(118).Value.vt=   8
      Columns(1).ValueItems(118).DisplayValue=   "レ"
      Columns(1).ValueItems(118).DisplayValue.vt=   8
      Columns(1).ValueItems(118)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(119)._DefaultItem=   0
      Columns(1).ValueItems(119).Value=   "309"
      Columns(1).ValueItems(119).Value.vt=   8
      Columns(1).ValueItems(119).DisplayValue=   "レ"
      Columns(1).ValueItems(119).DisplayValue.vt=   8
      Columns(1).ValueItems(119)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(120)._DefaultItem=   0
      Columns(1).ValueItems(120).Value=   "310"
      Columns(1).ValueItems(120).Value.vt=   8
      Columns(1).ValueItems(120).DisplayValue=   "　"
      Columns(1).ValueItems(120).DisplayValue.vt=   8
      Columns(1).ValueItems(120)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(121)._DefaultItem=   0
      Columns(1).ValueItems(121).Value=   "320"
      Columns(1).ValueItems(121).Value.vt=   8
      Columns(1).ValueItems(121).DisplayValue=   "　"
      Columns(1).ValueItems(121).DisplayValue.vt=   8
      Columns(1).ValueItems(121)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(122)._DefaultItem=   0
      Columns(1).ValueItems(122).Value=   "330"
      Columns(1).ValueItems(122).Value.vt=   8
      Columns(1).ValueItems(122).DisplayValue=   "　"
      Columns(1).ValueItems(122).DisplayValue.vt=   8
      Columns(1).ValueItems(122)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(123)._DefaultItem=   0
      Columns(1).ValueItems(123).Value=   "340"
      Columns(1).ValueItems(123).Value.vt=   8
      Columns(1).ValueItems(123).DisplayValue=   "　"
      Columns(1).ValueItems(123).DisplayValue.vt=   8
      Columns(1).ValueItems(123)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(124)._DefaultItem=   0
      Columns(1).ValueItems(124).Value=   "350"
      Columns(1).ValueItems(124).Value.vt=   8
      Columns(1).ValueItems(124).DisplayValue=   "　"
      Columns(1).ValueItems(124).DisplayValue.vt=   8
      Columns(1).ValueItems(124)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(125)._DefaultItem=   0
      Columns(1).ValueItems(125).Value=   "360"
      Columns(1).ValueItems(125).Value.vt=   8
      Columns(1).ValueItems(125).DisplayValue=   "　"
      Columns(1).ValueItems(125).DisplayValue.vt=   8
      Columns(1).ValueItems(125)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(126)._DefaultItem=   0
      Columns(1).ValueItems(126).Value=   "370"
      Columns(1).ValueItems(126).Value.vt=   8
      Columns(1).ValueItems(126).DisplayValue=   "　"
      Columns(1).ValueItems(126).DisplayValue.vt=   8
      Columns(1).ValueItems(126)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(127)._DefaultItem=   0
      Columns(1).ValueItems(127).Value=   "380"
      Columns(1).ValueItems(127).Value.vt=   8
      Columns(1).ValueItems(127).DisplayValue=   "　"
      Columns(1).ValueItems(127).DisplayValue.vt=   8
      Columns(1).ValueItems(127)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(128)._DefaultItem=   0
      Columns(1).ValueItems(128).Value=   "390"
      Columns(1).ValueItems(128).Value.vt=   8
      Columns(1).ValueItems(128).DisplayValue=   "　"
      Columns(1).ValueItems(128).DisplayValue.vt=   8
      Columns(1).ValueItems(128)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(129)._DefaultItem=   0
      Columns(1).ValueItems(129).Value=   "400"
      Columns(1).ValueItems(129).Value.vt=   8
      Columns(1).ValueItems(129).DisplayValue=   "　"
      Columns(1).ValueItems(129).DisplayValue.vt=   8
      Columns(1).ValueItems(129)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(130)._DefaultItem=   0
      Columns(1).ValueItems(130).Value=   "410"
      Columns(1).ValueItems(130).Value.vt=   8
      Columns(1).ValueItems(130).DisplayValue=   "　"
      Columns(1).ValueItems(130).DisplayValue.vt=   8
      Columns(1).ValueItems(130)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(131)._DefaultItem=   0
      Columns(1).ValueItems(131).Value=   "420"
      Columns(1).ValueItems(131).Value.vt=   8
      Columns(1).ValueItems(131).DisplayValue=   "　"
      Columns(1).ValueItems(131).DisplayValue.vt=   8
      Columns(1).ValueItems(131)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(132)._DefaultItem=   0
      Columns(1).ValueItems(132).Value=   "430"
      Columns(1).ValueItems(132).Value.vt=   8
      Columns(1).ValueItems(132).DisplayValue=   "　"
      Columns(1).ValueItems(132).DisplayValue.vt=   8
      Columns(1).ValueItems(132)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(133)._DefaultItem=   0
      Columns(1).ValueItems(133).Value=   "440"
      Columns(1).ValueItems(133).Value.vt=   8
      Columns(1).ValueItems(133).DisplayValue=   "　"
      Columns(1).ValueItems(133).DisplayValue.vt=   8
      Columns(1).ValueItems(133)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(134)._DefaultItem=   0
      Columns(1).ValueItems(134).Value=   "450"
      Columns(1).ValueItems(134).Value.vt=   8
      Columns(1).ValueItems(134).DisplayValue=   "　"
      Columns(1).ValueItems(134).DisplayValue.vt=   8
      Columns(1).ValueItems(134)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(135)._DefaultItem=   0
      Columns(1).ValueItems(135).Value=   "460"
      Columns(1).ValueItems(135).Value.vt=   8
      Columns(1).ValueItems(135).DisplayValue=   "　"
      Columns(1).ValueItems(135).DisplayValue.vt=   8
      Columns(1).ValueItems(135)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(136)._DefaultItem=   0
      Columns(1).ValueItems(136).Value=   "470"
      Columns(1).ValueItems(136).Value.vt=   8
      Columns(1).ValueItems(136).DisplayValue=   "　"
      Columns(1).ValueItems(136).DisplayValue.vt=   8
      Columns(1).ValueItems(136)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(137)._DefaultItem=   0
      Columns(1).ValueItems(137).Value=   "480"
      Columns(1).ValueItems(137).Value.vt=   8
      Columns(1).ValueItems(137).DisplayValue=   "　"
      Columns(1).ValueItems(137).DisplayValue.vt=   8
      Columns(1).ValueItems(137)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(138)._DefaultItem=   0
      Columns(1).ValueItems(138).Value=   "490"
      Columns(1).ValueItems(138).Value.vt=   8
      Columns(1).ValueItems(138).DisplayValue=   "　"
      Columns(1).ValueItems(138).DisplayValue.vt=   8
      Columns(1).ValueItems(138)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(139)._DefaultItem=   0
      Columns(1).ValueItems(139).Value=   "500"
      Columns(1).ValueItems(139).Value.vt=   8
      Columns(1).ValueItems(139).DisplayValue=   "　"
      Columns(1).ValueItems(139).DisplayValue.vt=   8
      Columns(1).ValueItems(139)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(140)._DefaultItem=   0
      Columns(1).ValueItems(140).Value=   "311"
      Columns(1).ValueItems(140).Value.vt=   8
      Columns(1).ValueItems(140).DisplayValue=   "確"
      Columns(1).ValueItems(140).DisplayValue.vt=   8
      Columns(1).ValueItems(140)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(141)._DefaultItem=   0
      Columns(1).ValueItems(141).Value=   "321"
      Columns(1).ValueItems(141).Value.vt=   8
      Columns(1).ValueItems(141).DisplayValue=   "確"
      Columns(1).ValueItems(141).DisplayValue.vt=   8
      Columns(1).ValueItems(141)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(142)._DefaultItem=   0
      Columns(1).ValueItems(142).Value=   "331"
      Columns(1).ValueItems(142).Value.vt=   8
      Columns(1).ValueItems(142).DisplayValue=   "確"
      Columns(1).ValueItems(142).DisplayValue.vt=   8
      Columns(1).ValueItems(142)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(143)._DefaultItem=   0
      Columns(1).ValueItems(143).Value=   "341"
      Columns(1).ValueItems(143).Value.vt=   8
      Columns(1).ValueItems(143).DisplayValue=   "確"
      Columns(1).ValueItems(143).DisplayValue.vt=   8
      Columns(1).ValueItems(143)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(144)._DefaultItem=   0
      Columns(1).ValueItems(144).Value=   "351"
      Columns(1).ValueItems(144).Value.vt=   8
      Columns(1).ValueItems(144).DisplayValue=   "確"
      Columns(1).ValueItems(144).DisplayValue.vt=   8
      Columns(1).ValueItems(144)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(145)._DefaultItem=   0
      Columns(1).ValueItems(145).Value=   "361"
      Columns(1).ValueItems(145).Value.vt=   8
      Columns(1).ValueItems(145).DisplayValue=   "確"
      Columns(1).ValueItems(145).DisplayValue.vt=   8
      Columns(1).ValueItems(145)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(146)._DefaultItem=   0
      Columns(1).ValueItems(146).Value=   "371"
      Columns(1).ValueItems(146).Value.vt=   8
      Columns(1).ValueItems(146).DisplayValue=   "確"
      Columns(1).ValueItems(146).DisplayValue.vt=   8
      Columns(1).ValueItems(146)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(147)._DefaultItem=   0
      Columns(1).ValueItems(147).Value=   "381"
      Columns(1).ValueItems(147).Value.vt=   8
      Columns(1).ValueItems(147).DisplayValue=   "確"
      Columns(1).ValueItems(147).DisplayValue.vt=   8
      Columns(1).ValueItems(147)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(148)._DefaultItem=   0
      Columns(1).ValueItems(148).Value=   "391"
      Columns(1).ValueItems(148).Value.vt=   8
      Columns(1).ValueItems(148).DisplayValue=   "確"
      Columns(1).ValueItems(148).DisplayValue.vt=   8
      Columns(1).ValueItems(148)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(149)._DefaultItem=   0
      Columns(1).ValueItems(149).Value=   "401"
      Columns(1).ValueItems(149).Value.vt=   8
      Columns(1).ValueItems(149).DisplayValue=   "確"
      Columns(1).ValueItems(149).DisplayValue.vt=   8
      Columns(1).ValueItems(149)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(150)._DefaultItem=   0
      Columns(1).ValueItems(150).Value=   "411"
      Columns(1).ValueItems(150).Value.vt=   8
      Columns(1).ValueItems(150).DisplayValue=   "確"
      Columns(1).ValueItems(150).DisplayValue.vt=   8
      Columns(1).ValueItems(150)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(151)._DefaultItem=   0
      Columns(1).ValueItems(151).Value=   "421"
      Columns(1).ValueItems(151).Value.vt=   8
      Columns(1).ValueItems(151).DisplayValue=   "確"
      Columns(1).ValueItems(151).DisplayValue.vt=   8
      Columns(1).ValueItems(151)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(152)._DefaultItem=   0
      Columns(1).ValueItems(152).Value=   "431"
      Columns(1).ValueItems(152).Value.vt=   8
      Columns(1).ValueItems(152).DisplayValue=   "確"
      Columns(1).ValueItems(152).DisplayValue.vt=   8
      Columns(1).ValueItems(152)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(153)._DefaultItem=   0
      Columns(1).ValueItems(153).Value=   "441"
      Columns(1).ValueItems(153).Value.vt=   8
      Columns(1).ValueItems(153).DisplayValue=   "確"
      Columns(1).ValueItems(153).DisplayValue.vt=   8
      Columns(1).ValueItems(153)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(154)._DefaultItem=   0
      Columns(1).ValueItems(154).Value=   "451"
      Columns(1).ValueItems(154).Value.vt=   8
      Columns(1).ValueItems(154).DisplayValue=   "確"
      Columns(1).ValueItems(154).DisplayValue.vt=   8
      Columns(1).ValueItems(154)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(155)._DefaultItem=   0
      Columns(1).ValueItems(155).Value=   "461"
      Columns(1).ValueItems(155).Value.vt=   8
      Columns(1).ValueItems(155).DisplayValue=   "確"
      Columns(1).ValueItems(155).DisplayValue.vt=   8
      Columns(1).ValueItems(155)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(156)._DefaultItem=   0
      Columns(1).ValueItems(156).Value=   "471"
      Columns(1).ValueItems(156).Value.vt=   8
      Columns(1).ValueItems(156).DisplayValue=   "確"
      Columns(1).ValueItems(156).DisplayValue.vt=   8
      Columns(1).ValueItems(156)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(157)._DefaultItem=   0
      Columns(1).ValueItems(157).Value=   "481"
      Columns(1).ValueItems(157).Value.vt=   8
      Columns(1).ValueItems(157).DisplayValue=   "確"
      Columns(1).ValueItems(157).DisplayValue.vt=   8
      Columns(1).ValueItems(157)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(158)._DefaultItem=   0
      Columns(1).ValueItems(158).Value=   "491"
      Columns(1).ValueItems(158).Value.vt=   8
      Columns(1).ValueItems(158).DisplayValue=   "確"
      Columns(1).ValueItems(158).DisplayValue.vt=   8
      Columns(1).ValueItems(158)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(159)._DefaultItem=   0
      Columns(1).ValueItems(159).Value=   "501"
      Columns(1).ValueItems(159).Value.vt=   8
      Columns(1).ValueItems(159).DisplayValue=   "確"
      Columns(1).ValueItems(159).DisplayValue.vt=   8
      Columns(1).ValueItems(159)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(160)._DefaultItem=   0
      Columns(1).ValueItems(160).Value=   "312"
      Columns(1).ValueItems(160).Value.vt=   8
      Columns(1).ValueItems(160).DisplayValue=   "再"
      Columns(1).ValueItems(160).DisplayValue.vt=   8
      Columns(1).ValueItems(160)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(161)._DefaultItem=   0
      Columns(1).ValueItems(161).Value=   "322"
      Columns(1).ValueItems(161).Value.vt=   8
      Columns(1).ValueItems(161).DisplayValue=   "再"
      Columns(1).ValueItems(161).DisplayValue.vt=   8
      Columns(1).ValueItems(161)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(162)._DefaultItem=   0
      Columns(1).ValueItems(162).Value=   "332"
      Columns(1).ValueItems(162).Value.vt=   8
      Columns(1).ValueItems(162).DisplayValue=   "再"
      Columns(1).ValueItems(162).DisplayValue.vt=   8
      Columns(1).ValueItems(162)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(163)._DefaultItem=   0
      Columns(1).ValueItems(163).Value=   "342"
      Columns(1).ValueItems(163).Value.vt=   8
      Columns(1).ValueItems(163).DisplayValue=   "再"
      Columns(1).ValueItems(163).DisplayValue.vt=   8
      Columns(1).ValueItems(163)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(164)._DefaultItem=   0
      Columns(1).ValueItems(164).Value=   "352"
      Columns(1).ValueItems(164).Value.vt=   8
      Columns(1).ValueItems(164).DisplayValue=   "再"
      Columns(1).ValueItems(164).DisplayValue.vt=   8
      Columns(1).ValueItems(164)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(165)._DefaultItem=   0
      Columns(1).ValueItems(165).Value=   "362"
      Columns(1).ValueItems(165).Value.vt=   8
      Columns(1).ValueItems(165).DisplayValue=   "再"
      Columns(1).ValueItems(165).DisplayValue.vt=   8
      Columns(1).ValueItems(165)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(166)._DefaultItem=   0
      Columns(1).ValueItems(166).Value=   "372"
      Columns(1).ValueItems(166).Value.vt=   8
      Columns(1).ValueItems(166).DisplayValue=   "再"
      Columns(1).ValueItems(166).DisplayValue.vt=   8
      Columns(1).ValueItems(166)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(167)._DefaultItem=   0
      Columns(1).ValueItems(167).Value=   "382"
      Columns(1).ValueItems(167).Value.vt=   8
      Columns(1).ValueItems(167).DisplayValue=   "再"
      Columns(1).ValueItems(167).DisplayValue.vt=   8
      Columns(1).ValueItems(167)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(168)._DefaultItem=   0
      Columns(1).ValueItems(168).Value=   "392"
      Columns(1).ValueItems(168).Value.vt=   8
      Columns(1).ValueItems(168).DisplayValue=   "再"
      Columns(1).ValueItems(168).DisplayValue.vt=   8
      Columns(1).ValueItems(168)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(169)._DefaultItem=   0
      Columns(1).ValueItems(169).Value=   "402"
      Columns(1).ValueItems(169).Value.vt=   8
      Columns(1).ValueItems(169).DisplayValue=   "再"
      Columns(1).ValueItems(169).DisplayValue.vt=   8
      Columns(1).ValueItems(169)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(170)._DefaultItem=   0
      Columns(1).ValueItems(170).Value=   "412"
      Columns(1).ValueItems(170).Value.vt=   8
      Columns(1).ValueItems(170).DisplayValue=   "再"
      Columns(1).ValueItems(170).DisplayValue.vt=   8
      Columns(1).ValueItems(170)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(171)._DefaultItem=   0
      Columns(1).ValueItems(171).Value=   "422"
      Columns(1).ValueItems(171).Value.vt=   8
      Columns(1).ValueItems(171).DisplayValue=   "再"
      Columns(1).ValueItems(171).DisplayValue.vt=   8
      Columns(1).ValueItems(171)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(172)._DefaultItem=   0
      Columns(1).ValueItems(172).Value=   "432"
      Columns(1).ValueItems(172).Value.vt=   8
      Columns(1).ValueItems(172).DisplayValue=   "再"
      Columns(1).ValueItems(172).DisplayValue.vt=   8
      Columns(1).ValueItems(172)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(173)._DefaultItem=   0
      Columns(1).ValueItems(173).Value=   "442"
      Columns(1).ValueItems(173).Value.vt=   8
      Columns(1).ValueItems(173).DisplayValue=   "再"
      Columns(1).ValueItems(173).DisplayValue.vt=   8
      Columns(1).ValueItems(173)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(174)._DefaultItem=   0
      Columns(1).ValueItems(174).Value=   "452"
      Columns(1).ValueItems(174).Value.vt=   8
      Columns(1).ValueItems(174).DisplayValue=   "再"
      Columns(1).ValueItems(174).DisplayValue.vt=   8
      Columns(1).ValueItems(174)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(175)._DefaultItem=   0
      Columns(1).ValueItems(175).Value=   "462"
      Columns(1).ValueItems(175).Value.vt=   8
      Columns(1).ValueItems(175).DisplayValue=   "再"
      Columns(1).ValueItems(175).DisplayValue.vt=   8
      Columns(1).ValueItems(175)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(176)._DefaultItem=   0
      Columns(1).ValueItems(176).Value=   "472"
      Columns(1).ValueItems(176).Value.vt=   8
      Columns(1).ValueItems(176).DisplayValue=   "再"
      Columns(1).ValueItems(176).DisplayValue.vt=   8
      Columns(1).ValueItems(176)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(177)._DefaultItem=   0
      Columns(1).ValueItems(177).Value=   "482"
      Columns(1).ValueItems(177).Value.vt=   8
      Columns(1).ValueItems(177).DisplayValue=   "再"
      Columns(1).ValueItems(177).DisplayValue.vt=   8
      Columns(1).ValueItems(177)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(178)._DefaultItem=   0
      Columns(1).ValueItems(178).Value=   "492"
      Columns(1).ValueItems(178).Value.vt=   8
      Columns(1).ValueItems(178).DisplayValue=   "再"
      Columns(1).ValueItems(178).DisplayValue.vt=   8
      Columns(1).ValueItems(178)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(179)._DefaultItem=   0
      Columns(1).ValueItems(179).Value=   "502"
      Columns(1).ValueItems(179).Value.vt=   8
      Columns(1).ValueItems(179).DisplayValue=   "再"
      Columns(1).ValueItems(179).DisplayValue.vt=   8
      Columns(1).ValueItems(179)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(180)._DefaultItem=   0
      Columns(1).ValueItems(180).Value=   "309"
      Columns(1).ValueItems(180).Value.vt=   8
      Columns(1).ValueItems(180).DisplayValue=   "レ"
      Columns(1).ValueItems(180).DisplayValue.vt=   8
      Columns(1).ValueItems(180)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(181)._DefaultItem=   0
      Columns(1).ValueItems(181).Value=   "319"
      Columns(1).ValueItems(181).Value.vt=   8
      Columns(1).ValueItems(181).DisplayValue=   "レ"
      Columns(1).ValueItems(181).DisplayValue.vt=   8
      Columns(1).ValueItems(181)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(182)._DefaultItem=   0
      Columns(1).ValueItems(182).Value=   "329"
      Columns(1).ValueItems(182).Value.vt=   8
      Columns(1).ValueItems(182).DisplayValue=   "レ"
      Columns(1).ValueItems(182).DisplayValue.vt=   8
      Columns(1).ValueItems(182)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(183)._DefaultItem=   0
      Columns(1).ValueItems(183).Value=   "339"
      Columns(1).ValueItems(183).Value.vt=   8
      Columns(1).ValueItems(183).DisplayValue=   "レ"
      Columns(1).ValueItems(183).DisplayValue.vt=   8
      Columns(1).ValueItems(183)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(184)._DefaultItem=   0
      Columns(1).ValueItems(184).Value=   "349"
      Columns(1).ValueItems(184).Value.vt=   8
      Columns(1).ValueItems(184).DisplayValue=   "レ"
      Columns(1).ValueItems(184).DisplayValue.vt=   8
      Columns(1).ValueItems(184)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(185)._DefaultItem=   0
      Columns(1).ValueItems(185).Value=   "359"
      Columns(1).ValueItems(185).Value.vt=   8
      Columns(1).ValueItems(185).DisplayValue=   "レ"
      Columns(1).ValueItems(185).DisplayValue.vt=   8
      Columns(1).ValueItems(185)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(186)._DefaultItem=   0
      Columns(1).ValueItems(186).Value=   "369"
      Columns(1).ValueItems(186).Value.vt=   8
      Columns(1).ValueItems(186).DisplayValue=   "レ"
      Columns(1).ValueItems(186).DisplayValue.vt=   8
      Columns(1).ValueItems(186)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(187)._DefaultItem=   0
      Columns(1).ValueItems(187).Value=   "379"
      Columns(1).ValueItems(187).Value.vt=   8
      Columns(1).ValueItems(187).DisplayValue=   "レ"
      Columns(1).ValueItems(187).DisplayValue.vt=   8
      Columns(1).ValueItems(187)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(188)._DefaultItem=   0
      Columns(1).ValueItems(188).Value=   "389"
      Columns(1).ValueItems(188).Value.vt=   8
      Columns(1).ValueItems(188).DisplayValue=   "レ"
      Columns(1).ValueItems(188).DisplayValue.vt=   8
      Columns(1).ValueItems(188)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(189)._DefaultItem=   0
      Columns(1).ValueItems(189).Value=   "399"
      Columns(1).ValueItems(189).Value.vt=   8
      Columns(1).ValueItems(189).DisplayValue=   "レ"
      Columns(1).ValueItems(189).DisplayValue.vt=   8
      Columns(1).ValueItems(189)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(190)._DefaultItem=   0
      Columns(1).ValueItems(190).Value=   "409"
      Columns(1).ValueItems(190).Value.vt=   8
      Columns(1).ValueItems(190).DisplayValue=   "レ"
      Columns(1).ValueItems(190).DisplayValue.vt=   8
      Columns(1).ValueItems(190)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(191)._DefaultItem=   0
      Columns(1).ValueItems(191).Value=   "419"
      Columns(1).ValueItems(191).Value.vt=   8
      Columns(1).ValueItems(191).DisplayValue=   "レ"
      Columns(1).ValueItems(191).DisplayValue.vt=   8
      Columns(1).ValueItems(191)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(192)._DefaultItem=   0
      Columns(1).ValueItems(192).Value=   "429"
      Columns(1).ValueItems(192).Value.vt=   8
      Columns(1).ValueItems(192).DisplayValue=   "レ"
      Columns(1).ValueItems(192).DisplayValue.vt=   8
      Columns(1).ValueItems(192)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(193)._DefaultItem=   0
      Columns(1).ValueItems(193).Value=   "439"
      Columns(1).ValueItems(193).Value.vt=   8
      Columns(1).ValueItems(193).DisplayValue=   "レ"
      Columns(1).ValueItems(193).DisplayValue.vt=   8
      Columns(1).ValueItems(193)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(194)._DefaultItem=   0
      Columns(1).ValueItems(194).Value=   "449"
      Columns(1).ValueItems(194).Value.vt=   8
      Columns(1).ValueItems(194).DisplayValue=   "レ"
      Columns(1).ValueItems(194).DisplayValue.vt=   8
      Columns(1).ValueItems(194)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(195)._DefaultItem=   0
      Columns(1).ValueItems(195).Value=   "459"
      Columns(1).ValueItems(195).Value.vt=   8
      Columns(1).ValueItems(195).DisplayValue=   "レ"
      Columns(1).ValueItems(195).DisplayValue.vt=   8
      Columns(1).ValueItems(195)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(196)._DefaultItem=   0
      Columns(1).ValueItems(196).Value=   "469"
      Columns(1).ValueItems(196).Value.vt=   8
      Columns(1).ValueItems(196).DisplayValue=   "レ"
      Columns(1).ValueItems(196).DisplayValue.vt=   8
      Columns(1).ValueItems(196)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(197)._DefaultItem=   0
      Columns(1).ValueItems(197).Value=   "479"
      Columns(1).ValueItems(197).Value.vt=   8
      Columns(1).ValueItems(197).DisplayValue=   "レ"
      Columns(1).ValueItems(197).DisplayValue.vt=   8
      Columns(1).ValueItems(197)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(198)._DefaultItem=   0
      Columns(1).ValueItems(198).Value=   "489"
      Columns(1).ValueItems(198).Value.vt=   8
      Columns(1).ValueItems(198).DisplayValue=   "レ"
      Columns(1).ValueItems(198).DisplayValue.vt=   8
      Columns(1).ValueItems(198)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(199)._DefaultItem=   0
      Columns(1).ValueItems(199).Value=   "499"
      Columns(1).ValueItems(199).Value.vt=   8
      Columns(1).ValueItems(199).DisplayValue=   "レ"
      Columns(1).ValueItems(199).DisplayValue.vt=   8
      Columns(1).ValueItems(199)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(200)._DefaultItem=   0
      Columns(1).ValueItems(200).Value=   "509"
      Columns(1).ValueItems(200).Value.vt=   8
      Columns(1).ValueItems(200).DisplayValue=   "レ"
      Columns(1).ValueItems(200).DisplayValue.vt=   8
      Columns(1).ValueItems(200)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems.Count=   201
      Columns(1).Caption=   "確定"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ｸｯｼｮﾝ"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "S№/F№"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "取順"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "銘柄"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "米坪／包装"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "製品寸法"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "听量"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "巻連数"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "本数"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "包装形態"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "品質"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "下札"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "行先反転"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "ﾕｰｻﾞ指定"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "補助ﾗﾍﾞﾙ"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "欠連引当"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "輸出C№"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "輸出ｺｰﾄﾞ"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   512
      Columns(20)._MaxComboItems=   2
      Columns(20).ValueItems(0)._DefaultItem=   0
      Columns(20).ValueItems(0).Value=   "0"
      Columns(20).ValueItems(0).Value.vt=   8
      Columns(20).ValueItems(0).DisplayValue=   " "
      Columns(20).ValueItems(0).DisplayValue.vt=   8
      Columns(20).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(20).ValueItems(1)._DefaultItem=   0
      Columns(20).ValueItems(1).Value=   "1"
      Columns(20).ValueItems(1).Value.vt=   8
      Columns(20).ValueItems(1).DisplayValue=   "継手×"
      Columns(20).ValueItems(1).DisplayValue.vt=   8
      Columns(20).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(20).ValueItems.Count=   2
      Columns(20).Caption=   "引当条件"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).ValueItems(0)._DefaultItem=   0
      Columns(21).ValueItems(0).Value=   ""
      Columns(21).ValueItems(0).Value.vt=   8
      Columns(21).ValueItems(0).DisplayValue=   "0"
      Columns(21).ValueItems(0).DisplayValue.vt=   8
      Columns(21).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(21).ValueItems.Count=   1
      Columns(21).Caption=   "備考"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "(非表示)注文総連数"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "(非表示)注文連数"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "(非表示)S№"
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   25
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).ShowCollapseExpandIcons=   0   'False
      Splits(0).Locked=   -1  'True
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   699
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=25"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=783"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=656"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
      Splits(0)._ColumnProps(5)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Merge=2"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=783"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=656"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=17"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(1).Merge=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1058"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=931"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=17"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2117"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1990"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=17"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=699"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=572"
      Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=17"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=5334"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=5207"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=16"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=1397"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1270"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=17"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=2942"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2815"
      Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=17"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=1588"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=1461"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=18"
      Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(49)=   "Column(9).Width=1588"
      Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=1461"
      Splits(0)._ColumnProps(52)=   "Column(9)._ColStyle=18"
      Splits(0)._ColumnProps(53)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(54)=   "Column(10).Width=1588"
      Splits(0)._ColumnProps(55)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(10)._WidthInPix=1461"
      Splits(0)._ColumnProps(57)=   "Column(10)._ColStyle=18"
      Splits(0)._ColumnProps(58)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(59)=   "Column(11).Width=4043"
      Splits(0)._ColumnProps(60)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(11)._WidthInPix=3916"
      Splits(0)._ColumnProps(62)=   "Column(11)._ColStyle=16"
      Splits(0)._ColumnProps(63)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(64)=   "Column(12).Width=1080"
      Splits(0)._ColumnProps(65)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(12)._WidthInPix=953"
      Splits(0)._ColumnProps(67)=   "Column(12)._ColStyle=17"
      Splits(0)._ColumnProps(68)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(69)=   "Column(13).Width=783"
      Splits(0)._ColumnProps(70)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(13)._WidthInPix=656"
      Splits(0)._ColumnProps(72)=   "Column(13)._ColStyle=532"
      Splits(0)._ColumnProps(73)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(74)=   "Column(14).Width=868"
      Splits(0)._ColumnProps(75)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(14)._WidthInPix=741"
      Splits(0)._ColumnProps(77)=   "Column(14)._ColStyle=532"
      Splits(0)._ColumnProps(78)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(79)=   "Column(15).Width=1228"
      Splits(0)._ColumnProps(80)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(15)._WidthInPix=1101"
      Splits(0)._ColumnProps(82)=   "Column(15)._ColStyle=529"
      Splits(0)._ColumnProps(83)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(84)=   "Column(16).Width=3514"
      Splits(0)._ColumnProps(85)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(16)._WidthInPix=3387"
      Splits(0)._ColumnProps(87)=   "Column(16)._ColStyle=528"
      Splits(0)._ColumnProps(88)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(89)=   "Column(17).Width=931"
      Splits(0)._ColumnProps(90)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(17)._WidthInPix=804"
      Splits(0)._ColumnProps(92)=   "Column(17)._ColStyle=17"
      Splits(0)._ColumnProps(93)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(94)=   "Column(18).Width=3514"
      Splits(0)._ColumnProps(95)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(96)=   "Column(18)._WidthInPix=3387"
      Splits(0)._ColumnProps(97)=   "Column(18)._ColStyle=528"
      Splits(0)._ColumnProps(98)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(99)=   "Column(19).Width=868"
      Splits(0)._ColumnProps(100)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(101)=   "Column(19)._WidthInPix=741"
      Splits(0)._ColumnProps(102)=   "Column(19)._ColStyle=16"
      Splits(0)._ColumnProps(103)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(104)=   "Column(20).Width=1503"
      Splits(0)._ColumnProps(105)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(106)=   "Column(20)._WidthInPix=1376"
      Splits(0)._ColumnProps(107)=   "Column(20)._ColStyle=131601"
      Splits(0)._ColumnProps(108)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(109)=   "Column(21).Width=9229"
      Splits(0)._ColumnProps(110)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(21)._WidthInPix=9102"
      Splits(0)._ColumnProps(112)=   "Column(21)._ColStyle=528"
      Splits(0)._ColumnProps(113)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(114)=   "Column(22).Width=127"
      Splits(0)._ColumnProps(115)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(116)=   "Column(22)._ColStyle=20"
      Splits(0)._ColumnProps(117)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(118)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(119)=   "Column(23).Width=127"
      Splits(0)._ColumnProps(120)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(23)._ColStyle=20"
      Splits(0)._ColumnProps(122)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(123)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(124)=   "Column(24).Width=4890"
      Splits(0)._ColumnProps(125)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(126)=   "Column(24)._WidthInPix=4763"
      Splits(0)._ColumnProps(127)=   "Column(24)._ColStyle=20"
      Splits(0)._ColumnProps(128)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(129)=   "Column(24).Order=25"
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
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ Ｐゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=1125,.italic=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.namedParent=34,.alignment=2"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.namedParent=36,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.bold=-1,.fontsize=1125"
      _StyleDefs(42)  =   ":id=29,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(43)  =   ":id=29,.fontname=ＭＳ ゴシック"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.bold=-1,.fontsize=1125"
      _StyleDefs(48)  =   ":id=43,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(49)  =   ":id=43,.fontname=ＭＳ ゴシック"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2,.bold=-1,.fontsize=1125"
      _StyleDefs(53)  =   ":id=50,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(54)  =   ":id=50,.fontname=ＭＳ ゴシック"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
      _StyleDefs(63)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=2,.bold=-1,.fontsize=1125"
      _StyleDefs(67)  =   ":id=62,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(68)  =   ":id=62,.fontname=ＭＳ ゴシック"
      _StyleDefs(69)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14,.bold=-1,.fontsize=900"
      _StyleDefs(70)  =   ":id=59,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(71)  =   ":id=59,.fontname=ＭＳ ゴシック"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=106,.parent=13,.alignment=2"
      _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=103,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=104,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=105,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.alignment=0"
      _StyleDefs(91)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(12).Style:id=122,.parent=13,.alignment=2"
      _StyleDefs(95)  =   "Splits(0).Columns(12).HeadingStyle:id=119,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(12).FooterStyle:id=120,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(12).EditorStyle:id=121,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(13).Style:id=130,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(13).HeadingStyle:id=127,.parent=14,.alignment=2"
      _StyleDefs(100) =   "Splits(0).Columns(13).FooterStyle:id=128,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(13).EditorStyle:id=129,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(14).Style:id=126,.parent=13,.alignment=3"
      _StyleDefs(103) =   "Splits(0).Columns(14).HeadingStyle:id=123,.parent=14,.alignment=2,.bold=-1"
      _StyleDefs(104) =   ":id=123,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(105) =   ":id=123,.fontname=ＭＳ ゴシック"
      _StyleDefs(106) =   "Splits(0).Columns(14).FooterStyle:id=124,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(14).EditorStyle:id=125,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(15).Style:id=82,.parent=13,.alignment=2"
      _StyleDefs(109) =   "Splits(0).Columns(15).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(110) =   "Splits(0).Columns(15).FooterStyle:id=80,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(15).EditorStyle:id=81,.parent=17"
      _StyleDefs(112) =   "Splits(0).Columns(16).Style:id=86,.parent=13,.alignment=0"
      _StyleDefs(113) =   "Splits(0).Columns(16).HeadingStyle:id=83,.parent=14,.alignment=2"
      _StyleDefs(114) =   "Splits(0).Columns(16).FooterStyle:id=84,.parent=15"
      _StyleDefs(115) =   "Splits(0).Columns(16).EditorStyle:id=85,.parent=17"
      _StyleDefs(116) =   "Splits(0).Columns(17).Style:id=90,.parent=13,.alignment=2"
      _StyleDefs(117) =   "Splits(0).Columns(17).HeadingStyle:id=87,.parent=14"
      _StyleDefs(118) =   "Splits(0).Columns(17).FooterStyle:id=88,.parent=15"
      _StyleDefs(119) =   "Splits(0).Columns(17).EditorStyle:id=89,.parent=17"
      _StyleDefs(120) =   "Splits(0).Columns(18).Style:id=94,.parent=13,.alignment=0"
      _StyleDefs(121) =   "Splits(0).Columns(18).HeadingStyle:id=91,.parent=14,.alignment=2"
      _StyleDefs(122) =   "Splits(0).Columns(18).FooterStyle:id=92,.parent=15"
      _StyleDefs(123) =   "Splits(0).Columns(18).EditorStyle:id=93,.parent=17"
      _StyleDefs(124) =   "Splits(0).Columns(19).Style:id=98,.parent=13,.alignment=0"
      _StyleDefs(125) =   "Splits(0).Columns(19).HeadingStyle:id=95,.parent=14"
      _StyleDefs(126) =   "Splits(0).Columns(19).FooterStyle:id=96,.parent=15"
      _StyleDefs(127) =   "Splits(0).Columns(19).EditorStyle:id=97,.parent=17"
      _StyleDefs(128) =   "Splits(0).Columns(20).Style:id=102,.parent=13,.alignment=2,.bold=-1"
      _StyleDefs(129) =   ":id=102,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(130) =   ":id=102,.fontname=ＭＳ ゴシック"
      _StyleDefs(131) =   "Splits(0).Columns(20).HeadingStyle:id=99,.parent=14,.alignment=2"
      _StyleDefs(132) =   "Splits(0).Columns(20).FooterStyle:id=100,.parent=15,.alignment=2"
      _StyleDefs(133) =   "Splits(0).Columns(20).EditorStyle:id=101,.parent=17"
      _StyleDefs(134) =   "Splits(0).Columns(21).Style:id=110,.parent=13,.alignment=0"
      _StyleDefs(135) =   "Splits(0).Columns(21).HeadingStyle:id=107,.parent=14,.alignment=2,.bold=-1"
      _StyleDefs(136) =   ":id=107,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(137) =   ":id=107,.fontname=ＭＳ ゴシック"
      _StyleDefs(138) =   "Splits(0).Columns(21).FooterStyle:id=108,.parent=15"
      _StyleDefs(139) =   "Splits(0).Columns(21).EditorStyle:id=109,.parent=17"
      _StyleDefs(140) =   "Splits(0).Columns(22).Style:id=114,.parent=13"
      _StyleDefs(141) =   "Splits(0).Columns(22).HeadingStyle:id=111,.parent=14"
      _StyleDefs(142) =   "Splits(0).Columns(22).FooterStyle:id=112,.parent=15"
      _StyleDefs(143) =   "Splits(0).Columns(22).EditorStyle:id=113,.parent=17"
      _StyleDefs(144) =   "Splits(0).Columns(23).Style:id=118,.parent=13"
      _StyleDefs(145) =   "Splits(0).Columns(23).HeadingStyle:id=115,.parent=14"
      _StyleDefs(146) =   "Splits(0).Columns(23).FooterStyle:id=116,.parent=15"
      _StyleDefs(147) =   "Splits(0).Columns(23).EditorStyle:id=117,.parent=17"
      _StyleDefs(148) =   "Splits(0).Columns(24).Style:id=134,.parent=13"
      _StyleDefs(149) =   "Splits(0).Columns(24).HeadingStyle:id=131,.parent=14"
      _StyleDefs(150) =   "Splits(0).Columns(24).FooterStyle:id=132,.parent=15"
      _StyleDefs(151) =   "Splits(0).Columns(24).EditorStyle:id=133,.parent=17"
      _StyleDefs(152) =   "Named:id=33:Normal"
      _StyleDefs(153) =   ":id=33,.parent=0,.fgcolor=&H0&"
      _StyleDefs(154) =   "Named:id=34:Heading"
      _StyleDefs(155) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(156) =   ":id=34,.wraptext=-1"
      _StyleDefs(157) =   "Named:id=35:Footing"
      _StyleDefs(158) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(159) =   "Named:id=36:Selected"
      _StyleDefs(160) =   ":id=36,.parent=33,.bgcolor=&H80000005&,.fgcolor=&H80000012&"
      _StyleDefs(161) =   "Named:id=37:Caption"
      _StyleDefs(162) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(163) =   "Named:id=38:HighlightRow"
      _StyleDefs(164) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(165) =   "Named:id=39:EvenRow"
      _StyleDefs(166) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(167) =   "Named:id=40:OddRow"
      _StyleDefs(168) =   ":id=40,.parent=33"
      _StyleDefs(169) =   "Named:id=41:RecordSelector"
      _StyleDefs(170) =   ":id=41,.parent=34"
      _StyleDefs(171) =   "Named:id=42:FilterBar"
      _StyleDefs(172) =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "メニュー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   13820
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   12630
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   11445
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   " 指示明細 　 印刷"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   8780
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "詳細表示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   7590
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6405
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修　正"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3740
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ｓ№指定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2550
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "日付・Ｇ№"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1365
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確　定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   180
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   5220
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   10260
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   8085
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "123.4"
      Top             =   540
      Width           =   930
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "μコートネオス"
      Top             =   525
      Width           =   3675
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "901279"
      Top             =   510
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "抄造号機"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   14280
      TabIndex        =   38
      Top             =   405
      Width           =   465
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F4)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   11
      Left            =   4080
      TabIndex        =   31
      Top             =   9945
      Width           =   540
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F3)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   10
      Left            =   2910
      TabIndex        =   30
      Top             =   9945
      Width           =   540
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F2)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   9
      Left            =   1725
      TabIndex        =   29
      Top             =   9960
      Width           =   540
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F8)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   8
      Left            =   9105
      TabIndex        =   28
      Top             =   9930
      Width           =   540
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F7)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   7
      Left            =   7935
      TabIndex        =   27
      Top             =   9930
      Width           =   540
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F6)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   6735
      TabIndex        =   26
      Top             =   9945
      Width           =   540
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F12)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   14040
      TabIndex        =   25
      Top             =   9945
      Width           =   675
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F11)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   12855
      TabIndex        =   24
      Top             =   9930
      Width           =   675
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F10)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   11700
      TabIndex        =   23
      Top             =   9945
      Width           =   675
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F9)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   5
      Left            =   10560
      TabIndex        =   22
      Top             =   9945
      Width           =   540
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F1)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   6
      Left            =   540
      TabIndex        =   21
      Top             =   9945
      Width           =   540
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(F5)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   4
      Left            =   5595
      TabIndex        =   20
      Top             =   9945
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "米坪"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   39
      Left            =   7440
      TabIndex        =   4
      Top             =   585
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "銘柄"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   38
      Left            =   2520
      TabIndex        =   3
      Top             =   585
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Ｇ№"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   37
      Left            =   60
      TabIndex        =   2
      Top             =   555
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "20XX年1月1日"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   36
      Left            =   45
      TabIndex        =   1
      Top             =   30
      Width           =   1755
   End
   Begin VB.Label LabMode 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "包　装　計　画"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   -15
      Visible         =   0   'False
      Width           =   15360
   End
End
Attribute VB_Name = "L9PK4000F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Des_Col         As Integer
Dim Des_Row         As Variant
Dim Now_BookMk      As Variant

Dim Disp_Kbn        As Integer

' 実行時に新しいスタイルを定義するためのオブジェクトです。
Dim Rstyle_Red      As New Style
Dim Rstyle_Black    As New Style
Dim Rstyle_Blue     As New Style

Dim NowGNo          As String
Dim NowDate         As String
Dim ppGouki     As String

Private Sub ListSet_GNo()
Dim adoRS       As New ADODB.Recordset
Dim strSQL      As String
Dim sBrk        As String
Dim sWk         As String
Dim sNum        As String
Dim yn          As Integer
Dim c           As String


    'Ｓ№，Ｆ№用ＤＢＯＰＥＮ
    adoCon.CursorLocation = adUseClient
    adoCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SHIAGE
    adoCon.Open

    '輸出データＤＢＯＰＥＮ
    YadoCon.CursorLocation = adUseClient
    YadoCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & YUSYUTU
    YadoCon.Open

    strSQL = "SELECT * FROM Gﾏｽﾀｰ WHERE GNo='" & NowGNo & "'"
    adoRS.Open (strSQL), adoCon

    If adoRS.EOF Then
        adoRS.Close
        Set adoRS = Nothing
        'Ｓ№，Ｆ№用ＤＢ ＣＬＯＳＥ
        adoCon.Close
        '輸出データＤＢ ＣＬＯＳＥ
        YadoCon.Close
        yn = zMsgBox("グレードマスター未登録 [" & NowGNo & "]", mgOKOnly)
        Exit Sub
    End If

    Option1(0).Value = True
    DoEvents

    Text1(0).Text = NowGNo
    pGNo_F0 = NowGNo
    pDate_F0 = NowDate
    Text1(1).Text = Meigara_Get(NowGNo, " ")
    pMeigara_F0 = Text1(1).Text
    Text1(2).Text = pBeitubo_F0
'    pBeitubo_F0 = adoRS!BeiTsubo

''    adoRS.Close
''    Set adoRS = Nothing

    If Load_ArX_F0(TDBGrid1.Columns.Count) Then
        'Ｓ№，Ｆ№用ＤＢ ＣＬＯＳＥ
        adoCon.Close
        '輸出データＤＢ ＣＬＯＳＥ
        YadoCon.Close
        yn = zMsgBox("該当グレードの仕上指図データ無し [" & NowGNo & "]", mgOKOnly)
        Exit Sub
    End If

    'Ｓ№，Ｆ№用ＤＢ ＣＬＯＳＥ
    adoCon.Close

    '輸出データＤＢ ＣＬＯＳＥ
    YadoCon.Close


    Frame1.Enabled = True

    Call Disp_Grid

    Label1(36).Caption = Format(Now, "yyyy年m月d日")

End Sub

Private Sub ListSet_Kaktei()
Dim adoRS       As New ADODB.Recordset
Dim strSQL      As String
Dim YadoRS      As New ADODB.Recordset
Dim YstrSQL     As String
Dim iMax        As Long
Dim i           As Long
Dim wSNo        As String
Dim sWk         As String
Dim wCHUNo      As String
Dim wGYONo      As String
Dim BCDSeq      As Long
Dim wRET        As Boolean

    If TDBGrid1.ApproxCount = 0 Then
        Exit Sub
    End If


    'Ｓ№，Ｆ№用ＤＢＯＰＥＮ
    adoCon.CursorLocation = adUseClient
    adoCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SHIAGE
    adoCon.Open

    '輸出データＤＢＯＰＥＮ
    YadoCon.CursorLocation = adUseClient
    YadoCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & YUSYUTU
    YadoCon.Open


    iMax = ArX_F0.UpperBound(1) - 2
'    iMax = ArX_F0.UpperBound(1)

    For i = 0 To iMax
        If Right(ArX_F0(i, Ax0_Kaku), 1) = Chk_No Then
            ArX_F0(i, Ax0_Kaku) = Left(ArX_F0(i, Ax0_Kaku), 1) & Kak_No

            If Len(Trim(ArX_F0(i, Ax0_SFNo))) > 2 Then
                wSNo = ArX_F0(i, Ax0_SFNo)
            Else
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2010/02/09 Upd
                If Hiki_Upd(wSNo, i, wRET) Then
                    adoCon.Close    'S№，F№用DB Close
                    YadoCon.Close   '輸出ﾃﾞｰﾀDB Close
                    Exit Sub
                End If

                '新規確定はﾊﾞｰｺｰﾄﾞ連番を初期設定
                If wRET = True Then
                    '輸出ﾊﾞｰｺｰﾄﾞ連番 初期設定
                    strSQL = "SELECT * FROM SHIAGE WHERE SNo='" & wSNo & "'" & _
                                                   " AND FNo='" & ArX_F0(i, Ax0_SFNo) & "'"
                    adoRS.Open (strSQL), adoCon
                    If adoRS.EOF = False Then
                        BCDSeq = 0
                        If Trim(adoRS!輸出品指定) = "1" Then    '製品ﾊﾞｰｺｰﾄﾞ連番(初期値)
                            sWk = adoRS!輸出注文番号_CNo
                            wCHUNo = Trim(Left(sWk, InStr(1, sWk, "-") - 1))
                            wGYONo = Trim(Right(sWk, Len(sWk) - InStr(1, sWk, "-")))
''                            YstrSQL = "SELECT * FROM Sippingu WHERE Gouki='" & ppGouki & "'" & _
''                                                                 "AND CNo='" & wCHUNo & "'" & _
''                                                              " AND GyoNo=" & wGYONo
                            YstrSQL = "SELECT * FROM Sippingu WHERE CNo='" & wCHUNo & "'" & _
                                                              " AND GyoNo=" & wGYONo
                            YadoRS.Open (YstrSQL), YadoCon
                            If YadoRS.EOF = False Then
                                BCDSeq = YadoRS!KKonNo1 - 1
                            End If
                            YadoRS.Close
                            Set YadoRS = Nothing
                        End If

                        strSQL = "UPDATE SHIAGE SET 製品バーコード連番=" & BCDSeq
                        strSQL = strSQL & " WHERE SNo='" & wSNo & _
                                           "' AND FNo='" & ArX_F0(i, Ax0_SFNo) & "';"
                        adoCon.Execute strSQL
                    End If
                    adoRS.Close
                End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< 2010/02/09 Upd
            End If
        End If
    Next i

    'Ｓ№，Ｆ№用ＤＢ ＣＬＯＳＥ
    adoCon.Close

    '輸出データＤＢ ＣＬＯＳＥ
    YadoCon.Close


    Call Disp_Grid

End Sub

Private Function Hiki_Upd(pSNo As String, pIdx As Long, pRESULT As Boolean) As Integer
'================================================================================================
'                   引当指図データ更新
'================================================================================================
Dim sts         As Integer
Dim com         As Integer


    Hiki_Upd = True


    '引当Ｆ№実績集計
    Call UniCode_Conv(K0_P9_HIKI.SNo, pSNo)
    Call UniCode_Conv(K0_P9_HIKI.FNo, ArX_F0(pIdx, Ax0_SFNo))
    Call UniCode_Conv(K0_P9_HIKI.FNo_SEQ, "")
    com = BtOpGetGreater
    sts = BTRCALL(com, P9_HIKI_POS, P9_HIKI_REC, Len(P9_HIKI_REC), K0_P9_HIKI, Len(K0_P9_HIKI), 0)
    If sts = BtNoErr Then
        If StrConv(P9_HIKI_REC.SNo, vbUnicode) = pSNo And _
           StrConv(P9_HIKI_REC.FNo, vbUnicode) = ArX_F0(pIdx, Ax0_SFNo) Then
            com = BtOpUpdate
        Else
            Exit Function
        End If
    Else
        Call File_Error(sts, com, "引当指図書データ")
        Exit Function
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2010/02/09 Upd
    'データ進捗セット
    If StrConv(P9_HIKI_REC.D_SINC, vbUnicode) = "0" Then
        Call UniCode_Conv(P9_HIKI_REC.D_SINC, "1")
        sts = BTRCALL(com, P9_HIKI_POS, P9_HIKI_REC, Len(P9_HIKI_REC), K0_P9_HIKI, Len(K0_P9_HIKI), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, com, "引当指図書データ")
            Exit Function
        End If
        pRESULT = True
''    ElseIf StrConv(P9_HIKI_REC.D_SINC, vbUnicode) = "1" Then
'''        Call UniCode_Conv(P9_HIKI_REC.D_SINC, "2")
    Else
        pRESULT = False
    End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< 2010/02/09 Upd


    Hiki_Upd = False

End Function

Private Sub Set_Cond(pSetGrp As String, pSetVal As String)
Dim iMax        As Long
Dim i           As Long

    If TDBGrid1.ApproxCount = 0 Then
        Exit Sub
    End If

    iMax = ArX_F0.UpperBound(1) - 2
'    iMax = ArX_F0.UpperBound(1)

    For i = 0 To iMax
        If Val(Left(ArX_F0(i, Ax0_Kaku), Len(ArX_F0(Des_Row, Des_Col)) - 1)) = Val(pSetGrp) Then
           ArX_F0(i, Ax0_Kaku) = Left(ArX_F0(i, Ax0_Kaku), Len(ArX_F0(Des_Row, Des_Col)) - 1) & pSetVal
        End If
    Next i

End Sub

Private Sub Disp_Grid()
Dim iMax        As Long
Dim i           As Long
Dim j           As Long
Dim ixRow       As Long


    If Frame1.Enabled = False Then
        Exit Sub
    End If

    DspAr_F0.Clear
    ixRow = -1

    iMax = ArX_F0.UpperBound(1) - 2
'    iMax = ArX_F0.UpperBound(1)

    For i = 0 To iMax
        If Disp_Kbn <> 2 And (Right(ArX_F0(i, Ax0_Kaku), 1) = Kak_No Or Right(ArX_F0(i, Ax0_Kaku), 1) = Sai_No) Then
            ixRow = ixRow + 1
            DspAr_F0.ReDim 0, ixRow, 0, TDBGrid1.Columns.Count - 1      'Upd 2010/07/01
            For j = 0 To ArX_F0.UpperBound(2)
                DspAr_F0(ixRow, j) = ArX_F0(i, j)
            Next j
        End If

        If Disp_Kbn <> 1 And (Right(ArX_F0(i, Ax0_Kaku), 1) = Mi_No Or Right(ArX_F0(i, Ax0_Kaku), 1) = Chk_No) Then
            ixRow = ixRow + 1
            DspAr_F0.ReDim 0, ixRow, 0, TDBGrid1.Columns.Count - 1      'Upd 2010/07/01
            For j = 0 To ArX_F0.UpperBound(2)
                DspAr_F0(ixRow, j) = ArX_F0(i, j)
            Next j
        End If
    Next i

    If ixRow < 0 Then
'        DspAr_F0.ReDim 0, -1, 0, 0
'        Set TDBGrid1.Array = DspAr_F0
'        TDBGrid1.ApproxCount = 0
        TDBGrid1.Close
    Else
        Set TDBGrid1.Array = DspAr_F0
        TDBGrid1.ReBind
    End If


End Sub

Private Sub Command1_Click(Index As Integer)
'----------------------------------------------------------------------------
'                   処理ファンクション　コントロール
'----------------------------------------------------------------------------
Dim Idx         As Integer
Dim yn          As Integer
Dim sWk         As String
Dim i           As Long

    Select Case Index

        Case 0              '確定
            If TDBGrid1.ApproxCount = 0 Then
                Exit Sub
            End If

            yn = zMsgBox("チェックマークの付いたＳ№を確定します。", _
                         mgYesNo, mgDefaultButton2, "確定")
            If yn = mgretNO Then
                Exit Sub
            End If

            Call ListSet_Kaktei

        Case 1              '日付･Ｇ№
            L9PK4000F2.Show vbModal

            If Rtn_Act <> RACT_OK Then
                Exit Sub
            End If

            NowGNo = strF2
            NowDate = strF2DT
            Call ListSet_GNo

        Case 2              'Ｓ№指定
            L9PK4000F3.Show vbModal

            If Rtn_Act <> RACT_OK Then
                Exit Sub
            End If

            NowGNo = strF3
            NowDate = strF3DT
            Call ListSet_GNo


        Case 3              '修正
            If TDBGrid1.ApproxCount = 0 Then
                Exit Sub
            End If

            pIdx_F4 = TDBGrid1.Bookmark

            L9PK4000F4.Show vbModal

            If Rtn_Act <> RACT_OK Then
                Exit Sub
            End If

            Call ListSet_GNo
'            TDBGrid1.Row = pIdx_F4
            DoEvents
'            TDBGrid1.Row = pIdx_F4

        Case 6              '詳細表示（画面移動）
'            MsgBox "詳細表示画面へ移動", , "画面移動"

            i = Shell("C:\N9LAB_PK\EXE\L9PK4010.exe", vbNormalFocus)
            Unload Me

        Case 7              '指示明細／印刷
            L9PK4000F6.Show vbModal

            If Rtn_Act <> RACT_OK Then
                Exit Sub
            End If


        Case 11             'メニュー
            Unload Me

    End Select
End Sub

Private Sub Form_DblClick()

    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            If Command1(KeyCode - vbKeyF1).Enabled = True Then
                Command1(KeyCode - vbKeyF1).Value = True
            End If
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()

'ｸﾞﾛ-ﾊﾞﾙ項目初期設定
    If Global_Init1 Then
        Unload Me
    End If

'Ｓ№，Ｆ№用ＤＢＯＰＥＮ
    Set adoCon = CreateObject("ADODB.Connection")
''    adoCon.CursorLocation = adUseClient
''    adoCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SHIAGE
''    adoCon.Open

'輸出データＤＢＯＰＥＮ
    Set YadoCon = CreateObject("ADODB.Connection")
''    YadoCon.CursorLocation = adUseClient
''    YadoCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & YUSYUTU
''    YadoCon.Open

'品名マスタＯＰＥＮ
    If HINMEI_Open(BtOpenNomal) Then Unload Me

'包装機品名マスタＯＰＥＮ
    If P9_HIN_Open(BtOpenNomal) Then Unload Me

'引当指図書データＯＰＥＮ
    If P9_HIKI_Open(BtOpenNomal) Then Unload Me

'欠連引当№データＯＰＥＮ
    If P9_KETUR_Open(BtOpenNomal) Then Unload Me



' 新しいスタイルを定義します。
                                    '黒
    Set Rstyle_Black = TDBGrid1.Styles.Add("Rstyle_Black")
    Rstyle_Black.BackColor = &H80000005         '背景色＝白
    Rstyle_Black.ForeColor = &H0                '文字色＝黒
                                    '青
    Set Rstyle_Blue = TDBGrid1.Styles.Add("Rstyle_Blue")
    Rstyle_Blue.BackColor = &H80000005          '背景色＝白
    Rstyle_Blue.ForeColor = &HFF0000            '文字色＝青
                                    '赤
    Set Rstyle_Red = TDBGrid1.Styles.Add("Rstyle_Red")
    Rstyle_Red.BackColor = &H80000005           '背景色＝白
    Rstyle_Red.ForeColor = &HFF                 '文字色＝赤


    Show

    DoEvents
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""

    NowGNo = ""
    NowDate = ""

    Disp_Kbn = 0

    Label1(36).Caption = Format(Now, "yyyy年m月d日")

    ppGouki = Text1(3).Text

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts     As Integer

'品名マスタＣＬＯＳＥ
    sts = BTRCALL(BtOpClose, HINMEI_POS, HINMEI_REC, Len(HINMEI_REC), K0_HINMEI, Len(K0_HINMEI), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品名マスタ")
        End If
    End If

'包装機品名マスタＣＬＯＳＥ
    sts = BTRCALL(BtOpClose, P9_HIN_POS, P9_HIN_REC, Len(P9_HIN_REC), K0_P9_HIN, Len(K0_P9_HIN), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "包装機品名マスタ")
        End If
    End If

'引当指図書データＣＬＯＳＥ
    sts = BTRCALL(BtOpClose, P9_HIKI_POS, P9_HIKI_REC, Len(P9_HIKI_REC), K0_P9_HIKI, Len(K0_P9_HIKI), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "引当指図書データ")
        End If
    End If

'欠連引当№データＣＬＯＳＥ
    sts = BTRCALL(BtOpClose, P9_KETUR_POS, P9_KETUR_REC, Len(P9_KETUR_REC), K0_P9_KETUR, Len(K0_P9_KETUR), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "欠連引当№データ")
        End If
    End If

    End

End Sub

Private Sub Option1_Click(Index As Integer)

    If TDBGrid1.ApproxCount = 0 Then
        Exit Sub
    End If

    Disp_Kbn = Index

    Call Disp_Grid

End Sub


Private Sub TDBGrid1_BeforeRowColChange(Cancel As Integer)

    If TDBGrid1.DestinationCol = 1 Then

        If Right(ArX_F0(TDBGrid1.DestinationRow, TDBGrid1.DestinationCol), Ax0_Kaku) <> Kak_No And _
           Right(ArX_F0(TDBGrid1.DestinationRow, TDBGrid1.DestinationCol), Ax0_Kaku) <> Sai_No Then

            Now_BookMk = TDBGrid1.Bookmark
            Des_Row = TDBGrid1.DestinationRow
            Des_Col = TDBGrid1.DestinationCol

            TDBGrid1.PostMsg 1

        End If

    End If

End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)

    If Right(DspAr_F0(Bookmark, Ax0_Kaku), Ax0_Kaku) = Kak_No Then
        RowStyle = "Rstyle_Black"
    ElseIf Right(DspAr_F0(Bookmark, Ax0_Kaku), Ax0_Kaku) = Sai_No Then
        RowStyle = "Rstyle_Blue"
    Else
        RowStyle = "Rstyle_Red"
    End If

End Sub

Private Sub TDBGrid1_PostEvent(ByVal MsgId As Integer)

    If Right(ArX_F0(Des_Row, Des_Col), Ax0_Kaku) = Mi_No Then
        Call Set_Cond(Left(ArX_F0(Des_Row, Des_Col), Len(ArX_F0(Des_Row, Des_Col)) - 1), Chk_No)
    Else
        Call Set_Cond(Left(ArX_F0(Des_Row, Des_Col), Len(ArX_F0(Des_Row, Des_Col)) - 1), Mi_No)
    End If

    Call Disp_Grid

    Tim_Disp.Enabled = False

    TDBGrid1.Col = 0
    TDBGrid1.Bookmark = Now_BookMk

End Sub
