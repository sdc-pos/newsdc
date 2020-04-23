VERSION 5.00
Begin VB.Form F1020121 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入出荷予定データ取込み「袋井用」"
   ClientHeight    =   6324
   ClientLeft      =   1908
   ClientTop       =   2388
   ClientWidth     =   11220
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
   ScaleHeight     =   6324
   ScaleWidth      =   11220
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   1
      Left            =   4800
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   0
      Left            =   3480
      TabIndex        =   37
      Top             =   1560
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ListBox LBox_Dup 
      ForeColor       =   &H00000080&
      Height          =   288
      Left            =   9840
      Sorted          =   -1  'True
      TabIndex        =   31
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "再取込"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      TabIndex        =   4
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "特売"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      TabIndex        =   3
      Top             =   2400
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.ListBox LBox_Hin 
      ForeColor       =   &H00000080&
      Height          =   288
      Left            =   9840
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox LBox_Etc 
      ForeColor       =   &H00000080&
      Height          =   288
      Left            =   9840
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "１便"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      TabIndex        =   1
      Top             =   1200
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "２便"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      TabIndex        =   2
      Top             =   1800
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "３便"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      TabIndex        =   0
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   3
      Left            =   8400
      TabIndex        =   42
      Top             =   4080
      Width           =   492
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   9000
      TabIndex        =   41
      Top             =   4080
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   6720
      TabIndex        =   40
      Top             =   4080
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "取込み処理中　取込み件数＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   1680
      TabIndex        =   39
      Top             =   4080
      Width           =   4932
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "特売り"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4920
      TabIndex        =   36
      Top             =   1320
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "SPIC"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   3720
      TabIndex        =   35
      Top             =   1320
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷重複分印刷中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   9
      Left            =   1920
      TabIndex        =   34
      Top             =   5040
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "（品番変更保留分・出荷重複分）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2640
      TabIndex        =   32
      Top             =   3240
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷重複"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   9840
      TabIndex        =   30
      Top             =   2880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "保留データ再処理中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   1920
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番変更リスト印刷中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   7
      Left            =   2520
      TabIndex        =   26
      Top             =   5040
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "赤伝、訂正、出荷確認一覧印刷中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   6
      Left            =   1440
      TabIndex        =   24
      Top             =   5040
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "受信エラーリスト印刷中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   5
      Left            =   2400
      TabIndex        =   22
      Top             =   5040
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番変更"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   9840
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "受信ｴﾗｰﾘｽﾄ"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   9840
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "３便取込み処理中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   3000
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "２便取込み処理中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   3240
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "１便取込み処理中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   3000
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "（「３便」「１便」「２便」反転表示は本日取込み済）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "前借り入荷、実績残チェック中！！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   8
      Left            =   1440
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "特売取込み処理中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   1920
      TabIndex        =   33
      Top             =   5040
      Visible         =   0   'False
      Width           =   4320
   End
End
Attribute VB_Name = "F1020121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#################################################################################################
'［テキストファイル処理の注意！！］
'
'　　このプログラムでは、テキストファイルをユーザー定義型の構造体で読書きしている。
'　　この形式によるＩ－Ｏは、「GET」および「PUT」ｽﾃｰﾄﾒﾝﾄにより行う事になるが、これらのｽﾃｰﾄﾒﾝﾄ
'　　を使用する為には、「RANDOM」または「BINARY」モードで「OPEN」しなければならない。
'　　但し、以下のロジックでは、テキストファイルの存在をチェックする目的で「INPUT」モードによる
'　　「OPEN」も行っており、存在チェック後、すぐに読込みを行う様な場合には、一旦「CLOSE」してから
'　　「BINARYモードでOPEN」している事に注意が必要。
'
'　　※．「INPUT」モードOPENでは、「INPUT#」「OUTPUT#」のみ使用できるが、これらは構造体のメンバ
'　　　　単位のＩ－Ｏになる。
'#################################################################################################
Private Type YUKO_SOKO_TBL1             '有効ﾎｽﾄ倉庫取り込みテーブル（洗濯機事業部）
    HS_SOKO As String * 2
    NAIGAI As String * 1
End Type
Dim SOKO_T1(ZERO To 9) As YUKO_SOKO_TBL1

Dim WS_NO As String * 2                 'ﾜｰｸｽﾃｰｼｮﾝ番号

Dim HS_NaiG As String                   '国内外（決定内容）････　ﾎｽﾄﾃﾞｰﾀ内容により設定
Dim BEF_GAI As String * 13              '変更前品番（外部）
Dim BEF_NAI As String * 13              '変 更前品番（内部）

Dim PRT_CAN As Boolean                  '印刷途中キャンセル要求
Dim NormalFont As New StdFont           '印刷フォント

Private Const LMAX% = 46                '頁内最大行数
Private Const MGN_L% = 1                '明細印刷開始桁位置（１から）
Private Const MGN_U% = 1                '上余白（行数：１から）
Private Const MGN_L2% = 20              '「過剰前借品ﾘｽﾄ」明細印刷開始桁位置（１から）
Private Const MGN_U2% = 1               '「過剰前借品ﾘｽﾄ」上余白（行数：１から）
Dim Pdate As String                     '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime As String                     '印刷開始時刻（ﾍｯﾀﾞｰ用）

Dim Proc_F As Integer                   '品番＆在庫有無　判定フラグ
Dim Last_Proc_F As Integer              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有無フラグ

Dim Shori_Mode  As Integer
Private Const Text_Max% = 1
                                    
Private Const Er_Soko_NoT% = True       'ホスト倉庫異常
Private Const Er_Item_NoT% = True       '伝票日付／品番／入出庫区分異常
Private Const Er_Dup_NoT% = True        '伝票重複
Private Const Er_Muke_NoT% = True       '向け先異常
                                    
                                    '画面初期表示（処理済「便」表示など）
Private Sub Scr_Init()
Dim i As Integer
Dim sts As Integer
Dim Work As String

    For i = 1 To 3
        Work = Format(Date, "ddd") & "." & Format(i, "0") & "00"
        sts = XX_SIJ_Open(Work, ZERO)
        If sts = False Then
            Close #XX_SIJ_No
            If Format(FileDateTime(Work), "yyyy/mm/dd") = Format(Date, "yyyy/mm/dd") Then
                SelCmd(i).Enabled = False
            Else
                SelCmd(i).Enabled = True
            End If
        Else
            SelCmd(i).Enabled = True
        End If
    Next i


End Sub
                                            'ホストデータ取込み処理
Private Function Data_Inport() As Integer
Dim sts         As Integer
Dim ans         As Integer
Dim Command     As Integer
Dim FPass       As String
Dim Work        As String
Dim FP_XX_SIJ   As String
Dim FP_ER_SIJ   As String
Dim FP_CHGHIN   As String
Dim FP_SYUDUP   As String

Dim In_Cnt      As Integer

    On Error Resume Next    'FileCopy / Kill ｽﾃｰﾄﾒﾝﾄでのﾌｧｲﾙ無しは次ｽﾃｯﾌﾟを実行

    Call Input_Lock                                 '画面項目ロック

    MsgLab(Shori_Mode).Visible = True       '更新中ﾒｯｾｰｼﾞ表示
    DoEvents

'ホスト便別ワーク　削除

    sts = GetIni("FILE", XX_SIJ_ID, "SYS", FPass)
    If sts <> False Then
        MsgBox "SYS.INI [FILE][" & XX_SIJ_ID & "]読み込みエラー "
        Call Log_Out(LOG_F, "SYS.INI [FILE][" & XX_SIJ_ID & "]読み込みエラー ")
        Unload Me
    End If
    FP_XX_SIJ = RTrim(FPass) & Format(Date, "ddd") & "." & Format(Shori_Mode, "0") & "00"
    Kill FP_XX_SIJ
    FP_ER_SIJ = RTrim(FPass) & Format(Date, "ddd") & "." & Format(Shori_Mode, "0") & "0E"
    Kill FP_XX_SIJ

'ﾌｧｲﾙ OPEN
    Work = Format(Date, "ddd") & "." & Format(Shori_Mode, "0") & "00"
    If XX_SIJ_Open(Work, 1) Then            '取込みﾜｰｸ
        Unload Me
    End If

    Work = Format(Date, "ddd") & "." & Format(Shori_Mode, "0") & "0E"
    If ER_SIJ_Open(Work, 1) Then            '取込みﾜｰｸ[対象外ﾃﾞｰﾀ用]
        Close #XX_SIJ_No
        Unload Me
    End If
    
    If CHGH_Open() Then                     '品番変更保留ﾃﾞｰﾀ
        Unload Me
    End If
    
    If SYUDUP_Open() Then                   '出荷予定重複ﾃﾞｰﾀ
        Unload Me
    End If


    Call Data_Load                          '便別取込みワークにﾃﾞｰﾀﾛｰﾄﾞ

        
    Close #CHGH_No                          '品番変更保留ﾃﾞｰﾀ CLOSE
    Close #SYUDUP_No                        '出荷予定重複ﾃﾞｰﾀ CLOSE
                                            
                                            
                                                '品番変更保留ﾃﾞｰﾀ削除
    sts = GetIni("FILE", CHGH_ID, "SYS", FP_CHGHIN)
    If sts <> False Then
        MsgBox "SYS.INI [FILE][" & CHGH_ID & "]読み込みエラー "
        Call Log_Out(LOG_F, "SYS.INI [FILE][" & CHGH_ID & "]読み込みエラー ")
        Unload Me
    End If
    FP_CHGHIN = RTrim(FP_CHGHIN)
    Kill FP_CHGHIN                          '品番変更保留ﾃﾞｰﾀ クリア

                                                '出荷予定重複ﾃﾞｰﾀ削除
    sts = GetIni("FILE", SYUDUP_ID, "SYS", FP_SYUDUP)
    If sts <> False Then
        MsgBox "SYS.INI [FILE][" & SYUDUP_ID & "]読み込みエラー "
        Call Log_Out(LOG_F, "SYS.INI [FILE][" & SYUDUP_ID & "]読み込みエラー ")
        Unload Me
    End If
    FP_SYUDUP = RTrim(FP_SYUDUP)
    Kill FP_SYUDUP                          '出荷予定重複ﾃﾞｰﾀ クリア

        
    If CHGH_Open() Then                     '品番変更保留ﾃﾞｰﾀ
        Unload Me
    End If
        
    If SYUDUP_Open() Then                   '出荷予定重複ﾃﾞｰﾀ
        Close #CHGH_No
        Unload Me
    End If
'ホストデータ　チェック＆取込み
    LBox_Etc.Clear      '印刷ﾃﾞｰﾀ用ﾘｽﾄﾎﾞｯｸｽ　クリア
    LBox_Hin.Clear
    LBox_Dup.Clear

    In_Cnt = ZERO

    Do
        If XX_SIJ_Get Then          '取込みﾜｰｸ 読込み
            Exit Do
        End If
        If Left(StrConv(XX_SIJREC.TEXT_NO, vbUnicode), 1) < " " Then    'EOF ?
            Exit Do
        End If

        sts = Data_Chk              'ﾃﾞｰﾀ重複／項目内容 チェック
        
        If sts = False Then
                                    '赤伝・出荷確認・訂正ﾃﾞｰﾀは印刷のみ 出荷実績・良品返品分追加
            If StrConv(XX_SIJREC.PM_KBN, vbUnicode) = "-" Or _
               StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "2" Or _
               StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "3" Or _
               StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "4" Or _
               StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "5" Then
                Call PDat_Etc_Add("1", " ")         '赤伝・出荷確認・訂正印刷ﾃﾞｰﾀ保存
            Else
                Call Proc_Sel(In_Cnt)               '売上(0)，入荷(1)ﾃﾞｰﾀ更新
            End If
        End If
    Loop

    Close #XX_SIJ_No            '取込みﾜｰｸ CLOSE
    Close #ER_SIJ_No            '取込みﾜｰｸ[対象外ﾃﾞｰﾀ用] CLOSE
    Close #CHGH_No              '品番変更保留ﾃﾞｰﾀ CLOSE
    Close #SYUDUP_No            '出荷重複保留ﾃﾞｰﾀ CLOSE

    MsgLab(Format(Shori_Mode, "0")).Visible = False      '更新中ﾒｯｾｰｼﾞ ｸﾘｱ

    Pdate = Date
    Ptime = Time
    Printer.Orientation = vbPRORLandscape       '用紙の長辺を上にして印刷
    
'受信エラーリスト印刷
    If LBox_Etc.ListCount > ZERO Then
        MsgLab(5).Visible = True                '印刷中ﾒｯｾｰｼﾞ表示
        DoEvents
        Set Printer.Font = NormalFont           '印刷フォント設定
        Call P_Etc_Proc(ZERO)                   '受信エラーリスト印刷
        MsgLab(5).Visible = False               '印刷中ﾒｯｾｰｼﾞクリア
'        DoEvents
    End If

'赤伝、訂正、出荷確認一覧表印刷
    If LBox_Etc.ListCount > ZERO Then
        MsgLab(6).Visible = True                '印刷中ﾒｯｾｰｼﾞ表示
        DoEvents
        Set Printer.Font = NormalFont           '印刷フォント設定
        Call P_Etc_Proc(1)                      '赤伝、訂正、出荷確認一覧表印刷
        MsgLab(6).Visible = False               '印刷中ﾒｯｾｰｼﾞクリア
'        DoEvents
    End If

'品番変更リスト印刷
    If LBox_Hin.ListCount > ZERO Then
        MsgLab(7).Visible = True                '印刷中ﾒｯｾｰｼﾞ表示
        DoEvents
        Set Printer.Font = NormalFont           '印刷フォント設定
        Call P_Hin_Proc                         '品番変更リスト印刷
        MsgLab(7).Visible = False               '印刷中ﾒｯｾｰｼﾞクリア
        DoEvents
    End If

'出荷重複リスト印刷
    If LBox_Dup.ListCount > ZERO Then
        MsgLab(9).Visible = True                '印刷中ﾒｯｾｰｼﾞ表示
        DoEvents
        Set Printer.Font = NormalFont           '印刷フォント設定
        Call P_Dup_Proc                         '出荷重複リスト印刷
        MsgLab(9).Visible = False               '印刷中ﾒｯｾｰｼﾞクリア
        DoEvents
    End If


    Call Input_UnLock                               '画面項目ロック解除

    Call Scr_Init                                   '画面クリア

End Function
                                            'データロード(各業部別ﾎｽﾄﾃﾞｰﾀ→便別取込みﾜｰｸ）
Private Sub Data_Load()
Dim sts As Integer
Dim Work As String

'便別取込みワークにデータロード（品番保留データ）
    Do
        If CHGH_Get Then            '保留データ 読込み
            Exit Do
        End If

        If Left(StrConv(CHGHREC.TEXT_NO, vbUnicode), 1) < " " Then    'EOF ?
            Exit Do
        End If

        If CHGH_Put(1) Then         '取込みﾜｰｸ書込み
            Close #CHGH_No
            Unload Me
        End If
    Loop

'便別取込みワークにデータロード（出荷予定データ）
    Do
        If SYUDUP_Get Then          '保留データ 読込み
            Exit Do
        End If

        If Left(StrConv(SYUDUPREC.TEXT_NO, vbUnicode), 1) < " " Then    'EOF ?
            Exit Do
        End If

        If SYUDUP_Put(1) Then       '取込みﾜｰｸ書込み
            Close #CHGH_No
            Close #SYUDUP_No
            Unload Me
        End If
    Loop
'便別取込みワークにデータロード（「1～3便」及び「特売り」指定時のみ ）
'重要　※　特売りは４便とみなす　※
    If Shori_Mode > ZERO Then
                                        '洗濯機事業部  ホスト受信データ取込み
        If HS_SIJ_Open1(ZERO, Format(Shori_Mode, "0")) = False Then
                                                        'ﾌｧｲﾙ無しなら処理しない
            Close #HS_SIJ_No
                                                        '洗濯機事業部ﾎｽﾄﾃﾞｰﾀ OPEN
            If HS_SIJ_Open1(1, Format(Shori_Mode, "0")) Then
                Unload Me
            End If

            Call Data_Load_Sub              'ﾃｷｽﾄ№

            Close #HS_SIJ_No                'ﾎｽﾄﾃﾞｰﾀ CLOSE
        End If
    End If

'取込みワーク 再ＯＰＥＮ
    Close #XX_SIJ_No
    Work = Format(Date, "ddd") & "." & Format(Shori_Mode, "0") & "00"
    If XX_SIJ_Open(Work, 1) Then
        Close #CHGH_No
        Close #SYUDUP_No
        Unload Me
    End If

End Sub
                                            '事業部別データロード(事業部別ﾎｽﾄﾃﾞｰﾀ→便別取込みﾜｰｸ）
Private Sub Data_Load_Sub()
Dim sts As Integer
Dim Put_Sel As Integer
Dim Work As String
Dim i As Integer

Dim In_Cnt  As Integer


    In_Cnt = ZERO
    
    Do
        DoEvents
        If HS_SIJ_Get Then          'ﾎｽﾄﾃﾞｰﾀ 読込み
            Exit Do
        End If
        If Left(StrConv(HS_SIJREC.TEXT_NO, vbUnicode), 1) < " " Then
            Exit Do
        End If

        In_Cnt = In_Cnt + 1
                                
        Label3(2).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
                                
                                'ﾃｷｽﾄ№
        Call UniCode_Conv(XX_SIJREC.TEXT_NO, StrConv(HS_SIJREC.TEXT_NO, vbUnicode))
                                '事業部区分
        Call UniCode_Conv(XX_SIJREC.JGYOBU, StrConv(HS_SIJREC.JGYOBU, vbUnicode))
                                '直送区分
        Call UniCode_Conv(XX_SIJREC.CYOK_KBN, StrConv(HS_SIJREC.CYOK_KBN, vbUnicode))
                                '伝票日付
        Call UniCode_Conv(XX_SIJREC.DEN_DT, StrConv(HS_SIJREC.DEN_DT, vbUnicode))
                                '入出庫区分
        Call UniCode_Conv(XX_SIJREC.IO_KBN, StrConv(HS_SIJREC.IO_KBN, vbUnicode))
                                '赤黒区分
        Call UniCode_Conv(XX_SIJREC.PM_KBN, StrConv(HS_SIJREC.PM_KBN, vbUnicode))
                                '伝票種別
        Call UniCode_Conv(XX_SIJREC.DEN_SYU, StrConv(HS_SIJREC.DEN_SYU, vbUnicode))
                                '伝票№
        Call UniCode_Conv(XX_SIJREC.DEN_NO, StrConv(HS_SIJREC.DEN_NO, vbUnicode))
                                '注文区分
        Call UniCode_Conv(XX_SIJREC.CYU_KBN, StrConv(HS_SIJREC.CYU_KBN, vbUnicode))
                                '品番（外部）
        Call UniCode_Conv(XX_SIJREC.HIN_GAI, StrConv(HS_SIJREC.HIN_GAI, vbUnicode))
                                '品番（内部）
        Call UniCode_Conv(XX_SIJREC.HIN_NAI, StrConv(HS_SIJREC.HIN_NAI, vbUnicode))
                                '品名
        Call UniCode_Conv(XX_SIJREC.HIN_NAME, StrConv(HS_SIJREC.HIN_NAME, vbUnicode))
                                '数量
        Call UniCode_Conv(XX_SIJREC.YOTEI_QTY, StrConv(HS_SIJREC.YOTEI_QTY, vbUnicode))
                                '予算単位（元）
        Call UniCode_Conv(XX_SIJREC.YOSAN_FROM, StrConv(HS_SIJREC.YOSAN_FROM, vbUnicode))
                                '予算単位（先）
        Call UniCode_Conv(XX_SIJREC.YOSAN_TO, StrConv(HS_SIJREC.YOSAN_TO, vbUnicode))
                                '倉庫区分（ﾎｽﾄ）
        Call UniCode_Conv(XX_SIJREC.HOST_SOKO, StrConv(HS_SIJREC.HOST_SOKO, vbUnicode))
                                '棚番（ﾎｽﾄ）
        Call UniCode_Conv(XX_SIJREC.HOST_TANA, StrConv(HS_SIJREC.HOST_TANA, vbUnicode))
                                '支給先／出荷先
        Call UniCode_Conv(XX_SIJREC.SYUK_CODE, StrConv(HS_SIJREC.SYUK_CODE, vbUnicode))
                                '支給先／出荷先名
        Call UniCode_Conv(XX_SIJREC.SYUK_NAME, StrConv(HS_SIJREC.SYUK_NAME, vbUnicode))
                                'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
        Call UniCode_Conv(XX_SIJREC.REC_END, StrConv(HS_SIJREC.REC_END, vbUnicode))
                                'CR.LF
        Call UniCode_Conv(XX_SIJREC.CR_LF, StrConv(HS_SIJREC.CR_LF, vbUnicode))

        Put_Sel = True
                                                '事業部区分　範囲外？
        For i = ZERO To UBound(JGYOBU_T) - 1
            If JGYOBU_T(i).Code = " " Then
                Put_Sel = False
                Exit For
            End If
            If JGYOBU_T(i).Code = StrConv(HS_SIJREC.JGYOBU, vbUnicode) Then
                Exit For
            End If
        Next i
                                                'ホスト倉庫　範囲外？
        For i = ZERO To UBound(SOKO_T1) - 1
            If SOKO_T1(i).HS_SOKO = "  " Then
                Put_Sel = False
                Exit For
            End If
            If RTrim(StrConv(HS_SIJREC.HOST_SOKO, vbUnicode)) = RTrim(SOKO_T1(i).HS_SOKO) Then
                Exit For
            End If
        Next i
        
        If Put_Sel = True Then
            sts = XX_SIJ_Put                    '取込みﾜｰｸ書込み（対象倉庫）
        Else
        '対象外ホスト倉庫！！   エラーログ
            If Er_Soko_NoT Then
                Call Err_Log_Out("Ｈ倉庫対象外")
            End If
            sts = ER_SIJ_Put                    '取込みﾜｰｸ書込み（対象外倉庫）
        End If
        If sts Then
            Close #HS_SIJ_No
            Close #XX_SIJ_No
            Unload Me
        End If
    Loop

End Sub
                                            'データ重複／項目内容 チェック
Private Function Data_Chk() As Integer

Dim sts     As Integer
Dim Work    As String
Dim i       As Integer

Dim MUKECHG As String * 10      '2001.07.04

    Data_Chk = False

'データ重複チェック（再処理データ（終端＝「？」「＊」はチェック無し）
    If StrConv(XX_SIJREC.REC_END, vbUnicode) <> "?" And StrConv(XX_SIJREC.REC_END, vbUnicode) <> "*" Then
        Call UniCode_Conv(K0_SEQCK.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
        If Shori_Mode = 4 Then
            Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "2")
        Else
            Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "1")
        End If
        sts = BTRV(BtOpGetEqual, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
        If sts Then
            If sts <> BtErrKeyNotFound Then
                Call File_Error(sts, BtOpGetEqual, "予定取込みチェック")
                Unload Me
            End If
            Call UniCode_Conv(SEQCKREC.LAST_TXTNO, "000000000")
        End If
                '
                '「前回ﾃｷｽﾄ№≧今回ﾃｷｽﾄ№」はエラー
                '（但し、前回ﾃｷｽﾄ№＝12(月)時は無条件ＯＫ）
                '
        If Shori_Mode = 4 Then
                '特売りデータはテキスト№の便＝５、注文区分＝４以外エラー
            If Mid(StrConv(XX_SIJREC.TEXT_NO, vbUnicode), 5, 1) <> "5" Or StrConv(XX_SIJREC.CYU_KBN, vbUnicode) <> CYU_KBN_TOK Then
                Data_Chk = True
                Exit Function
            End If
                '昨日以前のテキスト№の月日はエラー
            If Left(StrConv(XX_SIJREC.TEXT_NO, vbUnicode), 4) < Right(Format(Date, "yyyymmdd"), 4) Then
                Data_Chk = True
                Exit Function
            End If
        Else
'-----------テキスト№チェック無効
'            If Left(StrConv(SEQCKREC.LAST_TXTNO, vbUnicode), 2) <> "12" Then
'                If StrConv(SEQCKREC.LAST_TXTNO, vbUnicode) >= StrConv(XX_SIJREC.TEXT_NO, vbUnicode) Then
'                    Data_Chk = True
'                    Exit Function
'                End If
'            End If
'-----------テキスト№チェック無効
        End If
    Else
        If Shori_Mode > ZERO Then
                                '再処理要求じゃなければ
            Select Case StrConv(XX_SIJREC.REC_END, vbUnicode)
                Case "?"
                    If CHin_Put() Then            '外部品番変更保留ﾃﾞｰﾀ作成
                        Unload Me
                    End If
                
                Case "*"
                    Call MAKE_SYUDUP_Put                    '重複保留ﾃﾞｰﾀ作成
            End Select
            Data_Chk = True
            Exit Function
        End If
    End If
'取込みデータ 項目内容チェック
'   [ﾁｪｯｸ 項目]
'       1) 予定日　　　：日付ﾁｪｯｸ
'       2) 品番（外部）：≠空白
'       4) 入出庫区分　：範囲＝"0～3"or"E" ，"=0"の時、出荷先≠空白
'
'97.07.29  内部品番は「空白」も可とする。但し、「空白」の時はマスタの品番と置換えない。
'
    Work = StrConv(XX_SIJREC.DEN_DT, vbUnicode)     '伝票日付
    If IsDate(Left(Work, 4) & "/" & Mid(Work, 5, 2) & "/" & Right(Work, 2)) = False _
     Or StrConv(XX_SIJREC.HIN_GAI, vbUnicode) = Space(13) _
     Or StrConv(XX_SIJREC.IO_KBN, vbUnicode) < "0" _
     Or (StrConv(XX_SIJREC.IO_KBN, vbUnicode) > "3" _
       And StrConv(XX_SIJREC.IO_KBN, vbUnicode) <> "E") _
     Or (StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "0" _
       And StrConv(XX_SIJREC.SYUK_CODE, vbUnicode) = Space(5)) Then
        
        '伝票日付／品番／入出庫区分異常！！   エラーログ
        If Er_Item_NoT Then
            Call Err_Log_Out("伝票日付／品番／入出庫区分異常")
        End If
        
        Call PDat_Etc_Add("0", "0")                 'ｴﾗｰﾘｽﾄ印刷ﾃﾞｰﾀ保存
        Data_Chk = True
        Exit Function
    End If


'国内外区分の設定
    For i = ZERO To UBound(SOKO_T1)
        If SOKO_T1(i).HS_SOKO = "  " Then
            Exit For
        End If
        If RTrim(StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)) = RTrim(SOKO_T1(i).HS_SOKO) Then
            HS_NaiG = SOKO_T1(i).NAIGAI
            Exit For
        End If
    Next i


    If StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "0" Then
'同一出荷伝票存在チェック(事業部＋注文区分＋伝票№)
                                                    '事業部
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                                                    '注文区分
        If StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_TOK Then
            Call UniCode_Conv(K0_Y_SYU.HS_CYU_KBN, CYU_KBN_SPO)
        Else
            Call UniCode_Conv(K0_Y_SYU.HS_CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
        End If
                                                    '伝票№
        Call UniCode_Conv(K0_Y_SYU.DEN_NO, StrConv(XX_SIJREC.DEN_NO, vbUnicode))
                                                    'ＳＳ追番（空白固定）
        Call UniCode_Conv(K0_Y_SYU.SS_CODE, "")
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), ZERO)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_SYUREC.TOK_KBN, vbUnicode) = "1" Then
                    If StrConv(XX_SIJREC.SYUK_CODE, vbUnicode) = StrConv(Y_SYUREC.SYUK_CODE, vbUnicode) And _
                        StrConv(XX_SIJREC.HIN_GAI, vbUnicode) = StrConv(Y_SYUREC.HIN_GAI, vbUnicode) And _
                        CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)) = CLng(StrConv(Y_SYUREC.YOTEI_QTY, vbUnicode)) Then
                                '特売り時、向け先／品番／数量が等しいときは捨てる
                    Else
        '伝票日付／品番／入出庫区分異常！！   エラーログ
                        If Er_Dup_NoT Then
                            Call Err_Log_Out("重複伝票")
                        End If
                        
                        Call MAKE_SYUDUP_Put                    '重複保留ﾃﾞｰﾀ作成
                        Call PDat_DUP_Add("0", "2")             'ｴﾗｰﾘｽﾄ印刷ﾃﾞｰﾀ保存
                    End If
                Else
        '伝票日付／品番／入出庫区分異常！！   エラーログ
                    If Er_Dup_NoT Then
                        Call Err_Log_Out("重複伝票")
                    End If
                    
                    Call MAKE_SYUDUP_Put                    '重複保留ﾃﾞｰﾀ作成
                    Call PDat_DUP_Add("0", "2")             'ｴﾗｰﾘｽﾄ印刷ﾃﾞｰﾀ保存
                End If
                Data_Chk = True
                Exit Function
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定")
                Unload Me
        End Select
    
'向け先コード読替処理 2001.07.04
        Call UniCode_Conv(K0_MTSCHG.RYAKU, StrConv(XX_SIJREC.SYUK_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, MTSCHG_POS, MTSCHGREC, Len(MTSCHGREC), K0_MTSCHG, Len(K0_MTSCHG), ZERO)
        Select Case sts
            Case BtNoErr
                MUKECHG = StrConv(MTSCHGREC.MUKE_CODE, vbUnicode)
            Case BtErrKeyNotFound
                MUKECHG = StrConv(XX_SIJREC.SYUK_CODE, vbUnicode)
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先読替マスタ")
                Unload Me
        End Select
'向け先コード存在チェック
        Call UniCode_Conv(K0_MTS.MUKE_CODE, MUKECHG)
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), ZERO)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(K0_MTS.MUKE_CODE, ETS_MTS & HS_NaiG)
                sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), ZERO)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                                            'その他向け先対応しない
                        If CHin_Put() Then              '外部品番変更保留ﾃﾞｰﾀ作成
                            Unload Me
                        End If
                        
                        '向け先異常！！   エラーログ
                        If Er_Muke_NoT Then
                            Call Err_Log_Out("向け先異常")
                        End If
                        
                        
                        Call PDat_Etc_Add("0", "1")     'ｴﾗｰﾘｽﾄ印刷ﾃﾞｰﾀ保存
                        Data_Chk = True
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                        Unload Me
                End Select
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                Unload Me
        End Select
    
    End If
End Function
                                            'データ重複チェック用データ　更新
Private Function Seq_Update() As Boolean

Dim sts     As Integer
Dim Command As Integer
Dim ans     As Integer

    Seq_Update = True
'重複チェック用データ更新（再処理データ（終端＝「？」「＊」は更新しない）
    If StrConv(XX_SIJREC.REC_END, vbUnicode) <> "?" And StrConv(XX_SIJREC.REC_END, vbUnicode) <> "*" Then
        Call UniCode_Conv(K0_SEQCK.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
        If Shori_Mode = 4 Then
            Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "2")
        Else
            Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "1")
        End If
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
                        
            Select Case sts
                Case BtNoErr
                    Command = BtOpUpdate
                    Exit Do
                Case BtErrEOF, BtErrKeyNotFound
                    Command = BtOpInsert
                    Call UniCode_Conv(SEQCKREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                    If Shori_Mode = 4 Then
                        Call UniCode_Conv(SEQCKREC.SEQ_MODE, "2")
                    Else
                        Call UniCode_Conv(SEQCKREC.SEQ_MODE, "1")
                    End If
                    Call UniCode_Conv(SEQCKREC.LAST_TXTNO, "000000000")
                    Call UniCode_Conv(SEQCKREC.LAST_GET_DT, "00000000")
                    Call UniCode_Conv(SEQCKREC.LAST_GET_TM, "000000")
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<SEQCK.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "予定取込みチェック")
                    Exit Function
            End Select
        Loop
        
        Call UniCode_Conv(SEQCKREC.LAST_TXTNO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))   '最終テキスト№
        Call UniCode_Conv(SEQCKREC.LAST_GET_DT, Format(Date, "yyyymmdd"))               '最終取込み日付
        Call UniCode_Conv(SEQCKREC.LAST_GET_DT, Format(Time, "hhmm"))                   '最終取込み時刻
        
        Do
            sts = BTRV(Command, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<SEQCK.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, Command, "予定取込みチェック")
                    Exit Function
            End Select
        Loop
    End If

    Seq_Update = False

End Function
                                            '売上(0)，入荷(1)データ更新
Private Sub Proc_Sel(In_Cnt As Integer)
Dim sts As Integer
Dim ans As Integer
Dim Command As Integer
Dim i As Integer



    Proc_F = ZERO

    DoEvents
'品番マスタデータ（外部）有無チェック
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts = BtNoErr Then
        Proc_F = Proc_F + 1
        BEF_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
        BEF_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
    Else
        If sts <> BtErrEOF And sts <> BtErrKeyNotFound Then
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Unload Me
        End If
    End If

'品番マスタデータ（内部）有無チェック
    Call UniCode_Conv(K3_ITEM.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K3_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K3_ITEM.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K3_ITEM, Len(K3_ITEM), 3)
    If sts = BtNoErr Then
        Proc_F = Proc_F + 2
        BEF_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
        BEF_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
    Else
        If sts <> BtErrEOF And sts <> BtErrKeyNotFound Then
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Unload Me
        End If
    End If
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Unload Me
    End If

'データ処理判定
    Select Case Proc_F
        Case ZERO, 3                   '「外部＝無し，内部＝無し」「外部＝有り，内部＝有り」
            If Upd_Item() Then                                      '品目マスタ更新
                GoTo Abort_Tran
            End If
            
            If StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "1" Then
                If NyukaY_Put() Then        '入荷予定登録
                    GoTo Abort_Tran
                End If
            
                In_Cnt = In_Cnt + 1
                                
                Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
            
            Else
                If SyukaY_Put() Then        '出荷予定登録
                    GoTo Abort_Tran
                End If
                
                In_Cnt = In_Cnt + 1
                                
                Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
            
            End If

        Case 1                      '「外部＝有り，内部＝無し」→内部品番変更
            If Upd_Item() Then              '品目マスタ更新
                GoTo Abort_Tran
            End If
            
            If StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "1" Then
                If NyukaY_Put() Then        '入荷予定登録
                    GoTo Abort_Tran
                End If
                
                In_Cnt = In_Cnt + 1
                                
                Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
            
            Else
                If SyukaY_Put() Then        '出荷予定登録
                    GoTo Abort_Tran
                End If
                
                In_Cnt = In_Cnt + 1
                                
                Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
            
            End If
    
            Call PDat_Hin_Add("0")          '品番変更ﾘｽﾄﾃﾞｰﾀ保存（内部品番変更）

        Case Else                   '「外部＝無し，内部＝有り」→外部品番変更
            sts = Hin_Chg_Chk()            '外部品番変更　可否ﾁｪｯｸ
            Select Case sts
                Case False
                    If StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "1" Then      '入出荷予定登録
                        If NyukaY_Put() Then    '入荷予定登録
                            GoTo Abort_Tran
                        End If
                        
                        In_Cnt = In_Cnt + 1
                                
                        Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
                    
                    Else
                        If SyukaY_Put() Then    '出荷予定登録
                            GoTo Abort_Tran
                        End If
                        
                        In_Cnt = In_Cnt + 1
                                
                        Label3(1).Caption = StrConv(Format(In_Cnt, "#0"), vbWide)
                    
                    End If
                
                    Call PDat_Hin_Add("1")      '品番変更ﾘｽﾄﾃﾞｰﾀ保存（外部品番変更）
                Case True
                                
                    If CHin_Put() Then            '外部品番変更保留ﾃﾞｰﾀ作成
                        GoTo Abort_Tran
                    End If
                
                    Call PDat_Hin_Add("2")      '品番変更ﾘｽﾄﾃﾞｰﾀ保存（在庫有！品番変更不可）
                Case Else
                    GoTo Abort_Tran
            End Select
    End Select
            
    If Seq_Update() Then            'ﾃﾞｰﾀ重複ﾁｪｯｸ用ﾃﾞｰﾀ　更新
        GoTo Abort_Tran
    End If
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If

    If StrConv(XX_SIJREC.PM_KBN, vbUnicode) <> "-" And _
        StrConv(XX_SIJREC.IO_KBN, vbUnicode) = "1" Then
        Call PDat_Etc_Add("1", " ") '赤伝・出荷確認・訂正印刷ﾃﾞｰﾀ保存
    End If
    
    Exit Sub

Abort_Tran:
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Unload Me

End Sub
                                            '｢ｴﾗｰﾘｽﾄ｣｢確認ﾘｽﾄ｣印刷データ保存(→ List Box)
                                            '　引数　ﾘｽﾄ区分：０＝ｴﾗｰﾘｽﾄ　１＝赤伝、訂正、出荷確認
Private Sub PDat_Etc_Add(List_Kbn As String, Err_Kbn As String)
Dim sts As Integer
Dim Work As String

    Work = List_Kbn                                             'ﾘｽﾄ区分
    Work = Work & StrConv(XX_SIJREC.JGYOBU, vbUnicode)          '事業部区分
    Work = Work & StrConv(XX_SIJREC.HIN_GAI, vbUnicode)         '品番（外部）
    Work = Work & StrConv(XX_SIJREC.HIN_NAI, vbUnicode)         '品番（内部）
    Work = Work & StrConv(XX_SIJREC.DEN_DT, vbUnicode)          '伝票日付
    Work = Work & StrConv(XX_SIJREC.IO_KBN, vbUnicode)          '入出庫区分
    Work = Work & StrConv(XX_SIJREC.DEN_NO, vbUnicode)          '伝票№
    Work = Work & StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)       '数量
    Work = Work & StrConv(XX_SIJREC.PM_KBN, vbUnicode)          '赤黒区分
    Work = Work & StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)       '倉庫区分（ﾎｽﾄ）
    Work = Work & StrConv(XX_SIJREC.CYU_KBN, vbUnicode)         '注文区分
    Work = Work & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)        '直送区分
    Work = Work & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)      '予算単位（元）
    Work = Work & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)        '予算単位（先）
    Work = Work & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)       '棚番（ﾎｽﾄ）
'   Work = Work & StrConv(XX_SIJREC.SYUK_CODE, vbUnicode)       '支給先／出荷先
'   Work = Work & StrConv(XX_SIJREC.SYUK_NAME, vbUnicode)       '支給先／出荷先名
    Work = Work & Err_Kbn                                       'エラー区分
    
'    Work = Work & StrConv(XX_SIJREC.TEXT_NO, vbUnicode)     'ﾃｷｽﾄ№
'    Work = Work & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)    '直送区分
'    Work = Work & StrConv(XX_SIJREC.PM_KBN, vbUnicode)      '赤黒区分
'    Work = Work & StrConv(XX_SIJREC.DEN_SYU, vbUnicode)     '伝票種別
'    Work = Work & StrConv(XX_SIJREC.HIN_NAME, vbUnicode)    '品名
'    Work = Work & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)  '予算単位（元）
'    Work = Work & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)    '予算単位（先）
'    Work = Work & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)   '棚番（ﾎｽﾄ）

    LBox_Etc.AddItem Work

End Sub

                                            '｢ｴﾗｰﾘｽﾄ｣｢確認ﾘｽﾄ｣印刷データ保存(→ List Box)
                                            '　引数　ﾘｽﾄ区分：０＝ｴﾗｰﾘｽﾄ　１＝赤伝、訂正、出荷確認
Private Sub PDat_DUP_Add(List_Kbn As String, Err_Kbn As String)
Dim sts As Integer
Dim Work As String

    Work = List_Kbn                                             'ﾘｽﾄ区分
    Work = Work & StrConv(XX_SIJREC.JGYOBU, vbUnicode)          '事業部区分
    Work = Work & StrConv(XX_SIJREC.HIN_GAI, vbUnicode)         '品番（外部）
    Work = Work & StrConv(XX_SIJREC.HIN_NAI, vbUnicode)         '品番（内部）
    Work = Work & StrConv(XX_SIJREC.DEN_DT, vbUnicode)          '伝票日付
    Work = Work & StrConv(XX_SIJREC.IO_KBN, vbUnicode)          '入出庫区分
    Work = Work & StrConv(XX_SIJREC.DEN_NO, vbUnicode)          '伝票№
    Work = Work & StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)       '数量
    Work = Work & StrConv(XX_SIJREC.PM_KBN, vbUnicode)          '赤黒区分
    Work = Work & StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)       '倉庫区分（ﾎｽﾄ）
    Work = Work & StrConv(XX_SIJREC.CYU_KBN, vbUnicode)         '注文区分
    Work = Work & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)        '直送区分
    Work = Work & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)      '予算単位（元）
    Work = Work & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)        '予算単位（先）
    Work = Work & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)       '棚番（ﾎｽﾄ）
'   Work = Work & StrConv(XX_SIJREC.SYUK_CODE, vbUnicode)       '支給先／出荷先
'   Work = Work & StrConv(XX_SIJREC.SYUK_NAME, vbUnicode)       '支給先／出荷先名
    Work = Work & Err_Kbn                                       'エラー区分
    
'    Work = Work & StrConv(XX_SIJREC.TEXT_NO, vbUnicode)     'ﾃｷｽﾄ№
'    Work = Work & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)    '直送区分
'    Work = Work & StrConv(XX_SIJREC.PM_KBN, vbUnicode)      '赤黒区分
'    Work = Work & StrConv(XX_SIJREC.DEN_SYU, vbUnicode)     '伝票種別
'    Work = Work & StrConv(XX_SIJREC.HIN_NAME, vbUnicode)    '品名
'    Work = Work & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)  '予算単位（元）
'    Work = Work & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)    '予算単位（先）
'    Work = Work & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)   '棚番（ﾎｽﾄ）

    LBox_Dup.AddItem Work

End Sub

                                            '｢品番変更ﾘｽﾄ｣印刷データ保存(→ List Box)
                                            '　引数　変更区分：０＝内部品番変更　１＝外部品番変更
Private Sub PDat_Hin_Add(Chg_Kbn As String)
Dim sts As Integer
Dim Work As String

    Work = StrConv(XX_SIJREC.JGYOBU, vbUnicode)                 '事業部区分
    Work = Work & StrConv(XX_SIJREC.TEXT_NO, vbUnicode)         'ﾃｷｽﾄ№
    Work = Work & StrConv(XX_SIJREC.HIN_GAI, vbUnicode)         '品番（外部）
    Work = Work & BEF_GAI
    Work = Work & StrConv(XX_SIJREC.HIN_NAI, vbUnicode)         '品番（内部）
    Work = Work & BEF_NAI
    Work = Work & StrConv(XX_SIJREC.DEN_DT, vbUnicode)          '伝票日付
    Work = Work & StrConv(XX_SIJREC.IO_KBN, vbUnicode)          '入出庫区分
    Work = Work & StrConv(XX_SIJREC.DEN_NO, vbUnicode)          '伝票№
    Work = Work & StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)       '数量
    Work = Work & StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)       '倉庫区分（ﾎｽﾄ）
    Work = Work & StrConv(XX_SIJREC.CYU_KBN, vbUnicode)         '注文区分
    Work = Work & StrConv(XX_SIJREC.SYUK_CODE, vbUnicode)       '支給先／出荷先
    Work = Work & Chg_Kbn                                       '変更区分
    Work = Work & HS_NaiG                                       '国内外
    
'    Work = Work & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)    '直送区分
'    Work = Work & StrConv(XX_SIJREC.PM_KBN, vbUnicode)      '赤黒区分
'    Work = Work & StrConv(XX_SIJREC.DEN_SYU, vbUnicode)     '伝票種別
'    Work = Work & StrConv(XX_SIJREC.HIN_NAME, vbUnicode)    '品名
'    Work = Work & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)  '予算単位（元）
'    Work = Work & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)    '予算単位（先）
'    Work = Work & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)   '棚番（ﾎｽﾄ）
'    Work = Work & StrConv(XX_SIJREC.SYUK_NAME, vbUnicode)   '支給先／出荷先名

    LBox_Hin.AddItem Work

End Sub
                                            '外部品番変更　可否チェック
Private Function Hin_Chg_Chk() As Integer

Dim sts         As Integer
Dim Work        As String
Dim ZAIKO_QTY   As Long
Dim ans         As Integer

    Hin_Chg_Chk = False

    Call UniCode_Conv(K3_ITEM.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K3_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K3_ITEM.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K3_ITEM, Len(K3_ITEM), 3)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Hin_Chg_Chk = SYS_ERR
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Hin_Chg_Chk = SYS_ERR
                Exit Function
        End Select
    Loop

'在庫有無チェック
    If Zaiko_Syukei_Proc(ZAIKO_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Unload Me
    End If
    If ZAIKO_QTY <> ZERO Then
        Hin_Chg_Chk = True
        Exit Function
    End If

'品目マスタ　外部品番変更
    Do
    
        sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K3_ITEM, Len(K3_ITEM), 3)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Hin_Chg_Chk = SYS_ERR
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "品目マスタ")
                Hin_Chg_Chk = SYS_ERR
                Exit Function
        End Select
    
    Loop

    Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
    
    Do
        sts = BTRV(BtOpInsert, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
        
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Hin_Chg_Chk = SYS_ERR
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "品目マスタ")
                Hin_Chg_Chk = SYS_ERR
                Exit Function
        End Select
    
    Loop
End Function
                                            '品目マスタ更新
Private Function Upd_Item() As Boolean
Dim sts As Integer
Dim ans As Integer
Dim Command As Integer
Dim Work As String

    Upd_Item = True

    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
            Case BtNoErr
                Command = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                Command = BtOpInsert
                Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(ITEMREC.NAIGAI, HS_NaiG)
                Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")
                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")
                
                Call UniCode_Conv(ITEMREC.LOCK_F, "0")          '排他フラグ
                Call UniCode_Conv(ITEMREC.WEL_ID, "")           '使用中子機ＩＤ
                Call UniCode_Conv(ITEMREC.PRG_ID, "")           '使用中プログラム
                
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "0000000")
                
                Call UniCode_Conv(ITEMREC.FILLER, "")
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Exit Function
        End Select
    Loop
                                '品番（内部）(≠空白の時のみセット)
    If StrConv(XX_SIJREC.HIN_NAI, vbUnicode) <> Space(13) Then
        Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
    End If
                                
Debug.Print StrConv(ITEMREC.HIN_NAI, vbUnicode)
                                '品名（≠空白の時のみセット）
    If StrConv(XX_SIJREC.HIN_NAME, vbUnicode) <> Space(25) Then
        Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(XX_SIJREC.HIN_NAME, vbUnicode))
    End If
                                '備考：ﾎｽﾄ倉庫区分
                                '　　［倉庫区分の読替え　設定条件]
                                '　　　①受信データの倉庫区分≠空白
                                '　　　②　　〃　　　倉庫区分＞品目ﾏｽﾀの倉庫区分
    If StrConv(XX_SIJREC.HOST_SOKO, vbUnicode) <> Space(2) And _
       StrConv(XX_SIJREC.HOST_SOKO, vbUnicode) > StrConv(ITEMREC.BIKOU_SOKO, vbUnicode) Then
        Call UniCode_Conv(ITEMREC.BIKOU_SOKO, StrConv(XX_SIJREC.HOST_SOKO, vbUnicode))
    End If

                                '備考：ﾎｽﾄ棚番（≠空白の時）
    If StrConv(XX_SIJREC.HOST_TANA, vbUnicode) <> Space(8) Then
        Call UniCode_Conv(ITEMREC.BIKOU_TANA, StrConv(XX_SIJREC.HOST_TANA, vbUnicode))
    End If

    
    Do
        sts = BTRV(Command, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
            Case BtNoErr, BtErrEOF, BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, Command, "品目マスタ")
                Exit Function
        End Select
    Loop

    Upd_Item = False

End Function
                                            '入荷予定作成 ＆ 入荷更新
Private Function NyukaY_Put() As Boolean
Dim sts As Integer
Dim Work As String
Dim com As Integer
Dim W_Qty As Long
Dim W_Y_Qty As Long         '入荷予定数
Dim W_E_Qty As Long         '前借り入荷数
Dim W_Date As String        '処理日付

Dim ans     As Integer

    NyukaY_Put = True
    Call UniCode_Conv(K4_Y_NYU.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K4_Y_NYU.TEXT_NO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
    Select Case sts
        Case BtNoErr
            Work = "Y_NYUKA.DAT DUP ""事業部=" & StrConv(XX_SIJREC.JGYOBU, vbUnicode) & "TEXTNo=" & StrConv(XX_SIJREC.TEXT_NO, vbUnicode)
            Work = "伝票日付=" & StrConv(XX_SIJREC.DEN_DT, vbUnicode) & "伝票№=" & StrConv(XX_SIJREC.DEN_NO, vbUnicode)
            Work = "品番=" & StrConv(XX_SIJREC.HIN_GAI, vbUnicode)
            Call Log_Out(LOG_F, Work)
            NyukaY_Put = False
            Exit Function
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "入荷予定")
            Exit Function
    End Select
'////////////////////////////////////////////////////////　入荷データ排除ロジック
'
'
'                 （　建　設　予　定　地　）
'1997.08.22
    If RTrim(StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)) = "555" And _
        RTrim(StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)) = "0" Then
        Call UniCode_Conv(XX_SIJREC.CYOK_KBN, "*")          '排除データは「直送区分」に「＊」をセット
    End If

    If RTrim(StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)) = "111" And _
        RTrim(StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)) = "5" Then
        Call UniCode_Conv(XX_SIJREC.CYOK_KBN, "*")          '排除データは「直送区分」に「＊」をセット
    End If

    If RTrim(StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)) = "555" And _
        RTrim(StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)) = "1" Then
        Call UniCode_Conv(XX_SIJREC.CYOK_KBN, "*")          '排除データは「直送区分」に「＊」をセット
    End If

'
'
'        Call UniCode_Conv(XX_SIJREC.CYOK_KBN, "*")          '排除データは「直送区分」に「＊」をセット
'
'
'////////////////////////////////////////////////////////


    W_Date = Format(Date, "yyyymmdd")

'入荷予定作成
                                '完了区分
    Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_ON)
                                'データ種別
    Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                                '予定数量
    Call UniCode_Conv(Y_NYUREC.YOTEI_QTY, Format(CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)), "00000000"))
                                '確定数量
    Call UniCode_Conv(Y_NYUREC.FIX_QTY, "00000000")
                                '国内外
    Call UniCode_Conv(Y_NYUREC.NAIGAI, HS_NaiG)
                                '事業部区分
    Call UniCode_Conv(Y_NYUREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                                'ﾃｷｽﾄ№
    If StrConv(XX_SIJREC.CYOK_KBN, vbUnicode) = "*" Then        '排除データ？
                                '直送区分
        Call UniCode_Conv(Y_NYUREC.CYOK_KBN, "C")
    Else
                                '直送区分
        Call UniCode_Conv(Y_NYUREC.CYOK_KBN, StrConv(XX_SIJREC.CYOK_KBN, vbUnicode))
    End If
    
    Call UniCode_Conv(Y_NYUREC.TEXT_NO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))
                                '伝票日付
    Call UniCode_Conv(Y_NYUREC.DEN_DT, StrConv(XX_SIJREC.DEN_DT, vbUnicode))
                                '入出庫区分
    Call UniCode_Conv(Y_NYUREC.IO_KBN, StrConv(XX_SIJREC.IO_KBN, vbUnicode))
                                '赤黒区分
    Call UniCode_Conv(Y_NYUREC.PM_KBN, StrConv(XX_SIJREC.PM_KBN, vbUnicode))
                                '伝票種別
    
    Call UniCode_Conv(Y_NYUREC.DEN_SYU, StrConv(XX_SIJREC.DEN_SYU, vbUnicode))
                                '伝票№
    Call UniCode_Conv(Y_NYUREC.DEN_NO, StrConv(XX_SIJREC.DEN_NO, vbUnicode))
                                '注文区分
    Call UniCode_Conv(Y_NYUREC.CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
                                '品番（外部）
    Call UniCode_Conv(Y_NYUREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
                                '品番（内部）
    Call UniCode_Conv(Y_NYUREC.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
                                '品名
    Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(XX_SIJREC.HIN_NAME, vbUnicode))
                                '予算単位（元）
    Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode))
                                '予算単位（先）
    Call UniCode_Conv(Y_NYUREC.YOSAN_TO, StrConv(XX_SIJREC.YOSAN_TO, vbUnicode))
                                '倉庫区分（ﾎｽﾄ）
    Call UniCode_Conv(Y_NYUREC.HOST_SOKO, StrConv(XX_SIJREC.HOST_SOKO, vbUnicode))
                                '棚番（ﾎｽﾄ）←　標準入庫棚番（品目ﾏｽﾀ）
    Work = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
           StrConv(ITEMREC.ST_RETU, vbUnicode) & _
           StrConv(ITEMREC.ST_REN, vbUnicode) & _
           StrConv(ITEMREC.ST_DAN, vbUnicode)
    Call UniCode_Conv(Y_NYUREC.HOST_TANA, Work)
                                '支給先／出荷先
    Call UniCode_Conv(Y_NYUREC.SYUK_CODE, StrConv(XX_SIJREC.SYUK_CODE, vbUnicode))
                                '支給先／出荷先名
    Call UniCode_Conv(Y_NYUREC.SYUK_NAME, StrConv(XX_SIJREC.SYUK_NAME, vbUnicode))
                                '先行入荷数
    Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                                '完了日付
    Call UniCode_Conv(Y_NYUREC.KAN_DT, W_Date)
                                'FILLER
    Call UniCode_Conv(Y_NYUREC.FILLER, "")


'排除分の入荷データは、入荷予定登録のみ
    If StrConv(XX_SIJREC.CYOK_KBN, vbUnicode) = "*" Then        '排除対象データは入荷更新無し
        Do
            sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "入荷予定")
                    Exit Function
            End Select
        Loop
'直送排除分正常終了
        NyukaY_Put = False
        Exit Function
    End If
    
    If Shori_Mode = ZERO Then       '再取り込み指示時は、処理しない 01.05.03 **
        NyukaY_Put = False
        Exit Function
    End If

    W_Y_Qty = CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode))
    Last_Proc_F = True              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有り

'入荷ﾁｪｯｸﾃﾞｰﾀ更新
    Call UniCode_Conv(K0_J_NYU.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_J_NYU.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K0_J_NYU.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))

    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
        Select Case sts
            Case BtNoErr
                If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > W_Y_Qty Then
                    W_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - W_Y_Qty
                    Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(W_Qty, "00000000"))
                    
                    Do
                        sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "入荷ﾁｪｯｸﾃﾞｰﾀ")
                                Exit Function
                        End Select
                    
                    Loop
                    W_E_Qty = CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode))
                Else
                    Do
                        sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "入荷ﾁｪｯｸﾃﾞｰﾀ")
                                Exit Function
                        End Select
                    Loop
                    W_E_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
                End If
                
                Exit Do
            Case BtErrKeyNotFound
                W_E_Qty = ZERO
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "入荷ﾁｪｯｸﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop

'入荷予定データ追加（入荷分）
                                '先行入荷数（入荷実績数）
    Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(W_E_Qty, "00000000"))
    Do
        sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "入荷予定")
                Exit Function
        End Select
    Loop

'入荷数で在庫データ更新（＋）
    If Nyuko_Update_Proc(StrConv(XX_SIJREC.JGYOBU, vbUnicode), _
                            HS_NaiG, _
                            StrConv(XX_SIJREC.HIN_GAI, vbUnicode), _
                            StrConv(XX_SIJREC.DEN_DT, vbUnicode), _
                            (KASO_NYUKA_Soko & "01" & "01" & "01"), _
                            YOIN_TU_NYUKA, _
                            CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)), _
                            WS_NO) Then
        Exit Function
    
    End If


'前借り数で在庫データ更新（－）
    If W_E_Qty <> ZERO Then
'在庫データLOCK
        If Zaiko_Lock_Proc((KASO_NYUKA_Soko & "01" & "01" & "01"), _
                            StrConv(XX_SIJREC.JGYOBU, vbUnicode), _
                            HS_NaiG, _
                            StrConv(XX_SIJREC.HIN_GAI, vbUnicode), _
                            WS_NO) Then
            Exit Function

        End If
        
        
        If Syuko_Update_Proc(StrConv(XX_SIJREC.JGYOBU, vbUnicode), _
                                HS_NaiG, _
                                StrConv(XX_SIJREC.HIN_GAI, vbUnicode), _
                                StrConv(XX_SIJREC.DEN_DT, vbUnicode), _
                                (KASO_NYUKA_Soko & "01" & "01" & "01"), _
                                YOIN_MAE_SOUSAI, _
                                W_E_Qty, _
                                WS_NO) Then
            Exit Function
    
        End If

'在庫データUNLOCK
        If Zaiko_UNLock_Proc((KASO_NYUKA_Soko & "01" & "01" & "01"), _
                                StrConv(XX_SIJREC.JGYOBU, vbUnicode), _
                                HS_NaiG, _
                                StrConv(XX_SIJREC.HIN_GAI, vbUnicode), _
                            "") Then
            Exit Function
        End If
    End If
    
    NyukaY_Put = False

End Function
                                            '出荷予定作成
Private Function SyukaY_Put() As Boolean
Dim sts     As Integer
Dim Work    As String
Dim Command As Integer
                     
Dim ans     As Integer
    
    SyukaY_Put = True
                                '完了区分
    Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_SOFF_POFF_KOFF)
                                'データ種別
    Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                                '予定数量
    Call UniCode_Conv(Y_SYUREC.YOTEI_QTY, Format(CLng(StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)), "00000000"))
                                '確定数量
    Call UniCode_Conv(Y_SYUREC.FIX_QTY, "00000000")
                                '国内外
    Call UniCode_Conv(Y_SYUREC.NAIGAI, HS_NaiG)
                                '事業部区分
    Call UniCode_Conv(Y_SYUREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                                'ﾃｷｽﾄ№
    Call UniCode_Conv(Y_SYUREC.TEXT_NO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))
                                '直送区分
    Call UniCode_Conv(Y_SYUREC.CYOK_KBN, StrConv(XX_SIJREC.CYOK_KBN, vbUnicode))
                                '伝票日付
    Call UniCode_Conv(Y_SYUREC.DEN_DT, StrConv(XX_SIJREC.DEN_DT, vbUnicode))
                                '入出庫区分
    Call UniCode_Conv(Y_SYUREC.IO_KBN, StrConv(XX_SIJREC.IO_KBN, vbUnicode))
                                '赤黒区分
    Call UniCode_Conv(Y_SYUREC.PM_KBN, StrConv(XX_SIJREC.PM_KBN, vbUnicode))
                                '伝票種別
    Call UniCode_Conv(Y_SYUREC.DEN_SYU, StrConv(XX_SIJREC.DEN_SYU, vbUnicode))
                                '伝票№
    Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(XX_SIJREC.DEN_NO, vbUnicode))
                                '注文区分（補充とｽﾎﾟｯﾄは「補充・ｽﾎﾟｯﾄ」区分に置換え）
                                            '特売りも「補充・ｽﾎﾟｯﾄ」区分に置換え
    If StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_SPO Or _
       StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_HJU Or _
       StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_TOK Then
        Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_HSP)
    Else
        Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
    End If
                                '品番（外部）
    Call UniCode_Conv(Y_SYUREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
                                '品番（内部）
    Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
                                '品名
    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(XX_SIJREC.HIN_NAME, vbUnicode))
                                '予算単位（元）
    Call UniCode_Conv(Y_SYUREC.YOSAN_FROM, StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode))
                                '予算単位（先）
    Call UniCode_Conv(Y_SYUREC.YOSAN_TO, StrConv(XX_SIJREC.YOSAN_TO, vbUnicode))
                                '倉庫区分（ﾎｽﾄ）
    Call UniCode_Conv(Y_SYUREC.HOST_SOKO, StrConv(XX_SIJREC.HOST_SOKO, vbUnicode))
                                '棚番（ﾎｽﾄ）←　標準入庫棚番（品目ﾏｽﾀ）
    Work = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
           StrConv(ITEMREC.ST_RETU, vbUnicode) & _
           StrConv(ITEMREC.ST_REN, vbUnicode) & _
           StrConv(ITEMREC.ST_DAN, vbUnicode)
    Call UniCode_Conv(Y_SYUREC.HOST_TANA, Work)
                                '支給先／出荷先
    Call UniCode_Conv(Y_SYUREC.SYUK_CODE, StrConv(XX_SIJREC.SYUK_CODE, vbUnicode))
                                '支給先／出荷先名
    Call UniCode_Conv(Y_SYUREC.SYUK_NAME, StrConv(XX_SIJREC.SYUK_NAME, vbUnicode))
                                '完了日付
    Call UniCode_Conv(Y_SYUREC.KAN_DT, "")
                                '注文区分（ﾎｽﾄ）
    If StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_TOK Then
                                '特売りはスポットに置き換える2001.06.21
        Call UniCode_Conv(Y_SYUREC.HS_CYU_KBN, CYU_KBN_SPO)
    Else
        Call UniCode_Conv(Y_SYUREC.HS_CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
    End If
                                'ＳＳ追番
    Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
                                '使用子機ＩＤ
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                                '使用中プログラム
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                                '検品日付
    Call UniCode_Conv(Y_SYUREC.KENPIN_DT, "")
                                
                                
                                '向け先ｺｰﾄﾞ
    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(MTSREC.MUKE_CODE, vbUnicode))
                                '向け先読替えｺｰﾄﾞ
    Call UniCode_Conv(Y_SYUREC.MUKE_CHG_CD, StrConv(MTSREC.MUKE_CHG_CD, vbUnicode))
                                '特売りマーク   2001.06.21
    If StrConv(XX_SIJREC.CYU_KBN, vbUnicode) = CYU_KBN_TOK Then
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "1")
    Else
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, " ")
    End If
                                'FILLER
    Call UniCode_Conv(Y_SYUREC.FILLER, "")
    
    Do
        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), ZERO)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "出荷予定")
                Exit Function
        End Select
    Loop
    
    If SYUKA_LOG_ON Then
        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
    End If
        
    SyukaY_Put = False

End Function
                                            '外部品番変更保留データ作成
Private Function CHin_Put() As Integer
           
    CHin_Put = True
                                'ﾃｷｽﾄ№
    Call UniCode_Conv(CHGHREC.TEXT_NO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))
                                '事業部区分
    Call UniCode_Conv(CHGHREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                                '直送区分
    Call UniCode_Conv(CHGHREC.CYOK_KBN, StrConv(XX_SIJREC.CYOK_KBN, vbUnicode))
                                '伝票日付
    Call UniCode_Conv(CHGHREC.DEN_DT, StrConv(XX_SIJREC.DEN_DT, vbUnicode))
                                '入出庫区分
    Call UniCode_Conv(CHGHREC.IO_KBN, StrConv(XX_SIJREC.IO_KBN, vbUnicode))
                                '赤黒区分
    Call UniCode_Conv(CHGHREC.PM_KBN, StrConv(XX_SIJREC.PM_KBN, vbUnicode))
                                '伝票種別
    Call UniCode_Conv(CHGHREC.DEN_SYU, StrConv(XX_SIJREC.DEN_SYU, vbUnicode))
                                '伝票№
    Call UniCode_Conv(CHGHREC.DEN_NO, StrConv(XX_SIJREC.DEN_NO, vbUnicode))
                                '注文区分
    Call UniCode_Conv(CHGHREC.CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
                                '品番（外部）
    Call UniCode_Conv(CHGHREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
                                '品番（内部）
    Call UniCode_Conv(CHGHREC.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
                                '品名
    Call UniCode_Conv(CHGHREC.HIN_NAME, StrConv(XX_SIJREC.HIN_NAME, vbUnicode))
                                '数量
    Call UniCode_Conv(CHGHREC.YOTEI_QTY, StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode))
                                '予算単位（元）
    Call UniCode_Conv(CHGHREC.YOSAN_FROM, StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode))
                                '予算単位（先）
    Call UniCode_Conv(CHGHREC.YOSAN_TO, StrConv(XX_SIJREC.YOSAN_TO, vbUnicode))
                                '倉庫区分（ﾎｽﾄ）
    Call UniCode_Conv(CHGHREC.HOST_SOKO, StrConv(XX_SIJREC.HOST_SOKO, vbUnicode))
                                '棚番（ﾎｽﾄ）
    Call UniCode_Conv(CHGHREC.HOST_TANA, StrConv(XX_SIJREC.HOST_TANA, vbUnicode))
                                '支給先／出荷先
    Call UniCode_Conv(CHGHREC.SYUK_CODE, StrConv(XX_SIJREC.SYUK_CODE, vbUnicode))
                                '支給先／出荷先名
    Call UniCode_Conv(CHGHREC.SYUK_NAME, StrConv(XX_SIJREC.SYUK_NAME, vbUnicode))
                                'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
                                '　再処理時の重複ﾁｪｯｸ回避の為、品番変更データの終端マークには
                                '　「？」をセットする（重複ﾁｪｯｸでは「終端＝？」の時ﾁｪｯｸしない）
    Call UniCode_Conv(CHGHREC.REC_END, "?")
                                'CR & LF
    Call UniCode_Conv(CHGHREC.CR, Chr(13))
    Call UniCode_Conv(CHGHREC.LF, Chr(10))

    If CHGH_Put(ZERO) Then         '外部品番変更保留データ書込み
        Exit Function
    End If

    CHin_Put = False

End Function
                                            '重複出荷予定データ作成
Private Sub MAKE_SYUDUP_Put()
                                'ﾃｷｽﾄ№
    Call UniCode_Conv(SYUDUPREC.TEXT_NO, StrConv(XX_SIJREC.TEXT_NO, vbUnicode))
                                '事業部区分
    Call UniCode_Conv(SYUDUPREC.JGYOBU, StrConv(XX_SIJREC.JGYOBU, vbUnicode))
                                '直送区分
    Call UniCode_Conv(SYUDUPREC.CYOK_KBN, StrConv(XX_SIJREC.CYOK_KBN, vbUnicode))
                                '伝票日付
    Call UniCode_Conv(SYUDUPREC.DEN_DT, StrConv(XX_SIJREC.DEN_DT, vbUnicode))
                                '入出庫区分
    Call UniCode_Conv(SYUDUPREC.IO_KBN, StrConv(XX_SIJREC.IO_KBN, vbUnicode))
                                '赤黒区分
    Call UniCode_Conv(SYUDUPREC.PM_KBN, StrConv(XX_SIJREC.PM_KBN, vbUnicode))
                                '伝票種別
    Call UniCode_Conv(SYUDUPREC.DEN_SYU, StrConv(XX_SIJREC.DEN_SYU, vbUnicode))
                                '伝票№
    Call UniCode_Conv(SYUDUPREC.DEN_NO, StrConv(XX_SIJREC.DEN_NO, vbUnicode))
                                '注文区分
    Call UniCode_Conv(SYUDUPREC.CYU_KBN, StrConv(XX_SIJREC.CYU_KBN, vbUnicode))
                                '品番（外部）
    Call UniCode_Conv(SYUDUPREC.HIN_GAI, StrConv(XX_SIJREC.HIN_GAI, vbUnicode))
                                '品番（内部）
    Call UniCode_Conv(SYUDUPREC.HIN_NAI, StrConv(XX_SIJREC.HIN_NAI, vbUnicode))
                                '品名
    Call UniCode_Conv(SYUDUPREC.HIN_NAME, StrConv(XX_SIJREC.HIN_NAME, vbUnicode))
                                '数量
    Call UniCode_Conv(SYUDUPREC.YOTEI_QTY, StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode))
                                '予算単位（元）
    Call UniCode_Conv(SYUDUPREC.YOSAN_FROM, StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode))
                                '予算単位（先）
    Call UniCode_Conv(SYUDUPREC.YOSAN_TO, StrConv(XX_SIJREC.YOSAN_TO, vbUnicode))
                                '倉庫区分（ﾎｽﾄ）
    Call UniCode_Conv(SYUDUPREC.HOST_SOKO, StrConv(XX_SIJREC.HOST_SOKO, vbUnicode))
                                '棚番（ﾎｽﾄ）
    Call UniCode_Conv(SYUDUPREC.HOST_TANA, StrConv(XX_SIJREC.HOST_TANA, vbUnicode))
                                '支給先／出荷先
    Call UniCode_Conv(SYUDUPREC.SYUK_CODE, StrConv(XX_SIJREC.SYUK_CODE, vbUnicode))
                                '支給先／出荷先名
    Call UniCode_Conv(SYUDUPREC.SYUK_NAME, StrConv(XX_SIJREC.SYUK_NAME, vbUnicode))
                                'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
                                '　再処理時の重複ﾁｪｯｸ回避の為、品番変更データの終端マークには
                                '　「*」をセットする（重複ﾁｪｯｸでは「終端＝*」の時ﾁｪｯｸしない）
    Call UniCode_Conv(SYUDUPREC.REC_END, "*")
                                'CR & LF
    Call UniCode_Conv(SYUDUPREC.CR, Chr(13))
    Call UniCode_Conv(SYUDUPREC.LF, Chr(10))

    If SYUDUP_Put(ZERO) Then       '重複出荷予定データ書込み
        Unload Me
    End If
End Sub

                                            'ヘッダー印刷（「ｴﾗｰﾘｽﾄ」「赤伝、訂正、出荷確認一覧表」）
Private Sub P_Etc_Head(Lst_Kbn As Integer, Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    Printer.Print


    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        'ヘッダー（１）
    Printer.Print Tab(3);
    For i = ZERO To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).Code Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    Printer.Print Tab(37);
    If Lst_Kbn = ZERO Then
        Printer.Print "＊＊＊　受信エラーリスト　＊＊＊";
    Else
        Printer.Print "＊＊ 赤伝、出庫、訂正データ確認リスト　＊＊";
    End If
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        '明細ヘッダ印刷
    Printer.Print Tab(MGN_L);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 14);
    Printer.Print "品番（内部）";
    Printer.Print Tab(MGN_L + 28);
    Printer.Print "伝票日付";
    Printer.Print Tab(MGN_L + 39);
    Printer.Print "入出庫区分";
    Printer.Print Tab(MGN_L + 50);
    Printer.Print "伝票№";
    Printer.Print Tab(MGN_L + 58);
    Printer.Print "入庫数";
    Printer.Print Tab(MGN_L + 68);
    Printer.Print "出庫数";
    Printer.Print Tab(MGN_L + 77);
    Printer.Print "倉庫";
    Printer.Print Tab(MGN_L + 82);
    Printer.Print "区分";
    Printer.Print Tab(MGN_L + 92);
    Printer.Print "直送";
    Printer.Print Tab(MGN_L + 98);
    Printer.Print "予算単位　　 標準棚番"
    Printer.Print

    Lcnt = 6 + MGN_U

End Sub
                                            '明細印刷（「ｴﾗｰﾘｽﾄ」「赤伝、訂正、出荷確認一覧表」）
Private Sub P_Etc_Proc(Lst_Kbn As Integer)

Dim Lcnt As Integer
Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim i As Integer
Dim sts As Integer
Dim B_Jgyobu As String

    Lcnt = 99
    B_Jgyobu = Space(1)

    For i = ZERO To LBox_Etc.ListCount - 1
        If Left(LBox_Etc.List(i), 1) = Lst_Kbn Then

            Ldata = LBox_Etc.List(i)

                                        'ヘッダーコントロール
            If Lcnt > LMAX Or _
               Mid(Ldata, 2, 1) <> B_Jgyobu Then
                Call P_Etc_Head(Left(Ldata, 1), Lcnt, Mid(Ldata, 2, 1))
                B_Jgyobu = Mid(Ldata, 2, 1)
            End If
                                        '明細印刷
            Ldata = Mid(Ldata, 3, Len(Ldata) - 2)

            Printer.Print Tab(MGN_L);
            Printer.Print ChrCut(Ldata, 13);            '品番（外部）

            Printer.Print Tab(MGN_L + 14);
            Printer.Print ChrCut(Ldata, 13);            '品番（内部）

            Printer.Print Tab(MGN_L + 28);              '伝票日付
            Printer.Print ChrCut(Ldata, 4) & "/" & ChrCut(Ldata, 2) & "/" & ChrCut(Ldata, 2);

            Printer.Print Tab(MGN_L + 40);              '入出庫区分
            wk_IO = ChrCut(Ldata, 1)
            Select Case wk_IO
                Case IO_KBN_URI
                    Printer.Print wk_IO & " " & (IO_KBN_0);
                Case IO_KBN_NYU
                    Printer.Print wk_IO & " " & (IO_KBN_1);
                Case IO_KBN_SYU
                    Printer.Print wk_IO & " " & (IO_KBN_2);
                Case IO_KBN_ZAT
                    Printer.Print wk_IO & " " & (IO_KBN_3);
                Case IO_KBN_SYU_JITU
                    Printer.Print wk_IO & " " & (IO_KBN_4);
                Case IO_KBN_HENPIN
                    Printer.Print wk_IO & " " & (IO_KBN_5);
                Case Else
                    Printer.Print wk_IO;
            End Select

            Printer.Print Tab(MGN_L + 50);
            Printer.Print ChrCut(Ldata, 6);             '伝票№

                                                        '数量
            If wk_IO = IO_KBN_NYU Or wk_IO = IO_KBN_ZAT Or wk_IO = IO_KBN_HENPIN Then
                Printer.Print Tab(MGN_L + 57);          '入庫数
            Else
                Printer.Print Tab(MGN_L + 67);          '出庫数
            End If
            sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, ChrCut(Ldata, 6), Work)
            
            Printer.Print Work;
            
            Printer.Print ChrCut(Ldata, 1);             ' 赤黒区分
            
            Printer.Print Tab(MGN_L + 78);
            Printer.Print ChrCut(Ldata, 2);             '倉庫区分（ﾎｽﾄ）

            Printer.Print Tab(MGN_L + 83);              '注文区分
            Select Case Left(Ldata, 1)
                Case CYU_KBN_TUK
                    Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_1);
                Case CYU_KBN_SPO
                    Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_2);
                Case CYU_KBN_HJU
                    Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_3);
                Case CYU_KBN_TOK
                    Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_4);
                Case CYU_KBN_BOU
                    Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_E);
                Case Else
                    Printer.Print ChrCut(Ldata, 1);
            End Select

            Printer.Print Tab(MGN_L + 92);              ' 直送区分
            Select Case Left(Ldata, 1)
                Case "*"
                    Printer.Print ChrCut(Ldata, 1) & " 直";
                                    
                Case Else
                    Printer.Print ChrCut(Ldata, 1);
            End Select
            Printer.Print Tab(MGN_L + 98);
            Printer.Print ChrCut(Ldata, 5);             '相手先(元）
            Printer.Print " ";
            Printer.Print ChrCut(Ldata, 5);             '相手先(先）

            Printer.Print Tab(MGN_L + 111);
            Printer.Print ChrCut(Ldata, 8);             '標準棚番

            'Select Case Ldata
            '    Case "0"
            '        Printer.Print Tab(MGN_L + 115);
            '        Printer.Print "データエラー";
            '    Case "1"
            '        Printer.Print Tab(MGN_L + 115);
            '        Printer.Print "出荷先未登録";
            '    Case Else
            'End Select
                                                    '1997.10.30
'            Select Case Ldata
'                Case "2"
'                    Printer.Print "  伝票重複";
'            End Select
                                                    '1997.10.30
            
            Printer.Print
            Printer.Print

            Lcnt = Lcnt + 2
        End If
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
        Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    End If

End Sub
                                            'ヘッダー印刷（「品番変更リスト」）
Private Sub P_Hin_Head(Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    
    Printer.Print Tab(3);
    For i = ZERO To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).Code Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    Printer.Print Tab(37);
    Printer.Print "＊＊＊　品番変更リスト　＊＊＊";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print
                                        '明細ヘッダ印刷
    Printer.Print "------- 品番（外部）-------";
    Printer.Print Tab(30);
    Printer.Print "------- 品番（内部）-------";
    Printer.Print
    
    Printer.Print Tab(MGN_L);
    Printer.Print "受信データ";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "マスタ";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "受信データ";
    Printer.Print Tab(MGN_L + 44);
    Printer.Print "マスタ";
    Printer.Print Tab(MGN_L + 58);
    Printer.Print "伝票日付";
    Printer.Print Tab(MGN_L + 69);
    Printer.Print "入出庫区";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "伝票№";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "入出数";
    Printer.Print Tab(MGN_L + 93);
    Printer.Print "倉";
    Printer.Print Tab(MGN_L + 96);
    Printer.Print "注文区";
    Printer.Print Tab(MGN_L + 103);
    Printer.Print "出荷先"
    Printer.Print

    Lcnt = 7 + MGN_U

End Sub
                                            '明細印刷（「品番変更リスト」）
Private Sub P_Hin_Proc()

Dim Lcnt As Integer
Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim Emsg As String
Dim Wqty As Long
Dim i As Integer
Dim sts As Integer
Dim B_Jgyobu As String

    Lcnt = 99

    For i = ZERO To LBox_Hin.ListCount - 1
        
        Ldata = LBox_Hin.List(i)

                                        'ヘッダーコントロール
        If Lcnt > LMAX Or _
           B_Jgyobu <> Left(Ldata, 1) Then
            Call P_Hin_Head(Lcnt, Left(Ldata, 1))
            B_Jgyobu = Left(Ldata, 1)
        End If

                                        '明細印刷
        Ldata = Mid(Ldata, 11, Len(Ldata) - 11)                     '事業部，ﾃｷｽﾄ№，国内外　除外

        Printer.Print Tab(MGN_L);
        Printer.Print ChrCut(Ldata, 13);                            '受信ﾃﾞｰﾀ品番（外部）
        Work = ChrCut(Ldata, 13)
        If Right(Ldata, 1) = "1" Or Right(Ldata, 1) = "2" Then      '外部品番変更？
            Printer.Print Tab(MGN_L + 15);
            Printer.Print Work;                                     'マスタ品番（外部）
        End If

        Printer.Print Tab(MGN_L + 30);
        Printer.Print ChrCut(Ldata, 13);                            '受信ﾃﾞｰﾀ品番（内部）
        Work = ChrCut(Ldata, 13)
        If Right(Ldata, 1) = "0" Then                               '内部品番変更？
            Printer.Print Tab(MGN_L + 44);
            Printer.Print Work;                                     'マスタ品番（内部）
        End If

        Printer.Print Tab(MGN_L + 58);                              '伝票日付
        Printer.Print ChrCut(Ldata, 4) & "/" & ChrCut(Ldata, 2) & "/" & ChrCut(Ldata, 2);

        Printer.Print Tab(MGN_L + 69);                              '入出庫区分
        wk_IO = ChrCut(Ldata, 1)
        Select Case wk_IO
            Case IO_KBN_URI
                Printer.Print wk_IO & " " & (IO_KBN_0);
            Case IO_KBN_NYU
                Printer.Print wk_IO & " " & (IO_KBN_1);
            Case IO_KBN_SYU
                Printer.Print wk_IO & " " & (IO_KBN_2);
            Case IO_KBN_ZAT
                Printer.Print wk_IO & " " & (IO_KBN_3);
            Case Else
                Printer.Print wk_IO;
        End Select

        Printer.Print Tab(MGN_L + 78);
        Printer.Print ChrCut(Ldata, 6);                             '伝票№

        Printer.Print Tab(MGN_L + 85);                              '入出庫数
        Wqty = CLng(ChrCut(Ldata, 6))
        
        
        sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, Format(Wqty, "00000000"), Work)
        
        Printer.Print Work;

        Printer.Print Tab(MGN_L + 93);
        Printer.Print ChrCut(Ldata, 2);                             '倉庫区分（ﾎｽﾄ）

        Printer.Print Tab(MGN_L + 96);                              '注文区分
        Select Case Left(Ldata, 1)
            Case CYU_KBN_TUK
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_1);
            Case CYU_KBN_SPO
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_2);
            Case CYU_KBN_HJU
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_3);
            Case CYU_KBN_BOU
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_E);
            Case Else
                Printer.Print ChrCut(Ldata, 1);
        End Select

        Printer.Print Tab(MGN_L + 103);
        Printer.Print ChrCut(Ldata, 5);                             '支給先／出荷先

        Printer.Print Tab(MGN_L + 110);                             '変更メッセージ
        Select Case Left(Ldata, 1)
            Case "0"
                Printer.Print "内部変更 ﾏｽﾀ品番入替";
            Case "1"
                Printer.Print "外部変更 ﾏｽﾀ品番入替";
            Case "2"
                Printer.Print "在庫有！外部変更不可";
        End Select
        Printer.Print

        Printer.Print

        Lcnt = Lcnt + 2
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
        Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    End If

End Sub
                                            '明細印刷（出荷予定重複リスト）
Private Sub P_Dup_Proc()

Dim Lcnt As Integer
Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim i As Integer
Dim sts As Integer
Dim B_Jgyobu As String

    Lcnt = 99
    B_Jgyobu = Space(1)

    For i = ZERO To LBox_Dup.ListCount - 1

        Ldata = LBox_Dup.List(i)

                                        'ヘッダーコントロール
        If Lcnt > LMAX Or _
            Mid(Ldata, 2, 1) <> B_Jgyobu Then
            Call P_Dup_Head(Lcnt, Mid(Ldata, 2, 1))
            B_Jgyobu = Mid(Ldata, 2, 1)
        End If
                                        '明細印刷
        Ldata = Mid(Ldata, 3, Len(Ldata) - 2)

        Printer.Print Tab(MGN_L);
        Printer.Print ChrCut(Ldata, 13);            '品番（外部）

        Printer.Print Tab(MGN_L + 14);
        Printer.Print ChrCut(Ldata, 13);            '品番（内部）

        Printer.Print Tab(MGN_L + 28);              '伝票日付
        Printer.Print ChrCut(Ldata, 4) & "/" & ChrCut(Ldata, 2) & "/" & ChrCut(Ldata, 2);

        Printer.Print Tab(MGN_L + 40);              '入出庫区分
        wk_IO = ChrCut(Ldata, 1)
        Select Case wk_IO
            Case IO_KBN_URI
                Printer.Print wk_IO & " " & (IO_KBN_0);
            Case IO_KBN_NYU
                Printer.Print wk_IO & " " & (IO_KBN_1);
            Case IO_KBN_SYU
                Printer.Print wk_IO & " " & (IO_KBN_2);
            Case IO_KBN_ZAT
                Printer.Print wk_IO & " " & (IO_KBN_3);
            Case IO_KBN_SYU_JITU
                Printer.Print wk_IO & " " & (IO_KBN_4);
            Case IO_KBN_HENPIN
                Printer.Print wk_IO & " " & (IO_KBN_5);
            Case Else
                Printer.Print wk_IO;
        End Select

        Printer.Print Tab(MGN_L + 50);
        Printer.Print ChrCut(Ldata, 6);             '伝票№

        Printer.Print Tab(MGN_L + 67);              '出庫数
            
        
        sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, ChrCut(Ldata, 6), Work)
        
        Printer.Print Work;
            
        Printer.Print ChrCut(Ldata, 1);             ' 赤黒区分
            
        Printer.Print Tab(MGN_L + 78);
        Printer.Print ChrCut(Ldata, 2);             '倉庫区分（ﾎｽﾄ）

        Printer.Print Tab(MGN_L + 83);              '注文区分
        Select Case Left(Ldata, 1)
            Case CYU_KBN_TUK
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_1);
            Case CYU_KBN_SPO
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_2);
            Case CYU_KBN_HJU
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_3);
            Case CYU_KBN_TOK
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_4);
            Case CYU_KBN_BOU
                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_E);
            Case Else
                Printer.Print ChrCut(Ldata, 1);
        End Select

        Printer.Print Tab(MGN_L + 92);              ' 直送区分
        Select Case Left(Ldata, 1)
            Case "*"
                Printer.Print ChrCut(Ldata, 1) & " 直";
                                    
            Case Else
                Printer.Print ChrCut(Ldata, 1);
        End Select
        Printer.Print Tab(MGN_L + 98);
        Printer.Print ChrCut(Ldata, 5);             '相手先(元）
        Printer.Print " ";
        Printer.Print ChrCut(Ldata, 5);             '相手先(先）

        Printer.Print Tab(MGN_L + 111);
        Printer.Print ChrCut(Ldata, 8);             '標準棚番

            
        Printer.Print
        Printer.Print

        Lcnt = Lcnt + 2
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
        Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    End If

End Sub
                                            'ヘッダー印刷（出荷予定重複リスト）
Private Sub P_Dup_Head(Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    Printer.Print


    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        'ヘッダー（１）
    Printer.Print Tab(3);
    For i = ZERO To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).Code Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    Printer.Print Tab(37);
    Printer.Print "＊＊ 出荷予定重複リスト　＊＊";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        '明細ヘッダ印刷
    Printer.Print Tab(MGN_L);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 14);
    Printer.Print "品番（内部）";
    Printer.Print Tab(MGN_L + 28);
    Printer.Print "伝票日付";
    Printer.Print Tab(MGN_L + 39);
    Printer.Print "入出庫区分";
    Printer.Print Tab(MGN_L + 50);
    Printer.Print "伝票№";
    Printer.Print Tab(MGN_L + 58);
    Printer.Print "入庫数";
    Printer.Print Tab(MGN_L + 68);
    Printer.Print "出庫数";
    Printer.Print Tab(MGN_L + 77);
    Printer.Print "倉庫";
    Printer.Print Tab(MGN_L + 82);
    Printer.Print "区分";
    Printer.Print Tab(MGN_L + 92);
    Printer.Print "直送";
    Printer.Print Tab(MGN_L + 98);
    Printer.Print "予算単位　　 標準棚番"
    Printer.Print

    Lcnt = 6 + MGN_U

End Sub
                                            '文字列の切出し
Private Function ChrCut(Moto As String, Leng As Long) As String
    ChrCut = Left(Moto, Leng)

    If Len(Moto) <= Leng Then
        Moto = ""
        Exit Function
    End If

    Moto = Mid(Moto, Leng + 1, Len(Moto) - Leng)
End Function
                                            'ヘッダー印刷（「過剰前借品ﾘｽﾄ」）
Private Sub P_Last_Head(Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    Printer.Print


    For i = 1 To MGN_U2
        Printer.Print
    Next i

    Printer.Print
                                        'ヘッダー（１）
    Printer.Print Tab(3);
    For i = ZERO To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).Code Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    Printer.Print Tab(39);
    Printer.Print "＊＊＊ 過剰前借品リスト　＊＊＊";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        '明細ヘッダ印刷
    Printer.Print Tab(MGN_L2 + 1);
    Printer.Print "国内外";
    Printer.Print Tab(MGN_L2 + 11);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L2 + 29);
    Printer.Print "品　名";
    Printer.Print Tab(MGN_L2 + 62);
    Printer.Print "過剰数"
    Printer.Print

    Lcnt = 6 + MGN_U2

End Sub
                                            '前借り入荷、実績残チェック
Private Sub Last_Proc()
Dim sts         As Integer
Dim Lcnt        As Integer
Dim Command     As Integer
Dim RetBuf      As String
Dim B_Jgyobu    As String

Dim ans         As Integer

    Call Input_Lock           '画面項目ロック

    MsgLab(7).Visible = True       '更新中ﾒｯｾｰｼﾞ表示
    DoEvents

'    Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    Set Printer.Font = NormalFont           '印刷フォント設定
    Lcnt = 99
    B_Jgyobu = Space(1)

'入荷ﾁｪｯｸﾃﾞｰﾀ更新
    Command = BtOpGetFirst
    Do
        Do
            sts = BTRV(Command + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
            Select Case sts
                Case BtNoErr, BtErrKeyNotFound, BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Sub
                    End If
                Case Else
                    Call File_Error(sts, Command, "入荷ﾁｪｯｸﾃﾞｰﾀ")
                    Exit Sub
            End Select
        Loop

        If sts = BtErrKeyNotFound Or sts = BtErrEOF Then
            Exit Do
        End If

                                        'ヘッダーコントロール
        If Lcnt > LMAX Or _
           StrConv(J_NYUREC.JGYOBU, vbUnicode) <> B_Jgyobu Then
            Call P_Last_Head(Lcnt, StrConv(J_NYUREC.JGYOBU, vbUnicode))
            B_Jgyobu = StrConv(J_NYUREC.JGYOBU, vbUnicode)
        End If
                                        '明細印刷
        Printer.Print Tab(MGN_L2 + 2);
        If StrConv(J_NYUREC.NAIGAI, vbUnicode) = "1" Then
            Printer.Print NAIGAI1;         '国内
        Else
            Printer.Print NAIGAI2;         '海外
        End If

        Printer.Print Tab(MGN_L2 + 11);
        Printer.Print StrConv(J_NYUREC.HIN_GAI, vbUnicode);

        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(J_NYUREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(J_NYUREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(J_NYUREC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
            Case BtNoErr
                Printer.Print Tab(MGN_L2 + 29);
                Printer.Print StrConv(ITEMREC.HIN_NAME, vbUnicode);
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Sub
        End Select

        sts = Numeric_Check(EDIT_ONLY, 10, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, StrConv(J_NYUREC.JITU_QTY, vbUnicode), RetBuf)
        Printer.Print Tab(MGN_L2 + 58);
        Printer.Print RetBuf;
        Printer.Print

        Printer.Print

        Lcnt = Lcnt + 2
''袋井前借りデータは残す
''        Do
''            sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
''            Select Case sts
''                Case BtNoErr
''                    Exit Do
''                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
''                    Beep
''                    ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
''                    If ans = vbCancel Then
''                        Exit Sub
''                    End If
''                Case Else
''                    Call File_Error(sts, BtOpDelete, "入荷ﾁｪｯｸﾃﾞｰﾀ")
''                    Exit Sub
''            End Select
''
''        Loop
        
        Command = BtOpGetNext
    Loop

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If

End Sub
Private Sub RERUN_PROC()

Dim i   As Integer
Dim sts As Integer

'強制再実行処理
    SelCmd(3).Enabled = True
    SelCmd(1).Enabled = True
    SelCmd(2).Enabled = True
    
    Label2(ZERO).Visible = True
    Text1(ZERO).Visible = True
                                                '洗濯機 SPIC
    Call UniCode_Conv(K0_SEQCK.JGYOBU, "1")
    Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "1")
        
    sts = BTRV(BtOpGetEqual, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(SEQCKREC.LAST_TXTNO, "000000000")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "予定取込みチェック")
            Unload Me
    End Select

    Text1(ZERO).Text = StrConv(SEQCKREC.LAST_TXTNO, vbUnicode)

                                                '洗濯機 特売り
    Call UniCode_Conv(K0_SEQCK.JGYOBU, "1")
    Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "2")
        
    sts = BTRV(BtOpGetEqual, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(SEQCKREC.LAST_TXTNO, "000000000")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "予定取込みチェック")
            Unload Me
    End Select

    Text1(1).Text = StrConv(SEQCKREC.LAST_TXTNO, vbUnicode)


    Text1(ZERO).SetFocus
End Sub
Private Sub Command_Click(Index As Integer)

    Select Case Index
        Case 11
            Unload Me
        Case Else
            Beep
    End Select

End Sub
Private Sub Command_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF12
            Command(11).Value = True
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
            Command(KeyCode - vbKeyF1).Value = True
        Case vbKeyZ
            If Shift = 1 Then
                Call RERUN_PROC
            End If
    End Select


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = ZERO
    End If
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer


Dim sBuffer As String * 255
Dim com     As String
    
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
                                '出荷ログファイル名取り込み
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "出荷ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                
                                'システム予約済要因取り込み
    If SYSTEM_YOIN_Set() Then
        Beep
        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                '有効ﾎｽﾄ倉庫取り込み（掃除機）
    For i = ZERO To UBound(SOKO_T1) - 1
        SOKO_T1(i).HS_SOKO = "  "
        SOKO_T1(i).NAIGAI = " "
    Next i

    i = ZERO
    Do
        If GetIni("NYUSYU_OK_SOKO", "SOKO1" & RTrim(Format(i + 1, "#0")), "SYS", c) Then
            Beep
            MsgBox "システム異常が発生しました。処理を中止して下さい。"
            Call Log_Out(LOG_F, "[SYS.INI] [NYUSYU_OK_SOKO] [SOKO] READ ERROR")
            End
        End If
        If RTrim(c) = "**" Then
            Exit Do
        End If
        SOKO_T1(i).HS_SOKO = RTrim(c)
        If GetIni("NYUSYU_OK_SOKO", "NAIG1" & RTrim(Format(i + 1, "#0")), "SYS", c) Then
            Beep
            MsgBox "システム異常が発生しました。処理を中止して下さい。"
            Call Log_Out(LOG_F, "[SYS.INI] [NYUSYU_OK_SOKO] [NAIG] READ ERROR")
            End
        End If
        SOKO_T1(i).NAIGAI = RTrim(c)
        i = i + 1
    Loop

    If Kaso_Soko_No_Set() Then
        Beep
        MsgBox "仮想倉庫の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                'ﾜｰｸｽﾃｰｼｮﾝ番号取り込み
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> ZERO Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)

                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタ（更新用ワーク）ＯＰＥＮ
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先読替マスタＯＰＥＮ　2001.07.04
    If MTSCHG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因スタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入荷予定ＯＰＥＮ
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入荷ﾁｪｯｸﾃﾞｰﾀＯＰＥＮ
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '予定取込みﾁｪｯｸＯＰＥＮ
    If SEQCK_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F1020121.FontName
        .Size = F1020121.FontSize
    End With
    Set Printer.Font = NormalFont
                                '画面初期設定
    Call Scr_Init
    
    Last_Proc_F = False         '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有無フラグクリア

    Shori_Mode = 3
    Call Data_Inport            'ホストデータ取込み処理
    
    Unload Me



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

    If Last_Proc_F = True Then              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有り？
        Call Last_Proc
    End If

                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '品目マスタ（更新用ワーク）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '向け先読替マスタＣＬＯＳＥ　2001.07.04
    sts = BTRV(BtOpClose, MTSCHG_POS, MTSCHGREC, Len(MTSCHGREC), K0_MTSCHG, Len(K0_MTSCHG), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先読替マスタ")
        End If
    End If
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            '入荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷予定")
        End If
    End If
                                            '出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
                                            '入荷ﾁｪｯｸﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷ﾁｪｯｸﾃﾞｰﾀ")
        End If
    End If
                                            '予定取込みﾁｪｯｸＣＬＯＳＥ
    sts = BTRV(BtOpClose, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "予定取込みﾁｪｯｸ")
        End If
    End If
                                            'Ｂｔｒｉｅｖｅリセット
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), ZERO)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020121 = Nothing

    End
End Sub


Private Sub SelCmd_Click(Index As Integer)

Dim ans As Integer
        
    Beep
    ans = MsgBox("「" & SelCmd(Index).Caption & "」" & "　取込み処理　実行しますか？", vbYesNo + vbQuestion, "確認入力")
    If ans = vbYes Then
        If Text1(ZERO).Visible Then
            Call SEQCHEK_PUT
        End If
        Shori_Mode = Index
        Call Data_Inport            'ホストデータ取込み処理
    End If

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1020121.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020121)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020121)

    F1020121.MousePointer = vbDefault

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = ZERO
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If


End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim i   As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
            

    For i = Index + 1 To Text_Max
        If Text1(i).Enabled And Text1(i).Visible And Text1(i).TabStop Then
            Text1(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Sub SEQCHEK_PUT()
    
Dim sts As Integer
Dim com As Integer
Dim ans As Integer
                                    '洗濯機 SPIC
    Call UniCode_Conv(K0_SEQCK.JGYOBU, "1")
    Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "1")
        
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
    
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<SEQCHK.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Unload Me
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "予定取込みチェック")
                Unload Me
        End Select

    Loop

    If com = BtOpInsert Then
        Call UniCode_Conv(SEQCKREC.JGYOBU, "1")
        Call UniCode_Conv(SEQCKREC.SEQ_MODE, "1")
    End If

    Call UniCode_Conv(SEQCKREC.LAST_TXTNO, Text1(ZERO).Text)
    Call UniCode_Conv(SEQCKREC.LAST_GET_DT, Format(Date, "yyyymmdd"))
    Call UniCode_Conv(SEQCKREC.LAST_GET_TM, Format(Time, "HHmmss"))

    Do
        sts = BTRV(com, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<SEQCHK.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Unload Me
                End If
            Case Else
                Call File_Error(sts, com, "予定取込みチェック")
                Unload Me
        End Select
    
    Loop

                                    '洗濯機 特売り
    Call UniCode_Conv(K0_SEQCK.JGYOBU, "1")
    Call UniCode_Conv(K0_SEQCK.SEQ_MODE, "2")
        
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
    
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<SEQCHK.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Unload Me
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "予定取込みチェック")
                Unload Me
        End Select

    Loop

    If com = BtOpInsert Then
        Call UniCode_Conv(SEQCKREC.JGYOBU, "1")
        Call UniCode_Conv(SEQCKREC.SEQ_MODE, "2")
    End If

    Call UniCode_Conv(SEQCKREC.LAST_TXTNO, Text1(1).Text)
    Call UniCode_Conv(SEQCKREC.LAST_GET_DT, Format(Date, "yyyymmdd"))
    Call UniCode_Conv(SEQCKREC.LAST_GET_TM, Format(Time, "HHmmss"))

    Do
        sts = BTRV(com, SEQCK_POS, SEQCKREC, Len(SEQCKREC), K0_SEQCK, Len(K0_SEQCK), ZERO)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<SEQCHK.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Unload Me
                End If
            Case Else
                Call File_Error(sts, com, "予定取込みチェック")
                Unload Me
        End Select
    
    Loop

End Sub


Private Sub Err_Log_Out(Mesg As String)
Dim Work_Rec    As String

                                'ﾃｷｽﾄ№
        Work_Rec = StrConv(XX_SIJREC.TEXT_NO, vbUnicode)
                                '事業部区分
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.JGYOBU, vbUnicode)
                                '直送区分
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.CYOK_KBN, vbUnicode)
                                '伝票日付
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.DEN_DT, vbUnicode)
                                '入出庫区分
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.IO_KBN, vbUnicode)
                                '赤黒区分
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.PM_KBN, vbUnicode)
                                '伝票種別
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.DEN_SYU, vbUnicode)
                                '伝票№
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.DEN_NO, vbUnicode)
                                '注文区分
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.CYU_KBN, vbUnicode)
                                '品番（外部）
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.HIN_GAI, vbUnicode)
                                '品番（内部）
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.HIN_NAI, vbUnicode)
                                '品名
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.HIN_NAME, vbUnicode)
                                '数量
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.YOTEI_QTY, vbUnicode)
                                '予算単位（元）
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.YOSAN_FROM, vbUnicode)
                                '予算単位（先）
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.YOSAN_TO, vbUnicode)
                                '倉庫区分（ﾎｽﾄ）
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.HOST_SOKO, vbUnicode)
                                '棚番（ﾎｽﾄ）
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.HOST_TANA, vbUnicode)
                                '支給先／出荷先
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.SYUK_CODE, vbUnicode)
                                '支給先／出荷先名
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.SYUK_NAME, vbUnicode)
                                'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
        Work_Rec = Work_Rec & StrConv(XX_SIJREC.REC_END, vbUnicode)

        Call Log_Out(LOG_F, Mesg & " " & Work_Rec)

End Sub
