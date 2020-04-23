VERSION 5.00
Begin VB.Form F1200301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入出荷実績集計処理"
   ClientHeight    =   7080
   ClientLeft      =   2325
   ClientTop       =   2910
   ClientWidth     =   11295
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
   ScaleHeight     =   7080
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   14
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4440
      Width           =   492
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   13
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4440
      Width           =   492
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   12
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4440
      Width           =   852
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   11
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1812
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   10
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1812
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   9
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1812
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   8
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   7
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   6
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   5
      Left            =   5520
      MaxLength       =   2
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   4
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   3
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command 
      Caption         =   "終　了"
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
      Top             =   5880
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
      Top             =   5880
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   8
      Left            =   7800
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   7
      Left            =   6480
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "実　行"
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "日現在"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   14
      Left            =   8400
      TabIndex        =   43
      Top             =   4560
      Width           =   1092
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "月"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   13
      Left            =   7320
      TabIndex        =   42
      Top             =   4560
      Width           =   372
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   12
      Left            =   6240
      TabIndex        =   41
      Top             =   4560
      Width           =   372
      WordWrap        =   -1  'True
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
      TabIndex        =   40
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "集計処理実行中です。"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   39
      Top             =   3840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "出荷総数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   11
      Left            =   5760
      TabIndex        =   38
      Top             =   3240
      Width           =   1812
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "出荷品目数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   10
      Left            =   1800
      TabIndex        =   37
      Top             =   3240
      Width           =   1812
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "入荷総数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   9
      Left            =   5760
      TabIndex        =   36
      Top             =   2520
      Width           =   1812
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "入荷品目数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   8
      Left            =   1800
      TabIndex        =   35
      Top             =   2520
      Width           =   1812
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "総在庫数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   7
      Left            =   5760
      TabIndex        =   34
      Top             =   1800
      Width           =   1812
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "総品目数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   6
      Left            =   1800
      TabIndex        =   33
      Top             =   1800
      Width           =   1812
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   32
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   31
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "～"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   30
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   29
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   28
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "集計年月日"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   27
      Top             =   600
      Width           =   1215
      WordWrap        =   -1  'True
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
Attribute VB_Name = "F1200301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_DT_YY% = 0               '開始集計年月日 年
Private Const ptxS_DT_MM% = 1               '開始集計年月日 月
Private Const ptxS_DT_DD% = 2               '開始集計年月日 日
Private Const ptxE_DT_YY% = 3               '終了集計年月日 年
Private Const ptxE_DT_MM% = 4               '終了集計年月日 月
Private Const ptxE_DT_DD% = 5               '終了集計年月日 日
Private Const ptxZAIKO_ITEM_SU% = 6         '総品目数
Private Const ptxNYU_ITEM_SU% = 7           '入荷品目数
Private Const ptxSYU_ITEM_SU% = 8           '出荷品目数
Private Const ptxZAIKO_SU% = 9              '総在庫数
Private Const ptxNYU_SU% = 10               '入荷総数
Private Const ptxSYU_SU% = 11               '出荷総数
Private Const ptxNOW_DT_YY% = 12            '現在日付　年
Private Const ptxNOW_DT_MM% = 13            '現在日付　月
Private Const ptxNOW_DT_DD% = 14            '現在日付　日

Private Const Text_Max% = 14                '画面項目別最大ｲﾝﾃﾞｯｸｽ
Private Function Main_Proc() As Integer
'----------------------------------------------------------------------------
'                   入出荷実績ファイル作成＆表示処理
'----------------------------------------------------------------------------
                                 
Dim c               As String * 128

Dim sts             As Integer
Dim com             As Integer
Dim Upd_Com         As Integer
Dim ans             As Integer
                                 
Dim NYU_ITEM_SU     As Integer
Dim NYU_SU          As Long
Dim SYU_ITEM_SU     As Integer
Dim SYU_SU          As Long
                                 
Dim ZAIKO_ITEM_SU   As Integer
Dim ZAIKO_SU        As Long
                                 
Dim SAVE_NAIGAI     As String * 1
Dim SAVE_HIN_GAI    As String * 13
                                 
    
    Main_Proc = True
                                 
    Call Input_Lock
                                 
                                 
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, SUMJ_POS, SUMJREC, Len(SUMJREC), K0_SUMJ, Len(K0_SUMJ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                 
                                 
                                 
                                 '入出荷実績ファイル削除
    If GetIni("FILE", SUMJ_ID, "SYS", c) Then
        Beep
        MsgBox "入出荷実績ファイル情報の獲得に失敗しました。処理を中止して下さい。"
        Call Log_Out(LOG_F, "[SYS.INI] [FILE] [SUMJ] READ ERROR")
        Exit Function
    End If
    
    On Error Resume Next
    Kill RTrim(c)
                                
                                '入出荷実績ファイルＯＰＥＮ
    If SUMJ_Open(BtOpenNomal) Then
        Exit Function
    End If
                                            '移動歴より集計
    Call UniCode_Conv(K0_IDO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxS_DT_YY).Text & Text(ptxS_DT_MM).Text & Text(ptxS_DT_DD).Text)
    Call UniCode_Conv(K0_IDO.JITU_TM, "")
    
    com = BtOpGetGreater
    
    Do
        DoEvents
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
            
                If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxE_DT_YY).Text & Text(ptxE_DT_MM).Text & Text(ptxE_DT_DD).Text) Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫移動歴")
                Exit Function
        End Select
                                '入出荷実績ファイル読込み
        Call UniCode_Conv(K0_SUMJ.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_SUMJ.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_SUMJ.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
        
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, SUMJ_POS, SUMJREC, Len(SUMJREC), K0_SUMJ, Len(K0_SUMJ), 0)
            Select Case sts
                Case BtNoErr
                    Upd_Com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    Upd_Com = BtOpInsert
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE 'ここではない！
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<SUMJITU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "入出荷実績集計ファイル")
                    Exit Function
            End Select
        Loop
        
        If Upd_Com = BtOpInsert Then
            Call UniCode_Conv(SUMJREC.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(SUMJREC.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(SUMJREC.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(SUMJREC.NYUKA_QTY, "00000000")
            Call UniCode_Conv(SUMJREC.CHOKU_QTY, "00000000")
            Call UniCode_Conv(SUMJREC.TUK_QTY, "00000000")
            Call UniCode_Conv(SUMJREC.HSP_QTY, "00000000")
            Call UniCode_Conv(SUMJREC.BOU_QTY, "00000000")
            Call UniCode_Conv(SUMJREC.KIN_QTY, "00000000")
            Call UniCode_Conv(SUMJREC.ZAI_PURA, "00000000")
            Call UniCode_Conv(SUMJREC.ZAI_MINA, "00000000")
            Call UniCode_Conv(SUMJREC.FILLER, "")
        End If
        
        Select Case Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1)
            Case ACT_ZAITEI_IN      '在訂（＋）
                If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_TU_NYUKA Then
                                    '入荷（ホスト→ＰＣ分）
                    Call UniCode_Conv(SUMJREC.NYUKA_QTY, Format(CLng(StrConv(SUMJREC.NYUKA_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                Else
                    Call UniCode_Conv(SUMJREC.ZAI_PURA, Format(CLng(StrConv(SUMJREC.ZAI_PURA, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                End If
            
            Case ACT_ZAITEI_OUT     '在訂（－）
'                If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_CHOKUSO Then
'                                    '入荷直送分
'                    Call UniCode_Conv(SUMJREC.CHOKU_QTY, Format(CLng(StrConv(SUMJREC.CHOKU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
'                Else
                    Call UniCode_Conv(SUMJREC.ZAI_MINA, Format(CLng(StrConv(SUMJREC.ZAI_MINA, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
'                End If
            
            Case ACT_SYUKA_KEI          '計画出荷
                Select Case Right(StrConv(IDOREC.RIRK_ID, vbUnicode), 1)
                    Case CYU_KBN_HSP, CYU_KBN_SPO, CYU_KBN_HJU '補／ス
                        Call UniCode_Conv(SUMJREC.HSP_QTY, Format(CLng(StrConv(SUMJREC.HSP_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                    Case CYU_KBN_TUK    '月切り
                        Call UniCode_Conv(SUMJREC.TUK_QTY, Format(CLng(StrConv(SUMJREC.TUK_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                    Case CYU_KBN_BOU    '貿易
                        Call UniCode_Conv(SUMJREC.BOU_QTY, Format(CLng(StrConv(SUMJREC.BOU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                End Select
            Case ACT_SYUKA_HYO      '計画出荷
                Select Case Right(StrConv(IDOREC.RIRK_ID, vbUnicode), 1)
                    Case CYU_KBN_HSP, CYU_KBN_SPO, CYU_KBN_HJU '補／ス
                        Call UniCode_Conv(SUMJREC.HSP_QTY, Format(CLng(StrConv(SUMJREC.HSP_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                    Case CYU_KBN_TUK    '月切り
                        Call UniCode_Conv(SUMJREC.TUK_QTY, Format(CLng(StrConv(SUMJREC.TUK_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                    Case CYU_KBN_BOU    '貿易
                        Call UniCode_Conv(SUMJREC.BOU_QTY, Format(CLng(StrConv(SUMJREC.BOU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                End Select
            Case ACT_SYUKA_GAI      '計画外出荷
                Call UniCode_Conv(SUMJREC.KIN_QTY, Format(CLng(StrConv(SUMJREC.KIN_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
        End Select
            
        Do
            sts = BTRV(Upd_Com, SUMJ_POS, SUMJREC, Len(SUMJREC), K0_SUMJ, Len(K0_SUMJ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<SUMJITU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, Upd_Com, "入出荷実績集計ファイル")
                    Exit Function
            End Select
            
        Loop
        
        com = BtOpGetNext
    
    Loop
                                            
    SAVE_NAIGAI = ""
    SAVE_HIN_GAI = ""
                                            
    ZAIKO_ITEM_SU = 0
    ZAIKO_SU = 0
                                   '在庫集計開始
    Call UniCode_Conv(K4_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K4_ZAIKO.NAIGAI, "")
    Call UniCode_Conv(K4_ZAIKO.HIN_GAI, "")
    Call UniCode_Conv(K4_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K4_ZAIKO.Retu, "")
    Call UniCode_Conv(K4_ZAIKO.Ren, "")
    Call UniCode_Conv(K4_ZAIKO.Dan, "")
    
    com = BtOpGetGreater
    
    Do
        DoEvents
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
        Select Case sts
            Case BtNoErr
                If Last_JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目別集計データ")
                Main_Proc = SYS_ERR
                Exit Function
        End Select
        
        If Len(Trim(SAVE_NAIGAI)) = 0 Then
            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
            ZAIKO_ITEM_SU = ZAIKO_ITEM_SU + 1
        End If
        
        If SAVE_NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
            SAVE_HIN_GAI <> StrConv(ZAIKOREC.HIN_GAI, vbUnicode) Then
            ZAIKO_ITEM_SU = ZAIKO_ITEM_SU + 1
                    
            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
        
        End If
        
        ZAIKO_SU = ZAIKO_SU + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                
        com = BtOpGetNext
    
    Loop

    Text(ptxZAIKO_ITEM_SU).Text = Format(ZAIKO_ITEM_SU, "#,##0")

    Text(ptxZAIKO_SU).Text = Format(ZAIKO_SU, "#,##0")
                                            '入出荷集計
    NYU_ITEM_SU = 0
    NYU_SU = 0
    SYU_ITEM_SU = 0
    SYU_SU = 0
    
    Call UniCode_Conv(K0_SUMJ.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_SUMJ.NAIGAI, "")
    Call UniCode_Conv(K0_SUMJ.HIN_GAI, "")
    
    com = BtOpGetGreater
    Do
        DoEvents
        sts = BTRV(com, SUMJ_POS, SUMJREC, Len(SUMJREC), K0_SUMJ, Len(K0_SUMJ), 0)
        Select Case sts
            Case BtNoErr
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "入出荷実績集計データ")
                Exit Function
        End Select
        
        If CLng(StrConv(SUMJREC.NYUKA_QTY, vbUnicode)) <> 0 Then
            NYU_ITEM_SU = NYU_ITEM_SU + 1
            NYU_SU = NYU_SU + CLng(StrConv(SUMJREC.NYUKA_QTY, vbUnicode))
        End If
        
        If CLng(StrConv(SUMJREC.CHOKU_QTY, vbUnicode)) <> 0 Then
            NYU_ITEM_SU = NYU_ITEM_SU - 1
            NYU_SU = NYU_SU - CLng(StrConv(SUMJREC.CHOKU_QTY, vbUnicode))
        End If
                
        
        If (CLng(StrConv(SUMJREC.HSP_QTY, vbUnicode)) <> 0 Or _
            CLng(StrConv(SUMJREC.TUK_QTY, vbUnicode)) <> 0 Or _
            CLng(StrConv(SUMJREC.BOU_QTY, vbUnicode)) <> 0 Or _
            CLng(StrConv(SUMJREC.KIN_QTY, vbUnicode)) <> 0) Then
            SYU_ITEM_SU = SYU_ITEM_SU + 1
            SYU_SU = SYU_SU + CLng(StrConv(SUMJREC.HSP_QTY, vbUnicode)) _
                            + CLng(StrConv(SUMJREC.TUK_QTY, vbUnicode)) _
                            + CLng(StrConv(SUMJREC.BOU_QTY, vbUnicode)) _
                            + CLng(StrConv(SUMJREC.KIN_QTY, vbUnicode))
        
        End If
        
        com = BtOpGetNext
    
    Loop

    Text(ptxNYU_ITEM_SU).Text = Format(NYU_ITEM_SU, "#,##0")
    Text(ptxNYU_SU).Text = Format(NYU_SU, "#,##0")

    Text(ptxSYU_ITEM_SU).Text = Format(SYU_ITEM_SU, "#,##0")
    Text(ptxSYU_SU).Text = Format(SYU_SU, "#,##0")
    
    Text(ptxNOW_DT_YY).Text = Left(Format(Date, "yyyymmdd"), 4)
    Text(ptxNOW_DT_MM).Text = Mid(Format(Date, "yyyymmdd"), 5, 2)
    Text(ptxNOW_DT_DD).Text = Right(Format(Date, "yyyymmdd"), 2)

                                            '入出荷実績集計ＣＬＯＳＥ
    sts = BTRV(BtOpClose, SUMJ_POS, SUMJREC, Len(SUMJREC), K0_SUMJ, Len(K0_SUMJ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入出荷実績データ")
            Exit Function
        End If
    End If
    
    Call Input_UnLock

    Main_Proc = False
End Function

Private Function Err_Chk()
'----------------------------------------------------------------------------
'                   エラーチェック処理
'----------------------------------------------------------------------------
    
Dim i As Integer
    
    Err_Chk = True


    For i = ptxS_DT_YY To ptxE_DT_DD
        If Len(Text(i).Text) = 0 Then
            Select Case i
                Case ptxS_DT_YY
                    Text(i).Text = "0000"
                Case ptxE_DT_YY
                    Text(i).Text = "9999"
                Case ptxS_DT_MM, ptxS_DT_DD
                    Text(i).Text = "00"
                Case ptxE_DT_MM, ptxE_DT_DD
                    Text(i).Text = "99"
            End Select
        Else
            If IsNumeric(Text(i).Text) Then
                Select Case i
                    Case ptxS_DT_YY, ptxE_DT_YY
                        Text(i).Text = Format(CInt(Text(i).Text), "0000")
                    Case Else
                        Text(i).Text = Format(CInt(Text(i).Text), "00")
                End Select
            End If
        End If
    Next i
    
    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1200301.MousePointer = vbHourglass

    Call Ctrl_Lock(F1200301)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1200301)


    F1200301.MousePointer = vbDefault

End Sub

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 0                           '印刷
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("「入出荷集計処理」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Main_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Text(ptxS_DT_YY).SetFocus
                    
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    

End Sub

Private Sub Form_Activate()
    
    Text(ptxS_DT_YY).SetFocus

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
    End Select


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If

End Sub


Private Sub Form_Load()

Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer
        

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
    LOG_F = Trim(c)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1200301.Caption = "入出荷実績集計処理（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i

                                '通常入荷の要因取り込み
    If GetIni("YOIN", "YOIN_TU_NYUKA", "SYS", c) Then
        YOIN_TU_NYUKA = ""
    Else
        YOIN_TU_NYUKA = Trim(c)
    End If



                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
    
    sts = BTRV(BtOpReset, SUMJ_POS, SUMJREC, Len(SUMJREC), K0_SUMJ, Len(K0_SUMJ), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1200301 = Nothing

    End
End Sub



Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    
                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If
    
    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1200301.Caption = "入出荷実績集計処理（" + RTrim(JGYOBU_T(Index).NAME) + "）"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If



End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub


