VERSION 5.00
Begin VB.Form F1040351 
   BackColor       =   &H00FFFFFF&
   Caption         =   "棚別在庫一覧表印刷"
   ClientHeight    =   8250
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   11430
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
   ScaleHeight     =   8250
   ScaleWidth      =   11430
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   3
      Left            =   4410
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      Top             =   1680
      Width           =   2430
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   3885
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1680
      Width           =   540
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3885
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1080
      Width           =   1380
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   9450
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   360
      Width           =   852
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   6720
      MaxLength       =   13
      TabIndex        =   15
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   3885
      MaxLength       =   13
      TabIndex        =   14
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   13
      Top             =   4080
      Width           =   435
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   12
      Top             =   4080
      Width           =   435
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   11
      Top             =   3480
      Width           =   435
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3480
      Width           =   435
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2880
      Width           =   435
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2880
      Width           =   435
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   6720
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   360
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印刷中断"
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
      Left            =   4830
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2280
      Width           =   435
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   3360
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   360
      Width           =   1380
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2280
      Width           =   435
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
      Left            =   10290
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   9450
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   8610
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6960
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
      Index           =   8
      Left            =   7770
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6960
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
      Index           =   7
      Left            =   6510
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   5670
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   4830
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   3990
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   2625
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   1785
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   945
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6960
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
      Index           =   0
      Left            =   105
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "在庫収支"
      Height          =   240
      Index           =   14
      Left            =   2520
      TabIndex        =   46
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "以降"
      Height          =   240
      Index           =   13
      Left            =   5355
      TabIndex        =   45
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "前回移動日"
      Height          =   255
      Index           =   12
      Left            =   2310
      TabIndex        =   44
      Top             =   1200
      Width           =   1275
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   11
      Left            =   8610
      TabIndex        =   43
      Top             =   480
      Width           =   750
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
      Height          =   375
      Left            =   105
      TabIndex        =   42
      Top             =   7440
      Width           =   225
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（外部）"
      Height          =   255
      Index           =   10
      Left            =   2310
      TabIndex        =   41
      Top             =   4800
      Width           =   1485
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   9
      Left            =   5880
      TabIndex        =   40
      Top             =   4800
      Width           =   330
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "　　　段"
      Height          =   255
      Index           =   8
      Left            =   2310
      TabIndex        =   39
      Top             =   4200
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   38
      Top             =   4080
      Width           =   330
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "　　　連"
      Height          =   255
      Index           =   6
      Left            =   2310
      TabIndex        =   37
      Top             =   3600
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   36
      Top             =   3600
      Width           =   330
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "棚番　列"
      Height          =   255
      Index           =   4
      Left            =   2310
      TabIndex        =   35
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   34
      Top             =   3000
      Width           =   330
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "入荷日別明細"
      Height          =   255
      Index           =   2
      Left            =   5145
      TabIndex        =   33
      Top             =   480
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷中です"
      Height          =   255
      Left            =   4935
      TabIndex        =   32
      Top             =   5400
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   31
      Top             =   2400
      Width           =   330
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "空棚印刷"
      Height          =   255
      Index           =   33
      Left            =   2310
      TabIndex        =   30
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "倉庫№"
      Height          =   255
      Index           =   0
      Left            =   2310
      TabIndex        =   29
      Top             =   2400
      Width           =   750
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
Attribute VB_Name = "F1040351"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxDATE% = 0                  '前回移動日
Private Const ptxSYUSHI% = 1                '在庫収支


Private Const ptxS_SOKO_NO% = 2             '開始　倉庫№
Private Const ptxE_SOKO_NO% = 3             '終了　倉庫№
Private Const ptxS_RETU% = 4                '開始　棚番　列
Private Const ptxE_RETU% = 5                '終了　棚番　列
Private Const ptxS_REN% = 6                 '開始　棚番　連
Private Const ptxE_REN% = 7                 '終了　棚番　連
Private Const ptxS_DAN% = 8                 '開始　棚番　段
Private Const ptxE_DAN% = 9                 '終了　棚番　段
Private Const ptxS_HIN_GAI% = 10            '開始　品番（外部）
Private Const ptxE_HIN_GAI% = 11            '開始　品番（外部）



Private Const Text_Max% = 9                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbTANA_INF% = 0             '空棚印刷
Private Const pcmbDETA% = 1                 '入荷日別明細印刷
Private Const pcmbNAIGAI% = 2               '国内外

Private Const pcmbSYUSHI% = 3               '収支

Private Const LMAX% = 44                    '頁内最大行数
Private Const MGN_L% = 5                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Dim Pdate As String                         '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime As String                         '印刷開始時刻（ﾍｯﾀﾞｰ用）

Dim NormalFont As New StdFont               '印刷フォント

Dim PRT_CAN As Boolean                      '印刷途中キャンセル要求


Private Const TANA_INF_NO$ = "1"            '空棚印刷方法のリストボックス内容
Private Const TANA_INF_ALL$ = "2"
Private Const TANA_INF_ONLY$ = "3"
Private Const TANA_INF1$ = "空棚無し"
Private Const TANA_INF2$ = "空棚有り"
Private Const TANA_INF3$ = "空棚のみ"

Private Const DETA_ON$ = "0"                '明細印刷方法のリストボックス内容
Private Const DETA_OFF$ = "1"

Private Const DETA0$ = "明細有り"
Private Const DETA1$ = "明細無し"
Dim TZAIKO_DATA  As String                  '在庫データフルパス

Private Function Print_Proc() As Integer

Dim Soko_COM        As Integer
Dim TANA_COM        As Integer
Dim ZAIKO_COM       As Integer
Dim sts             As Integer

Dim RetBuf          As String

Dim Sum_Yuko_Z_Qty  As Long
Dim SAVE_NAIGAI     As String * 1
Dim SAVE_HIN_GAI    As String * 13

Dim PRI_TANA        As String * 8
Dim PRI_NAIGAI      As String * 1
Dim PRI_HIN_GAI     As String * 13

Dim LCNT            As Integer
    
Dim SKIP_F          As Boolean
    
    
    Print_Proc = True
'印刷中は「印刷中断」以外のイベント取得不可
    Call Input_Lock           '画面項目ロック
    Label1.Visible = True
    Command1.Visible = True
    Command1.Enabled = True

    PRT_CAN = False

'印刷開始
    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    Pdate = Date
    Ptime = Time

    LCNT = 99
    
    SAVE_NAIGAI = ""
    SAVE_HIN_GAI = ""
    Sum_Yuko_Z_Qty = 0

    Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxS_SOKO_NO).Text)
    
    Soko_COM = BtOpGetGreaterEqual

    Do
        DoEvents
        
        sts = BTRV(Soko_COM, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        
        Select Case sts
            Case BtNoErr
                If StrConv(SOKOREC.Soko_No, vbUnicode) > Text(ptxE_SOKO_NO).Text Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, Soko_COM, "倉庫マスタ")
                Exit Function
        End Select
        If (StrConv(SOKOREC.JGYOBU, vbUnicode) = Last_JGYOBU Or _
            StrConv(SOKOREC.JGYOBU, vbUnicode) = JGYOBU_NON) Then
            '印刷対象の倉庫？(事業部＝指定事業部／事業部無し)
            If StrConv(SOKOREC.NAIGAI, vbUnicode) = NAIGAI_NON Or _
                Right(Combo(pcmbNAIGAI).Text, 1) = NAIGAI_NON Then
            Else
                If StrConv(SOKOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                    Exit Do
                End If
            End If
            
            If LCNT <> 99 Then
                LCNT = LMAX + 1
            End If
            
            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(SOKOREC.Soko_No, vbUnicode))
            Call UniCode_Conv(K0_TANA.Retu, Text(ptxS_RETU).Text)
            Call UniCode_Conv(K0_TANA.Ren, Text(ptxS_REN).Text)
            Call UniCode_Conv(K0_TANA.Dan, Text(ptxS_DAN).Text)
            
            TANA_COM = BtOpGetGreaterEqual

            Do
                DoEvents

                sts = BTRV(TANA_COM, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                Select Case sts
                    Case BtNoErr
                        If (StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)) _
                            > (Text(ptxE_SOKO_NO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                            Exit Do
                        End If
                    
                    
                        If StrConv(SOKOREC.Soko_No, vbUnicode) <> StrConv(TANAREC.Soko_No, vbUnicode) Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, TANA_COM, "棚マスタ")
                        Exit Function
                End Select
                                            '在庫データ読み込み開始
                Call UniCode_Conv(K5_ZAIKO.Soko_No, StrConv(TANAREC.Soko_No, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Retu, StrConv(TANAREC.Retu, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Ren, StrConv(TANAREC.Ren, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Dan, StrConv(TANAREC.Dan, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K5_ZAIKO.NAIGAI, NAIGAI_NON)
                Call UniCode_Conv(K5_ZAIKO.HIN_GAI, "")
                Call UniCode_Conv(K5_ZAIKO.NYUKA_DT, "")
                                
                Sum_Yuko_Z_Qty = 0
                SAVE_NAIGAI = ""
                SAVE_HIN_GAI = ""
                                
                ZAIKO_COM = BtOpGetGreater
                
                Do
                    DoEvents
                
                    If PRT_CAN Then
                        Printer.KillDoc
                        Call Input_UnLock   '画面項目ロック解除
                        Label1.Visible = False
                        Command1.Visible = False
                        Print_Proc = False
                        Exit Function
                    End If
                
                    sts = BTRV(ZAIKO_COM, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> StrConv(TANAREC.Soko_No, vbUnicode) Or _
                                StrConv(ZAIKOREC.Retu, vbUnicode) <> StrConv(TANAREC.Retu, vbUnicode) Or _
                                StrConv(ZAIKOREC.Ren, vbUnicode) <> StrConv(TANAREC.Ren, vbUnicode) Or _
                                StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(TANAREC.Dan, vbUnicode) Or _
                                StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                                            '棚番／事業部ブレーク
                                If Len(Trim(SAVE_NAIGAI)) = 0 Then
                                            '在庫が無かった
                                    If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_NO Then
                                        If LCNT > LMAX Then
                                            Call Print_Head(LCNT)
                                            PRI_TANA = ""
                                        End If
                                        Printer.Print Tab(MGN_L);
                                        Printer.Print StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode)
                                        Printer.Print
                                        LCNT = LCNT + 2
                                    End If
                                Else
                                    If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_ONLY And _
                                        Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                        If TOTAL_PRINT(LCNT, PRI_TANA, SAVE_NAIGAI, SAVE_HIN_GAI, Sum_Yuko_Z_Qty) Then
                                            Exit Function
                                        End If
                                    End If
                                        
                                    Printer.Print       '１行改行
                                    LCNT = LCNT + 1
                                End If
                                
                                Exit Do
                            
                            End If
                        Case BtErrEOF
                            If Len(Trim(SAVE_NAIGAI)) = 0 Then
                                            '在庫が無かった
                                If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_NO Then
                                    If LCNT > LMAX Then
                                        Call Print_Head(LCNT)
                                        PRI_TANA = ""
                                    End If
                                    Printer.Print Tab(MGN_L);
                                    Printer.Print StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode)
                                    Printer.Print
                                    LCNT = LCNT + 2
                                End If
                            Else
                                If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_ONLY And _
                                    Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                    If TOTAL_PRINT(LCNT, PRI_TANA, SAVE_NAIGAI, SAVE_HIN_GAI, Sum_Yuko_Z_Qty) Then
                                        Exit Function
                                    End If
                                End If
                                    
                                Printer.Print       '１行改行
                                LCNT = LCNT + 1
                            
                            End If
                            
                            Exit Do
                        Case Else
                            Call File_Error(sts, ZAIKO_COM, "在庫データ")
                            Exit Function
                    End Select
                
                                    
                
                    SKIP_F = False
                
                
                    If Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) < Trim(Text(ptxS_HIN_GAI).Text) Or _
                        Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) > Trim(Text(ptxE_HIN_GAI).Text) Then
                        SKIP_F = True
                    End If
                
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
                
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                            SKIP_F = True
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
                
                
                    If Trim(Text(ptxDATE).Text) <> "" Then
                        If StrConv(ITEMREC.LAST_NYU_DT, vbUnicode) < Format(Text(ptxDATE).Text, "YYYYMMDD") And _
                            StrConv(ITEMREC.LAST_SYU_DT, vbUnicode) < Format(Text(ptxDATE).Text, "YYYYMMDD") Then
                            SKIP_F = True
                        End If
                    End If
                
                
                    If Trim(Text(ptxSYUSHI).Text) <> "" Then
                        If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) <> Trim(Text(ptxSYUSHI).Text) Then
                            SKIP_F = True
                        End If
                    End If
                
                
                
                
                
                
                    If (Right(Combo(pcmbNAIGAI).Text, 1) <> NAIGAI_NON And _
                        Right(Combo(pcmbNAIGAI).Text, 1) <> StrConv(ZAIKOREC.NAIGAI, vbUnicode)) Or _
                        SKIP_F Then
                                                '内外対象外
                    
                                            
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    Else
                            
                        If Right(Combo(pcmbTANA_INF).Text, 1) = TANA_INF_ONLY Then
                                                '空棚のみ
                            Exit Do
                            
                        End If
                            
                        If Len(Trim(SAVE_NAIGAI)) = 0 Then
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                        End If
                
                        If SAVE_NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                            SAVE_HIN_GAI <> Left(StrConv(ZAIKOREC.HIN_GAI, vbUnicode), 13) Then
                            If Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                If TOTAL_PRINT(LCNT, PRI_TANA, SAVE_NAIGAI, SAVE_HIN_GAI, Sum_Yuko_Z_Qty) Then
                                    Exit Function
                                End If
                            End If
                            
                            Printer.Print           '1行改行
                            LCNT = LCNT + 1
                            
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                                                    
                            Sum_Yuko_Z_Qty = 0
                                                     
                            PRI_NAIGAI = ""
                            PRI_HIN_GAI = ""
                            
                        End If
                                                    
                                                    
                        Sum_Yuko_Z_Qty = Sum_Yuko_Z_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                                    
                        If Right(Combo(pcmbDETA).Text, 1) = DETA_ON Then
                                                    '明細印刷
                            If LCNT > LMAX Then
                                Call Print_Head(LCNT)
                                PRI_TANA = ""
                            End If
                                '棚番
                            If PRI_TANA <> (StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)) Then
                                Printer.Print Tab(MGN_L);
                                Printer.Print StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode);
                                PRI_TANA = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                            End If
                                '国内外
                            Printer.Print Tab(MGN_L + 10);
                            If SAVE_NAIGAI = NAIGAI_NAI Then
                                Printer.Print NAIGAI1;
                            Else
                                Printer.Print NAIGAI2;
                            End If
                                '品番
                            Printer.Print Tab(MGN_L + 18);
                            Printer.Print SAVE_HIN_GAI;
                                '品名
                            Printer.Print Tab(MGN_L + 39);
                            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                            Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                    Printer.Print LeftB(Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), 44);
                                
                                
                                
                                
                                Case BtErrKeyNotFound
                                
                                
                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                                
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Exit Function
                            End Select
                                '入荷日
                            Printer.Print Tab(MGN_L + 66);
                            Printer.Print Left(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 4) & "/" & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & Right(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 2);
                                '品番（内部）
                            Printer.Print Tab(MGN_L + 78);
                            Printer.Print Left(StrConv(ZAIKOREC.HIN_NAI, vbUnicode), 13);
                                
                            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                Printer.Print "(済)";
                            Else
                                Printer.Print "(未)";
                            End If
                                                        
                                '有効在庫数
                            Printer.Print Tab(MGN_L + 99);
                            RetBuf = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#,##0")
                            If Len(RetBuf) < 9 Then
                                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                            End If
                            Printer.Print RetBuf;
                                               
                                
                                '累計有効在庫数
                            Printer.Print Tab(MGN_L + 110);
                            RetBuf = Format(Sum_Yuko_Z_Qty, "#,##0")
                            If Len(RetBuf) < 9 Then
                                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                            End If
                            Printer.Print RetBuf;
                                '標準棚番
                            Printer.Print Tab(MGN_L + 120);
                            Printer.Print StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
                            
                            
                    
                            LCNT = LCNT + 1
                        End If
                    End If
                    
                    ZAIKO_COM = BtOpGetNext
                
                Loop
                
                
                TANA_COM = BtOpGetNext

            Loop

        End If
    
        Soko_COM = BtOpGetNext
    
    Loop

    If LCNT <> 99 Then
        Printer.EndDoc
    End If

    Call Input_UnLock               '画面項目ロック解除
    Label1.Visible = False
    Command1.Visible = False

    Print_Proc = False

End Function
Private Sub Print_Head(LCNT As Integer)
'ヘッダ印刷
Dim i As Integer

    If LCNT < 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        'ヘッダー（１）
    Printer.Print Tab(3);
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    Printer.Print Tab(36);
    Printer.Print "＊＊＊  棚別在庫一覧表  ＊＊＊";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        'ヘッダー（２）
    Printer.Print Tab(MGN_L);
    Printer.Print "倉庫：";
    Printer.Print StrConv(SOKOREC.Soko_No, vbUnicode);
    Printer.Print " ";
    Printer.Print StrConv(SOKOREC.SOKO_NAME, vbUnicode);
    
    Printer.Print " ";
    
    If Trim(Text(ptxSYUSHI).Text) <> "" Then
        Printer.Print Text(ptxSYUSHI).Text & ":" & Left(Combo(pcmbSYUSHI).Text, Len(Combo(pcmbSYUSHI).Text) - 3);
    End If
    
    Printer.Print
    Printer.Print

                                        'ヘッダー（３）
    Printer.Print Tab(MGN_L);
    Printer.Print "棚番";
    Printer.Print Tab(MGN_L + 10);
    Printer.Print "国内外";
    Printer.Print Tab(MGN_L + 18);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 39);
    Printer.Print "品  名  ";
    Printer.Print Tab(MGN_L + 66);
    Printer.Print "入荷日";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "品番（内部）";
    Printer.Print Tab(MGN_L + 102);
    Printer.Print "在庫数";
    
    If Right(Combo(pcmbDETA).Text, 1) = DETA_ON Then
        Printer.Print Tab(MGN_L + 113);
        Printer.Print "累計数";
    End If
    
    Printer.Print Tab(MGN_L + 120);
    Printer.Print "標準棚番";
    
    
    
    Printer.Print

    Printer.Print

    LCNT = 7 + MGN_U

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1040351.MousePointer = vbHourglass

    Call Ctrl_Lock(F1040351)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1040351)


    F1040351.MousePointer = vbDefault

End Sub


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   コンボボックス入力（ＫｅｙＤｏｗｎ）処理
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case pcmbTANA_INF           '空棚印刷
            If Right(Combo(pcmbTANA_INF).Text, 1) = TANA_INF_ONLY Then
                Combo(pcmbDETA).Enabled = False
                Combo(pcmbDETA).TabStop = False
                Combo(pcmbNAIGAI).SetFocus
            Else
                Combo(pcmbDETA).Enabled = True
                Combo(pcmbDETA).TabStop = True
                Combo(pcmbDETA).SetFocus
            End If
        Case pcmbDETA               '入荷日別明細
            Combo(pcmbNAIGAI).SetFocus
        Case pcmbNAIGAI             '国内外
            Text(ptxDATE).SetFocus
    
        Case pcmbSYUSHI
            Text(ptxSYUSHI).Text = Right(Combo(pcmbSYUSHI).Text, 3)
            Text(ptxS_SOKO_NO).SetFocus
    
    
    End Select


End Sub


Private Sub Combo_LostFocus(Index As Integer)
    Select Case Index
        Case pcmbTANA_INF           '空棚印刷
        Case pcmbDETA               '入荷日別明細
        Case pcmbNAIGAI             '国内外
    
        Case pcmbSYUSHI
            Text(ptxSYUSHI).Text = Right(Combo(pcmbSYUSHI).Text, 3)
    
    
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        
        
        Case 7                              'データ
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("「棚別在庫一覧表」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If OUTPUT_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Combo(pcmbTANA_INF).SetFocus
        
        
        Case 8                              '印刷
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("「棚別在庫一覧表」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Combo(pcmbTANA_INF).SetFocus
                    
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
End Sub
Private Sub Command1_Click()
    PRT_CAN = True
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
    LOG_F = Trim(c)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1040351.Caption = "棚番別在庫一覧表印刷（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
    
    
    

                                '棚別在庫ファイル名取り込み
    If GetIni("FILE", "TZAIKO_DATA", "SYS", c) Then
        Beep
        MsgBox "棚別在庫ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    TZAIKO_DATA = Trim(c)
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '棚マスタＯＰＥＮ
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
        
    '収支セット
    If Code_Set_Proc(pcmbSYUSHI, P_KBN03_CD, 1) Then
        Unload Me
    End If
                                
                                
                                
                                '印刷フォント設定
    With NormalFont
        .NAME = F1040351.FontName
        .Size = F1040351.FontSize
    End With
    Set Printer.Font = NormalFont
                                '画面初期設定
    Combo(pcmbTANA_INF).AddItem TANA_INF1 & "   " & TANA_INF_NO
    Combo(pcmbTANA_INF).AddItem TANA_INF2 & "   " & TANA_INF_ALL
    Combo(pcmbTANA_INF).AddItem TANA_INF3 & "   " & TANA_INF_ONLY
    Combo(pcmbTANA_INF).ListIndex = 0
    
    Combo(pcmbDETA).AddItem DETA0 & "   " & DETA_ON
    Combo(pcmbDETA).AddItem DETA1 & "   " & DETA_OFF
    Combo(pcmbDETA).ListIndex = 0
    
    Combo(pcmbNAIGAI).AddItem NAIGAI0 & "   " & NAIGAI_NON
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
    
    Combo(pcmbTANA_INF).SetFocus
    
    

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '棚マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1040351 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1040351.Caption = "棚別在庫一覧表印刷（" + RTrim(JGYOBU_T(Index).NAME) + ")"
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



Private Function Err_Chk()
    
Dim i As Integer
    
    Err_Chk = True



'日付
    If Trim(Text(ptxDATE).Text) = "" Then
    Else
        If Not IsDate(Text(ptxDATE).Text) Then
            Beep
            MsgBox "入力した項目はエラーです。"
            Text(ptxS_SOKO_NO).SetFocus
            Exit Function
        End If
    End If
'倉庫番号

    If Len(Text(ptxE_SOKO_NO).Text) = 0 Then
        Text(ptxE_SOKO_NO).Text = "zz"
    End If

    If Text(ptxS_SOKO_NO).Text > Text(ptxE_SOKO_NO).Text Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxS_SOKO_NO).SetFocus
        Exit Function
    End If

'棚番
    For i = ptxS_RETU To ptxE_DAN
        Select Case i
            Case ptxS_RETU, ptxS_REN, ptxS_DAN
                If Len(Text(i).Text) = 0 Then
                    Text(i).Text = "00"
                End If
            Case ptxE_RETU, ptxE_REN, ptxE_DAN
                If Len(Text(i).Text) = 0 Then
                    Text(i).Text = "99"
                End If
        End Select
        If IsNumeric(Text(i).Text) Then
            Text(i).Text = Format(CInt(Text(i).Text), "00")
        End If
    Next i


    If Text(ptxS_RETU).Text & Text(ptxS_REN).Text & Text(ptxS_DAN).Text _
        > Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxS_RETU).SetFocus
        Exit Function
    End If
'品番(外部)
    If Len(Text(ptxE_HIN_GAI).Text) = 0 Then
        Text(ptxE_HIN_GAI).Text = String(Len(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)), "z")
    End If

    If Text(ptxS_HIN_GAI).Text > Text(ptxE_HIN_GAI).Text Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxS_HIN_GAI).SetFocus
        Exit Function
    End If
    
    Err_Chk = False

End Function

Private Function TOTAL_PRINT(LCNT As Integer, _
                                PRI_TANA As String, _
                                SAVE_NAIGAI As String, _
                                SAVE_HIN_GAI As String, _
                                Sum_Yuko_Z_Qty As Long) As Integer

Dim sts     As Integer
Dim RetBuf  As String
    
    TOTAL_PRINT = True
    
    If LCNT > LMAX Then
        Call Print_Head(LCNT)
        PRI_TANA = ""
    End If
                                '棚番
    If PRI_TANA <> (StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)) Then
        Printer.Print Tab(MGN_L);
        Printer.Print StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode);
        PRI_TANA = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
    End If
                                '国内外
    Printer.Print Tab(MGN_L + 10);
    If SAVE_NAIGAI = NAIGAI_NAI Then
        Printer.Print NAIGAI1;
    Else
        Printer.Print NAIGAI2;
    End If
                                '品番
    Printer.Print Tab(MGN_L + 18);
    Printer.Print SAVE_HIN_GAI;
                                '品名
    Printer.Print Tab(MGN_L + 39);
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Printer.Print Left(StrConv(ITEMREC.HIN_NAME, vbUnicode), 25);
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
                                '有効在庫数
    Printer.Print Tab(MGN_L + 99);
    RetBuf = Format(Sum_Yuko_Z_Qty, "#,##0")
    If Len(RetBuf) < 9 Then
        RetBuf = Space(9 - Len(RetBuf)) & RetBuf
    End If
    Printer.Print RetBuf;
                                '標準棚番
    Printer.Print Tab(MGN_L + 120);
    Printer.Print StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)

    LCNT = LCNT + 1
                    
    TOTAL_PRINT = False
                    
                    
End Function
Private Function OUTPUT_Proc() As Integer
    
Dim sts             As Integer
Dim Soko_COM        As Integer
Dim TANA_COM        As Integer
Dim ZAIKO_COM       As Integer
Dim Ret             As Integer
    
Dim Sum_Yuko_Z_Qty  As Long
Dim SAVE_HIN_GAI    As String * 13
Dim SAVE_NAIGAI     As String * 1

Dim FileNo          As Long
Dim fileName        As String

Dim c               As String * 128
Dim Soko_No         As String * 2

Dim SKIP_F          As Boolean

    
    OUTPUT_Proc = True
'実行中中はイベント取得不可
    Call Input_Lock         '画面項目ロック

    FileNo = FreeFile
    fileName = TZAIKO_DATA
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo

    If Trim(Text(ptxSYUSHI).Text) <> "" Then
        Write #FileNo, " " & Trim(Text(ptxSYUSHI).Text) & "：" & Trim(Left(Combo(pcmbSYUSHI).Text, Len(Combo(pcmbSYUSHI).Text) - 3))
    End If



    Write #FileNo, "棚番", "国内外", "品番（外）", "品名", "入荷日", "品番（内）", "商／未商", "在庫数", "累計数", "標準棚番"



    Sum_Yuko_Z_Qty = 0

    Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxS_SOKO_NO).Text)
    
    Soko_COM = BtOpGetGreaterEqual

    Do
        DoEvents
        
        sts = BTRV(Soko_COM, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        
        Select Case sts
            Case BtNoErr
                If StrConv(SOKOREC.Soko_No, vbUnicode) > Text(ptxE_SOKO_NO).Text Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, Soko_COM, "倉庫マスタ")
                Exit Function
        End Select
        If (StrConv(SOKOREC.JGYOBU, vbUnicode) = Last_JGYOBU Or _
            StrConv(SOKOREC.JGYOBU, vbUnicode) = JGYOBU_NON) Then
            '印刷対象の倉庫？(事業部＝指定事業部／事業部無し)
            If StrConv(SOKOREC.NAIGAI, vbUnicode) = NAIGAI_NON Or _
                Right(Combo(pcmbNAIGAI).Text, 1) = NAIGAI_NON Then
            Else
                If StrConv(SOKOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                    Exit Do
                End If
            End If
            
            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(SOKOREC.Soko_No, vbUnicode))
            Call UniCode_Conv(K0_TANA.Retu, Text(ptxS_RETU).Text)
            Call UniCode_Conv(K0_TANA.Ren, Text(ptxS_REN).Text)
            Call UniCode_Conv(K0_TANA.Dan, Text(ptxS_DAN).Text)
            
            TANA_COM = BtOpGetGreaterEqual

            Do
                DoEvents

                sts = BTRV(TANA_COM, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                Select Case sts
                    Case BtNoErr
                        If (StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)) _
                            > (Text(ptxE_SOKO_NO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                            Exit Do
                        End If
                    
                    
                        If StrConv(SOKOREC.Soko_No, vbUnicode) <> StrConv(TANAREC.Soko_No, vbUnicode) Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, TANA_COM, "棚マスタ")
                        Exit Function
                End Select
                                            '在庫データ読み込み開始
                Call UniCode_Conv(K5_ZAIKO.Soko_No, StrConv(TANAREC.Soko_No, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Retu, StrConv(TANAREC.Retu, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Ren, StrConv(TANAREC.Ren, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Dan, StrConv(TANAREC.Dan, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K5_ZAIKO.NAIGAI, NAIGAI_NON)
                Call UniCode_Conv(K5_ZAIKO.HIN_GAI, "")
                Call UniCode_Conv(K5_ZAIKO.NYUKA_DT, "")
                                
                Sum_Yuko_Z_Qty = 0
                SAVE_NAIGAI = ""
                SAVE_HIN_GAI = ""
                                
                ZAIKO_COM = BtOpGetGreater
                
                Do
                    DoEvents
                
                
                    sts = BTRV(ZAIKO_COM, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> StrConv(TANAREC.Soko_No, vbUnicode) Or _
                                StrConv(ZAIKOREC.Retu, vbUnicode) <> StrConv(TANAREC.Retu, vbUnicode) Or _
                                StrConv(ZAIKOREC.Ren, vbUnicode) <> StrConv(TANAREC.Ren, vbUnicode) Or _
                                StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(TANAREC.Dan, vbUnicode) Or _
                                StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                                            '棚番／事業部ブレーク
                                If Len(Trim(SAVE_NAIGAI)) = 0 Then
                                            '在庫が無かった
                                    If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_NO Then
                                    
                                                                        
                                        Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode)
                                    
                                    
                                    
                                    End If
                                Else
                                    If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_ONLY And _
                                        Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                                    
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                                                
                                                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                                Call UniCode_Conv(ITEMREC.ST_REN, "")
                                                Call UniCode_Conv(ITEMREC.ST_DAN, "")
                                            
                                            
                                                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                                Exit Function
                                        End Select
                                                    
                                        
                                        
                                        
                                        Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode), SAVE_NAIGAI, SAVE_HIN_GAI, StrConv(ITEMREC.HIN_NAME, vbUnicode), StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
                                    
                                    End If
                                End If
                                
                                Exit Do
                            
                            End If
                        Case BtErrEOF
                            If Len(Trim(SAVE_NAIGAI)) = 0 Then
                                            '在庫が無かった
                                If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_NO Then
                                    
                                                                        
                                    Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode)
                                End If
                            Else
                                If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_ONLY And _
                                    Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                    
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        Case BtErrKeyNotFound
                                        
                                            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                        
                                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                            Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                            Call UniCode_Conv(ITEMREC.ST_REN, "")
                                            Call UniCode_Conv(ITEMREC.ST_DAN, "")
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                            Exit Function
                                    End Select
                                                
                                    
                                        
                                    Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode), SAVE_NAIGAI, SAVE_HIN_GAI, StrConv(ITEMREC.HIN_NAME, vbUnicode), StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
                                    
                                    
                                    
                                End If
                            
                            End If
                            
                            Exit Do
                        Case Else
                            Call File_Error(sts, ZAIKO_COM, "在庫データ")
                            Exit Function
                    End Select
                
                
                
                    SKIP_F = False
If Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) = "B015" Then
 Debug.Print StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
End If
                
                    If Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) < Trim(Text(ptxS_HIN_GAI).Text) Or _
                        Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) > Trim(Text(ptxE_HIN_GAI).Text) Then
                        SKIP_F = True
                    End If
                
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
                
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                            SKIP_F = True
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
                
                
                    If Trim(Text(ptxDATE).Text) <> "" Then
                        If StrConv(ITEMREC.LAST_NYU_DT, vbUnicode) < Format(Text(ptxDATE).Text, "YYYYMMDD") And _
                            StrConv(ITEMREC.LAST_SYU_DT, vbUnicode) < Format(Text(ptxDATE).Text, "YYYYMMDD") Then
                            SKIP_F = True
                        End If
                    End If
                
                
                    If Trim(Text(ptxSYUSHI).Text) <> "" Then
                        If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) <> Trim(Text(ptxSYUSHI).Text) Then
                            SKIP_F = True
                        End If
                    End If
                
                
                
                
                
                
                
                    If (Right(Combo(pcmbNAIGAI).Text, 1) <> NAIGAI_NON And _
                        Right(Combo(pcmbNAIGAI).Text, 1) <> StrConv(ZAIKOREC.NAIGAI, vbUnicode)) Or _
                        SKIP_F Then
                                                '内外対象外
                    Else
                            
                        If Right(Combo(pcmbTANA_INF).Text, 1) = TANA_INF_ONLY Then
                                                '空棚のみ
                            Exit Do
                            
                        End If
                            
                        If Len(Trim(SAVE_NAIGAI)) = 0 Then
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                        End If
                
                        If Trim(SAVE_NAIGAI) <> Trim(StrConv(ZAIKOREC.NAIGAI, vbUnicode)) Or _
                            Trim(SAVE_HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                            If Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                                Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
                                Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound
                                        Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                        Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                        Call UniCode_Conv(ITEMREC.ST_REN, "")
                                        Call UniCode_Conv(ITEMREC.ST_DAN, "")
                                        Call UniCode_Conv(ITEMREC.HIN_NAME, "")

                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                        Exit Function
                                End Select


                                
''                                Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode), SAVE_NAIGAI, SAVE_HIN_GAI, StrConv(ITEMREC.HIN_NAME, vbUnicode), StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
                            
                            End If
                            
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                                                
                            Sum_Yuko_Z_Qty = 0
                                                     
                            
                                                    
                                                    
                        End If
                    End If
                    
                    
                    If (Right(Combo(pcmbNAIGAI).Text, 1) <> NAIGAI_NON And _
                        Right(Combo(pcmbNAIGAI).Text, 1) <> StrConv(ZAIKOREC.NAIGAI, vbUnicode)) Or _
                        SKIP_F Then
                                                '内外対象外
                    Else
                            
                        If Right(Combo(pcmbTANA_INF).Text, 1) = TANA_INF_ONLY Then
                                                '空棚のみ
                            Exit Do
                            
                        End If
                            
                        If Len(Trim(SAVE_NAIGAI)) = 0 Then
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                        End If
                
                        If SAVE_NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                            SAVE_HIN_GAI <> Left(StrConv(ZAIKOREC.HIN_GAI, vbUnicode), 13) Then
                            If Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
'                                If TOTAL_OUTPUT(FileNo, SAVE_NAIGAI, SAVE_HIN_GAI, Sum_Yuko_Z_Qty) Then
'                                    Exit Function
'                                End If
                            End If
                            
                            
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                                                    
                            Sum_Yuko_Z_Qty = 0
                                                     
                            
                        End If
                                                    
                                                    
                        Sum_Yuko_Z_Qty = Sum_Yuko_Z_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                                    
                        If Right(Combo(pcmbDETA).Text, 1) = DETA_ON Then
                                '棚番
                            Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode),
                                '国内外
                            If SAVE_NAIGAI = NAIGAI_NAI Then
                                Write #FileNo, NAIGAI1,
                            Else
                                Write #FileNo, NAIGAI2,
                            End If
                                '品番
                            Write #FileNo, SAVE_HIN_GAI,
                                '品名
                            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                            Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                    Write #FileNo, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)),
                                
                                
                                
                                
                                Case BtErrKeyNotFound
                                
                                
                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                                    Write #FileNo, ,
                                
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Exit Function
                            End Select
                                '入荷日
                            Write #FileNo, Left(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 4) & "/" & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & Right(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 2),
                                '品番（内部）
                            Write #FileNo, Left(StrConv(ZAIKOREC.HIN_NAI, vbUnicode), 13),
                                
                            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                Write #FileNo, "(済)",
                            Else
                                Write #FileNo, "(未)",
                            End If
                                                        
                                '有効在庫数
                            Write #FileNo, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#,##0"),
                                '累計有効在庫数
                            Write #FileNo, Format(Sum_Yuko_Z_Qty, "#,##0"),
                                '標準棚番
                            Write #FileNo, StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
                            
                            
                    
                        End If
                    End If
                    
                    
                    
                    
                    
                    ZAIKO_COM = BtOpGetNext
                
                Loop
                
                
                TANA_COM = BtOpGetNext

            Loop

        End If
    
        Soko_COM = BtOpGetNext
    
    Loop
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Close #FileNo
    
    Call Input_UnLock         '画面項目ロック解除
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"

    OUTPUT_Proc = False
    
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If
End Function


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
    
    Combo(Index).Clear
    
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
        Combo(Index).AddItem Space(Key_Len)
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
        
        
        
        Combo(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function


