VERSION 5.00
Begin VB.Form F9000301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "品目マスタ（事業部用）出力[F900030]2015.07.02 08:00"
   ClientHeight    =   6315
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   11220
   ClipControls    =   0   'False
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
   ScaleHeight     =   6315
   ScaleWidth      =   11220
   StartUpPosition =   2  '画面の中央
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   10
      Left            =   9480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   9
      Left            =   8640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   8
      Left            =   7800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   7
      Left            =   6480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   6
      Left            =   5640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   5
      Left            =   4800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   4
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   3
      Left            =   2640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   1
      Left            =   960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "実  行"
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   19
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "件"
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   18
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "品目読込件数"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   17
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "品目マスタ出力処理が終了しました。"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "品目マスタ出力処理実行中です。"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "件"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   14
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   13
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "品目作成件数"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   12
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "F9000301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Const Out_File$ = "\\W1\newsdc\work\ITEM.TXT"  '2014.08.20

Dim Out_File    As String                       '2014.08.20

Dim READ_QTY As Long                    '読込データ件数（掃除機）   2015.04.10


Dim OK_QTY As Long                      '処理データ件数（掃除機）

                                    '画面初期表示（処理済「便」表示など）
Private Sub Scr_Init()

    Label2(0).Visible = False
    Label2(1).Visible = False
    
    Command(0).SetFocus

End Sub
                                        '品目マスタ登録 処理
Private Sub SYS_HIN_Main()
Dim sts As Integer
Dim Work As String
Dim com As Integer
    
Dim Err_moji    As Boolean      '2015.07.01
Dim i           As Integer      '2015.07.01
Dim wkHIN_GAI   As String * 15  '2015.07.01
    
    On Error GoTo Err_Exit
    Open Trim(Out_File) For Output As #1
    
    Call Scr_Lock                       '画面項目ロック
    Label2(0).Visible = True            '取込み中ﾒｯｾｰｼﾞ表示

    READ_QTY = 0        '2015.04.10
    OK_QTY = 0          '2015.04.10


    com = BtOpGetFirst

    Do
        
        DoEvents        '2015.04.10
        
        
        Do
            sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
            
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "品目マスタ")
                    Unload Me
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        READ_QTY = READ_QTY + 1                             '2015.04.10
        Label1(5).Caption = Format(READ_QTY, "#####0")      '2015.04.10
        
        '>2015.07.01
        Err_moji = False
        wkHIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
        For i = 1 To Len(wkHIN_GAI)
            If Mid(wkHIN_GAI, i, 1) < " " Or Mid(wkHIN_GAI, i, 1) > "z" Then
                Err_moji = True
                Exit For
            End If
        Next i
        '>2015.07.01
        
        
        If StrConv(ITEMREC.JGYOBU, vbUnicode) <> SOJIKI Or Err_moji Then
        Else
        
            OK_QTY = OK_QTY + 1
            Label1(1).Caption = Format(OK_QTY, "#####0")
            DoEvents
            Work = ""
                                            '事業部
            Work = Work & StrConv(ITEMREC.JGYOBU, vbUnicode)
                                            '国内外
            Work = Work & StrConv(ITEMREC.NAIGAI, vbUnicode)
                                            '品番（外部）
            Work = Work & Left(StrConv(ITEMREC.HIN_GAI, vbUnicode), 13)
                                            '品名
            Work = Work & StrConv(ITEMREC.HIN_NAME, vbUnicode)              '2015.07.01
'            Work = Work & Space(25)                                        '2015.07.01
                                            '標準倉庫設定日付
            Work = Work & StrConv(ITEMREC.ST_SET_DT, vbUnicode)
                                            '標準入庫棚 倉庫
            Work = Work & StrConv(ITEMREC.ST_SOKO, vbUnicode)
                                            '標準入庫棚 列
            Work = Work & StrConv(ITEMREC.ST_RETU, vbUnicode)
                                            '標準入庫棚 連
            Work = Work & StrConv(ITEMREC.ST_REN, vbUnicode)
                                            '標準入庫棚 段
            Work = Work & StrConv(ITEMREC.ST_DAN, vbUnicode)
                                            '前回入庫棚 倉庫
            Work = Work & StrConv(ITEMREC.BEF_SOKO, vbUnicode)
                                            '標準入庫棚 列
            Work = Work & StrConv(ITEMREC.BEF_RETU, vbUnicode)
                                            '標準入庫棚 連
            Work = Work & StrConv(ITEMREC.BEF_REN, vbUnicode)
                                            '標準入庫棚 段
            Work = Work & StrConv(ITEMREC.BEF_DAN, vbUnicode)
                                            '最終入庫日付
            Work = Work & StrConv(ITEMREC.LAST_NYU_DT, vbUnicode)
                                            '最終出庫日付
            Work = Work & StrConv(ITEMREC.LAST_SYU_DT, vbUnicode)
                                            '品番（内部）
            Work = Work & Left(StrConv(ITEMREC.HIN_NAI, vbUnicode), 13)
                                            '備考 ホスト倉庫
            Work = Work & StrConv(ITEMREC.BIKOU_SOKO, vbUnicode)
                                            '備考 ホスト棚番
            Work = Work & StrConv(ITEMREC.BIKOU_TANA, vbUnicode)
                                            '資材コード
    '        Work = Work & StrConv(ITEMREC.SIZAI_CD, vbUnicode)
            Work = Work & Space(5)
                                            '補充点
            Work = Work & StrConv(ITEMREC.HOJYU_P, vbUnicode)
                                            '月平均出荷数
            Work = Work & StrConv(ITEMREC.AVE_SYUKA, vbUnicode)
                                            'サンプル数
            Work = Work & StrConv(ITEMREC.SAMPLE_QTY, vbUnicode)
                                            '最終入荷日付
            Work = Work & StrConv(ITEMREC.LAST_INP_DT, vbUnicode)
            Work = Work & Space(11)
            
            Work = "144," & Work
            Print #1, Work
        
        End If
        com = BtOpGetNext
        
    Loop
    Print #1, Chr(26)         '1A
    Close #1
    Label2(0).Visible = False           '実行中ﾒｯｾｰｼﾞ ｸﾘｱ
    Label2(1).Visible = True            '終了表示
    
    Call Scr_UnLock
    Exit Sub
Err_Exit:
    MsgBox "出力エラー発生！", vbExclamation
    Call Scr_UnLock
End Sub
                                    '画面項目ロック（イベント取得不可）
Private Sub Scr_Lock()

Dim i As Integer

    F9000301.MousePointer = vbHourglass

    For i = 0 To 11
        Command(i).Enabled = False
    Next i

End Sub
                                    '画面項目ロック解除（イベント取得可）
Private Sub Scr_UnLock()

Dim i As Integer

    For i = 0 To 11
        Command(i).Enabled = True
    Next i

    F9000301.MousePointer = vbDefault

End Sub
Private Sub Command_Click(Index As Integer)
Dim yn As Integer

    Select Case Index
        Case 0
            yn = MsgBox("品目マスタ出力処理　実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                Call SYS_HIN_Main            '品目マスタ出力処理
                Command(0).Enabled = False
                Command(11).SetFocus
            End If
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
End Sub
Private Sub Command_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case KeyCode
        Case vbKeyF1
            Command(0).Value = True
        Case vbKeyF12
            Command(11).Value = True
        Case Else
            Beep
            If Command(0).Enabled = True Then
                Command(0).SetFocus
            Else
                Command(11).SetFocus
            End If
    End Select
End Sub

Private Sub Form_Activate()
    
    Call SYS_HIN_Main            '品目マスタ出力処理
    
    Unload Me

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
Dim i As Integer
Dim j As Integer
Dim c As String * 128
Dim sts As Integer


'    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                '新 品目マスタＯＰＥＮ
    If ITEM_Open(0) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    
    
'----------------   2014.08.20
    If GetIni(App.EXEName, "Out_File", App.EXEName, c) Then
    
        Out_File$ = "ITEM.TXT"
    Else
        Out_File = RTrim(c)
    
    End If
'----------------   2014.08.20

    
    
'    Call SYS_HIN_Main            '品目マスタ出力処理
'    Unload Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '新 品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
                                            'Ｂｔｒｉｅｖｅリセット
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "品目マスタ")
        Beep
        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
    End If

    End
End Sub

