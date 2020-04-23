VERSION 5.00
Begin VB.Form F1070301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "消費リスト印刷([F107030] 2012.04.19 14:00)"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2250
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
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "処理中断"
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
      Left            =   4680
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
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
      Left            =   7800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "ﾃﾞｰﾀ"
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "～"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "印刷日付範囲"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   16
      Top             =   2040
      Width           =   1455
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
      TabIndex        =   15
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1070301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_DATE% = 0                '開始　日付
Private Const ptxE_DATE% = 1                '終了　日付

Private Const Text_Max% = 2                 '画面項目別最大ｲﾝﾃﾞｯｸｽ


Private Print_Jgyobu        As Variant      '印刷対象事業部
Private Print_Jgyobu_T()    As String * 1


Private Print_Yoin          As Variant      '印刷対象要因
Private Print_Yoin_T()      As String * 2


Private Const LMAX% = 44                    '頁内最大行数
Private Const MGN_L% = 3                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Private Pdate           As String           '印刷開始日付（ﾍｯﾀﾞｰ用）
Private Ptime           As String           '印刷開始時刻（ﾍｯﾀﾞｰ用）

Private NormalFont      As New StdFont      '印刷フォント

Private PRT_CAN         As Boolean          '印刷途中キャンセル要求

Private F107030CSV      As String           'CSV出力ファイル


Private Function Print_Proc() As Integer
    
Dim sts             As Integer
Dim com             As Integer
    
Dim LCNT            As Integer

Dim i               As Integer
Dim j               As Integer


Dim RetBuf          As String

Dim Print_F         As Boolean



    Print_Proc = True


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "消費リスト印刷中", Me.hwnd, 0)


'印刷中は「印刷中断」以外のイベント取得不可
    Call Input_Lock         '画面項目ロック
    
    
    
    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    Pdate = Date
    Ptime = Time
    
    
    
    
    Command1.Visible = True
    Command1.Enabled = True


    Pdate = Date
    Ptime = Time





    PRT_CAN = False

    For i = 0 To UBound(Print_Jgyobu_T)
        Call UniCode_Conv(K0_IDO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_IDO.JITU_DT, Format(Text1(ptxS_DATE).Text, "YYYYMMDD"))
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
        com = BtOpGetGreaterEqual
        LCNT = 99
    
    
        Do
            DoEvents
                                                '印刷中断要求
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '画面項目ロック解除
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "消費リスト印刷中断", Me.hwnd, 0)
                Command1.Visible = False
                Print_Proc = False
                Exit Function
            End If
    
            
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> Print_Jgyobu_T(i) Then
                        Exit Do
                    End If
                    
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Format(Text1(ptxE_DATE).Text, "YYYYMMDD") Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫移動歴")
                    Exit Function
            End Select
    
    
            Print_F = False
            For j = 0 To UBound(Print_Yoin_T)
                If Print_Yoin_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    Print_F = True
                    Exit For
                End If
            
            Next j
    
    
    
            If Print_F Then
    
                'ヘッダーコントロール
                If LCNT > LMAX Then
                    Call Print_Head(LCNT)
                End If
        
        
        
                '事業部
                For j = 0 To UBound(JGYOBU_T)
                    If JGYOBU_T(j).CODE = StrConv(IDOREC.JGYOBU, vbUnicode) Then
                        Exit For
                    End If
                Next j
        
        
                Printer.Print Tab(MGN_L);
                If j <= UBound(JGYOBU_T) Then
                    Call Moji_Cut_Proc(JGYOBU_T(j).NAME, RetBuf, 10)
                    Printer.Print RetBuf;
                End If
                Printer.Print Tab(MGN_L + 10);
                Printer.Print Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2);
        
                Printer.Print Tab(MGN_L + 22);
                Printer.Print Left(StrConv(IDOREC.HIN_GAI, vbUnicode), 14);
                Printer.Print Tab(MGN_L + 37);
                Call Moji_Cut_Proc(StrConv(IDOREC.HIN_NAME, vbUnicode), RetBuf, 35)
                Printer.Print RetBuf;
                
                Printer.Print Tab(MGN_L + 75);
                Printer.Print StrConv(IDOREC.FROM_SOKO, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_RETU, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_REN, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_DAN, vbUnicode);
                RetBuf = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)), "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print Tab(MGN_L + 90);
                Printer.Print RetBuf;
                
                RetBuf = Format(CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
                Printer.Print Tab(MGN_L + 100);
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf;
                
                Printer.Print Tab(MGN_L + 112);
                Call Moji_Cut_Proc(StrConv(IDOREC.MEMO, vbUnicode), RetBuf, 20)
                Printer.Print RetBuf;
                LCNT = LCNT + 1
            End If
            com = BtOpGetNext
    
        Loop
    Next i

    If LCNT <> 99 Then
        Printer.EndDoc
    End If
    
    
    Call Input_UnLock         '画面項目ロック解除
    Command1.Visible = False


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "消費リスト印刷終了", Me.hwnd, 0)

    Print_Proc = False
End Function

Private Sub Print_Head(LCNT As Integer)
                                        
Dim i As Integer
Dim RetBuf As String
Dim sts As Integer

    If LCNT <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        'ヘッダー（１）
    Printer.Print Tab(3);
    Printer.Print Tab(36);
    Printer.Print "＊＊＊  消費リスト  ＊＊＊";
    Printer.Print Tab(100);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    Printer.Print
                                        '明細印刷
    Printer.Print Tab(MGN_L);
    Printer.Print "事業部";
    Printer.Print Tab(MGN_L + 10);
    Printer.Print "消費日";
    Printer.Print Tab(MGN_L + 22);
    Printer.Print "品  番";
    Printer.Print Tab(MGN_L + 37);
    Printer.Print "　　品             名";
    Printer.Print Tab(MGN_L + 75);
    Printer.Print "棚    番";
    Printer.Print Tab(MGN_L + 92);
    Printer.Print "商済み";
    Printer.Print Tab(MGN_L + 102);
    Printer.Print "未商品";
    Printer.Print Tab(MGN_L + 112);
    Printer.Print "メ　モ(指図票№)"
    
    Printer.Print

    LCNT = 6 + MGN_U

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1070301.MousePointer = vbHourglass

    Call Ctrl_Lock(F1070301)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1070301)


    F1070301.MousePointer = vbDefault

End Sub
Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ＣＳＶ出力
        Case 7
            If Not IsDate(Text1(ptxS_DATE).Text) Then
                Text1(ptxS_DATE).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Not IsDate(Text1(ptxE_DATE).Text) Then
                Text1(ptxE_DATE).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Text1(ptxS_DATE).Text > Text1(ptxE_DATE).Text Then
                MsgBox "入力した項目はエラーです。（日付範囲）"
                Text1(ptxS_DATE).SetFocus
                Exit Sub
            End If
            
            
            Beep
            ans = MsgBox("「消費リスト」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Output_Proc() Then
                    Unload Me
                End If
                Text1(ptxS_DATE).SetFocus
            End If
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ＣＳＶ出力
        Case 8                              '印刷
            
            If Not IsDate(Text1(ptxS_DATE).Text) Then
                Text1(ptxS_DATE).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Not IsDate(Text1(ptxE_DATE).Text) Then
                Text1(ptxE_DATE).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Text1(ptxS_DATE).Text > Text1(ptxE_DATE).Text Then
                MsgBox "入力した項目はエラーです。（日付範囲）"
                Text1(ptxS_DATE).SetFocus
                Exit Sub
            End If
            
            
            Beep
            ans = MsgBox("「消費リスト」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
                Text1(ptxS_DATE).SetFocus
            End If
                    
        Case 11                             '終了
            Unload Me
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
Dim i   As Integer
Dim c   As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
    

    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "消費リスト印刷", Me.hwnd, 0)
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
    LOG_F = Trim(c)
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                
                                
                                '印刷対象事業部
    If GetIni(App.EXEName, "JGYOBU_CODE", App.EXEName, c) Then
        MsgBox "印刷対象事業部の獲得に失敗しました(JGYOBU_CODE=)。処理を中止します。"
        End
    Else
        Print_Jgyobu = Split(Trim(c), ",", -1)
        Erase Print_Jgyobu_T
        
        For i = 0 To UBound(Print_Jgyobu)
        
            ReDim Preserve Print_Jgyobu_T(0 To i)
            Print_Jgyobu_T(i) = Print_Jgyobu(i)
        Next i
        
        
    End If
                                '印刷対象要因
    If GetIni(App.EXEName, "YOIN_CODE", App.EXEName, c) Then
        MsgBox "印刷対象要因の獲得に失敗しました(YOIN_CODE=)。処理を中止します。"
        End
    Else
        Print_Yoin = Split(Trim(c), ",", -1)
    
        Erase Print_Yoin_T
        
        For i = 0 To UBound(Print_Yoin)
        
            ReDim Preserve Print_Yoin_T(0 To i)
            Print_Yoin_T(i) = Print_Yoin(i)
        Next i
    
    
    End If
                                
                                'ＣＳＶﾌｧｲﾙ
    If GetIni(App.EXEName, "F107030CSV", App.EXEName, c) Then
    Else
        F107030CSV = Trim(c)
        Command(7).Enabled = True
    End If
                                
                                
                                
                                
                                
                                
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F1070301.FontName
        .Size = F1070301.FontSize
    End With
    Set Printer.Font = NormalFont
    
    Text1(ptxS_DATE).Text = Format(Now, "YYYY/MM/DD")
    Text1(ptxE_DATE).Text = Format(Now, "YYYY/MM/DD")
    
    Text1(ptxS_DATE).SetFocus

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
Dim yn  As Integer
    
    
    
    yn = MsgBox("[消費リスト印刷]処理を終了しますか？", vbYesNo, "確認入力")
    If yn = vbNo Then
        Cancel = True
        Exit Sub
    End If
    
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
    
    sts = BTRV(BtOpReset, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1070301 = Nothing

    End
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
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

Private Function Output_Proc() As Integer
    
Dim sts             As Integer
Dim com             As Integer
    
Dim LCNT            As Integer

Dim i               As Integer
Dim j               As Integer


Dim RetBuf          As String

Dim Print_F         As Boolean

Dim FileNo          As Integer


    Output_Proc = True


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "消費リストデータ出力中", Me.hwnd, 0)


'印刷中は「印刷中断」以外のイベント取得不可
    Call Input_Lock         '画面項目ロック
    Command1.Visible = True
    Command1.Enabled = True


    Pdate = Date
    Ptime = Time


    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    
    Open (F107030CSV) For Output As FileNo



    LCNT = 99

    PRT_CAN = False

    For i = 0 To UBound(Print_Jgyobu_T)
        Call UniCode_Conv(K0_IDO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_IDO.JITU_DT, Format(Text1(ptxS_DATE).Text, "YYYYMMDD"))
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
        com = BtOpGetGreaterEqual
    
    
        Do
            DoEvents
                                                '印刷中断要求
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '画面項目ロック解除
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "消費リストデータ出力中断", Me.hwnd, 0)
                Command1.Visible = False
                Output_Proc = False
                Exit Function
            End If
    
            
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> Print_Jgyobu_T(i) Then
                        Exit Do
                    End If
                    
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Format(Text1(ptxE_DATE).Text, "YYYYMMDD") Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫移動歴")
                    Exit Function
            End Select
    
    
            Print_F = False
            For j = 0 To UBound(Print_Yoin_T)
                If Print_Yoin_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    Print_F = True
                    Exit For
                End If
            
            Next j
    
    
    
            If Print_F Then
    
                'ヘッダーコントロール
                If LCNT = 99 Then
                    Write #FileNo, "事業部", "消費日", "品番", "品名", "棚番", "商済み", "未商品", "メモ（指図票№）"
                    LCNT = 0
                End If
                '事業部
                For j = 0 To UBound(JGYOBU_T)
                    If JGYOBU_T(j).CODE = StrConv(IDOREC.JGYOBU, vbUnicode) Then
                        Exit For
                    End If
                Next j
                If j <= UBound(JGYOBU_T) Then
                    Write #FileNo, RTrim(JGYOBU_T(j).NAME),
                End If
                '消費日
                Write #FileNo, Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2),
        
                '品番
                Write #FileNo, RTrim(StrConv(IDOREC.HIN_GAI, vbUnicode)),
                '品名
                Write #FileNo, RTrim(StrConv(IDOREC.HIN_NAME, vbUnicode)),
                '棚番
                Write #FileNo, StrConv(IDOREC.FROM_SOKO, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_RETU, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_REN, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_DAN, vbUnicode),
                '商品化済み
                Write #FileNo, Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)), "#,##0"),
                '未商品
                Write #FileNo, Format(CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0"),
                'メモ（指図票№）
                Write #FileNo, RTrim(StrConv(IDOREC.MEMO, vbUnicode)),
                                    
                Write #FileNo,
            
            End If
            com = BtOpGetNext
    
        Loop
    Next i

    
    
    Close #FileNo
    
    MsgBox "「" & F107030CSV & "」は正常に出力されました。"
    
    
    
    
    Call Input_UnLock         '画面項目ロック解除
    Command1.Visible = False

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "消費リストデータ出力終了", Me.hwnd, 0)

    Output_Proc = False
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox F107030CSV & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        Output_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
    End If

End Function


