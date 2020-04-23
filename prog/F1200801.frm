VERSION 5.00
Begin VB.Form F1200801 
   BackColor       =   &H8000000A&
   Caption         =   "仮想在庫廃棄処理([F120080] 2012.04.14 11:00)"
   ClientHeight    =   7860
   ClientLeft      =   2130
   ClientTop       =   2835
   ClientWidth     =   11595
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
   ScaleHeight     =   7860
   ScaleWidth      =   11595
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox CmbSoko_No 
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      ItemData        =   "F1200801.frx":0000
      Left            =   4305
      List            =   "F1200801.frx":0002
      TabIndex        =   15
      Top             =   2280
      Width           =   2010
   End
   Begin VB.ComboBox CmbSoko_No 
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      ItemData        =   "F1200801.frx":0004
      Left            =   4305
      List            =   "F1200801.frx":0006
      TabIndex        =   0
      Top             =   1650
      Width           =   2850
   End
   Begin VB.CommandButton Command 
      Caption         =   "終　了"
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
      Index           =   11
      Left            =   10425
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   10
      Left            =   9585
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   9
      Left            =   8745
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Cancel          =   -1  'True
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
      Left            =   7905
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   7
      Left            =   6585
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   6
      Left            =   5745
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   5
      Left            =   4905
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   4
      Left            =   4065
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   2745
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   2
      Left            =   1905
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   1
      Left            =   1065
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "実  行"
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
      Left            =   225
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
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
      Left            =   4305
      TabIndex        =   17
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000A&
      Caption         =   "出庫件数→"
      Height          =   255
      Index           =   1
      Left            =   2730
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000A&
      Caption         =   "出庫要因"
      Height          =   255
      Index           =   0
      Left            =   3045
      TabIndex        =   14
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000A&
      Caption         =   "倉庫"
      Height          =   255
      Index           =   5
      Left            =   3585
      TabIndex        =   13
      Top             =   1800
      Width           =   615
   End
End
Attribute VB_Name = "F1200801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private WS_NO As String * 3



Private Sub CmbSoko_No_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)


    If KeyCode <> vbKeyReturn And KeyCode <> vbKeyTab Then
        Exit Sub
    End If


End Sub

Private Sub Command_Click(Index As Integer)

Dim ans     As Integer
Dim sts     As Integer
    
Dim Ren     As Integer
Dim Dan     As Integer
    
    Select Case Index
        
        Case 0                  '保存
            
            ans = MsgBox("指定の要因で出庫処理を行いますか？", vbYesNo, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            
            
            End If
            

        Case 11                 '終了
    
            Unload Me
    
    End Select

End Sub

Private Sub Form_DblClick()
'    PrintForm
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

Dim c       As String * 128
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


'端末番号取り込み
                                'ﾜｰｸｽﾃｰｼｮﾝ番号取り込み
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)


                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
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
                                '品目マスタ(ワーク)ＯＰＥＮ
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '移動歴データＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '倉庫情報をコンボにセットする
    If Soko_Tbl_Set() Then
        Unload Me
    End If
                                '出庫要因をコンボにセットする
    If Yoin_Tbl_Set() Then
        Unload Me
    End If

    CmbSoko_No(0).SetFocus




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
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
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
                                            '品目マスタ(ワーク)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ(ワーク)")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '移動歴データＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1200801 = Nothing

    End




End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1200801.MousePointer = vbHourglass

    Call Ctrl_Lock(F1200801)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1200801)


    F1200801.MousePointer = vbDefault

End Sub
Private Function Soko_Tbl_Set() As Integer
'------------------------------------   倉庫名称＆ｺｰﾄﾞをコンボにセットする
Dim sts As Integer
Dim com As Integer


    Soko_Tbl_Set = True
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "倉庫マスタ")
        End Select
    
    
    
    
        If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
            
            CmbSoko_No(0).AddItem StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "[" & StrConv(SOKOREC.Soko_No, vbUnicode) & "]"
        
        End If
    
    
        com = BtOpGetNext
    
    Loop

    CmbSoko_No(0).ListIndex = 0




    Soko_Tbl_Set = False



End Function
Private Function Yoin_Tbl_Set() As Integer
'------------------------------------   出庫要因名称＆ｺｰﾄﾞをコンボにセットする
Dim sts As Integer
Dim com As Integer


    Yoin_Tbl_Set = True
    
    
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, ACT_ZAITEI_OUT)
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, "")
    
    
    
    com = BtOpGetGreater
    Do
        sts = BTRV(com, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
        Select Case sts
            Case BtNoErr
            
            
                If StrConv(YOINREC.CODE_TYPE, vbUnicode) <> ACT_ZAITEI_OUT Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "要因マスタ")
        End Select
    
            
            
            
        CmbSoko_No(1).AddItem StrConv(YOINREC.YOIN_DNAME, vbUnicode) & "[" & StrConv(YOINREC.CODE_TYPE, vbUnicode) & StrConv(YOINREC.YOIN_CODE, vbUnicode) & "]"
    
    
    
        com = BtOpGetNext
    
    Loop

    CmbSoko_No(1).ListIndex = 0




    Yoin_Tbl_Set = False



End Function


Private Function Update_Proc() As Integer
'---------------------------------- 一括出庫処理
Dim sts         As Integer
Dim ans         As Integer


Dim JGYOBU      As String * 1
Dim NAIGAI      As String * 1
Dim HIN_GAI     As String * 13
Dim NYUKA_DT    As String * 8
Dim LOCATION    As String * 8
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim Upd_Cnt     As Long



    Update_Proc = True


    Upd_Cnt = 0

    
    Do
        
        DoEvents
        
        Call UniCode_Conv(K0_ZAIKO.Soko_No, Mid(Right(CmbSoko_No(0), 4), 2, 2))
        Call UniCode_Conv(K0_ZAIKO.Retu, "01")
        Call UniCode_Conv(K0_ZAIKO.Ren, "01")
        Call UniCode_Conv(K0_ZAIKO.Dan, "01")
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, "")
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, "")
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, "")
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
        
        
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetGreater, "在庫データ")
                Exit Function
        End Select
    
        If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> Mid(Right(CmbSoko_No(0), 4), 2, 2) Then
            
            Exit Do
        End If
    
    
        If StrConv(ZAIKOREC.LOCK_F, vbUnicode) = LOCK_ON And _
            (Trim(StrConv(ZAIKOREC.WEL_ID, vbUnicode)) <> WS_NO Or _
            Trim(StrConv(ZAIKOREC.PRG_ID, vbUnicode)) <> App.EXEName) Then
            
            
            ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
            If ans = vbCancel Then
                Exit Function
            End If
        Else
                                            'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            If sts <> BtNoErr Then
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
                Exit Function
            End If
    
            JGYOBU = StrConv(ZAIKOREC.JGYOBU, vbUnicode)
            NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
            HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
            NYUKA_DT = StrConv(ZAIKOREC.NYUKA_DT, vbUnicode)
            LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
            SUMI_QTY = 0
            MI_QTY = 0
            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            Else
                MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            End If

            sts = Zaiko_Lock_Proc(StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode), _
                            StrConv(ZAIKOREC.JGYOBU, vbUnicode), _
                            StrConv(ZAIKOREC.NAIGAI, vbUnicode), _
                            StrConv(ZAIKOREC.HIN_GAI, vbUnicode), _
                            WS_NO)
            Select Case sts
                Case False
                Case True, SYS_CANCEL
                    GoTo Abort_Tran
                Case SYS_ERR
                    GoTo Abort_Tran
            End Select
    


            sts = Syuko_Update_Proc(JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    NYUKA_DT, _
                                    (Mid(Right(CmbSoko_No(0), 4), 2, 2) & "01" & "01" & "01"), _
                                    Mid(Right(CmbSoko_No(1), 4), 2, 2), _
                                    SUMI_QTY, _
                                    MI_QTY, _
                                    0, _
                                    WS_NO, _
                                    WS_NO, , _
                                    "仮想倉庫一括出庫")
            Select Case sts
                Case False
                Case Else
                    GoTo Abort_Tran
            End Select
    
    
    
            sts = BTRV(BtOpEndTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            If sts <> BtNoErr Then
                Call File_Error(sts, BtOpEndTransaction, "")
                GoTo Abort_Tran
            End If
    
    
            Upd_Cnt = Upd_Cnt + 1
    
            Label2.Caption = Format(Upd_Cnt, "#,##0")
    
        End If
    Loop


    MsgBox "出庫処理は正常終了しました。"




    Update_Proc = False
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

End Function
