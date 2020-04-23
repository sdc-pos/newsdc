VERSION 5.00
Begin VB.Form F1010701 
   BackColor       =   &H00FFFFFF&
   Caption         =   "[作業管理マスタ]担当者別メニュー登録"
   ClientHeight    =   6300
   ClientLeft      =   2130
   ClientTop       =   2430
   ClientWidth     =   11280
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
   MaxButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11280
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      ItemData        =   "F1010701.frx":0000
      Left            =   2760
      List            =   "F1010701.frx":0007
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   16
      Top             =   600
      Width           =   3375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      ItemData        =   "F1010701.frx":0021
      Left            =   6240
      List            =   "F1010701.frx":0028
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   15
      Top             =   600
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   4140
      ItemData        =   "F1010701.frx":0042
      Left            =   2760
      List            =   "F1010701.frx":0044
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   6015
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
      Left            =   10320
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
      Index           =   10
      Left            =   9480
      TabIndex        =   11
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "共 通"
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
      Index           =   7
      Left            =   6480
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "MENU"
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "設 定"
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "更  新"
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label lblALL_Sel 
      BorderStyle     =   1  '実線
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "メニュー"
      Height          =   240
      Index           =   0
      Left            =   6240
      TabIndex        =   14
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "担当者"
      Height          =   240
      Index           =   2
      Left            =   2760
      TabIndex        =   13
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "F1010701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcmbTANTO% = 0
Private Const pcmbMENU% = 1

Private Const Command_Max% = 11

Private Const MENU_NON$ = "**"
Private Const MENU_NON_N$ = "なし　　　　　　　　"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1010701.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010701)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010701)

    F1010701.MousePointer = vbDefault

End Sub
Private Function List_Proc()
'----------------------------------------------------------------------------
'                   担当者別メニュー表示
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim com     As Integer
Dim ans     As Integer
Dim Edit    As String


    List_Proc = True
    
    
    List1.Clear
    
    com = BtOpGetFirst
    Do
        '担当者マスタ読み込み
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "担当者マスタ")
                Exit Function
        End Select
        '担当者別メニュー読み込み
        Call UniCode_Conv(K0_TMENU.TANTO_CODE, StrConv(TANTOREC.TANTO_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(TMENUREC.MENU_GRP_NO, MENU_NON)
                Call UniCode_Conv(MENUREC.MENU_GRP, MENU_NON_N)
            Case Else
                Call File_Error(sts, BtOpGetGreater, "担当者別メニュー")
                Exit Function
        End Select
        'メニュー管理読み込み
        If Trim(StrConv(TMENUREC.MENU_GRP_NO, vbUnicode)) = MENU_NON Then
        Else
            Call UniCode_Conv(K0_MENU.MENU_GRP_NO, StrConv(TMENUREC.MENU_GRP_NO, vbUnicode))
            Call UniCode_Conv(K0_MENU.MENU_LV1, "")
            Call UniCode_Conv(K0_MENU.MENU_LV2, "")
            Call UniCode_Conv(K0_MENU.MENU_LV3, "")
            
            sts = BTRV(BtOpGetGreater, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(MENUREC.MENU_GRP_NO, vbUnicode) <> StrConv(TMENUREC.MENU_GRP_NO, vbUnicode) Then
                    
                        Call UniCode_Conv(TMENUREC.MENU_GRP_NO, MENU_NON)
                        Call UniCode_Conv(MENUREC.MENU_GRP, MENU_NON_N)
                    
                    End If
                Case BtErrEOF
                    Call UniCode_Conv(TMENUREC.MENU_GRP_NO, MENU_NON)
                    Call UniCode_Conv(MENUREC.MENU_GRP, MENU_NON_N)
                Case Else
                    Call File_Error(sts, com, "メニュー管理マスタ")
                    Exit Function
            End Select
        
        End If
    
        Edit = StrConv(TANTOREC.TANTO_CODE, vbUnicode) & " "
        Edit = Edit & StrConv(TANTOREC.TANTO_NAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MENUREC.MENU_GRP, vbUnicode) & "     "
        Edit = Edit & StrConv(TMENUREC.MENU_GRP_NO, vbUnicode)
    
        List1.AddItem Edit
    
    
        com = BtOpGetNext
    Loop
    
    List_Proc = False
    
End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   担当者別メニューの更新
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
    
Dim Edit    As String
Dim i       As Integer

    
    Update_Proc = True
        
    Call Input_Lock
        
        
    Do                      '全件削除後。再構築
        DoEvents
        Do
            sts = BTRV(BtOpGetFirst + BtSNoWait, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
            Select Case sts
                Case BtNoErr
                    Do
                        sts = BTRV(BtOpDelete, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "担当者別メニュー")
                                Exit Function
                        End Select
                    Loop
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "担当者別メニュー")
                    Exit Function
            End Select
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
        
    Loop
    
    For i = 0 To List1.ListCount - 1
        DoEvents
        
        Edit = List1.List(i)
        
        If MENU_NON = Right(Edit, 2) Then
        Else
            
            Call UniCode_Conv(TMENUREC.TANTO_CODE, Trim(Left(Edit, 5)))
            Call UniCode_Conv(TMENUREC.MENU_GRP_NO, Trim(Right(Edit, 2)))
            Call UniCode_Conv(TMENUREC.FILLER, "")
    
            Do
                sts = BTRV(BtOpInsert, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "担当者別メニュー")
                        Exit Function
                End Select
            Loop
        End If
    Next i
        
    If List_Proc() Then
        Exit Function
    End If
    
    Call Input_UnLock
    
    Update_Proc = False
End Function

Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0
            Beep
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
        
        Case 1
            
            Call List_Update_Proc
        
        Case 4
        
            F1010702.Show vbModal
            If Form_RTN Then
                Unload Me
            End If
        
        Case 8
            sts = All_Tanto_Chk_Proc()
            Select Case sts
                Case False  '現在無し
                
                    Beep
                    yn = MsgBox("全担当者共通メニューを登録しますか？（担当者個別は無効になります）", vbYesNo + vbQuestion, "確認入力")
                    If yn = vbYes Then
                        If ALL_Update_Proc(0) Then
                            Unload Me
                        End If
                    End If
                
                Case True   '現在有り
                    Beep
                    yn = MsgBox("全担当者共通メニューを削除しますか？（担当者個別を登録して下さい）", vbYesNo + vbQuestion, "確認入力")
                    If yn = vbYes Then
                        If ALL_Update_Proc(1) Then
                            Unload Me
                        End If
                    End If
                
                Case SYS_ERR
                    Unload Me
            End Select
        Case 11
            Unload Me
    End Select

End Sub


Private Sub Form_Activate()

Dim com                 As Integer
Dim sts                 As Integer
Dim Edit                As String
Dim Sv_MENU_GRP_No      As String * 2
                                        
                                        
                                        
                                        'メニュー設定
    Combo(pcmbMENU).Clear
    Combo(pcmbMENU).AddItem MENU_NON_N & " " & MENU_NON
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "メニュー管理マスタ")
                Unload Me
        End Select
        
        If com = BtOpGetFirst Then
            
            Edit = StrConv(MENUREC.MENU_GRP, vbUnicode) & " " & StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
            Combo(pcmbMENU).AddItem Edit
            Sv_MENU_GRP_No = StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
        
        End If
        
        
        If Sv_MENU_GRP_No <> StrConv(MENUREC.MENU_GRP_NO, vbUnicode) Then
            Edit = StrConv(MENUREC.MENU_GRP, vbUnicode) & " " & StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
            Combo(pcmbMENU).AddItem Edit
            Sv_MENU_GRP_No = StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
        End If
        
        com = BtOpGetNext
    
    Loop
    Combo(pcmbMENU).ListIndex = 0
    
    If List_Proc() Then
        Unload Me
    End If
        
    If List1.ListCount = 0 Then
        Combo(pcmbTANTO).SetFocus
    Else
        List1.ListIndex = 0
        List1.SetFocus
    End If

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
Dim com         As Integer

Dim Sv_MENU_GRP As String * 10

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。"
        End
    End If

    
    F1010702.cboJIGYOBU.Clear
    For i = 0 To UBound(JGYOBU_T)
        If Trim(JGYOBU_T(i).CODE) = "" Then
            Exit For
        End If
        F1010702.cboJIGYOBU.AddItem JGYOBU_T(i).NAME & " " & JGYOBU_T(i).CODE
    Next i
    F1010702.cboJIGYOBU.ListIndex = 0
    If F1010702.cboJIGYOBU.ListCount = 1 Then
        F1010702.cboJIGYOBU.Enabled = False
    End If
    
    
    '国内外情報設定
    i = 0
    Do
        i = i + 1
        If GetIni(App.EXEName, "NAIGAI" & Format(i, "0"), "SYS", c) Then
            Exit Do
        End If
        ReDim Preserve NAIGAI_CODE(i - 1)
        NAIGAI_CODE(i - 1) = Trim(c)
    
    Loop
    If i = 1 Then
        Beep
        MsgBox "国内外の獲得に失敗しました。"
        End
    End If
                                
                                
    F1010702.cboNAIGAI.Clear
    For i = 0 To UBound(NAIGAI_CODE)
        
        Select Case NAIGAI_CODE(i)
            Case NAIGAI_NAI
                F1010702.cboNAIGAI.AddItem NAIGAI1 & " " & NAIGAI_CODE(i)
        
            Case NAIGAI_GAI
                F1010702.cboNAIGAI.AddItem NAIGAI2 & " " & NAIGAI_CODE(i)
        End Select
                    
    Next i
    F1010702.cboNAIGAI.ListIndex = 0
    If F1010702.cboNAIGAI.ListCount = 1 Then
        F1010702.cboNAIGAI.Enabled = False
    End If
                                
                                
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'メニューマスタＯＰＥＮ
    If MENU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'メニューマスタ（一時）ＯＰＥＮ
    If tmpMENU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者別メニューＯＰＥＮ
    If TMENU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
    Form_RTN = False
    Load F1010702
    If Form_RTN Then
        Unload Me
    End If
                                        
                                        '共通メニュー有無判定
    Call UniCode_Conv(K0_TMENU.TANTO_CODE, ALL_TANTO_CODE)
    sts = BTRV(BtOpGetEqual, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
    Select Case sts
        
        Case BtNoErr
            lblALL_Sel.Caption = "共通メニュー指定"
        Case BtErrKeyNotFound
            lblALL_Sel.Caption = "個別メニュー指定"
        Case Else
            Call File_Error(sts, com, "メニュー管理マスタ")
            Unload Me
    
    End Select
                                        
                                        '担当者設定
    Combo(pcmbTANTO).Clear
    com = BtOpGetFirst
    Do
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "担当者マスタ")
                Unload Me
        End Select
        Combo(pcmbTANTO).AddItem StrConv(TANTOREC.TANTO_CODE, vbUnicode) & " " & StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        com = BtOpGetNext
    Loop
    
    If Combo(pcmbTANTO).ListCount <> 0 Then
        Combo(pcmbTANTO).ListIndex = 0
    End If
                                        'メニュー設定
    Combo(pcmbMENU).Clear
    Combo(pcmbMENU).AddItem MENU_NON
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "メニュー管理マスタ")
                Unload Me
        End Select
        If com = BtOpGetFirst Then
            Combo(pcmbMENU).AddItem StrConv(MENUREC.MENU_GRP, vbUnicode)
            Sv_MENU_GRP = Trim(StrConv(MENUREC.MENU_GRP, vbUnicode))
        End If
        
        
        If Trim(Sv_MENU_GRP) <> Trim(StrConv(MENUREC.MENU_GRP, vbUnicode)) Then
            Combo(pcmbMENU).AddItem StrConv(MENUREC.MENU_GRP, vbUnicode)
        End If
        
        Sv_MENU_GRP = Trim(StrConv(MENUREC.MENU_GRP, vbUnicode))
        com = BtOpGetNext
    
    Loop
    If Combo(pcmbMENU).ListCount <> 0 Then
        Combo(pcmbMENU).ListIndex = 0
    End If
    
    If List_Proc() Then
        Unload Me
    End If
        
    If List1.ListCount = 0 Then
        Combo(pcmbTANTO).SetFocus
    Else
        List1.ListIndex = 0
        List1.SetFocus
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            'メニュー管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "メニュー管理マスタ")
        End If
    End If
                                            'メニュー管理マスタ（一時ファイル）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "メニュー管理マスタ（一時ファイル）")
        End If
    End If
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '担当者別メニューＣＬＯＳＥ
    sts = BTRV(BtOpClose, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者別メニュー")
        End If
    End If
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "メニュー管理マスタ")
    End If
    Set F1010701 = Nothing
    Set F1010702 = Nothing
    End
End Sub

Private Sub List1_GotFocus()
    
    If List1.ListCount > 0 Then
        If List1.ListIndex <= 0 Then
            List1.ListIndex = 0
        End If
    End If

End Sub

Private Sub List_Update_Proc()
'----------------------------------------------------------------------------
'                   リストボックス更新
'----------------------------------------------------------------------------
Dim i       As Integer
Dim Edit    As String


    For i = 0 To List1.ListCount - 1
        
        If Trim(Left(Combo(pcmbTANTO).Text, 5)) = Trim(Left(List1.List(i), 5)) Then
            List1.RemoveItem i
        End If
    
    Next i
    
    Edit = Combo(pcmbTANTO).Text & "   "
    Edit = Edit & Combo(pcmbMENU).Text
         
         
    List1.AddItem Edit

End Sub

Private Function All_Tanto_Chk_Proc() As Integer
'----------------------------------------------------------------------------
'                   全担当者共通メニュー作成／開放チェック
'----------------------------------------------------------------------------

Dim sts     As Integer
    
    
    Call UniCode_Conv(K0_TMENU.TANTO_CODE, ALL_TANTO_CODE)
    sts = BTRV(BtOpGetEqual, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
    Select Case sts
        Case BtNoErr
            All_Tanto_Chk_Proc = True
        Case BtErrKeyNotFound
            All_Tanto_Chk_Proc = False
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者別メニュー")
            All_Tanto_Chk_Proc = SYS_ERR
    End Select


    
End Function

Private Function ALL_Update_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   全担当者共通メニュー登録／削除
'                   0:追加　1:削除
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim Rec_Flg As Boolean


    ALL_Update_Proc = True

    Call UniCode_Conv(K0_TMENU.TANTO_CODE, ALL_TANTO_CODE)
    
    Do
        
        sts = BTRV(BtOpGetEqual + BtSNoWait, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
        Select Case sts
            Case BtNoErr
                Rec_Flg = True
                Exit Do
            Case BtErrKeyNotFound
                Rec_Flg = False
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "メニュー管理マスタ")
                Exit Function
        End Select
    
    Loop

    Select Case Mode
        Case 0              '追加
            
            lblALL_Sel.Caption = "共通メニュー指定"
            
            Select Case Rec_Flg
                Case True
                
                    sts = BTRV(BtOpUnlock, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
                    If sts Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "メニュー管理マスタ")
                        Exit Function
                    End If
                
                Case False
                    
                    Call UniCode_Conv(TMENUREC.TANTO_CODE, ALL_TANTO_CODE)  '担当者コード
                    Call UniCode_Conv(TMENUREC.MENU_GRP_NO, "")             'メニューグループ
                    Call UniCode_Conv(TMENUREC.FILLER, "")
                    
                    Do
                        sts = BTRV(BtOpInsert, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "メニュー管理マスタ")
                                Exit Function
                        End Select
                    Loop
            
            End Select
        Case 1              '削除
            
            lblALL_Sel.Caption = "個別メニュー指定"
            
            Select Case Rec_Flg
                Case True
                
                    Do
                        sts = BTRV(BtOpDelete, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "メニュー管理マスタ")
                                Exit Function
                        End Select
                    Loop
                
                
                Case False
                    
            End Select
    
    End Select


    ALL_Update_Proc = False

End Function
