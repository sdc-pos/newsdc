VERSION 5.00
Begin VB.Form F1010801 
   BackColor       =   &H00FFFFFF&
   Caption         =   "担当者マスタメンテナンス"
   ClientHeight    =   11955
   ClientLeft      =   2130
   ClientTop       =   2430
   ClientWidth     =   17055
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
   ScaleHeight     =   11955
   ScaleWidth      =   17055
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   20
      Top             =   960
      Width           =   372
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   18
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   1
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   15
      Top             =   960
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   9150
      ItemData        =   "F1010801.frx":0000
      Left            =   2160
      List            =   "F1010801.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   5415
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   0
      Top             =   960
      Width           =   735
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   11160
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
      Top             =   11160
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
      Top             =   11160
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
      Index           =   8
      Left            =   7800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   11160
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   11160
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
      Top             =   11160
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
      Top             =   11160
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
      Top             =   11160
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "削  除"
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
      Top             =   11160
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
      Top             =   11160
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
      Top             =   11160
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   11160
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(空白：出勤予定表示対象外　以外：対象)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6840
      TabIndex        =   21
      Top             =   1440
      Width           =   3732
   End
   Begin VB.Label Label 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "区分"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   6840
      TabIndex        =   19
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "所属"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   6000
      TabIndex        =   17
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "担当者名称"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   3240
      TabIndex        =   16
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "担当者"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   2160
      TabIndex        =   14
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "F1010801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxTANTO_CODE% = 0
Private Const ptxTANTO_NAME% = 1
Private Const ptxPOST_CODE% = 2
'2011.09.06
Private Const ptxKUBUN% = 3



Private Const Text_Max% = 3                     '画面項目別最大ｲﾝﾃﾞｯｸｽ
Private Const Command_Max% = 11

Private TANTO_CSV As String

'Private Const LAST_UPDATE_DAY$ = "[F101080] 2011.09.30 10:00 [商品化ｼｽﾃﾑ対応]" 2011.09.06
Private Const LAST_UPDATE_DAY$ = "[F101080] 2019.06.25 11:15"  '2019.06.25 画面サイズ拡張



Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1010801.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010801)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010801)

    F1010801.MousePointer = vbDefault

End Sub
Private Function List_Proc()
'----------------------------------------------------------------------------
'                   担当者マスタ表示
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim com     As Integer
Dim yn      As Integer
Dim Edit    As String


    List_Proc = True
    
    
    List1.Clear
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "担当者マスタ")
                Exit Function
        End Select
        
        Edit = StrConv(TANTOREC.TANTO_CODE, vbUnicode) & "    "
        Edit = Edit & StrConv(TANTOREC.TANTO_NAME, vbUnicode) & "   "
        
        '2011.09.06
'        Edit = Edit & StrConv(TANTOREC.POST_CODE, vbUnicode)
        Edit = Edit & StrConv(TANTOREC.POST_CODE, vbUnicode) & "     "
        Edit = Edit & StrConv(TANTOREC.KUBUN, vbUnicode)
        '2011.09.06
         
        List1.AddItem Edit
         
        com = BtOpGetNext
    Loop
    
    List_Proc = False
    
End Function
Private Sub Clear_Field(Optional Mode As Integer = 0)
'----------------------------------------------------------------------------
'                   画面内容初期設定
'----------------------------------------------------------------------------

Dim i As Integer
    
    For i = Mode To Text_Max
        Text(i) = ""
    Next i

End Sub
Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   入力内容のチェック
'----------------------------------------------------------------------------
    Err_Chk = True
    
    If Len(Text(ptxTANTO_CODE).Text) = 0 Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxTANTO_CODE).SetFocus
        Exit Function
    End If
        
    Err_Chk = False
End Function

Private Sub Item_Dsp()
'----------------------------------------------------------------------------
'                   明細表示
'----------------------------------------------------------------------------
    Text(ptxTANTO_CODE).Text = Trim(StrConv(TANTOREC.TANTO_CODE, vbUnicode))
    Text(ptxTANTO_NAME).Text = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
    Text(ptxPOST_CODE).Text = Trim(StrConv(TANTOREC.POST_CODE, vbUnicode))
    '2011.09.06
    Text(ptxKUBUN).Text = Trim(StrConv(TANTOREC.KUBUN, vbUnicode))

End Sub
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   担当者マスタの追加／修正
'----------------------------------------------------------------------------
Dim sts As Integer
Dim ans As Integer
Dim com As Integer

    Update_Proc = True
    
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTANTO_CODE).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<TANTO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Call Clear_Field(0)
                    Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                Exit Function
        End Select
    Loop
                                            'レコード内容編集
    Call UniCode_Conv(TANTOREC.TANTO_CODE, Text(ptxTANTO_CODE).Text)
    Call UniCode_Conv(TANTOREC.TANTO_NAME, Text(ptxTANTO_NAME).Text)
    Call UniCode_Conv(TANTOREC.POST_CODE, Text(ptxPOST_CODE).Text)
    
    '2011.09.06
    Call UniCode_Conv(TANTOREC.KUBUN, Text(ptxKUBUN).Text)
    '2011.09.06
    
    
    Call UniCode_Conv(TANTOREC.FILLER, "")

    Do
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                
                Beep
                ans = MsgBox("他端末でデータ使用中です。<TANTO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                    Call Clear_Field(0)
                    Update_Proc = False
                    Exit Function
                End If
                        
            Case Else
                Call File_Error(sts, com, "担当者マスタ")
                Exit Function
        End Select
    Loop
    
    Call List_Update_Proc(0)                'リストボックス更新

    Call Clear_Field(0)                     '画面クリアー
    
    Update_Proc = False
End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   担当者マスタの削除
'----------------------------------------------------------------------------
Dim sts As Integer
Dim ans As Integer

    Delete_Proc = True
    
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTANTO_CODE).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Delete_Proc = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<TANTO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Call Clear_Field(0)
                    Delete_Proc = False
                    
                    Exit Function
                
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                Exit Function
        End Select
    Loop
        
    Do
        sts = BTRV(BtOpDelete, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です<MTS.DAT>。", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                    Call Clear_Field(0)
                    Delete_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "担当者マスタ")
                Exit Function
        End Select
    Loop
    
    Call List_Update_Proc(1)                'リストボックス更新
    
    Call Clear_Field(0)                     '画面クリアー

    Delete_Proc = False
End Function


Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0
                                            'エラーチェック
            sts = Err_Chk()
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            Beep
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
            Text(ptxTANTO_CODE) = ""
        Case 3
            If Trim(Text(ptxTANTO_CODE).Text) = "" Then
                Beep
                MsgBox "削除するコードを指定して下さい。", vbExclamation
            Else
                Beep
                yn = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
                If yn = vbYes Then
                    If Delete_Proc() Then
                        Unload Me
                    End If
                End If
            End If
        Case 8                  'データ出力
            Beep
            yn = MsgBox("データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Data_Proc() Then
                    Unload Me
                End If
            End If
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
    
    Text(ptxTANTO_CODE).SetFocus

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
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                'ＣＳＶファイル名取り込み
    If GetIni("FILE", "TANTO_CSV", "SYS", c) Then
        Beep
        MsgBox "担当者マスタデータ出力用ファイル[TANTO_CSV]の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    
    TANTO_CSV = Trim(c)
    Me.Caption = Me.Caption & " " & LAST_UPDATE_DAY '2019.06.25 タイトルバー用でF101080→me.に変更
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
    
    Call List_Proc
    Text(ptxTANTO_CODE).SetFocus
    
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
    sts = BTRV(BtOpReset, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "担当者マスタ")
    End If
    Set F1010801 = Nothing
    End
End Sub

Private Sub List1_DblClick()
Dim sts     As Integer

    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Left(List1.List(List1.ListIndex), 5))
    
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            MsgBox "マスタ内容が変更されています。最新情報を再表示します。"
            If List_Proc() Then
                Unload Me
            End If
            Exit Sub
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Unload Me
    End Select
    
    Call Item_Dsp
    Text(ptxTANTO_CODE).SetFocus
        
End Sub


Private Sub List1_GotFocus()
    
    If List1.ListCount > 0 Then
        If List1.ListIndex <= 0 Then
            List1.ListIndex = 0
        End If
    End If

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
            
        Call List1_DblClick
    
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
Dim RetBuf As String
Dim sts As Integer
Dim i   As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub

    Select Case Index
        
        Case ptxTANTO_CODE
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTANTO_CODE).Text)
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call Clear_Field(1)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Unload Me
            End Select
    
    End Select
    For i = Index + 1 To Text_Max
        If Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i
End Sub

Private Sub List_Update_Proc(Mode As Integer)
'----------------------------------------------------------------------------
'                   リストボックス更新
'----------------------------------------------------------------------------
Dim i       As Integer
Dim Edit    As String


    For i = 0 To List1.ListCount - 1
        
        
        If Trim(Text(ptxTANTO_CODE).Text) = Trim(Left(List1.List(i), 5)) Then
            List1.RemoveItem i
        End If
    
    Next i

    If Mode = 0 Then
        Edit = StrConv(TANTOREC.TANTO_CODE, vbUnicode) & "    "
        Edit = Edit & StrConv(TANTOREC.TANTO_NAME, vbUnicode) & "   "
        
        '2011.09.06
'        Edit = Edit & StrConv(TANTOREC.POST_CODE, vbUnicode)
        Edit = Edit & StrConv(TANTOREC.POST_CODE, vbUnicode) & "     "
        Edit = Edit & StrConv(TANTOREC.KUBUN, vbUnicode)
        '2011.09.06
        
        List1.AddItem Edit
    End If
End Sub
Private Function Data_Proc() As Integer

Dim FileNo          As Integer
Dim FileName        As String
Dim ret             As Integer

Dim com             As Integer
Dim sts             As Integer

    Call Input_Lock

    FileNo = FreeFile
    FileName = TANTO_CSV
    
    ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), ret) & Right(Trim(FileName), Len(Trim(FileName)) - ret)

    On Error GoTo Error_Proc

    Open (FileName) For Output As FileNo
    
    '2011.09.06
'    Write #FileNo, "担当者ｺｰﾄﾞ", "担当者名称", "所属"
    Write #FileNo, "担当者ｺｰﾄﾞ", "担当者名称", "所属", "区分"
    '2011.09.06

    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "担当者マスタ")
                Exit Function
        End Select
    
        Write #FileNo, StrConv(TANTOREC.TANTO_CODE, vbUnicode),
        Write #FileNo, StrConv(TANTOREC.TANTO_NAME, vbUnicode),
        
        '2011.09.06
'        Write #FileNo, StrConv(TANTOREC.POST_CODE, vbUnicode)
        Write #FileNo, StrConv(TANTOREC.POST_CODE, vbUnicode),
        Write #FileNo, StrConv(TANTOREC.KUBUN, vbUnicode)
        '2011.09.06
    
    
        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock
    
    Beep
    MsgBox "「" & FileName & "」は正常に出力されました。"


    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName & "が使用中です。"
        Data_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        Data_Proc = True
    End If

    Call Input_UnLock



End Function


