VERSION 5.00
Begin VB.Form SEK00401 
   BackColor       =   &H00C0C0C0&
   Caption         =   "邸別注文データ繰越処理 [SEK0040] 2011.06.29 14:00"
   ClientHeight    =   4704
   ClientLeft      =   1920
   ClientTop       =   2436
   ClientWidth     =   7860
   ControlBox      =   0   'False
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
   MousePointer    =   11  '砂時計
   ScaleHeight     =   4704
   ScaleWidth      =   7860
   StartUpPosition =   2  '画面の中央
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "邸別注文データ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "「繰越」更新中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   3840
   End
End
Attribute VB_Name = "SEK00401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'「ﾊﾟﾗﾒｰﾀ1 0:画面確認無し(デフォルト)　1;有り」
'「ﾊﾟﾗﾒｰﾀ2 0:未照合は残す(デフォルト)　1:未照合でも削除」
'「ﾊﾟﾗﾒｰﾀ3 0:未検品は残す(デフォルト)　1:未検品でも削除」
'「ﾊﾟﾗﾒｰﾀ4 未定義(デフォルト)　YYYYMMDD:定義した場合その日以前のデータ作成日分を削除」

Private Option_Mode As Variant


Private Sub Form_Activate()
Dim ans As Integer

    
    
    If Option_Mode(0) = 1 Then
                        '手動実行
        Beep
        ans = MsgBox("「邸別注文データ繰越処理　実行しますか？", vbYesNo + vbDefaultButton2, "確認処理")
        If ans = vbYes Then
            
            
            Call Y_SYU_TEI_DEL_PROC
        End If
    
    Else
                        '自動実行
        Call Y_SYU_TEI_DEL_PROC
    End If

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

Dim c As String
Dim Today       As String * 8


    c = Command

    
    Today = Format(Now, "YYYYMMDD")
''    Today = "99999999"
    
    If Trim(c) = "" Then
        c = "0,0,0," & Today
    End If


    If Len(Trim(c)) = 1 Then
        c = Trim(c) & ",0,0," & Today
    End If

    If Len(Trim(c)) = 3 Then
        c = Trim(c) & ",0," & Today
    End If

    If Len(Trim(c)) = 5 Then
        c = Trim(c) & "," & Today
    End If


    Option_Mode = Split(Trim(c), ",", -1)


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
    LOG_F = RTrim(c)
                                


    Show
End Sub
Private Sub Form_Unload(CANCEL As Integer)
    
    Set SEK00401 = Nothing
        
    End
End Sub

Private Sub Y_SYU_TEI_DEL_PROC()

Dim sts         As Integer
Dim com         As Integer
        
Dim ans         As Integer
        
Dim Undo        As Boolean
Dim i           As Integer
        
        
        
        
    MsgLab(0).Visible = True
    MsgLab(1).Visible = True
        
    DoEvents
    If Y_SYU_TEI_Open(BtOpenNomal) Then                 '出荷予定データ
        Exit Sub
    End If
        
    If DEL_SYU_TEI_Open(BtOpenNomal) Then               '削除済出荷予定データ
        Exit Sub
    End If
        
    com = BtOpGetFirst

    Do
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Do
                    End If
                
                    com = BtOpGetEqual
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "邸別注文データ")
                    Exit Do
            End Select
        Loop
    
        If sts = BtNoErr Then
            
            
            'データ作成日チェック
            If Option_Mode(3) >= StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode) Then
                
                Undo = False
                
                '照合チェック
                If Option_Mode(1) = 0 Then
                    If Trim(StrConv(Y_SYU_TEI_REC.SHOGO_TANTO, vbUnicode)) = "" Then
                        Undo = True
                    End If
                End If
                '検品チェック
                If Option_Mode(2) = 0 Then
                    If Trim(StrConv(Y_SYU_TEI_REC.KENPIN_TANTO, vbUnicode)) = "" Then
                        Undo = True
                    End If
                End If
            
            
                If StrConv(Y_SYU_TEI_REC.CANCEL_F, vbUnicode) = "1" Then
                    Undo = False
                End If
            
                If Undo Then
                Else
                    Do
                        DoEvents
                        sts = BTRV(BtOpInsert, DEL_SYU_TEI_POS, Y_SYU_TEI_REC, Len(DEL_SYU_TEI_REC), K0_DEL_SYU_TEI, Len(K0_DEL_SYU_TEI), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<DEL_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Do
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "削除済邸別注文データ")
                                Exit Do
                        End Select
                    Loop
                        
                    Do
                        DoEvents
                        sts = BTRV(BtOpDelete, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Do
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "邸別注文データ")
                                Exit Do
                        End Select
                    Loop
                    
                    
                End If
            End If
        End If
            
        If sts <> BtNoErr Then
            Exit Do
        End If
            
        com = BtOpGetNext
    
    Loop
        
                                                    '邸別注文データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "邸別注文データ")
    End If
                                                    '削除済出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, DEL_SYU_TEI_POS, DEL_SYU_TEI_REC, Len(DEL_SYU_TEI_REC), K0_DEL_SYU_TEI, Len(K0_DEL_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "削除済邸別注文データ")
    End If
        
    
    
    sts = BTRV(BtOpReset, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

End Sub

