VERSION 5.00
Begin VB.Form F1100701 
   BackColor       =   &H00C0C0C0&
   Caption         =   "出荷予定強制完了 "
   ClientHeight    =   4710
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   7320
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
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '画面の中央
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "本日以前の出荷予定データ"
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
      Width           =   5760
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "全件「削除」更新中！"
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
      Width           =   4800
   End
End
Attribute VB_Name = "F1100701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Limit_day   As String           '強制完了猶予期間
Private Shori_Mode  As Integer          '実行モード 0:手動 1:自動

Private JYOGAI_MTS  As Variant          '(ini) 除外倉庫 配列
Private OSAKA_PC    As Boolean          '2006.12.06

'2011.07.27
Private SEK_MTS_TBL As Variant          '積水向け先
Private SEK_MTS_FLG As Boolean          '積水向け先有無
'2011.07.27


Private KENPIN_CHECK    As Integer      '検品のﾁｪｯｸ 2012.11.19


'Private Const LAST_UPDATE_Day$ = "[F110070]2016.07.06 15:30"
Private Const LAST_UPDATE_Day$ = "[F110070]2018.08.30 11:15"




Private Sub Form_Activate()
Dim ans As Integer

    
    
    If Shori_Mode = 0 Then
                        '手動実行
        Beep
        ans = MsgBox("「出荷予定強制削除（繰越し更新）」処理　実行しますか？", vbYesNo + vbDefaultButton2, "確認処理")
        If ans = vbYes Then
            Call Y_SYU_DEL_PROC
        End If
    
    Else
                        '自動実行
        Call Y_SYU_DEL_PROC
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
Dim c As String * 128

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                
                                
   F1100701.Caption = F1100701.Caption & LAST_UPDATE_Day
                                
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
                                
                                '削除対象日付算出
    If GetIni(App.EXEName, "CMPLT_DAY", App.EXEName, c) Then
        Beep
        MsgBox "日付の獲得に失敗しました。処理を中止して下さい。"
'        Call LOG_OUT(LOG_F, "[SYS.INI] [SYSTEM] [CMPLT_DAY] READ ERROR")           '2016.06.30
        Call LOG_OUT(LOG_F, "[F110070.INI] [F110070] [CMPLT_DAY] READ ERROR")        '2016.06.30
        End
    End If
    
    If Not IsNumeric(Trim(c)) Then
                                '取り込み異常は今日とする
        Limit_day = Format(Date, "yyyymmdd")
    Else
        Limit_day = Format(DateAdd("d", -CInt(Trim(c)), Date), "yyyymmdd")
    End If
                                '処理モード取り込み
    If GetIni(App.EXEName, "AUTO", App.EXEName, c) Then
        Shori_Mode = 0          '取り込み異常は手動操作
    Else
        If IsNumeric(Trim(c)) Then
            Shori_Mode = CInt(Trim(c))
        Else
            Shori_Mode = 0      '取り込み異常は手動操作
        End If
    End If
                                '常に消し込み対象のＭＴＳ 取り込み  2005/06/15
    If GetIni(App.EXEName, "MTS", App.EXEName, c) Then
        c = " "
    End If
    JYOGAI_MTS = Split(Trim(c), ",", -1)


                                '大阪ＰＣ？         2006.12.06
    OSAKA_PC = False
    If GetIni(App.EXEName, "OSAKA_PC", App.EXEName, c) Then
    Else
        If Trim(c) = "1" Then
            OSAKA_PC = True
        End If
    End If


'2011.07.27     積水受け先
    
    SEK_MTS_FLG = False
    If GetIni(App.EXEName, "SEK_MTS", App.EXEName, c) Then
        If OSAKA_PC Then                                                                '2016.07.04
            Call LOG_OUT(LOG_F, "[F110070.INI] [F110070] [SEK_MTS] READ ERROR")         '2016.07.04
        End If                                                                          '2016.07.04
    
    
    Else
        SEK_MTS_FLG = True
        SEK_MTS_TBL = Split(Trim(c), ",", -1)
    End If
'2011.07.27

    '-------------------------  未検品のﾁｪｯｸ    2012.11.19
    KENPIN_CHECK = 0
    If GetIni(App.EXEName, "KENPIN_CHECK", App.EXEName, c) Then
    Else
        If Trim(c) = "1" Then
            KENPIN_CHECK = 1
        End If
    End If
    '-------------------------  未検品のﾁｪｯｸ    2012.11.19

    Show
End Sub
Private Sub Form_Unload(CANCEL As Integer)
    
    Set F1100701 = Nothing
        
    End
End Sub

Private Sub Y_SYU_DEL_PROC()

Dim sts         As Integer
Dim com         As Integer
        
Dim ans         As Integer
        
Dim Undo        As Boolean
Dim i           As Integer
        
        
Dim DEN_NO      As String   '2006.12.06
Dim SEQ_NO      As String   '2006.12.06
        
        
Dim Next_Flg    As Boolean  '2009.02.17
        
        
    MsgLab(0).Visible = True
    MsgLab(1).Visible = True
        
    DoEvents
    If Y_SYU_Open(BtOpenNomal) Then                 '出荷予定データ
        Exit Sub
    End If
        
    If DEL_SYU_Open(BtOpenNomal) Then               '削除済出荷予定データ
        Exit Sub
    End If
        
'----------------
    If OSAKA_PC Then
        If Y_SYU_H_Open(BtOpenNomal) Then                 '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ
            Exit Sub
        End If
            
        If DEL_SYU_H_Open(BtOpenNomal) Then               '削除済出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ
            Exit Sub
        End If
    End If

'----------------
        
        
    com = BtOpGetFirst

    Do
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'この時点でのファイル使用中は無限ループとする。キャンセルで異常終了
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Do
                    End If
                
                    com = BtOpGetEqual
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "出荷予定データ")
                    Exit Do
            End Select
        Loop
    
        If sts = BtNoErr Then
            
'            If Limit_day >= StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then            '2018.08.30
            If Limit_day >= StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then         '2018.08.30
                
                
                
                Undo = False
                
                
                                '2009.08.24
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) = SETSUBI Then
                '-------------------------  積水対応    2011.06.29
                    If KENPIN_CHECK_PROC(Undo) Then
                        Exit Do
                    End If
                '-------------------------  積水対応    2011.06.29
                
                
                '-------------------------  未検品のﾁｪｯｸ    2012.11.19
                    If KENPIN_CHECK = 1 Then
                        If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
                            Undo = True
                        End If
                    End If
                '-------------------------  未検品のﾁｪｯｸ    2012.11.19
                
                Else
                
                
                    '貿易のﾁｪｯｸ
                    
                                    
                    If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) = CYU_KBN_BOU Then
                        If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
                            Undo = True
                        End If
                    Else
                        For i = 0 To UBound(JYOGAI_MTS)
                            If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = JYOGAI_MTS(i) Then
                                Exit For
                            End If
                        Next i
                    
                        If UBound(JYOGAI_MTS) >= 0 Then
                            If i > UBound(JYOGAI_MTS) Then
                                If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
                                    Undo = True
                                End If
                            End If
                        End If
                    End If
                    
                    
                    If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then        '2008.01.10
                    
                        '出荷実績ﾃﾞｰﾀに対応 2006.08.11
                        If (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
                            If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> CYU_KBN_BOU Then
                                
                                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" Or StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "3" Then
                                
                                    If Trim(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) = "" Then
                                        Undo = True
                                    End If
                                End If
                            End If
                        End If
                    
                    
                    '2008.02.22
                    Else
                        
                        If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
                            Undo = True
                        End If
                    
                    
                    End If
                                    
                    '2009.02.09
                    If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) <> KAN_KBN_FIN Then
                        Undo = True
                    End If
                
                
                '-------------------------  未検品のﾁｪｯｸ    2012.11.19
                    If KENPIN_CHECK = 1 Then
                        If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
                            Undo = True
                        End If
                    End If
                '-------------------------  未検品のﾁｪｯｸ    2012.11.19
                
                
                End If
                                    
                
                
                
                If Undo Then
                Else
                    '日付期限切れ
                    Do
                        DoEvents
                        sts = BTRV(BtOpInsert, DEL_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<DEL_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Do
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "削除済出荷予定")
                                Exit Do
                        End Select
                    Loop
                    
                    Do
                        DoEvents
                        sts = BTRV(BtOpDelete, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
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
                                Call File_Error(sts, BtOpDelete, "出荷予定")
                                Exit Do
                        End Select
                    Loop
                
                
                    '-------------------------------------  ﾎｽﾄｲﾒｰｼﾞﾃﾞｰﾀの処理  2006.12.06
                    If OSAKA_PC Then
'                        DEN_NO = Left(Trim(StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode)), Len(Trim(StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))) - 1)
'                        SEQ_NO = Right(Trim(StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode)), 1)
                    
                    
'                        Call UniCode_Conv(K0_Y_SYU_H.DEN_NO, DEN_NO)
'                        Call UniCode_Conv(K0_Y_SYU_H.SEQ_NO, SEQ_NO)
                    
                    
                    
                        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                    
                        Do
                            DoEvents
'                            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Do
                                    End If
                                
                                
                                Case BtErrKeyNotFound
                                    
                                    Next_Flg = True
                                    
                                    Exit Do
                                Case Else
                                    Call File_Error(BtOpGetEqual, BtOpGetEqual + BtSNoWait, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
                                    Exit Do
                            End Select
                        Loop
                    
                    
                        If sts = BtNoErr Then
                    
                    
                            Do
                                DoEvents
                                sts = BTRV(BtOpInsert, DEL_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_DEL_SYU_H, Len(K0_DEL_SYU_H), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        Beep
                                        ans = MsgBox("他端末でデータ使用中です。<DEL_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                        If ans = vbCancel Then
                                            Exit Do
                                        End If
                                    Case Else
                                        Call File_Error(sts, BtOpInsert, "削除済出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                                        Exit Do
                                End Select
                            Loop
                            
                            Do
                                DoEvents
                                sts = BTRV(BtOpDelete, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        Beep
                                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                        If ans = vbCancel Then
                                            Exit Do
                                        End If
                                    Case Else
                                        Call File_Error(sts, BtOpDelete, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                                        Exit Do
                                End Select
                            Loop
                                                
                        Else
                            If sts = BtErrKeyNotFound Then
                                sts = BtNoErr
                            End If
                        End If
                    
                    
                    End If
                
                
                
                
                
                    '-------------------------------------  ﾎｽﾄｲﾒｰｼﾞﾃﾞｰﾀの処理  2006.12.06
                
                
                
                    If SYUKA_LOG_ON Then
                        
                        If OSAKA_PC Then
                            If Not Next_Flg Then
                        
                                Call SYUKA_LOG_OUT_PROC("DEL", "AFT")
                            End If
                    
                        Else
                            Call SYUKA_LOG_OUT_PROC("DEL", "AFT")
                    
                        End If
                    
                    End If
                End If
            End If
        End If
            
        If sts <> BtNoErr Then
            Exit Do
        End If
            
        com = BtOpGetNext
    
    Loop
        
                                                    '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "出荷予定データ")
    End If
                                                    '削除済出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "削除済出荷予定データ")
    End If
        
    
    If OSAKA_PC Then                                '2006.12.06
                                                        '出荷予定データＣＬＯＳＥ
        sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
        If sts Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
        End If
                                                        '削除済出荷予定データＣＬＯＳＥ
        sts = BTRV(BtOpClose, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K0_DEL_SYU_H, Len(K0_DEL_SYU_H), 0)
        If sts Then
            Call File_Error(sts, BtOpClose, "削除済出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

End Sub


Private Function KENPIN_CHECK_PROC(Undo As Boolean) As Integer
'-----------------------------------------------------------------------------------
'
'   大阪ＰＣ向け　検品済み　＆　キャンセルのチェック
'
'               2011.06.29
'               2011.07.27 積水以外は日付で削除する
'
'-----------------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer      '2011.07.27
    
    
    
    KENPIN_CHECK_PROC = True
    
    
    If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "Z0001" Then
        Debug.Print
    End If
    '------------------------------------------ 2011.07.27
    If SEK_MTS_FLG Then
        For i = 0 To UBound(SEK_MTS_TBL)
            If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = SEK_MTS_TBL(i) Then
                Exit For
            End If
        Next i
        
        If i > UBound(SEK_MTS_TBL) Then     '2016.07.06
            KENPIN_CHECK_PROC = False       '2016.07.06
            Exit Function                   '2016.07.06
        End If                              '2016.07.06
    
    
    
    End If
        
'    If i > UBound(SEK_MTS_TBL) Then        '2016.07.06
'        KENPIN_CHECK_PROC = False          '2016.07.06
'        Exit Function                      '2016.07.06
'    End If                                 '2016.07.06
    
    
    
    
    '------------------------------------------ 2011.07.28
    
    If Trim(StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode)) = "" Then
        
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
        
        sts = BTRV(BtOpGetEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(Y_SYU_HREC.CANCEL_F, "1")
            Case Else
                Call File_Error(BtOpGetEqual, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
                Exit Function
        End Select
        
        
        If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) = "1" Then
            Debug.Print
        Else
            Undo = True
        End If
    End If
    
    KENPIN_CHECK_PROC = False

End Function
