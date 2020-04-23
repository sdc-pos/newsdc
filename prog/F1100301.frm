VERSION 5.00
Begin VB.Form F1100301 
   BackColor       =   &H00C0C0C0&
   Caption         =   "不要データ削除処理 2009.10.05 11:00"
   ClientHeight    =   4704
   ClientLeft      =   2328
   ClientTop       =   2628
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
   ScaleHeight     =   4704
   ScaleWidth      =   7320
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   372
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   3480
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   372
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   3480
      Width           =   972
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2280
      Y1              =   3600
      Y2              =   3720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "在庫移動歴"
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
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "出荷予定"
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
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "不要データ削除処理"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1320
      Width           =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "入荷予定"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   1920
   End
End
Attribute VB_Name = "F1100301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DEL_IDO     As String * 8
Dim DEL_SYU     As String * 8
Dim DEL_NYU     As String * 8

Dim OSAKA_PC    As Boolean          '2006.12.06
Dim DEL_SYU_H   As String * 8       '2006.12.06

Private Function IDO_Delete() As Integer
'----------------------------------------------------------------------------
'                   「在庫移動歴」消し込み処理
'----------------------------------------------------------------------------
Dim sts As Integer
Dim com As Integer
Dim i   As Integer
Dim ans As Integer

    IDO_Delete = True
    
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        
        Call UniCode_Conv(K0_IDO.JGYOBU, JGYOBU_T(i).CODE)
        Call UniCode_Conv(K0_IDO.JITU_DT, "")
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
        
        com = BtOpGetGreater
        Do
            DoEvents
            Do
                sts = BTRV(com + BtSNoWait, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'ない
                        DoEvents
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<IDOREKI.DAT", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, com, "在庫移動歴")
                        Exit Function
                End Select
            Loop
                        
            If sts Then
                'EOF
                Exit Do
            End If
            
            If StrConv(IDOREC.JGYOBU, vbUnicode) <> JGYOBU_T(i).CODE Then
                '事業部ブレーク
                Exit Do
            End If
    
            If StrConv(IDOREC.JITU_DT, vbUnicode) > DEL_IDO Then
                '実績日付ブレーク
                Exit Do
            End If
    
    
            Do
               sts = BTRV(BtOpDelete, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'ない
                        DoEvents
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpDelete, "在庫移動歴")
                        Exit Function
                End Select
            Loop
    
            com = BtOpGetNext
        
        Loop
    
    Next i

    IDO_Delete = False

End Function
Private Function Y_Nyu_Delete() As Integer
'----------------------------------------------------------------------------
'                   「入荷予定」消し込み処理
'----------------------------------------------------------------------------

Dim sts As Integer
Dim com As Integer
Dim ans As Integer
    
    Y_Nyu_Delete = True
    
    com = BtOpGetFirst
    
    Do
        DoEvents
        Do
            sts = BTRV(com + BtSNoWait, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K3_Y_NYU, Len(K3_Y_NYU), 3)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'ない
                    DoEvents
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com, "入荷予定")
                    Exit Function
            End Select
        Loop
                    
        If sts Then
            'EOF
            Exit Do
        End If
                        
        If StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode) > DEL_NYU Then
            '日付ブレーク
            Exit Do
        End If
                        
        Do
            sts = BTRV(BtOpDelete, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K3_Y_NYU, Len(K3_Y_NYU), 3)
                        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'ない
                    DoEvents
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "入荷予定")
                    Exit Function
            End Select
        Loop
    Loop

    Y_Nyu_Delete = False

End Function
Private Function DEL_Syu_Delete() As Integer
'----------------------------------------------------------------------------
'                   「削除済み出荷予定」消し込み処理
'----------------------------------------------------------------------------

Dim sts As Integer
Dim com As Integer
Dim ans As Integer
    
    
Dim in_cnt  As Long
Dim out_cnt As Long
    
    
    DEL_Syu_Delete = True
    
    com = BtOpGetFirst
    in_cnt = 0
    out_cnt = 0
    Do
        DoEvents
        Do
            sts = BTRV(com + BtSNoWait, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K1_DEL_SYU, Len(K1_DEL_SYU), 1)
            
in_cnt = in_cnt + 1
Text1(1).Text = Format(in_cnt, "#0")
DoEvents
                        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'ない
                    DoEvents
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<DEL_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com, "出荷予定")
                    Exit Function
            End Select
        Loop
                    
        If sts Then
            'EOF
            Exit Do
        End If
                        
        If StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode) > DEL_SYU Then
            '日付ブレーク
            Exit Do
        End If
                        
        Do
            sts = BTRV(BtOpDelete, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K1_DEL_SYU, Len(K1_DEL_SYU), 1)
                        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'ない
                    DoEvents
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<DEL_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "削除済出荷予定")
                    Exit Function
            End Select
        Loop
    
out_cnt = out_cnt + 1
Text1(0).Text = Format(out_cnt, "#0")
DoEvents
    
    
    Loop

    DEL_Syu_Delete = False

End Function


Private Function DEL_Syu_H_Delete() As Integer
'----------------------------------------------------------------------------
'                   「削除済み出荷予定(ﾎｽﾄｲﾒｰｼﾞ)」消し込み処理
'                   2006.12.06
'----------------------------------------------------------------------------

Dim sts As Integer
Dim com As Integer
Dim ans As Integer
    
    
    DEL_Syu_H_Delete = True
    
    com = BtOpGetFirst
    
    Do
        DoEvents
        Do
            sts = BTRV(com + BtSNoWait, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K3_DEL_SYU_H, Len(K3_DEL_SYU_H), 3)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'ない
                    DoEvents
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<DEL_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    Exit Function
            End Select
        Loop
                    
        If sts Then
            'EOF
            Exit Do
        End If
                        
        If StrConv(DEL_SYU_HREC.SYUKA_YMD, vbUnicode) > DEL_SYU_H Then
            '日付ブレーク
            Exit Do
        End If
                        
        Do
            sts = BTRV(BtOpDelete, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K3_DEL_SYU_H, Len(K3_DEL_SYU_H), 3)
                        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'ない
                    DoEvents
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<DEL_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "削除済出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    Exit Function
            End Select
        Loop
    Loop

    DEL_Syu_H_Delete = False

End Function


Private Sub Form_Activate()
    
    Label1(1).Visible = True
    DoEvents

    If Y_Nyu_Delete() Then      '入荷予定の削除
        Unload Me
    End If

    Label1(1).Visible = False
    Label1(2).Visible = True
    DoEvents

    If DEL_Syu_Delete() Then    '出荷予定の削除
        Unload Me
    End If

    
    If OSAKA_PC Then
        If DEL_Syu_H_Delete() Then  '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)の削除   '2006.12.06
            Unload Me
        End If
    End If

    Label1(2).Visible = False
    Label1(3).Visible = True
    DoEvents

    If IDO_Delete() Then        '在庫移動歴の削除
        Unload Me
    End If
    
    Unload Me


End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()

Dim c As String * 128
Dim i As Integer
    
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
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                'データ有効期間取り込み（在庫移動歴）
    If GetIni(App.EXEName, "SAVE_DATA_IDO", "SYS", c) Then
        Beep
        MsgBox "日付の獲得に失敗しました。処理を中止して下さい。"
        Call LOG_OUT(LOG_F, "[SYSTEM][SAVE_DATA_IDO][SYS]読み込みエラー")
        End
    End If
    DEL_IDO = Format(DateAdd("d", -CInt(RTrim(c)), Date), "yyyymmdd")
                                'データ有効期間取り込み（入荷予定）
    If GetIni(App.EXEName, "SAVE_DATA_NYU", "SYS", c) Then
        Beep
        MsgBox "日付の獲得に失敗しました。処理を中止して下さい。"
        Call LOG_OUT(LOG_F, "[SYSTEM][SAVE_DATA_NYU][SYS]読み込みエラー")
        End
    End If
    DEL_NYU = Format(DateAdd("d", -CInt(RTrim(c)), Date), "yyyymmdd")
                                'データ有効期間取り込み（出荷予定）
    If GetIni(App.EXEName, "SAVE_DATA_SYU", "SYS", c) Then
        Beep
        MsgBox "日付の獲得に失敗しました。処理を中止して下さい。"
        Call LOG_OUT(LOG_F, "[SYSTEM][SAVE_DATA_SYU][SYS]読み込みエラー")
        End
    End If
    DEL_SYU = Format(DateAdd("d", -CInt(RTrim(c)), Date), "yyyymmdd")
                                
    
                                '大阪ＰＣ？         2006.12.06
    OSAKA_PC = False
    If GetIni(App.EXEName, "OSAKA_PC", "SYS", c) Then
    Else
        If Trim(c) = "1" Then
            OSAKA_PC = True
        End If
    End If
    
    '2006.12.06                 データ有効期間取り込み（出荷予定(ﾎｽﾄｲﾒｰｼﾞ)）
    
    If OSAKA_PC Then
        If GetIni(App.EXEName, "SAVE_DATA_SYU_H", "SYS", c) Then
            DEL_SYU_H = Format(Now, "yyyymmdd")
        Else
            DEL_SYU_H = Format(DateAdd("d", -CInt(RTrim(c)), Date), "yyyymmdd")
        End If
    End If
                                
                                
                                '入荷予定データＯＰＥＮ
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '削除済出荷予定データＯＰＥＮ
    If DEL_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '削除済出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データＯＰＥＮ 2006.12.06
    If OSAKA_PC Then
        If DEL_SYU_H_Open(BtOpenNomal) Then
            Unload Me
        End If
    End If
                                '在庫移動歴データＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '入荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷予定データ")
        End If
    End If
                                            '削除済出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "削除済出荷予定データ")
        End If
    End If
                                            
    If OSAKA_PC Then
                                            '削除済出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データＣＬＯＳＥ
        sts = BTRV(BtOpClose, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K0_DEL_SYU_H, Len(K0_DEL_SYU_H), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "削除済出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
            End If
        End If
    End If
                                            '在庫移動歴データＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴データ")
        End If
    End If
    
    sts = BTRV(BtOpReset, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1100301 = Nothing

    End

End Sub




