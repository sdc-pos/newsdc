VERSION 5.00
Begin VB.Form F1100401 
   BackColor       =   &H00C0C0C0&
   Caption         =   "システム起動処理"
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
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   1236
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "システム起動処理中です。                しばらくお待ち下さい。"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   22.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1092
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   7452
   End
End
Attribute VB_Name = "F1100401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WS_NO       As String * 2           '自端末番号
Dim SERVER_ID   As String * 2           'サーバーＩＤ

 
Private Sub Form_Activate()
Dim sts  As Integer
    
    
    If WS_NO = SERVER_ID Then
'---------------------------'サーバー上の処理
        MsgBox "周辺機器の電源状態を確認後、「Ｅｎｔｅｒ」キーを押して下さい。", vbSystemModal
                            '出荷予定の開放
'''        sts = Y_SYUKA_UNLOCK_PROC()
'''        If sts Then
'''            End
'''        End If
                            '在庫の開放
'''        sts = Zaiko_UNLOCK_Proc()
'''        If sts Then
'''            End
'''        End If

        DoEvents

        sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        If sts Then
            Call File_Error(sts, BtOpReset, "")
        End If

        
        sts = Shell("..\exe\f110010.exe", vbNormalFocus)
        If sts = 0 Then
            MsgBox "[F110010]スキャナ制御の起動に失敗しました｡ "
            Call Log_Out(LOG_F, "[F110010]スキャナ制御の起動に失敗しました｡")
        End If
        
'        sts = Shell("..\exe\f120050.exe", vbNormalFocus)
'        If sts = 0 Then
'            MsgBox "[F120050]月平均出荷数算出処理の起動に失敗しました｡ "
'            Call Log_Out(LOG_F, "[F120050]月平均出荷数算出処理の起動に失敗しました｡")
'            End
'        End If

    Else
'---------------------------'クライアント上の処理
        MsgBox "サーバーＰＣの立上りを確認後、「Ｅｎｔｅｒ」キーを押して下さい。", vbSystemModal
        Call FILE_BACKUP_PROC
    End If

    Unload Me
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()

Dim c           As String * 128
Dim sts         As Integer
    
Dim sBuffer     As String * 255
Dim com         As String
    
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
  
                                
    Label1.Visible = True
'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
    
'自端末番号取り込み
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = StrConv(RTrim(com), vbUpperCase)

'サーバーＩＤ取り込み
    If GetIni("SYSTEM", "SERVER_ID", "SYS", c) Then
        Beep
        MsgBox "サーバーＩＤの獲得に失敗しました。処理を中止して下さい。"
        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [SERVER_ID] READ ERROR")
        End
    End If
    SERVER_ID = StrConv(RTrim(c), vbUpperCase)
    
    
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
    
    Set F1100401 = Nothing

    End
End Sub


Public Function Y_SYUKA_UNLOCK_PROC() As Integer
        
Dim sts As Integer
Dim com As Integer
       
Dim ans As Integer
       
    Y_SYUKA_UNLOCK_PROC = False
                                
    If Y_SYU_Open(BtOpenNomal) Then                 '出荷予定データ
        Exit Function
    End If
        
        
    Call UniCode_Conv(K4_Y_SYU.WEL_ID, "")
    Call UniCode_Conv(K4_Y_SYU.PRG_ID, "")
        
    com = BtOpGetGreater

    Do
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'この時点でのファイル使用中は無限ループとする。キャンセルで異常終了
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Y_SYUKA_UNLOCK_PROC = True
                        Exit Do
                    End If
                
                    com = BtOpGetEqual
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "出荷予定データ")
                    Y_SYUKA_UNLOCK_PROC = True
                    Exit Do
            End Select
        Loop
    
        If sts = BtNoErr Then
            
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")          '使用子機ID
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")          '使用ﾌﾟﾛｸﾞﾗﾑ
    
            Do
                DoEvents
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), BtNCC)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Y_SYUKA_UNLOCK_PROC = True
                            Exit Do
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定")
                        Y_SYUKA_UNLOCK_PROC = True
                        Exit Do
                End Select
            Loop
                
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
        Y_SYUKA_UNLOCK_PROC = True
        Exit Function
    End If



End Function

Private Function Zaiko_UNLOCK_Proc() As Integer
        
Dim sts As Integer
Dim com As Integer
       
Dim ans As Integer
       
    Zaiko_UNLOCK_Proc = False
                                
    If ZAIKO_Open(BtOpenNomal) Then                 '在庫データ
        Exit Function
    End If
        
        
    Call UniCode_Conv(K3_ZAIKO.WEL_ID, "")
    Call UniCode_Conv(K3_ZAIKO.PRG_ID, "")
        
    com = BtOpGetGreater

    Do
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'この時点でのファイル使用中は無限ループとする。キャンセルで異常終了
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Zaiko_UNLOCK_Proc = True
                        Exit Do
                    End If
                
                    com = BtOpGetEqual
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "在庫データ")
                    Zaiko_UNLOCK_Proc = True
                    Exit Do
            End Select
        Loop
    
        If sts = BtNoErr Then
            
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)    '排他ﾌﾗｸﾞ（OFF）
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")          '使用子機ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")          '使用ﾌﾟﾛｸﾞﾗﾑ
    
            Do
                DoEvents
                sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), BtNCC)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Zaiko_UNLOCK_Proc = True
                            Exit Do
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "在庫データ")
                        Zaiko_UNLOCK_Proc = True
                        Exit Do
                End Select
            Loop
                
        End If
            
        If sts <> BtNoErr Then
            Exit Do
        End If
            
        com = BtOpGetNext
    
    Loop
    
                                                '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "在庫データ")
        Zaiko_UNLOCK_Proc = True
        Exit Function
    End If


End Function

Public Sub FILE_BACKUP_PROC()

Dim FROM_DIR    As String
Dim TO_DIR      As String
Dim FILE_NAME   As String
Dim c           As String * 128
                                    'バックアップ元フォルダ取り込み
    If GetIni("FILE", "BACK_FROM", "SYS", c) Then
        Beep
        MsgBox "バックアップ元フォルダの獲得に失敗しました。処理を中止して下さい。"
        Exit Sub
    End If
    FROM_DIR = RTrim(c)

    If GetIni("FILE", "BACK_TO", "SYS", c) Then
        Beep
        MsgBox "バックアップ先フォルダの獲得に失敗しました。処理を中止して下さい。"
        Exit Sub
    End If
    TO_DIR = RTrim(c)

    Label2.Visible = True
    
    On Error GoTo Err_Proc
    ChDir FROM_DIR

    FILE_NAME = Dir(FROM_DIR, vbNormal)

    Do While FILE_NAME <> ""
        DoEvents
        Label2.Caption = "「" & FILE_NAME & "」バックアップ中です。"
        On Error Resume Next
        FileCopy FROM_DIR & FILE_NAME, TO_DIR & FILE_NAME
        FILE_NAME = Dir
    Loop
    Exit Sub

Err_Proc:
    If Err.Number = 76 Then
        MsgBox "ネットワークへの接続が不正です。再立上げを行ってください。"
        Exit Sub
    Else
        Resume Next
    End If
End Sub
