VERSION 5.00
Begin VB.Form CONV20071 
   BackColor       =   &H00C0C0C0&
   Caption         =   "データコンバート処理"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   10095
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
   ScaleHeight     =   7230
   ScaleWidth      =   10095
   StartUpPosition =   2  '画面の中央
   Begin VB.Label Label1 
      Caption         =   "受入履歴"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "指図票（子）"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "指図票（親）＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "データコンバート処理"
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
      TabIndex        =   0
      Top             =   960
      Width           =   4800
   End
End
Attribute VB_Name = "CONV20071"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim DISP_INTERVAL   As Long



Dim c               As String * 128

    Update_Proc = True

'---------------------------------------------  商品化指図票データ（親）のコンバート
    MsgLab(1) = "商品化指図票データ（親）コンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        DoEvents
        
        sts = BTRV(com, OLD_P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_OLD_P_SSHIJI_O, Len(K0_OLD_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）商品化指図票データ（親）")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_NO, Format(CLng(StrConv(P_SSHIJI_O_REC.xSHIJI_NO, vbUnicode)), "00000000"))
        
        
        Do
            sts = BTRV(BtOpInsert, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "商品化指図票データ（親）")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")


'---------------------------------------------  商品化指図票データ（子）のコンバート
    MsgLab(1) = "商品化指図票データ（子）コンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(1).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, OLD_P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_OLD_P_SSHIJI_K, Len(K0_OLD_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）商品化指図票データ（子）")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(1).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        
        Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_NO, Format(CLng(StrConv(P_SSHIJI_K_REC.xSHIJI_NO, vbUnicode)), "00000000"))
        
        
        Do
            sts = BTRV(BtOpInsert, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "商品化指図票データ（親）")
                    Exit Function
            End Select
        Loop
        
        
        
        com = BtOpGetNext
    
    Loop

    Cnt(1).Caption = Format(Count, "#0")


'---------------------------------------------  受入履歴データのコンバート
    MsgLab(1) = "受入履歴データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(2).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, OLD_P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_OLD_P_SUKEIRE, Len(K0_OLD_P_SUKEIRE), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）受入履歴")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(2).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        
        Call UniCode_Conv(P_SUKEIRE_REC.SHIJI_NO, Format(CLng(StrConv(P_SUKEIRE_REC.xSHIJI_NO, vbUnicode)), "00000000"))
        
        
        Do
            sts = BTRV(BtOpInsert, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_SUKEIRE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "受入履歴")
                    Exit Function
            End Select
        Loop
        
        
        
        com = BtOpGetNext
    
    Loop

    Cnt(2).Caption = Format(Count, "#0")


'---------------------------------------------  終了
Update_End:
    
    Update_Proc = False

End Function

Private Sub Form_Activate()

Dim ans As Integer
                                
                                '処理選択
    Beep
    ans = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認入力")
    If ans = vbYes Then
        If Update_Proc() Then
            Unload Me
        End If
    End If
    MsgBox "終了しました。"
    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
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
    LOG_F = RTrim(c)
    
    
    If OLD_P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
    If OLD_P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
    If OLD_P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = BTRV(BtOpClose, OLD_P_SSHIJI_O_POS, OLD_P_SSHIJI_O_REC, Len(OLD_P_SSHIJI_O_REC), K0_OLD_P_SSHIJI_O, Len(K0_OLD_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "旧　指図票（親）")
        End If
    End If
    sts = BTRV(BtOpClose, OLD_P_SSHIJI_K_POS, OLD_P_SSHIJI_K_REC, Len(OLD_P_SSHIJI_K_REC), K0_OLD_P_SSHIJI_K, Len(K0_OLD_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "旧　指図票（子）")
        End If
    End If
    sts = BTRV(BtOpClose, OLD_P_SUKEIRE_POS, OLD_P_SUKEIRE_REC, Len(OLD_P_SUKEIRE_REC), K0_OLD_P_SUKEIRE, Len(K0_OLD_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "旧　受入履歴")
        End If
    End If
    
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "指図票（親）")
        End If
    End If
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "指図票（子）")
        End If
    End If
    sts = BTRV(BtOpClose, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受入履歴")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV20071 = Nothing

    End
End Sub

