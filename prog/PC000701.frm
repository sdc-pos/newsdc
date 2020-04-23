VERSION 5.00
Begin VB.Form PC000701 
   BackColor       =   &H00C0C0C0&
   Caption         =   "品目マスタコンバート処理"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   9120
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
   ScaleWidth      =   9120
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   5
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "実行"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　受入　＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　指図票＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
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
      Top             =   960
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
      Top             =   240
      Width           =   4800
   End
End
Attribute VB_Name = "PC000701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function Update_Proc(mode As Integer) As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim i               As Integer

Dim DISP_INTERVAL   As Long




    Update_Proc = True


    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                '(旧)指示ﾃﾞｰﾀＯＰＥＮ
    If old_P_SSHIJI_O_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If mode = 0 Then
                GoTo UKEIRE_CONV
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, old_P_SSHIJI_O_POS, old_P_SSHIJI_O_REC, Len(old_P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        
        
        
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        
        
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_NO, StrConv(old_P_SSHIJI_O_REC.SHIJI_NO, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.HAKKO_DT, StrConv(old_P_SSHIJI_O_REC.HAKKO_DT, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.PRINT_DATETIME, StrConv(old_P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.TANTO_CODE, StrConv(old_P_SSHIJI_O_REC.TANTO_CODE, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.SHONIN_CODE, StrConv(old_P_SSHIJI_O_REC.SHONIN_CODE, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIMUKE_CODE, StrConv(old_P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.JGYOBU, StrConv(old_P_SSHIJI_O_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.NAIGAI, StrConv(old_P_SSHIJI_O_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_GAI, StrConv(old_P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_QTY, StrConv(old_P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.UKEHARAI_CODE, StrConv(old_P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.S_CLASS_CODE, StrConv(old_P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.F_CLASS_CODE, StrConv(old_P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.N_CLASS_CODE, StrConv(old_P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.S_TANTO, StrConv(old_P_SSHIJI_O_REC.S_TANTO, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.SAMPLE_F, StrConv(old_P_SSHIJI_O_REC.SAMPLE_F, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_F, StrConv(old_P_SSHIJI_O_REC.SHIJI_F, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.TORI_KBN, StrConv(old_P_SSHIJI_O_REC.TORI_KBN, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_SHIJI, StrConv(old_P_SSHIJI_O_REC.PRI_SHIJI, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_PARTS, StrConv(old_P_SSHIJI_O_REC.PRI_PARTS, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_GAISOU, StrConv(old_P_SSHIJI_O_REC.PRI_GAISOU, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_KISHU, StrConv(old_P_SSHIJI_O_REC.PRI_KISHU, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.BIKOU, StrConv(old_P_SSHIJI_O_REC.BIKOU, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.KAN_F, StrConv(old_P_SSHIJI_O_REC.KAN_F, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.KAN_DT, StrConv(old_P_SSHIJI_O_REC.KAN_DT, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.BUNNOU_CNT, StrConv(old_P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode))
        
        Call UniCode_Conv(P_SSHIJI_O_REC.UKEIRE_QTY, StrConv(old_P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))
    
        For i = 0 To 2
    
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(i).NIN, Format(CDbl(StrConv(old_P_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)), "0.0"))
            
            If IsNumeric(StrConv(old_P_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) Then
                Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(i).TIMES, "000.00")
            Else
                Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(i).TIMES, Format(CDbl(StrConv(old_P_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)), "000.00"))
            End If
        Next i

        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NAME, StrConv(old_P_SSHIJI_O_REC.JISEKI_NAME, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NIN, Format(CDbl(StrConv(old_P_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)), "0.0"))
            
        If IsNumeric(StrConv(old_P_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) Then
            Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_TIMES, "000.00")
        Else
            Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_TIMES, Format(CDbl(StrConv(old_P_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "000.00"))
        End If
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NAME, StrConv(old_P_SSHIJI_O_REC.TASEKI_NAME, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NIN, Format(CDbl(StrConv(old_P_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)), "0.0"))
            
        If IsNumeric(StrConv(old_P_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) Then
            Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_TIMES, "000.00")
        Else
            Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_TIMES, Format(CDbl(StrConv(old_P_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "000.00"))
        End If
    
        Call UniCode_Conv(P_SSHIJI_O_REC.CANCEL_F, StrConv(old_P_SSHIJI_O_REC.CANCEL_F, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.CANCEL_DATETIME, StrConv(old_P_SSHIJI_O_REC.CANCEL_DATETIME, vbUnicode))
        Call UniCode_Conv(P_SSHIJI_O_REC.FILLER, "")
        Call UniCode_Conv(P_SSHIJI_O_REC.UPD_DATETIME, StrConv(old_P_SSHIJI_O_REC.UPD_DATETIME, vbUnicode))
        
        Do
            sts = BTRV(BtOpInsert, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "品目マスタ")
                    Exit Function
            End Select
        Loop
        
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(0).Caption = Format(Count, "#0")

UKEIRE_CONV:

                                '(旧)受入ﾃﾞｰﾀＯＰＥＮ
    If old_P_SUKEIRE_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If mode = 0 Then
                GoTo UKEIRE_CONV
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, old_P_SUKEIRE_POS, old_P_SUKEIRE_REC, Len(old_P_SUKEIRE_REC), K0_old_P_SUKEIRE, Len(K0_old_P_SUKEIRE), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        
        
        
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        Call UniCode_Conv(P_SUKEIRE_REC.SHIJI_NO, StrConv(old_P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
        Call UniCode_Conv(P_SUKEIRE_REC.SEQNO, StrConv(old_P_SUKEIRE_REC.SEQNO, vbUnicode))
        Call UniCode_Conv(P_SUKEIRE_REC.SHIMUKE_CODE, StrConv(old_P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
        
        Call UniCode_Conv(P_SUKEIRE_REC.UKEIRE_DT, StrConv(old_P_SUKEIRE_REC.UKEIRE_DT, vbUnicode))
        Call UniCode_Conv(P_SUKEIRE_REC.UKEIRE_QTY, StrConv(old_P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode))
        
        
    
        For i = 0 To 2
    
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(i).NIN, Format(CDbl(StrConv(old_P_SUKEIRE_REC.GENKA_TBL(i).NIN, vbUnicode)), "0.0"))
            
            If IsNumeric(StrConv(old_P_SUKEIRE_REC.GENKA_TBL(i).NIN, vbUnicode)) Then
                Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(i).TIMES, "000.00")
            Else
                Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(i).TIMES, Format(CDbl(StrConv(old_P_SUKEIRE_REC.GENKA_TBL(i).NIN, vbUnicode)), "000.00"))
            End If
        Next i

        Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_NAME, StrConv(old_P_SUKEIRE_REC.JISEKI_NAME, vbUnicode))
        Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_NIN, Format(CDbl(StrConv(old_P_SUKEIRE_REC.JISEKI_NIN, vbUnicode)), "0.0"))
            
        If IsNumeric(StrConv(old_P_SUKEIRE_REC.JISEKI_TIMES, vbUnicode)) Then
            Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_TIMES, "000.00")
        Else
            Call UniCode_Conv(P_SUKEIRE_REC.JISEKI_TIMES, Format(CDbl(StrConv(old_P_SUKEIRE_REC.JISEKI_TIMES, vbUnicode)), "000.00"))
        End If
        
        Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_NAME, StrConv(old_P_SUKEIRE_REC.TASEKI_NAME, vbUnicode))
        Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_NIN, Format(CDbl(StrConv(old_P_SUKEIRE_REC.TASEKI_NIN, vbUnicode)), "0.0"))
            
        If IsNumeric(StrConv(old_P_SUKEIRE_REC.TASEKI_TIMES, vbUnicode)) Then
            Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_TIMES, "000.00")
        Else
            Call UniCode_Conv(P_SUKEIRE_REC.TASEKI_TIMES, Format(CDbl(StrConv(old_P_SUKEIRE_REC.TASEKI_TIMES, vbUnicode)), "000.00"))
        End If
        
        
    
        Call UniCode_Conv(P_SUKEIRE_REC.LAST_F, StrConv(old_P_SUKEIRE_REC.LAST_F, vbUnicode))
        Call UniCode_Conv(P_SUKEIRE_REC.TORI_CODE, StrConv(old_P_SUKEIRE_REC.TORI_CODE, vbUnicode))
        
        
        
        Call UniCode_Conv(P_SUKEIRE_REC.FILLER, "")
        Call UniCode_Conv(P_SUKEIRE_REC.UPD_DATETIME, StrConv(old_P_SUKEIRE_REC.UPD_DATETIME, vbUnicode))
        
        Do
            sts = BTRV(BtOpInsert, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "品目マスタ")
                    Exit Function
            End Select
        Loop
        
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(0).Caption = Format(Count, "#0")



    Me.MousePointer = vbDefault
    MsgBox "コンバート終了"
    Update_Proc = False


End Function


Private Sub Command1_Click(Index As Integer)


Dim ans As Integer
                                
    If Index = 1 Then
        Unload Me
    End If
                                '処理選択
    Beep
    ans = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認入力")
    If ans = vbYes Then
        
        
        
        If Update_Proc(0) Then
            Unload Me
        End If
    End If
'    Unload Me



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
    
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "指示ﾃﾞｰﾀ")
        End If
    End If
    
    sts = BTRV(BtOpClose, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受入")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PC000701 = Nothing

    End
End Sub

