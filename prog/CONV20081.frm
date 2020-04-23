VERSION 5.00
Begin VB.Form CONV20081 
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
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   2940
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "構成マスタ（子）＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   630
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2940
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "構成マスタ（親）＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   630
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
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
Attribute VB_Name = "CONV20081"
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


Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8

Dim c               As String * 128

    Update_Proc = True

'---------------------------------------------  構成マスタのコンバート
    MsgLab(1) = "構成マスタ（親）コンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_P_COMPO_POS, OLD_P_COMPO_O_REC, Len(OLD_P_COMPO_O_REC), K0_OLD_P_COMPO, Len(K0_OLD_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）構成マスタ")
                Exit Function
        End Select
        
        
        
        If StrConv(OLD_P_COMPO_O_REC.SEQNO, vbUnicode) <> "000" Then
        Else
            Count = Count + 1
            DISP_INTERVAL = DISP_INTERVAL + 1
            If DISP_INTERVAL = 100 Then
                Cnt(0).Caption = Format(Count, "#0")
                DISP_INTERVAL = 0
            End If
        
        
            Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, StrConv(OLD_P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode))
            Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, StrConv(OLD_P_COMPO_O_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, StrConv(OLD_P_COMPO_O_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, StrConv(OLD_P_COMPO_O_REC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, StrConv(OLD_P_COMPO_O_REC.DATA_KBN, vbUnicode))
            Call UniCode_Conv(P_COMPO_O_REC.SEQNO, StrConv(OLD_P_COMPO_O_REC.SEQNO, vbUnicode))
            
            Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, StrConv(OLD_P_COMPO_O_REC.CLASS_CODE, vbUnicode))
            Call UniCode_Conv(P_COMPO_O_REC.BIKOU, StrConv(OLD_P_COMPO_O_REC.BIKOU, vbUnicode))
            Call UniCode_Conv(P_COMPO_O_REC.F_CLASS_CODE, StrConv(OLD_P_COMPO_O_REC.F_CLASS_CODE, vbUnicode))
            Call UniCode_Conv(P_COMPO_O_REC.N_CLASS_CODE, StrConv(OLD_P_COMPO_O_REC.N_CLASS_CODE, vbUnicode))
            Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")
            Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, StrConv(OLD_P_COMPO_O_REC.UPD_TANTO, vbUnicode))
            Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, StrConv(OLD_P_COMPO_O_REC.UPD_DATETIME, vbUnicode))
        
            Do
                sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ（親）")
                        Exit Function
                End Select
            Loop
        
        End If
        
        com = BtOpGetNext
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")
'---------------------------------------------  構成マスタ（子）のコンバート
    MsgLab(1) = "構成マスタ（子）コンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(1).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_P_COMPO_POS, OLD_P_COMPO_K_REC, Len(OLD_P_COMPO_K_REC), K0_OLD_P_COMPO, Len(K0_OLD_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）構成マスタ")
                Exit Function
        End Select
        
        
        
        If StrConv(OLD_P_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
        Else
            Count = Count + 1
            DISP_INTERVAL = DISP_INTERVAL + 1
            If DISP_INTERVAL = 100 Then
                Cnt(1).Caption = Format(Count, "#0")
                DISP_INTERVAL = 0
            End If
        
        
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, StrConv(OLD_P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, StrConv(OLD_P_COMPO_K_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, StrConv(OLD_P_COMPO_K_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, StrConv(OLD_P_COMPO_K_REC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, StrConv(OLD_P_COMPO_K_REC.DATA_KBN, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, StrConv(OLD_P_COMPO_K_REC.SEQNO, vbUnicode))
            
            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, StrConv(OLD_P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, StrConv(OLD_P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, StrConv(OLD_P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, StrConv(OLD_P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, StrConv(OLD_P_COMPO_K_REC.KO_QTY, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, StrConv(OLD_P_COMPO_K_REC.KO_BIKOU, vbUnicode))
            
            Call UniCode_Conv(P_COMPO_K_REC.CLASS_CODE, "")
            
            
            
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
            Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, StrConv(OLD_P_COMPO_K_REC.UPD_TANTO, vbUnicode))
            Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, StrConv(OLD_P_COMPO_K_REC.UPD_DATETIME, vbUnicode))
        
        
        
        
        
        
        
        
            Do
                sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ（子）")
                        Exit Function
                End Select
            Loop
        
        End If
        
        com = BtOpGetNext
    
    Loop

    Cnt(1).Caption = Format(Count, "#0")



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
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）構成マスタＯＰＥＮ
    If OLD_P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
                                            '(旧)在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_P_COMPO_POS, OLD_P_COMPO_O_REC, Len(OLD_P_COMPO_O_REC), K0_OLD_P_COMPO, Len(K0_OLD_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)在庫移動歴")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV20081 = Nothing

    End
End Sub

