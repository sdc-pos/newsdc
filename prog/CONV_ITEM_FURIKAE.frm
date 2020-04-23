VERSION 5.00
Begin VB.Form CONV_FURIKAE1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "振替マスタ作成処理（CONV_FORIAKE 2012.03.24 16:15)"
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
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Height          =   435
      Index           =   1
      Left            =   7290
      TabIndex        =   11
      Top             =   1680
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始"
      Height          =   435
      Index           =   0
      Left            =   7290
      TabIndex        =   10
      Top             =   1080
      Width           =   1860
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   4725
      TabIndex        =   9
      Top             =   1620
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   2250
      TabIndex        =   8
      Top             =   1620
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   5940
      TabIndex        =   7
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "書換件数"
      Height          =   315
      Index           =   2
      Left            =   5940
      TabIndex        =   6
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "対象件数"
      Height          =   315
      Index           =   1
      Left            =   4455
      TabIndex        =   5
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "読込み件数"
      Height          =   315
      Index           =   0
      Left            =   2925
      TabIndex        =   4
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2925
      TabIndex        =   3
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "振替品番マスタ"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   990
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4455
      TabIndex        =   1
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "振替品番マスタ作成処理"
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
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   5280
   End
End
Attribute VB_Name = "CONV_FURIKAE1"
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
Dim sel_count       As Long
Dim upd_count       As Long

Dim DISP_INTERVAL   As Long

Dim wkKISHU1        As String * 25
Dim wkKISHU2        As String * 52
Dim wkKISHU3        As String * 150
Dim wkKISHU_BIKOU   As String * 450

Dim c               As String * 128

Dim i               As Integer

Dim CHG_FLG         As Boolean

Dim Start_Now       As String

    Update_Proc = True


'---------------------------------------------  受入履歴データのコンバート
    Me.MousePointer = vbHourglass
    Count = 0
    sel_count = 0
    upd_count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
    Start_Now = Format(Now, "YYYYMMDDHHMMSS")
                                        
                                        
                                        
                                        
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, OLD_FURIKAE_POS, OLD_FURIKAEREC, Len(OLD_FURIKAEREC), K0_OLD_FURIKAE, Len(K0_OLD_FURIKAE), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        Cnt(0).Caption = Format(Count, "#0")
        Cnt(1).Caption = Format(Count, "#0")
        Cnt(2).Caption = Format(Count, "#0")
        
        
        
        
        Call UniCode_Conv(FURIKAEREC.JGYOBU_MAE, "B")
        Call UniCode_Conv(FURIKAEREC.NAIGAI_MAE, "1")
        Call UniCode_Conv(FURIKAEREC.HIN_MAE, StrConv(OLD_FURIKAEREC.HIN_MAE, vbUnicode))
        
        Call UniCode_Conv(FURIKAEREC.JGYOBU_GO, "B")
        Call UniCode_Conv(FURIKAEREC.NAIGAI_GO, "1")
        Call UniCode_Conv(FURIKAEREC.HIN_GO, StrConv(OLD_FURIKAEREC.HIN_GO, vbUnicode))
        
        Call UniCode_Conv(FURIKAEREC.BIKOU, StrConv(OLD_FURIKAEREC.BIKOU, vbUnicode))
        
        Call UniCode_Conv(FURIKAEREC.INS_TANTO, StrConv(OLD_FURIKAEREC.INS_TANTO, vbUnicode))
        Call UniCode_Conv(FURIKAEREC.Ins_DateTime, StrConv(OLD_FURIKAEREC.Ins_DateTime, vbUnicode))
        
        Call UniCode_Conv(FURIKAEREC.INS_TANTO, "CONV")
        Call UniCode_Conv(FURIKAEREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
        
        
        
        
        Do
            sts = BTRV(BtOpInsert, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrDuplicates
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                    Exit Function
            End Select
        Loop
        
        
        
        
        com = BtOpGetNext
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")


'---------------------------------------------  終了
Update_End:
    
    Update_Proc = False

End Function

Private Sub Command1_Click(Index As Integer)
    
Dim ans As Integer
    
    Select Case Index
        Case 0
            ans = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            
                MsgBox "終了しました"
                Unload Me
            
            End If


        Case 1
            Unload Me
    End Select

End Sub

Private Sub Form_Activate()

Dim ans As Integer
                                
                                
    Text1(0).Text = "20100716164000"
    Text1(1).Text = "20100802235900"
                                
                                

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyZ Then
        Text1(0).Visible = True
        Text1(1).Visible = True
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
   
    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
    
    If FURIKAE_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
    If OLD_FURIKAE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = BTRV(BtOpClose, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品番振替マスタ")
        End If
    End If
    
    
    sts = BTRV(BtOpClose, OLD_FURIKAE_POS, OLD_FURIKAEREC, Len(OLD_FURIKAEREC), K0_OLD_FURIKAE, Len(K0_OLD_FURIKAE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "旧　品番振替マスタ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV_FURIKAE1 = Nothing

    End
End Sub

