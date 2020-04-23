VERSION 5.00
Begin VB.Form CONV2006_SYUKA1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "データコンバート処理"
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
      Height          =   615
      Index           =   1
      Left            =   6120
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始"
      Height          =   615
      Index           =   0
      Left            =   6120
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "対象向け先"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　出荷予定＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   4080
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
Attribute VB_Name = "CONV2006_SYUKA1"
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

Dim i               As Integer
Dim wk_MTS          As String * 2

    Update_Proc = True

'---------------------------------------------  出荷予定のコンバート
syuka_upd:
    
    MsgLab(1) = "出荷予定データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(3).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）出荷予定データ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
'        If DISP_INTERVAL = 100 Then
            Cnt(3).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
'        End If
        
        If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) = CYU_KBN_TUK Then
        
            sts = BTRV(BtOpDelete, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            If sts Then
                Call File_Error(sts, com, "（旧）出荷予定データ")
                Exit Function
            
            End If
            
            wk_MTS = ""
            
            For i = 1 To Len(Trim(StrConv(Y_SYUREC.MUKE_NAME, vbUnicode)))
                If Mid(StrConv(Y_SYUREC.MUKE_NAME, vbUnicode), i, 1) = "(" Then
                    wk_MTS = Mid(StrConv(Y_SYUREC.MUKE_NAME, vbUnicode), i + 1, 1) & Mid(StrConv(Y_SYUREC.MUKE_NAME, vbUnicode), i + 2, 1)
                    Exit For
                End If
            Next i
            
            If Text1.Text = wk_MTS Then
            
                Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, wk_MTS)
                Call UniCode_Conv(Y_SYUREC.MUKE_CODE, wk_MTS)
            
                Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                Call UniCode_Conv(Y_SYUREC.SS_CODE, "")

            
            End If
            
            
            
            Do
                sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定")
                        Exit Function
                End Select
            Loop
        End If
        
        com = BtOpGetNext
    
    Loop

    Cnt(3).Caption = Format(Count, "#0")

    Me.MousePointer = vbDefault

'---------------------------------------------  終了
Update_End:
    
    Update_Proc = False

End Function

Private Sub Command1_Click(Index As Integer)

Dim ans As Integer


    Select Case Index
        Case 0
    
                                
                                '処理選択
            Beep
            ans = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
            MsgBox "終了しました。"
    
        Case 1
            Unload Me
    
    End Select

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
                                '出荷予定データＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV2006_SYUKA1 = Nothing

    End
End Sub

