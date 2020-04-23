VERSION 5.00
Begin VB.Form PC000301 
   BackColor       =   &H00C0C0C0&
   Caption         =   "クラスマスタコンバート処理"
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
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "クラスマスタ＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
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
Attribute VB_Name = "PC000301"
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


Dim FileNo          As Long
Dim fileName        As String


Dim CLASS_REC     As Variant
Dim RecordBuf       As String

Dim c               As String * 128

    Update_Proc = True

    FileNo = FreeFile
    
                                'ログファイル名取り込み
    If GetIni("FILE", "CLASS_TXT", "CONV2006", c) Then
        Beep
        MsgBox "[CLASS_TXT]の獲得に失敗しました。処理を中止して下さい。"
        Unload Me
    End If
    fileName = RTrim(c)
    
        
    Open fileName For Input As FileNo
    
    
    
    
    
    MsgLab(1) = "クラスマスタコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                        
                                        
                                        
    Do Until EOF(FileNo)
        
        DoEvents
        
        Line Input #FileNo, RecordBuf
        
        CLASS_REC = Split(RecordBuf, vbTab, -1)
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(P_CLASSREC.SHIMUKE_CODE, "01")                                 '仕向け先ｺｰﾄﾞ（掃除機は固定）
        Call UniCode_Conv(P_CLASSREC.CLASS_CODE, CStr(CLASS_REC(0)))                'クラス
        Call UniCode_Conv(P_CLASSREC.CLASS_NAME, CStr(CLASS_REC(3)))                '呼び名
                                                                                    '商品化価格
        
        If IsNumeric(CLASS_REC(1)) Then
            Call UniCode_Conv(P_CLASSREC.TANKA, Format(CDbl(CLASS_REC(1)), "00000000.00"))
        Else
            Call UniCode_Conv(P_CLASSREC.TANKA, "00000000.00")
        End If
                                                                                    
        If IsNumeric(CLASS_REC(9)) Then                                             '工数
            Call UniCode_Conv(P_CLASSREC.KOUSU, Format(CDbl(CLASS_REC(9)), "000.000"))
        Else
            Call UniCode_Conv(P_CLASSREC.KOUSU, "000.000")
        End If
                                                                                    '工料
        If IsNumeric(CLASS_REC(10)) Then
            Call UniCode_Conv(P_CLASSREC.KOURYOU, Format(CDbl(CLASS_REC(10)), "00000000.00"))
        Else
            Call UniCode_Conv(P_CLASSREC.KOURYOU, "00000000.00")
        End If
                                            
        Call UniCode_Conv(P_CLASSREC.ETC, "00000000.00")                            'その他
        
        
        
        Call UniCode_Conv(P_CLASSREC.UPD_TANTO, "CONV")                             '更新担当者
                                                                                    '更新日時
        Call UniCode_Conv(P_CLASSREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
        Do
            sts = BTRV(BtOpInsert, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_CLASS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case BtErrDuplicates
                    
                    Exit Do
                
                Case Else
                    
                    
                    
                    Call File_Error(sts, BtOpInsert, "クラスマスタ")
                    Exit Function
            End Select
        Loop
        
    
    Loop
'---------------------------------------------  終了

    Cnt(0).Caption = Format(Count, "#0")
    
    Close #FileNo

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
                                'クラスマスタＯＰＥＮ
    If P_Class_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            'クラスマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "クラスマスタ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PC000301 = Nothing

    End
End Sub

