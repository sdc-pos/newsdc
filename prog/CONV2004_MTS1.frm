VERSION 5.00
Begin VB.Form CONV2004_MTS1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "向け先管理マスタセットアップ処理"
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
   Begin VB.Label In_Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "向け先ＣＳＶ＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1200
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
      Caption         =   "向け先管理マスタセットアップ処理"
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
      Width           =   7680
   End
End
Attribute VB_Name = "CONV2004_MTS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim In_Count        As Long
Dim Out_Count       As Long

Dim DISP_INTERVAL   As Long

Dim fileName        As String
Dim FileNo          As Integer

Dim c               As String * 128

Dim In_NAIGAI               As String       '国内外
Dim In_DATA_KBN             As String       'データ区分
Dim In_MUKE_CODE            As String       '得意先コード
Dim In_SS_CODE              As String       '倉庫／ＳＳコード
Dim In_MUKE_NAME            As String       '得意先名称
Dim In_SS_NAME              As String       'ＳＳ名称
Dim In_MUKE_DNAME           As String       '表示略称
Dim In_DISPLAY_RANKING      As String       '表示略称
Dim In_EOD                  As String

    Update_Proc = True
'---------------------------------------------  向け先マスタ追加項目セットアップ
    MsgLab(1) = "向け先マスタセットアップ処理中！！"
    Me.MousePointer = vbHourglass
                                                '向け先ＣＳＶデータフルパス取込み
    sts = GetIni("FILE", "MTS_CSV", "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI [MTS_CSV]読み込みエラー ")
        Exit Function
    End If
    fileName = Trim(c)
    
        
    
    On Error GoTo Error_Proc
        
    FileNo = FreeFile
    Open fileName For Input As #FileNo
    
    On Error GoTo 0
    
    
    
    In_Count = 0
    DISP_INTERVAL = 0
    In_Cnt(0).Caption = Format(In_Count, "#0")
                                        
                                        
    Do
        
        DoEvents
            
        On Error GoTo Error_Proc
        
        
        Input #FileNo, In_NAIGAI, In_DATA_KBN, In_MUKE_CODE, In_SS_CODE, _
                        In_MUKE_NAME, In_SS_NAME, In_MUKE_DNAME, In_DISPLAY_RANKING, In_EOD
        On Error GoTo 0
        
        In_Count = In_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            In_Cnt(0).Caption = Format(In_Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(MTSREC.NAIGAI, In_NAIGAI)
        Call UniCode_Conv(MTSREC.DATA_KBN, In_DATA_KBN)
        Call UniCode_Conv(MTSREC.MUKE_CODE, In_MUKE_CODE)
        Call UniCode_Conv(MTSREC.SS_CODE, In_SS_CODE)
        Call UniCode_Conv(MTSREC.MUKE_NAME, In_MUKE_NAME)
        Call UniCode_Conv(MTSREC.SS_NAME, In_SS_NAME)
        Call UniCode_Conv(MTSREC.MUKE_DNAME, In_MUKE_DNAME)
        
        If Not IsNumeric(In_DISPLAY_RANKING) Then
            Call UniCode_Conv(MTSREC.DISPLAY_RANKING, "999")
        Else
            Call UniCode_Conv(MTSREC.DISPLAY_RANKING, Format(CInt(In_DISPLAY_RANKING), "000"))
        End If
        
        
        Call UniCode_Conv(MTSREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<MTS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "向け先マスタ")
                    Exit Function
            End Select
        Loop
        
            
    
    Loop

    In_Cnt(0).Caption = Format(In_Count, "#0")
    DoEvents

    MsgBox "正常終了しました"
'---------------------------------------------  終了
    Update_Proc = False
    
    Exit Function

Error_Proc:
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case 62
            In_Cnt(0).Caption = Format(In_Count, "#0")
            MsgBox "正常終了しました"
            Update_Proc = False
            Exit Function
        Case ErrDiskNotReady
            Beep
            ans = MsgBox("ドライブを確認して下さい", vbYesNo + vbExclamation + vbDefaultButton1, "確認入力")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable
            Beep
            ans = MsgBox("ドライブが見つかりません" & fileName, vbExclamation)
        Case ErrNotFound
            Beep
            ans = MsgBox("ファイルが見つかりません" & fileName, vbExclamation)
        Case 76
            Beep
            ans = MsgBox("ファイルパスが見つかりません" & fileName, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [PACKING_CSV Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select

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
                                '向け先マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '向け先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先マスタ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV2004_MTS1 = Nothing

    End
End Sub

