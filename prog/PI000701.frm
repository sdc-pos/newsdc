VERSION 5.00
Begin VB.Form PI000701 
   Caption         =   "資材 前借削除処理 PI00070"
   ClientHeight    =   11145
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   17025
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
   ScaleHeight     =   11145
   ScaleWidth      =   17025
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   8220
      Index           =   0
      ItemData        =   "PI000701.frx":0000
      Left            =   240
      List            =   "PI000701.frx":0002
      TabIndex        =   1
      Top             =   1800
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   240
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   13
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "削 除"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   5
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "前借数"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   18
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "前借日"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblBikou 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   4560
      TabIndex        =   15
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "品番"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "PI000701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'テキスト用添字
Private Const ptxHIN_GAI% = 0               '品番
Private Const ptxNYUKA_DT% = 1              '前借日
Private Const ptxNYUKA_QTY% = 2             '前借数
'リストＢＯＸ用添字
Private Const plstNYU% = 0
Private Const LAST_UPDATE_DAY$ = "(資材)前借削除処理 [PI00070] 2019.06.25 11:15"  '2019.06.25 画面サイズ拡張



Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PI000701.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000701)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000701)


    PI000701.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim com As Integer
Dim sts As Integer

    
    Error_Check_Proc = True
    
        
    Select Case Mode
    
        Case ptxHIN_GAI
    
    
            Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
            Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_P_NYU.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
            sts = BTRV(BtOpGetEqual, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
            Select Case sts
                Case BtNoErr
                    
                    Text1(ptxNYUKA_DT).Text = Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 7, 2)
                    Text1(ptxNYUKA_QTY).Text = Format(CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
                
                
                Case BtErrKeyNotFound
                
                    Text1(ptxNYUKA_DT).Text = ""
                    Text1(ptxNYUKA_QTY).Text = ""
                
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxHIN_GAI).SetFocus
                    Exit Function
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "資材前借ﾃﾞｰﾀ")
                    Exit Function
            End Select
    
    
    
    
    
    End Select
        
        
        
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc(CODE As String) As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Item_Disp_Proc = True
    
    Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_NYU.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    Select Case sts
        Case BtNoErr
            
            Text1(ptxNYUKA_DT).Text = Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 7, 2)
            Text1(ptxNYUKA_QTY).Text = Format(CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
        
        
        Case BtErrKeyNotFound
        
            Text1(ptxNYUKA_DT).Text = ""
            Text1(ptxNYUKA_QTY).Text = ""
        
            MsgBox "入力した項目はエラーです。"
            Text1(ptxHIN_GAI).SetFocus
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "資材前借ﾃﾞｰﾀ")
            Exit Function
    End Select

    Item_Disp_Proc = False

End Function
Private Function List_Disp_Proc() As Integer


'----------------------------------------------------------------------------
'                   リストボックス表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim wkQty       As String

    List_Disp_Proc = True
    
    List1(plstNYU).Clear
    
    '資材前借ﾃﾞｰﾀ読み込み
    
    com = BtOpGetFirst
    
    
    Do
    
        sts = BTRV(com, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材前借ﾃﾞｰﾀ")
                Exit Function
        
        End Select
        
        wkQty = Format(CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
        If Len(wkQty) < 8 Then
            wkQty = Space(7 - Len(wkQty)) & wkQty
        End If
        
        List1(plstNYU).AddItem StrConv(P_NYUREC.HIN_GAI, vbUnicode) & "   " & _
                                    Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 7, 2) & "  " & _
                                    wkQty
        
        com = BtOpGetNext
    Loop
        
    DoEvents

    List_Disp_Proc = False

End Function

Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   前借ﾃﾞｰﾀ削除
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer

    Delete_Proc = True
    
    Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_NYU.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_NYU.NYUKA_DT, Format(Text1(ptxNYUKA_DT).Text, "YYYYMMDD"))

    sts = BTRV(BtOpGetEqual + BtSNoWait, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    
    Do
        Select Case sts
            Case BtNoErr
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Delete_Proc = False
                    Exit Do
                End If
            
            Case BtErrKeyNotFound
                
                For i = ptxHIN_GAI To ptxNYUKA_QTY
                    Text1(i).Text = ""
                Next i
                
                Delete_Proc = False
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材前借ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
        Select Case sts
            Case BtNoErr
                
                '2018.04.11
                Call LOG_OUT(LOG_F, "前借削除　DELETE　HIN_GAI=" & StrConv(P_NYUREC.JGYOBU, vbUnicode) & StrConv(P_NYUREC.NAIGAI, vbUnicode) & StrConv(P_NYUREC.HIN_GAI, vbUnicode) & " " & StrConv(P_NYUREC.NYUKA_DT, vbUnicode) & " " & StrConv(P_NYUREC.UPD_DATETIME, vbUnicode))
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                For i = ptxHIN_GAI To ptxNYUKA_QTY
                    Text1(i).Text = ""
                Next i
                ans = MsgBox("他端末でデータ使用中です。<P_NYU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                    Delete_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "資材前借ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop

    For i = ptxHIN_GAI To ptxNYUKA_QTY
        Text1(i).Text = ""
    Next i




    If List_Disp_Proc() Then
        Exit Function
    End If
    
    If List1(plstNYU).ListCount > 0 Then
        List1(plstNYU).SetFocus
        List1(plstNYU).ListIndex = 0
    Else
        Text1(ptxHIN_GAI).SetFocus
    End If


    Delete_Proc = False


End Function

Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer

    Select Case Index
        Case P_CMD_Upd                      '更新
                    
        
        
        Case P_CMD_DEL                      '削除
            ans = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
            End If
        Case P_CMD_DSP                      '検索/表示
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
        
        Case P_CMD_End                      '終了
            Unload Me
    End Select

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c       As String * 128
Dim i       As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If

                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                '資材前借ＯＰＥＮ
    If P_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
   
    'Me.Caption = "(" + RTrim(JGYOBU_T(i).NAME) + ")" & LAST_UPDATE_DAY
    Me.Caption = LAST_UPDATE_DAY
    
    
    If List_Disp_Proc() Then
        Unload Me
    End If
    
    Show
    
    If List1(plstNYU).ListCount > 0 Then
        List1(plstNYU).ListIndex = 0
        List1(plstNYU).SetFocus
    End If
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                '資材前借ＯＰＥＮ
    sts = BTRV(BtOpClose, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材前借ﾃﾞｰﾀ")
        End If
    End If
    sts = BTRV(BtOpReset, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI000701 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)

Dim W_KEY   As String
Dim sts     As Integer


    W_KEY = Left(List1(Index).List(List1(Index).ListIndex), 20)

    Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_NYU.HIN_GAI, W_KEY)
    Call UniCode_Conv(K0_P_NYU.NYUKA_DT, Format(Mid(List1(Index).List(List1(Index).ListIndex), 24, 10), "YYYYMMDD"))


    sts = BTRV(BtOpGetEqual, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    Select Case sts
        Case BtNoErr
            
            Text1(ptxHIN_GAI).Text = W_KEY
            
            
            Text1(ptxNYUKA_DT).Text = Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 7, 2)
            Text1(ptxNYUKA_QTY).Text = Format(CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
        
        
        Case BtErrKeyNotFound
            Text1(ptxHIN_GAI).Text = ""
        
        
            Text1(ptxNYUKA_DT).Text = ""
            Text1(ptxNYUKA_QTY).Text = ""
        
            MsgBox "入力した項目はエラーです。"
            Text1(ptxHIN_GAI).SetFocus
            Exit Sub
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "資材前借ﾃﾞｰﾀ")
            Exit Sub
    End Select

End Sub

Private Sub List1_GotFocus(Index As Integer)
    
    If List1(plstNYU).ListCount > 0 And _
       List1(plstNYU).ListIndex < 0 Then
        List1(plstNYU).ListIndex = 0
    End If

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim W_KEY   As String
Dim sts     As Integer
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Shift = vbShiftMask Then
        Call Tab_Ctrl(Shift)        '移動
    Else
        
        W_KEY = Left(List1(Index).List(List1(Index).ListIndex), 20)
        
        Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
        Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K0_P_NYU.HIN_GAI, W_KEY)
        Call UniCode_Conv(K0_P_NYU.NYUKA_DT, Format(Mid(List1(Index).List(List1(Index).ListIndex), 24, 10), "YYYYMMDD"))

        sts = BTRV(BtOpGetEqual, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
        Select Case sts
            Case BtNoErr
                Text1(ptxHIN_GAI).Text = W_KEY
                
                Text1(ptxNYUKA_DT).Text = Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                            Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                            Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 7, 2)
                Text1(ptxNYUKA_QTY).Text = Format(CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
            
            
            Case BtErrKeyNotFound
                Text1(ptxHIN_GAI).Text = ""
            
                Text1(ptxNYUKA_DT).Text = ""
                Text1(ptxNYUKA_QTY).Text = ""
            
                MsgBox "入力した項目はエラーです。"
                Text1(ptxHIN_GAI).SetFocus
                Exit Sub
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材前借ﾃﾞｰﾀ")
                Exit Sub
        End Select
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

