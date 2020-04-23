VERSION 5.00
Begin VB.Form PM000701 
   Caption         =   "受払先マスタメンテナンス"
   ClientHeight    =   10200
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   11655
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
   ScaleHeight     =   10200
   ScaleWidth      =   11655
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txSEL_KEY 
      Height          =   375
      Left            =   7920
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   5520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   360
      Width           =   2805
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   7980
      Index           =   0
      ItemData        =   "PM000701.frx":0000
      Left            =   2040
      List            =   "PM000701.frx":0002
      TabIndex        =   3
      Top             =   1320
      Width           =   7935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   0
      Top             =   360
      Width           =   735
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
      Left            =   10440
      TabIndex        =   15
      Top             =   9720
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
      Left            =   9600
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   8760
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   7920
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   6600
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   5760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   4920
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "検 索"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   9720
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
      Index           =   3
      Left            =   2760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新 規"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   9720
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
      Left            =   1080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "受払先名称"
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   21
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "収支ｺｰﾄﾞ"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   20
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "受払先ｺｰﾄﾞ"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   19
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "収支ｺｰﾄﾞ"
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   17
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "受払先ｺｰﾄﾞ"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   16
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "PM000701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'テキスト用添字
Private Const ptxUKEHARAI_CODE% = 0         '受払先ｺｰﾄﾞ
Private Const ptxSYUSHI_CODE% = 1           '収支ｺｰﾄﾞ

'リスト用添字
Private Const plstUKEHARAI% = 0

'コンボ用添え字
Private Const pcmbSYUSHI% = 0               '収支名称

Private W_Index As Integer


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000701.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000701)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000701)


    PM000701.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim i           As Integer
Dim Item_Key    As String * 5
Dim sts         As Integer
    
    Error_Check_Proc = True
    
    
    Select Case Mode
        
        Case ptxUKEHARAI_CODE
            
            '受払先ﾏｽﾀ読み込み
            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)
        
    
            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                
            Select Case sts
                Case BtNoErr
                
                    Item_Key = Text1(ptxUKEHARAI_CODE).Text
                
                    txSEL_KEY.Text = Item_Key
                
                    
                    G_SCREEN_FLG = G_SCREEN_UPD
                    If Item_Input_Proc() Then           '明細入力
                        Unload Me
                    End If
                        
                        
                Case BtErrKeyNotFound
                    If List_Disp_Proc() Then
                        Unload Me
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                    Exit Function
                    
            End Select
            
            
        
        
        Case ptxSYUSHI_CODE
            For i = 0 To Combo1(pcmbSYUSHI).ListCount - 1
                    
                If Text1(ptxSYUSHI_CODE).Text = Right(Combo1(pcmbSYUSHI).List(i), 3) Then
                    Combo1(pcmbSYUSHI).ListIndex = i
                    Exit For
                End If
            
            Next i
    End Select
        
    Error_Check_Proc = False
    

End Function
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   リストボックス表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer


    List_Disp_Proc = True
    
    
    Text1(ptxSYUSHI_CODE).Text = Right(Combo1(pcmbSYUSHI).Text, 3)
    
    
    List1(plstUKEHARAI).Clear
    
    '受払先ﾏｽﾀ読み込み
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)
    
    com = BtOpGetGreaterEqual
    
    
    Do
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "受払先マスタ")
                Exit Function
        
        End Select

        
        If Trim(Text1(ptxSYUSHI_CODE).Text) <> "" Then
            If Trim(Text1(ptxSYUSHI_CODE).Text) = Trim(StrConv(P_UKEHARAIREC.SYUSHI_CODE, vbUnicode)) Then
'                List1(plstUKEHARAI).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode) & " " & _
'                                            StrConv(P_UKEHARAIREC.SYUSHI_CODE, vbUnicode) & " " & _
'                                            StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode)
            
            
            
                List1(plstUKEHARAI).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode) & "      " & _
                                            StrConv(P_UKEHARAIREC.SYUSHI_CODE, vbUnicode) & "      " & _
                                            StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode)
            
            
            End If
        Else
'            List1(plstUKEHARAI).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode) & " " & _
'                                        StrConv(P_UKEHARAIREC.SYUSHI_CODE, vbUnicode) & " " & _
'                                        StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode)
        
            List1(plstUKEHARAI).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode) & "      " & _
                                        StrConv(P_UKEHARAIREC.SYUSHI_CODE, vbUnicode) & "      " & _
                                        StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode)
        
        
        End If
        
        com = BtOpGetNext
    
    Loop
        
    DoEvents

    If List1(plstUKEHARAI).ListCount = 0 Then
        
        W_Index = -1
        Text1(ptxUKEHARAI_CODE).SetFocus
    
    End If

    List_Disp_Proc = False

End Function

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbSYUSHI     '収支
            Text1(ptxSYUSHI_CODE).Text = Right(Combo1(pcmbSYUSHI).Text, 3)
    
    End Select
    
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Command1_Click(Index As Integer)

Dim i   As Integer




    Select Case Index
        Case P_CMD_Upd                      '更新
        Case P_CMD_Ins                      '新規
        
            G_SCREEN_FLG = G_SCREEN_INS
            If Item_Input_Proc() Then
                Unload Me
            End If
        
        Case P_CMD_DEL                      '削除
        Case P_CMD_DSP                      '検索/表示
        
                    
            For i = ptxUKEHARAI_CODE To ptxSYUSHI_CODE
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            Next i
            
            If List_Disp_Proc() Then
                Unload Me
            End If
        
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
        Case P_CMD_End                      '終了
            Unload Me
    End Select

End Sub

Private Sub Form_DblClick()
'    PrintForm
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

'    If App.PrevInstance Then
'        Beep
'        MsgBox "同一プログラム実行中です。"
'        End
'    End If

                                'ログファイル名取り込み
'    If GetIni("FILE", "LOGF", "SYS", c) Then
    If GetIni(App.EXEName, "LOGF", App.EXEName, c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
                                
    PM000701.Caption = PM000701.Caption & LAST_UPDATE_DAY
                                
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    Call P_CODE_TBL_Proc
                                
    
    Load PM000702
                                
                                
    W_Index = -1
    
    
    '収支をコンボにセットする
    If Code_Set_Proc(pcmbSYUSHI, P_KBN03_CD, 1) Then
        Unload Me
    End If
    
    Show
    
    Text1(ptxUKEHARAI_CODE).SetFocus
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000701 = Nothing
    Set PM000702 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)


    Select Case Index
        Case plstUKEHARAI
        
            W_Index = List1(plstUKEHARAI).ListIndex
            txSEL_KEY.Text = Left(List1(plstUKEHARAI).List(List1(plstUKEHARAI).ListIndex), 5)
        
            
            G_SCREEN_FLG = G_SCREEN_UPD
            If Item_Input_Proc() Then           '明細入力
                Unload Me
            End If
    End Select

End Sub

Private Sub List1_GotFocus(Index As Integer)
    
    If List1(Index).ListCount > 0 And _
       List1(Index).ListIndex < 0 Then
        List1(Index).ListIndex = 0
    End If

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim W_KEY   As String
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Shift = vbShiftMask Then
        Call Tab_Ctrl(Shift)        '移動
    Else
        If List1(Index).ListIndex = -1 Then
            Exit Sub
        Else
            Select Case Index
                Case plstUKEHARAI
                
                    W_Index = List1(plstUKEHARAI).ListIndex
                    txSEL_KEY.Text = Left(List1(plstUKEHARAI).List(List1(plstUKEHARAI).ListIndex), 10)
                
                    
                    G_SCREEN_FLG = G_SCREEN_UPD
                    If Item_Input_Proc() Then           '明細入力
                        Unload Me
                    End If
            End Select
        End If
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
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
        
    Select Case Index
        Case ptxUKEHARAI_CODE
            Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)
    End Select
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub


Private Function Syushi_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   収支の内容を画面にセットする
'----------------------------------------------------------------------------
Dim sts As Integer
Dim com As Integer


    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    Combo1(pcmbSYUSHI).Clear

    Do
        DoEvents
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN03_CD Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "コードマスタ")
                Exit Function
        
        End Select
    
        Combo1(pcmbSYUSHI).AddItem StrConv(P_CODEREC.C_NAME, vbUnicode) & "  " & Left(StrConv(P_CODEREC.C_Code, vbUnicode), P_KBN03_CD)
    
        com = BtOpGetNext
    
    Loop

    



End Function

Private Function Item_Input_Proc() As Integer
'----------------------------------------------------------------------------
'                   作業管理明細入力画面　表示
'----------------------------------------------------------------------------
    Item_Input_Proc = True

    If G_SCREEN_FLG <> G_SCREEN_INS Then
        
'        If List1(plstUKEHARAI).ListCount = 0 Then
'            Exit Function                           'ﾃﾞｰﾀ無し→即ﾘﾀｰﾝ
'        End If
    
    End If
    
    PM000702.Show vbModal                       '明細入力フォーム表示
    If G_SCREEN_FLG = SYS_ERR Then
        Exit Function
    End If

    If List_Disp_Proc() Then                        'ﾘｽﾄﾎﾞｯｸｽ再表示
        Exit Function
    End If
    
    If List1(plstUKEHARAI).ListCount = 0 Then
        Text1(ptxUKEHARAI_CODE).SetFocus
    Else
        List1(plstUKEHARAI).SetFocus
        If (List1(plstUKEHARAI).ListCount - 1) < W_Index Then
            List1(plstUKEHARAI).ListIndex = List1(plstUKEHARAI).ListCount - 1
        Else
            List1(plstUKEHARAI).ListIndex = W_Index
        End If
    End If

    Item_Input_Proc = False

End Function

Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   コードマスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
Dim OPTION2     As Integer

Dim wkOption    As String



Dim i           As Integer
    
    Code_Set_Proc = True
    
    Combo1(Index).Clear
    
    For i = 0 To UBound(P_KBN_TBL)
    
        If KBN = P_KBN_TBL(i).KBN_CD Then
            Key_Len = P_KBN_TBL(i).KBN_Len
            Exit For
        End If
    
    Next i
    
    If i > UBound(P_KBN_TBL) Then
        Exit Function
    End If
    
    If Mode = 1 Then
        Combo1(Index).AddItem Space(Key_Len)
    End If
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> KBN Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "コードマスタ")
                Exit Function
        
        End Select

        wkOption = ""
        If P_KBN_TBL(i).KBN_OP1 Then
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    




End Function


