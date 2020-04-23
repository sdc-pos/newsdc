VERSION 5.00
Begin VB.Form PM000501 
   Caption         =   "商品化システム　構成マスタメンテナンス"
   ClientHeight    =   6840
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   11295
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
   ScaleHeight     =   6840
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   360
      Index           =   0
      Left            =   1680
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   18
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txSEL_KEY 
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   4380
      Index           =   0
      ItemData        =   "PM000501.frx":0000
      Left            =   1680
      List            =   "PM000501.frx":0002
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
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
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
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
      Left            =   3960
      TabIndex        =   6
      Top             =   5880
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
      Left            =   2640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
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
      Left            =   1800
      TabIndex        =   4
      Top             =   5880
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
      Top             =   5880
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "仕向け先"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Width           =   180
   End
   Begin VB.Label Label 
      Caption         =   "品番"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   14
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "PM000501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'テキスト用添字
Private Const ptxHIN_GAI% = 0               '資材品番

'リスト用添字
Private Const plstP_COMPO% = 0

'コンボ用添え字
Private Const pcmbSHIMUKE% = 0              '国内外

Private W_Index As Integer


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000501.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000501)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000501)


    PM000501.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim Item_Key    As String * 20
Dim sts         As Integer
    
    
    Error_Check_Proc = True
    
    
    Select Case Mode
    
        Case ptxHIN_GAI
    
            Text1(Mode).Text = StrConv(Trim(Text1(Mode).Text), vbUpperCase)
            
            Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2)) '仕向け先
            Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1)) '事業部
            Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1)) '国内外
            
            Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)                       '品番（外部）
            
            Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
            Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
            
            sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                
            
                    Item_Key = Text1(Mode).Text
                    
                    
                    
                    
                    
                    txSEL_KEY.Text = Right(Combo1(pcmbSHIMUKE).Text, 4) & Item_Key
                
                    
                    G_SCREEN_FLG = G_SCREEN_UPD
                    If Item_Input_Proc() Then           '明細入力
                        Unload Me
                    End If
            
                
                
                Case BtErrKeyNotFound
                    If List_Disp_Proc() Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
        
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
    
    PM000501.MousePointer = vbHourglass
    
    List1(plstP_COMPO).Clear
    
    
    '構成マスタ読み込み
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2)) '仕向け先
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1)) '事業部
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1)) '国内外
    
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)                       '品番（外部）
    
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    
    com = BtOpGetGreaterEqual
    
    Do
    
        sts = BTRV(com, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2) Or _
                    StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1) Or _
                    StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1) Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "構成マスタ")
                Exit Function
        
        End Select


        If StrConv(P_COMPO_O_REC.DATA_KBN, vbUnicode) = P_HEAD Then
            'ﾍｯﾀﾞｰ情報なら
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode))
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select

        
            List1(plstP_COMPO).AddItem StrConv(ITEMREC.HIN_GAI, vbUnicode) & " " & _
                                        StrConv(ITEMREC.HIN_NAME, vbUnicode)
        
        
        
            If List1(plstP_COMPO).ListCount > 50 Then
                Exit Do
            End If
        End If
        
        
        

        
        
        com = BtOpGetNext
    
    Loop
    
    DoEvents

    If List1(plstP_COMPO).ListCount = 0 Then
        
        W_Index = -1
        Text1(pcmbSHIMUKE).SetFocus
    
    Else
        List1(plstP_COMPO).SetFocus
        List1(plstP_COMPO).ListIndex = 0
            
    End If

    PM000501.MousePointer = vbDefault
    
    List_Disp_Proc = False

End Function
Private Function Item_Input_Proc() As Integer
'----------------------------------------------------------------------------
'                   作業管理明細入力画面　表示
'----------------------------------------------------------------------------
    Item_Input_Proc = True

    If G_SCREEN_FLG <> G_SCREEN_INS Then
        
'        If List1(plstP_COMPO).ListCount = 0 Then
'            Exit Function                           'ﾃﾞｰﾀ無し→即ﾘﾀｰﾝ
'        End If
    
    End If
    
    PM000502.Show vbModal                           '明細入力フォーム表示
    If G_SCREEN_FLG = SYS_ERR Then
        Exit Function
    End If

    If List_Disp_Proc() Then                        'ﾘｽﾄﾎﾞｯｸｽ再表示
        Exit Function
    End If
    
    If List1(plstP_COMPO).ListCount = 0 Then
        Text1(pcmbSHIMUKE).SetFocus
    Else
        List1(plstP_COMPO).SetFocus
        If (List1(plstP_COMPO).ListCount - 1) < W_Index Then
            List1(plstP_COMPO).ListIndex = List1(plstP_COMPO).ListCount - 1
        Else
            List1(plstP_COMPO).ListIndex = W_Index
        End If
    End If

    Item_Input_Proc = False

End Function


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
        Case P_CMD_Upd                      '更新
        Case P_CMD_DEL                      '削除
        Case P_CMD_Ins                      '新規
        
            G_SCREEN_FLG = G_SCREEN_INS
            If Item_Input_Proc() Then
                Unload Me
            End If
        
        Case P_CMD_DSP                      '検索/表示
        
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
    Show
    MsgBox "未使用プログラム！"
    End
    
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
                                
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'クラスマスタＯＰＥＮ
    If P_Class_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
    Call P_CODE_TBL_Proc
                                
    Load PM000502
                                
    W_Index = -1
    
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD) Then
        Unload Me
    End If
    
    
    If Combo1(pcmbSHIMUKE).ListCount = 0 Then
        MsgBox "仕向け先の登録がありません。"
        Unload Me
    End If
    
    Combo1(pcmbSHIMUKE).ListIndex = 0
       
    Combo1(pcmbSHIMUKE).SetFocus
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
                                            
                                            
                                            
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
                                            'クラスマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "クラスマスタ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000501 = Nothing
    Set PM000502 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)
    Select Case Index
        Case plstP_COMPO
        
            W_Index = List1(plstP_COMPO).ListIndex
            txSEL_KEY.Text = Right(Combo1(pcmbSHIMUKE).Text, 4) & Left(List1(plstP_COMPO).List(List1(plstP_COMPO).ListIndex), 20)
        
            
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

    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Shift = vbShiftMask Then
        Call Tab_Ctrl(Shift)        '移動
    Else
        Select Case Index
            Case plstP_COMPO
            
                W_Index = List1(plstP_COMPO).ListIndex
                txSEL_KEY.Text = Right(Combo1(pcmbSHIMUKE).Text, 4) & Left(List1(plstP_COMPO).List(List1(plstP_COMPO).ListIndex), 20)
            
                
                G_SCREEN_FLG = G_SCREEN_UPD
                If Item_Input_Proc() Then           '明細入力
                    Unload Me
                End If
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


Private Function Code_Set_Proc(Index As Integer, KBN As String) As Integer
'----------------------------------------------------------------------------
'                   コードマスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim Option1     As Integer
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
            wkOption = Trim(StrConv(P_CODEREC.Option1, vbUnicode))
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
