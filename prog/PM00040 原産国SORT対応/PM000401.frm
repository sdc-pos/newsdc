VERSION 5.00
Begin VB.Form PM000401 
   Caption         =   "商品化システム　品目マスタメンテナンス（商品ラベル項目）"
   ClientHeight    =   6840
   ClientLeft      =   1920
   ClientTop       =   2730
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
   Begin VB.TextBox txSEL_KEY 
      Height          =   375
      Left            =   8640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   885
   End
   Begin VB.ListBox List1 
      Height          =   4380
      Index           =   0
      ItemData        =   "PM000401.frx":0000
      Left            =   2040
      List            =   "PM000401.frx":0002
      TabIndex        =   2
      Top             =   1080
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   1
      Top             =   480
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      Index           =   8
      Left            =   7800
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
      Index           =   7
      Left            =   6480
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
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
      TabIndex        =   18
      Top             =   6360
      Width           =   180
   End
   Begin VB.Label Label 
      Caption         =   "国内外"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "品番"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "PM000401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'テキスト用添字
Private Const ptxHIN_GAI% = 0               '資材品番

'リスト用添字
Private Const plstITEM% = 0

'コンボ用添え字
Private Const pcmbNAIGAI% = 0               '国内外

Private W_Index As Integer


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000401.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000401)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000401)


    PM000401.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim Item_Key    As String * 20
Dim sts         As Integer
Dim i           As Integer
Dim j           As Integer


    Error_Check_Proc = True


    Select Case Mode
        Case ptxHIN_GAI


            '========================================================= 2007/03/19 =====
''            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
''            Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI), 1))
''            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
''
''            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
''            Select Case sts
''                Case BtNoErr
''                    Item_Key = Text1(Mode).Text
''
''                    txSEL_KEY.Text = Last_JGYOBU & Right(Combo1(pcmbNAIGAI).Text, 1) & Item_Key
''
''                    G_SCREEN_FLG = G_SCREEN_UPD
''                    If Item_Input_Proc() Then           '明細入力
''                        Unload Me
''                    End If
''
''                Case BtErrKeyNotFound
''                    If List_Disp_Proc() Then
''                        Exit Function
''                    End If
''                Case Else
''                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
''                    Exit Function
''            End Select


            If Trim(Text1(Mode).Text) = "" Then
                If List_Disp_Proc() Then
                    Exit Function
                End If
            Else
                Text1(Mode).Text = StrConv(Trim(Text1(Mode).Text), vbUpperCase)


                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI), 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Item_Key = Text1(Mode).Text
    
                        txSEL_KEY.Text = Last_JGYOBU & Right(Combo1(pcmbNAIGAI).Text, 1) & Item_Key
    
                        G_SCREEN_FLG = G_SCREEN_UPD
                        If Item_Input_Proc() Then           '明細入力
                            Unload Me
                        End If
    
                    Case BtErrKeyNotFound
                    
                    
                    
                        For i = 0 To UBound(JGYOBU_T)
                            SubMenu(i).Checked = False
                        Next i
        
        
                        For i = 0 To UBound(JGYOBU_T)
                            For j = 0 To Combo1(pcmbNAIGAI).ListCount - 1
                                Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU_T(i).CODE)
                                Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).List(j), 1))
                                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
        
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                        PM000401.Caption = "商品化システム　品目マスタメンテナンス（商品ラベル項目）（" + RTrim(JGYOBU_T(i).NAME) + ")"
                                        SubMenu(i).Checked = True
                                        LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
                                        LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
                                        Last_JGYOBU = JGYOBU_T(i).CODE
        
                                        Combo1(pcmbNAIGAI).ListIndex = j
        
                                        Item_Key = Text1(Mode).Text
                                        txSEL_KEY.Text = Last_JGYOBU & Right(Combo1(pcmbNAIGAI).Text, 1) & Item_Key
        
                                        G_SCREEN_FLG = G_SCREEN_UPD
                                        If Item_Input_Proc() Then           '明細入力
                                            Unload Me
                                        End If
        
                                        Exit For
        
                                    Case BtErrKeyNotFound
                                        Exit For
        
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                        Exit Function
                                End Select
        
                            Next j
        
                            If sts = BtNoErr Then
                                Exit For
                            End If
        
                        Next i
        
                        If sts <> BtNoErr Then
                            If List_Disp_Proc() Then
                                Exit Function
                            End If
                        End If
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select







            End If
            '==========================================================================


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
    
    PM000401.MousePointer = vbHourglass
    
    List1(plstITEM).Clear
    
    '品目ﾏｽﾀ読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)                          '事業部
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))    '国内外
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)              '品番（外部）
    
    com = BtOpGetGreaterEqual
    
    
    Do
    
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo1(pcmbNAIGAI).Text, 1) Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        
        End Select

        
        List1(plstITEM).AddItem StrConv(ITEMREC.HIN_GAI, vbUnicode) & " " & _
                                    StrConv(ITEMREC.HIN_NAME, vbUnicode)
        
        
        If List1(plstITEM).ListCount > 100 Then
            Exit Do
        End If
        
        com = BtOpGetNext
    
    Loop

    DoEvents
    
    If List1(plstITEM).ListCount = 0 Then
        
        W_Index = -1
        Text1(ptxHIN_GAI).SetFocus
    
    Else
        List1(plstITEM).SetFocus
        List1(plstITEM).ListIndex = 0
            
    End If
    PM000401.MousePointer = vbDefault

    List_Disp_Proc = False

End Function
Private Function Item_Input_Proc() As Integer
'----------------------------------------------------------------------------
'                   作業管理明細入力画面　表示
'----------------------------------------------------------------------------
    Item_Input_Proc = True

    If G_SCREEN_FLG <> G_SCREEN_INS Then
        
'        If List1(plstITEM).ListCount = 0 Then
'            Exit Function                           'ﾃﾞｰﾀ無し→即ﾘﾀｰﾝ
'        End If
    
    End If
    
    PM000402.Show vbModal                           '明細入力フォーム表示
    If G_SCREEN_FLG = SYS_ERR Then
        Exit Function
    End If

    If List_Disp_Proc() Then                        'ﾘｽﾄﾎﾞｯｸｽ再表示
        Exit Function
    End If
    
    If List1(plstITEM).ListCount = 0 Then
        Text1(ptxHIN_GAI).SetFocus
    Else
        List1(plstITEM).SetFocus
        If (List1(plstITEM).ListCount - 1) < W_Index Then
            List1(plstITEM).ListIndex = List1(plstITEM).ListCount - 1
        Else
            List1(plstITEM).ListIndex = W_Index
        End If
    End If

    Item_Input_Proc = False

End Function


Private Sub Command1_Click(Index As Integer)

Dim yn As Integer

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
                                
                                
                                'ラベル印刷用コントロールＦ獲得2008.05.30
    If GetIni("FILE", "labelprint", "SYS", c) Then
        Beep
        MsgBox "ラベル印刷用コントロールＦの獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LabelPrint_F = RTrim(c)
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
        
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            PM000401.Caption = "商品化システム　品目マスタメンテナンス（商品ラベル項目）（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                
    Unload SubMenu(i)
                                
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '品目マスタＯＰＥＮ
    If L_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスタ（仕入先）ＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    Call P_CODE_TBL_Proc
                                
    Load PM000402
                                
    W_Index = -1
    
    
    Combo1(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo1(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    
    Show
    
    Combo1(pcmbNAIGAI).ListIndex = 0
       
    Text1(ptxHIN_GAI).SetFocus
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
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
    
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000401 = Nothing
    Set PM000402 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)
    Select Case Index
        Case plstITEM
        
            W_Index = List1(plstITEM).ListIndex
            txSEL_KEY.Text = Last_JGYOBU & Right(Combo1(pcmbNAIGAI).Text, 1) & Left(List1(plstITEM).List(List1(plstITEM).ListIndex), 20)
        
            
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
            Case plstITEM
            
                W_Index = List1(plstITEM).ListIndex
                txSEL_KEY.Text = Last_JGYOBU & Right(Combo1(pcmbNAIGAI).Text, 1) & Left(List1(plstITEM).List(List1(plstITEM).ListIndex), 20)
            
                
                G_SCREEN_FLG = G_SCREEN_UPD
                If Item_Input_Proc() Then           '明細入力
                    Unload Me
                End If
        End Select
    End If

End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
    
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    PM000401.Caption = "商品化システム　品目マスタメンテナンス（商品ラベル項目））（" + RTrim(JGYOBU_T(Index).NAME) + "）"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

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

