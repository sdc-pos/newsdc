VERSION 5.00
Begin VB.Form F9001101 
   Caption         =   "出荷予定リカバリー処理"
   ClientHeight    =   4125
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   11820
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   1
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   960
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   0
      Left            =   5250
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   960
      Width           =   1170
   End
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
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
      Left            =   10530
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   9690
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   8850
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "実 行"
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
      Left            =   8010
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   6690
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   5850
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   5010
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   4170
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   2850
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   2010
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   1170
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   330
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      Height          =   495
      Left            =   3675
      TabIndex        =   17
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "／"
      Height          =   375
      Left            =   6405
      TabIndex        =   15
      Top             =   1080
      Width           =   330
   End
   Begin VB.Label Label1 
      Caption         =   "更新件数/ 読み込み件数＝"
      Height          =   375
      Left            =   2205
      TabIndex        =   13
      Top             =   1080
      Width           =   3060
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
      Left            =   330
      TabIndex        =   12
      Top             =   3480
      Width           =   180
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Checked         =   -1  'True
         Index           =   0
      End
   End
End
Attribute VB_Name = "F9001101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



                






Private Sub Command_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
        Case 8                              'データ出力
        
            Beep
            ans = MsgBox("「出荷予定」コンバート処理　実行しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If ans = vbYes Then
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
            
            
            
            End If
        
        
        
        
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
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
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer

Dim Start_Pos   As Integer


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
                    
                    
                    
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    ReDim Preserve JGYOBU_T(UBound(JGYOBU_T) + 1)
    JGYOBU_T(UBound(JGYOBU_T)).CODE = "*"
    JGYOBU_T(UBound(JGYOBU_T)).NAME = "全BU"
    JGYOBU_T(UBound(JGYOBU_T)).COLOR = 12


    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F9001101.Caption = "出荷予定リカバリー処理（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F9001101.Caption = "出荷予定リカバリー処理（" + RTrim(JGYOBU_T(Index).NAME) + ")"

    SubMenu(Index).Checked = True
    If Last_JGYOBU <> JGYOBU_T(Index).CODE Then
        Last_JGYOBU = JGYOBU_T(Index).CODE
        LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
        LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

    End If

End Sub
Private Function List_Disp_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim DEN_MAISU   As Long
Dim KAN_MAISU   As Long
    
Dim Skip_Flg    As Boolean
    
    
    List_Disp_Proc = True
                                    
    
    
'    Call Input_Lock
    F9001101.MousePointer = vbHourglass
    F9001101.Enabled = False
                                    
    
                                    
                                    
    DEN_MAISU = 0
    KAN_MAISU = 0
    
    
    If Last_JGYOBU = "*" Then
        Call UniCode_Conv(K2_Y_SYU.JGYOBU, "") '事業部
    Else
        Call UniCode_Conv(K2_Y_SYU.JGYOBU, Last_JGYOBU) '事業部
    End If
                                                    '注文区分
    Call UniCode_Conv(K2_Y_SYU.KEY_CYU_KBN, "")
                                                    '向け先
    Call UniCode_Conv(K2_Y_SYU.KEY_MUKE_CODE, "")
    Call UniCode_Conv(K2_Y_SYU.KEY_SS_CODE, "")
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)
    
    
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定")
                List_Disp_Proc = SYS_ERR
'                Call Input_UnLock
                F9001101.MousePointer = vbDefault
                F9001101.Enabled = True
                Exit Function
        End Select
                                '事業部 KEYﾌﾞﾚｰｸ
        
        If Last_JGYOBU = "*" Then
        Else
            If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Exit Do
            End If
        End If
                                
If StrConv(Y_SYUREC.ID_NO, vbUnicode) = "400047343808" Then
    Debug.Print
End If
        
        DEN_MAISU = DEN_MAISU + 1
        
        Label3.Caption = StrConv(Y_SYUREC.ID_NO, vbUnicode)
        
        If StrConv(Y_SYUREC.TANABAN1, vbUnicode) = StrConv(Y_SYUREC.TANABAN2, vbUnicode) Then
            
            If Trim(StrConv(Y_SYUREC.TANABAN1, vbUnicode)) = "" Then
            Else
                Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                Call UniCode_Conv(Y_SYUREC.TANABAN2, StrConv(Y_SYUREC.TANABAN3, vbUnicode))
                Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                KAN_MAISU = KAN_MAISU + 1
            End If
        Else
            If StrConv(Y_SYUREC.TANABAN1, vbUnicode) = StrConv(Y_SYUREC.TANABAN3, vbUnicode) Then
                If Trim(StrConv(Y_SYUREC.TANABAN1, vbUnicode)) = "" Then
                Else
                    Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                    KAN_MAISU = KAN_MAISU + 1
                End If
            Else
                If StrConv(Y_SYUREC.TANABAN2, vbUnicode) = StrConv(Y_SYUREC.TANABAN3, vbUnicode) Then
                    If Trim(StrConv(Y_SYUREC.TANABAN2, vbUnicode)) = "" Then
                    Else
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                        KAN_MAISU = KAN_MAISU + 1
                    End If
                End If
            End If
        End If
        
        
        If StrConv(Y_SYUREC.TANABAN1, vbUnicode) = StrConv(Y_SYUREC.TANABAN2, vbUnicode) Then
            
            If Trim(StrConv(Y_SYUREC.TANABAN1, vbUnicode)) = "" Then
            Else
                Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                Call UniCode_Conv(Y_SYUREC.TANABAN2, StrConv(Y_SYUREC.TANABAN3, vbUnicode))
                Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                KAN_MAISU = KAN_MAISU + 1
            End If
        Else
            If StrConv(Y_SYUREC.TANABAN1, vbUnicode) = StrConv(Y_SYUREC.TANABAN3, vbUnicode) Then
                If Trim(StrConv(Y_SYUREC.TANABAN1, vbUnicode)) = "" Then
                Else
                    Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                    KAN_MAISU = KAN_MAISU + 1
                End If
            Else
                If StrConv(Y_SYUREC.TANABAN2, vbUnicode) = StrConv(Y_SYUREC.TANABAN3, vbUnicode) Then
                    If Trim(StrConv(Y_SYUREC.TANABAN2, vbUnicode)) = "" Then
                    Else
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                        KAN_MAISU = KAN_MAISU + 1
                    End If
                End If
            End If
        End If
        
        If StrConv(Y_SYUREC.TANABAN1, vbUnicode) = StrConv(Y_SYUREC.TANABAN2, vbUnicode) Then
            
            If Trim(StrConv(Y_SYUREC.TANABAN1, vbUnicode)) = "" Then
            Else
                Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                Call UniCode_Conv(Y_SYUREC.TANABAN2, StrConv(Y_SYUREC.TANABAN3, vbUnicode))
                Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                KAN_MAISU = KAN_MAISU + 1
            End If
        Else
            If StrConv(Y_SYUREC.TANABAN1, vbUnicode) = StrConv(Y_SYUREC.TANABAN3, vbUnicode) Then
                If Trim(StrConv(Y_SYUREC.TANABAN1, vbUnicode)) = "" Then
                Else
                    Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                    KAN_MAISU = KAN_MAISU + 1
                End If
            Else
                If StrConv(Y_SYUREC.TANABAN2, vbUnicode) = StrConv(Y_SYUREC.TANABAN3, vbUnicode) Then
                    If Trim(StrConv(Y_SYUREC.TANABAN2, vbUnicode)) = "" Then
                    Else
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                        KAN_MAISU = KAN_MAISU + 1
                    End If
                End If
            End If
        End If
        
        
        
        
        
        
        
        
        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)
        If sts Then
            Call File_Error(sts, BtOpUpdate, "出荷予定")
            List_Disp_Proc = SYS_ERR
'                Call Input_UnLock
            F9001101.MousePointer = vbDefault
            F9001101.Enabled = True
            Exit Function
        End If
        
        
        
        com = BtOpGetNext
        
    Loop
    
    
    
    Text1(0).Text = KAN_MAISU
    
    Text1(1).Text = DEN_MAISU
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
    
    
    F9001101.MousePointer = vbDefault
    F9001101.Enabled = True
    
    Command(7).Enabled = True
'    Call Input_UnLock
    
    If DEN_MAISU > 0 Then
        Command(8).Enabled = True
    End If
    
    List_Disp_Proc = False

    
End Function

