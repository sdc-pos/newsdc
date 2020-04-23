VERSION 5.00
Begin VB.Form F1010301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "棚マスタメンテナンス"
   ClientHeight    =   11775
   ClientLeft      =   2130
   ClientTop       =   2835
   ClientWidth     =   16695
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
   ScaleHeight     =   11775
   ScaleWidth      =   16695
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   2052
   End
   Begin VB.ListBox List1 
      Height          =   9420
      Index           =   0
      ItemData        =   "F1010301.frx":0000
      Left            =   1200
      List            =   "F1010301.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   5415
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   4440
      MaxLength       =   2
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   0
      Top             =   360
      Width           =   375
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
      Left            =   10320
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   11160
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
      Left            =   9480
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   11160
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
      Left            =   8640
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   11160
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
      Index           =   8
      Left            =   7800
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   11160
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
      Left            =   6480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   11160
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
      Left            =   5640
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   11160
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
      Left            =   4800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   11160
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
      Left            =   3960
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   11160
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
      Left            =   2640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   11160
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
      Left            =   1800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   11160
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
      Left            =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   11160
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "更  新"
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   11160
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   4215
      Left            =   7440
      TabIndex        =   24
      Top             =   1440
      Width           =   3255
      Begin VB.Frame Frame2 
         Caption         =   "在庫照合"
         Height          =   1455
         Left            =   360
         TabIndex        =   28
         Top             =   2160
         Width           =   2535
         Begin VB.OptionButton Option2 
            Caption         =   "対　象"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "対象外"
            ForeColor       =   &H000000FF&
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   29
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "棚使用可否"
         Height          =   1455
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "使用可"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "使用不可"
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   26
            Top             =   840
            Width           =   1455
         End
      End
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "照合"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   23
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "棚使用"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   22
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "段"
      Height          =   252
      Index           =   11
      Left            =   6480
      TabIndex        =   21
      Top             =   480
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "連"
      Height          =   252
      Index           =   9
      Left            =   5280
      TabIndex        =   20
      Top             =   480
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "列"
      Height          =   252
      Index           =   7
      Left            =   4080
      TabIndex        =   19
      Top             =   480
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "倉庫№"
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   18
      Top             =   480
      Width           =   732
   End
End
Attribute VB_Name = "F1010301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxSoko_No% = 0       '倉庫№
Private Const ptxSoko_Name% = 1     '列
Private Const ptxRetu% = 2          '列
Private Const ptxRen% = 3           '連
Private Const ptxDan% = 4           '段

Private Const Text_Max% = 4

Private Const pLstTana% = 0         '棚情報

Private Ing_Index   As Integer
Private Const LAST_UPDATE_DAY$ = "棚マスタメンテナンス [F101030] 2019.07.17 16:00" '画面拡張 タイトルバー編集

Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim sts     As Integer
    
    Select Case Index
        Case 0
            
            Beep
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                sts = Update_Proc
                Select Case sts
                    Case False
                    Case True
                        Unload Me
                    Case SYS_CANCEL
                End Select
            
                List1(pLstTana).Clear
            
                Frame3.Caption = ""
            
            
            End If
            
            
            Text(ptxSoko_No).SetFocus
        
        
        Case 11
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
            Command(KeyCode - vbKeyF1).Value = True
    End Select


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
Dim c As String * 128
Dim sts As Integer
Dim Work As String


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
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '棚マスタＯＰＥＮ
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If

    Me.Caption = LAST_UPDATE_DAY

    Ing_Index = -1
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
                                            '棚マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
    
    
    sts = BTRV(BtOpReset, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
        Beep
        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
    End If
    
    Set F1010301 = Nothing

    End

End Sub

Private Sub List1_DblClick(Index As Integer)
Dim Edit            As String
Dim KAHI_KBN        As String * 1
Dim ZAIKO_SHOGO_FLG As String * 1
    
Dim Work            As String * 2
        
    
    Ing_Index = List1(pLstTana).ListIndex
    
    Edit = List1(pLstTana).List(List1(pLstTana).ListIndex)
    
    
    Frame3.Caption = Left(Edit, 8)
    
    Work = Right(Edit, 2)
    
    If Left(Work, 1) = KAHI_KBN_OK Then
        Option1(0).Value = True
        Option1(1).Value = False
    Else
        Option1(0).Value = False
        Option1(1).Value = True
    End If
    
    If Right(Work, 1) = ZAIKO_SHOGO_FLG_OK Then
        Option2(0).Value = True
        Option2(1).Value = False
    Else
        Option2(0).Value = False
        Option2(1).Value = True
    End If
    
End Sub

Private Sub Option1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim Edit            As String
Dim KAHI_KBN        As String * 1
Dim ZAIKO_SHOGO_FLG As String * 1
    
    
    If Ing_Index < 0 Then
    '行未選択
        Exit Sub
    End If
    
    List1(pLstTana).RemoveItem Ing_Index
        
    Edit = Frame3.Caption & "  "
    If Option1(0).Value Then
        KAHI_KBN = KAHI_KBN_OK
        Edit = Edit & KAHI_KBN0 & " "
    Else
        KAHI_KBN = KAHI_KBN_NG
        Edit = Edit & KAHI_KBN1 & " "
    End If
     
    If Option2(0).Value Then
        ZAIKO_SHOGO_FLG = ZAIKO_SHOGO_FLG_OK
        Edit = Edit & ZAIKO_SHOGO0 & "  "
    Else
        ZAIKO_SHOGO_FLG = ZAIKO_SHOGO_FLG_NG
        Edit = Edit & ZAIKO_SHOGO1 & "  "
    End If
     
    Edit = Edit & KAHI_KBN & ZAIKO_SHOGO_FLG
     
    List1(pLstTana).AddItem Edit
End Sub

Private Sub Option2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim Edit            As String
Dim KAHI_KBN        As String * 1
Dim ZAIKO_SHOGO_FLG As String * 1
    
    
    If Ing_Index < 0 Then
    '行未選択
        Exit Sub
    End If

    
    
    
    List1(pLstTana).RemoveItem Ing_Index
    
        
    Edit = Frame3.Caption & "  "
    If Option1(0).Value Then
        KAHI_KBN = KAHI_KBN_OK
        Edit = Edit & KAHI_KBN0 & " "
    Else
        KAHI_KBN = KAHI_KBN_NG
        Edit = Edit & KAHI_KBN1 & " "
    End If
     
    If Option2(0).Value Then
        ZAIKO_SHOGO_FLG = ZAIKO_SHOGO_FLG_OK
        Edit = Edit & ZAIKO_SHOGO0 & "  "
    Else
        ZAIKO_SHOGO_FLG = ZAIKO_SHOGO_FLG_NG
        Edit = Edit & ZAIKO_SHOGO1 & "  "
    End If
     
    Edit = Edit & KAHI_KBN & ZAIKO_SHOGO_FLG
     
    List1(pLstTana).AddItem Edit

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim i   As Integer
Dim sts As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub
            
    Text(Index).Text = RTrim(StrConv(Text(Index).Text, vbUpperCase)) '大文字読替え追加 2019/07/17
            
    Select Case Index
        Case ptxSoko_No                 '倉庫№
            sts = Soko_Read_Proc
            Select Case sts
                Case False
                    Text(ptxSoko_Name).Text = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                Case True
                    Text(ptxSoko_Name).Text = ""
                    MsgBox "入力した項目はエラーです｡ (未登録エラー)"
                    Text(Index).SetFocus
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
        Case ptxRetu, ptxRen, ptxDan
            If Len(Trim(Text(Index).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(Index).Text) Then
                    MsgBox "入力した項目はエラーです｡ (未登録エラー)"
                    Text(Index).SetFocus
                    Exit Sub
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
    
            If Index = ptxDan Then
                sts = List_Disp_Proc()
                Select Case sts
                    Case False
                        If List1(pLstTana).ListCount = 0 Then
                            Text(ptxSoko_No).SetFocus
                            Exit Sub
                        Else
                            List1(pLstTana).ListIndex = 0
                            List1(pLstTana).SetFocus
                            Exit Sub
                        End If
                    Case Else
                        Unload Me
                End Select
            End If
    End Select
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i


End Sub

Private Function List_Disp_Proc() As Integer

Dim com     As Integer
Dim sts     As Integer
Dim Edit    As String

    List_Disp_Proc = True
    
    List1(pLstTana).Clear

    Call UniCode_Conv(K0_TANA.SOKO_NO, Text(ptxSoko_No).Text)
    Call UniCode_Conv(K0_TANA.Retu, Text(ptxRetu).Text)
    Call UniCode_Conv(K0_TANA.Ren, Text(ptxRen).Text)
    Call UniCode_Conv(K0_TANA.Dan, Text(ptxDan).Text)
    
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
                If Text(ptxSoko_No).Text <> StrConv(TANAREC.SOKO_NO, vbUnicode) Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "棚マスタ")
                Exit Function
        End Select
    
        Edit = StrConv(TANAREC.Retu, vbUnicode) & "-"
        Edit = Edit & StrConv(TANAREC.Ren, vbUnicode) & "-"
        Edit = Edit & StrConv(TANAREC.Dan, vbUnicode) & "  "
        If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_OK Then
            Edit = Edit & KAHI_KBN0 & " "
        Else
            Edit = Edit & KAHI_KBN1 & " "
        End If
     
        If StrConv(TANAREC.ZAIKO_SHOGO_FLG, vbUnicode) = ZAIKO_SHOGO_FLG_OK Then
            Edit = Edit & ZAIKO_SHOGO0 & "  "
        Else
            Edit = Edit & ZAIKO_SHOGO1 & "  "
        End If
     
        Edit = Edit & StrConv(TANAREC.KAHI_KBN, vbUnicode) & StrConv(TANAREC.ZAIKO_SHOGO_FLG, vbUnicode)
     
        List1(pLstTana).AddItem Edit
     
        com = BtOpGetNext
     Loop

    List_Disp_Proc = False

End Function

Private Function Soko_Read_Proc() As Integer

Dim sts     As Integer

    Soko_Read_Proc = True
    
    Call UniCode_Conv(K0_SOKO.SOKO_NO, Text(ptxSoko_No).Text)
    
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Exit Function
        Case Else
            Soko_Read_Proc = SYS_ERR
            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
            Exit Function
    End Select

    Soko_Read_Proc = False
    
End Function


Private Function Update_Proc() As Integer

Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer
Dim Work    As String * 2
                                            
                                            
    Update_Proc = True
                                        
    Call Input_Lock
    
                                        
    For i = 0 To List1(pLstTana).ListCount - 1
                                        '棚マスタ読み込み
        Call UniCode_Conv(K0_TANA.SOKO_NO, Text(ptxSoko_No).Text)           '倉庫№
        Call UniCode_Conv(K0_TANA.Retu, Mid(List1(pLstTana).List(i), 1, 2)) '列
        Call UniCode_Conv(K0_TANA.Ren, Mid(List1(pLstTana).List(i), 4, 2))  '連
        Call UniCode_Conv(K0_TANA.Dan, Mid(List1(pLstTana).List(i), 7, 2))  '段
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<TANA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Update_Proc = SYS_CANCEL
                        Call Input_UnLock
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "棚マスタ")
                    Exit Function
            End Select
        
        
                    
        Loop
    
        If sts = BtNoErr Then
            
            Work = Right(List1(pLstTana).List(i), 2)
            Call UniCode_Conv(TANAREC.KAHI_KBN, Left(Work, 1))
            Call UniCode_Conv(TANAREC.ZAIKO_SHOGO_FLG, Right(Work, 1))
    
            Do
                sts = BTRV(BtOpUpdate, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<TANA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Update_Proc = SYS_CANCEL
                            Call Input_UnLock
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "棚マスタ")
                        Exit Function
                End Select
            Loop
        
        End If
    
        DoEvents
    Next i
    
    Call Input_UnLock

    Update_Proc = False
End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1010301.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010301)


    F1010301.MousePointer = vbDefault

End Sub

