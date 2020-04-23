VERSION 5.00
Begin VB.Form F1010601 
   BackColor       =   &H00FFFFFF&
   Caption         =   "向け先管理マスタメンテナンス"
   ClientHeight    =   11325
   ClientLeft      =   2130
   ClientTop       =   2430
   ClientWidth     =   16875
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
   MaxButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   16875
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   7875
      MaxLength       =   2
      TabIndex        =   28
      Top             =   1440
      Width           =   390
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   3
      Top             =   960
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   3
      Left            =   8040
      MaxLength       =   20
      TabIndex        =   4
      Top             =   960
      Width           =   4935
   End
   Begin VB.ListBox List1 
      Height          =   7980
      ItemData        =   "F1010601.frx":0000
      Left            =   600
      List            =   "F1010601.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Width           =   12375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   4
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1440
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   1
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   2
      Top             =   960
      Width           =   4935
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   600
      MaxLength       =   8
      TabIndex        =   1
      Top             =   960
      Width           =   1092
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   600
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   360
      Width           =   975
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   10560
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   10560
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "データ"
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   10560
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   10560
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   10560
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   10560
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "削  除"
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   10560
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   10560
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   10560
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   8295
      TabIndex        =   29
      Top             =   1560
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷区分コード"
      Height          =   240
      Index           =   6
      Left            =   6195
      TabIndex        =   27
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "表示順序"
      Height          =   240
      Index           =   7
      Left            =   4440
      TabIndex        =   26
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ＳＳｺｰﾄﾞ"
      Height          =   240
      Index           =   5
      Left            =   6840
      TabIndex        =   25
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ＳＳ名称"
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   24
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "スキャナ表示用名称"
      Height          =   240
      Index           =   3
      Left            =   600
      TabIndex        =   23
      Top             =   1560
      Width           =   2160
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "得意先名称"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   22
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "得意先ｺｰﾄﾞ"
      Height          =   240
      Index           =   2
      Left            =   600
      TabIndex        =   21
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "F1010601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Combo_Max% = 0
Private Const Command_Max% = 11

Private Const ptxMUKE_CODE% = 0
Private Const ptxMUKE_NAME% = 1
Private Const ptxSS_CODE% = 2
Private Const ptxSS_NAME% = 3
Private Const ptxMUKE_DNAME% = 4

Private Const ptxRANKING% = 5
Private Const ptxSYUKA_KBN% = 6


Private Const Text_Max% = 6


Private Const pcmbNaiGai% = 0

Private MTS_CSV As String

Private wkMTS_CHG_CD(0 To 1295) As String * 2
Private Const LAST_UPDATE_DAY$ = "[F101060] 2019.06.25 11:15"  '2019.06.25 画面サイズ拡張


Private Function List_Proc() As Integer
'----------------------------------------------------------------------------
'                   リストボックス表示処理
'----------------------------------------------------------------------------
Dim sts As Integer
Dim com As Integer

    List_Proc = True
    
    List1.Clear
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "向け先管理マスタ")
                Exit Function
        End Select
        
        Call List_Edit_Proc
         
        com = BtOpGetNext
    Loop
    
    List_Proc = False
    
End Function
Private Sub Clear_Field(Mode As Integer)
'----------------------------------------------------------------------------
'                   画面消去処理
'----------------------------------------------------------------------------
Dim i As Integer

    
    For i = Mode To Text_Max
            Text(i) = ""
    Next i

End Sub
Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim sts     As Integer
    
    Err_Chk = True
    If Len(Text(ptxMUKE_CODE).Text) = 0 Then
        Beep
        MsgBox "入力した項目はエラーです。(必須入力)"
        Text(ptxMUKE_CODE).SetFocus
        Exit Function
    End If
        
    If IsNumeric(Text(ptxRANKING).Text) Then
        Text(ptxRANKING).Text = Format(CInt(Text(ptxRANKING).Text), "000")
    End If
        
    If Trim(Text(ptxSYUKA_KBN).Text) = "" Then
        Label(8).Caption = ""
    Else
        Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, Text(ptxSYUKA_KBN).Text)
        sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
        Select Case sts
            Case BtNoErr
                Label(8).Caption = Trim(StrConv(SE_SHIP_TANKA_M_REC.SE_SYUKA_NAME, vbUnicode))
            Case BtErrKeyNotFound
                Label(8).Caption = ""
                Beep
                MsgBox "入力した項目はエラーです。(単価設定未登録)"
                Text(ptxSYUKA_KBN).SetFocus
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定マスタ")
                Exit Function
        End Select
    End If
        
        
    Err_Chk = False
End Function
Private Function Dislpay_Proc() As Integer
'----------------------------------------------------------------------------
'                   レコード内容の表示
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

    Dislpay_Proc = True

    Call UniCode_Conv(K0_MTS.MUKE_CODE, Text(ptxMUKE_CODE).Text)
    Call UniCode_Conv(K0_MTS.SS_CODE, Text(ptxSS_CODE).Text)
    
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Dislpay_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
            Exit Function
    End Select
    
    
    For i = 0 To 1
        
        If Right(Combo(pcmbNaiGai).List(i), 1) = StrConv(MTSREC.NAIGAI, vbUnicode) Then
            Combo(pcmbNaiGai).ListIndex = i
            Exit For
        End If
    
    Next i
    
    Text(ptxMUKE_NAME).Text = StrConv(MTSREC.MUKE_NAME, vbUnicode)
    Text(ptxSS_NAME).Text = StrConv(MTSREC.SS_NAME, vbUnicode)
    Text(ptxMUKE_DNAME).Text = StrConv(MTSREC.MUKE_DNAME, vbUnicode)

    Text(ptxRANKING).Text = StrConv(MTSREC.DISPLAY_RANKING, vbUnicode)

    Text(ptxSYUKA_KBN).Text = StrConv(MTSREC.SYUKA_KBN, vbUnicode)
    Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, Text(ptxSYUKA_KBN).Text)
    sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
    Select Case sts
        Case BtNoErr
            Label(8).Caption = Trim(StrConv(SE_SHIP_TANKA_M_REC.SE_SYUKA_NAME, vbUnicode))
        Case BtErrKeyNotFound
            Label(8).Caption = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定マスタ")
            Exit Function
    End Select


    Dislpay_Proc = False
End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   追加／変更処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim wkMUKE_CHG_CD   As String * 2
    
    Update_Proc = True

    Call UniCode_Conv(K0_MTS.MUKE_CODE, Text(ptxMUKE_CODE).Text)
    Call UniCode_Conv(K0_MTS.SS_CODE, Text(ptxSS_CODE).Text)


    Do
        
        sts = BTRV(BtOpGetEqual + BtSNoWait, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<MTS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "向け先管理マスタ")
                Exit Function
        End Select
    
    Loop

    If com = BtOpInsert Then
        Call UniCode_Conv(MTSREC.MUKE_CODE, Text(ptxMUKE_CODE).Text)
        Call UniCode_Conv(MTSREC.SS_CODE, Text(ptxSS_CODE).Text)
        Call UniCode_Conv(MTSREC.DATA_KBN, "")
        Call UniCode_Conv(MTSREC.FILLER, "")
    End If

    Call UniCode_Conv(MTSREC.NAIGAI, Right(Combo(pcmbNaiGai).Text, 1))
    Call UniCode_Conv(MTSREC.MUKE_NAME, Text(ptxMUKE_NAME).Text)
    Call UniCode_Conv(MTSREC.SS_NAME, Text(ptxSS_NAME).Text)
    Call UniCode_Conv(MTSREC.MUKE_DNAME, Text(ptxMUKE_DNAME).Text)
    Call UniCode_Conv(MTSREC.DISPLAY_RANKING, Text(ptxRANKING).Text)
    Call UniCode_Conv(MTSREC.SYUKA_KBN, Text(ptxSYUKA_KBN).Text)


    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<MTS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "向け先管理マスタ")
                Exit Function
        End Select
    Loop

    Call List_Update_Proc(0)

    Call Clear_Field(0)

    Update_Proc = False

End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   削除処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

    
    Delete_Proc = True

    Call UniCode_Conv(K0_MTS.MUKE_CODE, Text(ptxMUKE_CODE).Text)
    Call UniCode_Conv(K0_MTS.SS_CODE, Text(ptxSS_CODE).Text)


    Do
        
        sts = BTRV(BtOpGetEqual + BtSNoWait, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
                Do
                    sts = BTRV(BtOpDelete, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<MTS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Delete_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpDelete, "向け先管理マスタ")
                            Exit Function
                    End Select
                Loop
                Exit Do
            Case BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<MTS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Delete_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "向け先管理マスタ")
                Exit Function
        End Select
    
    Loop

    Call List_Update_Proc(1)

    Call Clear_Field(0)

    Delete_Proc = False

End Function

Private Sub Command_Click(Index As Integer)

Dim yn  As Integer
Dim sts As Integer

    Select Case Index
        Case 0
                                            
            sts = Err_Chk()             'エラーチェック
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            Beep
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If yn = vbYes Then
                sts = Update_Proc()
                Select Case sts
                    Case False
                    Case True
                        Unload Me
                    Case SYS_CANCEL
                End Select
            End If
            
            Text(ptxMUKE_CODE) = ""
        
        Case 3
            If Trim(Text(ptxMUKE_CODE)) = "" Then
                Beep
                MsgBox "削除するコードを指定して下さい。", vbExclamation
            Else
                Beep
                yn = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
                If yn = vbYes Then
                    If Delete_Proc() Then
                        Unload Me
                    End If
                End If
            End If
        Case 8                  'データ出力
            Beep
            yn = MsgBox("データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Data_Proc() Then
                    Unload Me
                End If
            End If
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
    
    Text(ptxMUKE_CODE).SetFocus

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
Dim i       As Integer
Dim j       As Integer
Dim k       As Integer
Dim c       As String * 128
Dim sts     As Integer

Dim Wk      As String * 36


    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                'ＣＳＶファイル名取り込み
    If GetIni("FILE", "MTS_CSV", "SYS", c) Then
        Beep
        MsgBox "向け先管理マスタデータ出力用ファイル[MTS_CSV]の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    MTS_CSV = Trim(c)
    
    Me.Caption = Me.Caption & " " & LAST_UPDATE_DAY '2019.06.25 タイトルバー表示用で追加
    
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '出荷別単価設定マスタＯＰＥＮ
    If SE_SHIP_TANKA_M_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
    Wk = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    k = 0
    For i = 1 To 36
        
        For j = 1 To 36
            wkMTS_CHG_CD(k) = Mid(Wk, i, 1) & Mid(Wk, j, 1)
            k = k + 1
        Next j
    Next i
                                
                                
                                '国内外設定
    Combo(pcmbNaiGai).Clear
    Combo(pcmbNaiGai).AddItem NAIGAI1 & Space(4) & NAIGAI_NAI   '国内
    Combo(pcmbNaiGai).AddItem NAIGAI2 & Space(4) & NAIGAI_GAI   '海外
    Combo(pcmbNaiGai).ListIndex = 0
    
    Show
                                
    If List_Proc() Then
        Unload Me
    End If
                                '画面初期設定
    Clear_Field (0)
    
    Text(ptxMUKE_CODE).SetFocus
    
    End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '出荷別単価設定マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷別単価設定")
        End If
    End If
    sts = BTRV(BtOpReset, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "向け先管理マスタ")
    End If
    Set F1010601 = Nothing
    End
End Sub
Private Sub List1_DblClick()

Dim i       As Integer
Dim CODE    As String * 16

        CODE = Right(List1.List(List1.ListIndex), 16)

        Text(ptxMUKE_CODE).Text = Left(CODE, 8)
        Text(ptxSS_CODE).Text = Right(CODE, 8)

        If Dislpay_Proc() Then
            Unload Me
        End If

        Text(ptxMUKE_CODE).SetFocus

End Sub

Private Sub List1_GotFocus()
    
    If List1.ListCount > 0 Then
        If List1.ListIndex <= 0 Then
            List1.ListIndex = 0
        End If
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sts As Integer

    Select Case KeyCode
        Case vbKeyReturn
            
            Call List1_DblClick
        Case vbKeyF12
            Command(11).Value = True
    End Select

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

    Select Case Index
        Case ptxSS_CODE
            If Len(Trim(Text(ptxMUKE_CODE).Text)) = 0 Then
                Beep
                MsgBox "入力した項目はエラーです。（必須入力）"
                Text(ptxMUKE_CODE).SetFocus
                Exit Sub
            End If
    
            If Dislpay_Proc() Then
                Unload Me
            End If
    
    End Select
    
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i
    
End Sub
Private Function Data_Proc() As Integer

Dim FileNo          As Integer
Dim FileName        As String
Dim Ret             As Integer

Dim com             As Integer
Dim sts             As Integer

    Call Input_Lock

    FileNo = FreeFile
    FileName = MTS_CSV
    
    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

    On Error GoTo Error_Proc

    Open (FileName) For Output As FileNo
    
    Write #FileNo, "国内外", "得意先ｺｰﾄﾞ", "得意先名称", "倉庫／ＳＳｺｰﾄﾞ", "倉庫／ＳＳ名称", "表示略称（スキャナ用）", "読替えｺｰﾄﾞ（内部処理用）"

    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "向け先管理マスタ")
                Exit Function
        End Select
    
        Write #FileNo, StrConv(MTSREC.NAIGAI, vbUnicode),
        Write #FileNo, StrConv(MTSREC.MUKE_CODE, vbUnicode),
        Write #FileNo, StrConv(MTSREC.MUKE_NAME, vbUnicode),
        Write #FileNo, StrConv(MTSREC.SS_CODE, vbUnicode),
        Write #FileNo, StrConv(MTSREC.SS_NAME, vbUnicode),
        Write #FileNo, StrConv(MTSREC.MUKE_DNAME, vbUnicode)
    
    
        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock
    
    Beep
    MsgBox "「" & FileName & "」は正常に出力されました。"


    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName & "が使用中です。"
        Data_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        Data_Proc = True
    End If

    Call Input_UnLock



End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1010601.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010601)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010601)

    F1010601.MousePointer = vbDefault

End Sub


Public Sub List_Edit_Proc()
'----------------------------------------------------------------------------
'                   リストボックス明細表示
'----------------------------------------------------------------------------

Dim Edit    As String

    
        
    Select Case StrConv(MTSREC.NAIGAI, vbUnicode)
        Case NAIGAI_NAI
            Edit = NAIGAI1 & " "
        Case NAIGAI_GAI
            Edit = NAIGAI2 & " "
    End Select
    Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & " "
    Edit = Edit & StrConv(MTSREC.MUKE_NAME, vbUnicode) & " "
    Edit = Edit & StrConv(MTSREC.SS_CODE, vbUnicode) & " "
    Edit = Edit & StrConv(MTSREC.SS_NAME, vbUnicode) & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
    
    List1.AddItem Edit

End Sub


Private Sub List_Update_Proc(Mode As Integer)
'----------------------------------------------------------------------------
'                   リストボックス更新
'----------------------------------------------------------------------------
Dim i       As Integer
Dim CODE    As String * 16


    For i = 0 To List1.ListCount - 1
        
        CODE = Right(List1.List(i), 16)
        
        If Trim(Text(ptxMUKE_CODE).Text) = Trim(Left(CODE, 8)) And _
            Trim(Text(ptxSS_CODE).Text) = Trim(Right(CODE, 8)) Then
                List1.RemoveItem i
        End If
    
    Next i

    If Mode = 0 Then
        Call List_Edit_Proc
    End If
End Sub
