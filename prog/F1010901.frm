VERSION 5.00
Begin VB.Form F1010901 
   BackColor       =   &H00FFFFFF&
   Caption         =   "箱マスタメンテナンス"
   ClientHeight    =   6315
   ClientLeft      =   2130
   ClientTop       =   2835
   ClientWidth     =   11640
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
   ScaleHeight     =   6315
   ScaleWidth      =   11640
   StartUpPosition =   2  '画面の中央
   Begin VB.ListBox List1 
      Height          =   4140
      Index           =   0
      ItemData        =   "F1010901.frx":0000
      Left            =   360
      List            =   "F1010901.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   26
      Top             =   1440
      Width           =   10935
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   6
      Left            =   10200
      MaxLength       =   5
      TabIndex        =   25
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   5
      Left            =   8640
      MaxLength       =   5
      TabIndex        =   24
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   4
      Left            =   7080
      MaxLength       =   5
      TabIndex        =   21
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   3
      Left            =   5520
      MaxLength       =   5
      TabIndex        =   20
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   2
      Left            =   3960
      MaxLength       =   5
      TabIndex        =   17
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   1
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   16
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   360
      MaxLength       =   4
      TabIndex        =   13
      Top             =   960
      Width           =   615
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "確 認"
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '実線
      Caption         =   "Ｃ－２"
      Height          =   375
      Index           =   5
      Left            =   9360
      TabIndex        =   23
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '実線
      Caption         =   "Ｃ－１"
      Height          =   375
      Index           =   4
      Left            =   7800
      TabIndex        =   22
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '実線
      Caption         =   "Ｂ－２"
      Height          =   375
      Index           =   3
      Left            =   6240
      TabIndex        =   19
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '実線
      Caption         =   "Ｂ－１"
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   18
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '実線
      Caption         =   "Ａ－２"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   15
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '実線
      Caption         =   "Ａ－１"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   14
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "個装箱№"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "F1010901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxPack_No% = 0       '倉庫№
Private Const ptxRANK_A1% = 1       'ランク　Ａ－１
Private Const ptxRANK_A2% = 2       'ランク　Ａ－２
Private Const ptxRANK_B1% = 3       'ランク　Ｂ－１
Private Const ptxRANK_B2% = 4       'ランク　Ｂ－２
Private Const ptxRANK_C1% = 5       'ランク　Ｃ－１
Private Const ptxRANK_C2% = 6       'ランク　Ｃ－２

Private Const Text_Max% = 6

Private Const pLstPack% = 0         '棚情報

Private PACKING_CSV As String

Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim sts     As Integer
    
    Select Case Index
        Case 0                  '追加／変更
            
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
            
            
            End If
            
        Case 3                  '削除
            
            Beep
            yn = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                sts = Delete_Proc
                Select Case sts
                    Case False
                    Case True
                        Unload Me
                    Case SYS_CANCEL
                End Select
            
            
            End If
                    
        Case 4                  '確認画面
                    
                    
            F1010902.Show vbModal
                    
        Case 8                  'データ出力
            Beep
            yn = MsgBox("データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Data_Proc() Then
                    Unload Me
                End If
            End If
                    
        Case 11                 '終了
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
                                'ＣＳＶファイル名取り込み
    If GetIni("FILE", "PACKING_CSV", "SYS", c) Then
        Beep
        MsgBox "個装箱マスタデータ出力用ファイル[PACKING_CSV]の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    PACKING_CSV = Trim(c)
                                '個装箱マスタＯＰＥＮ
    If PACKING_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '棚別個装箱マスタＯＰＥＮ
    If TPACKING_Open(BtOpenNomal) Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '個装箱マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "個装箱マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
                                            '棚個装箱マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚別個装箱マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
    
    
    sts = BTRV(BtOpReset, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
        Beep
        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
    End If
    
    Set F1010901 = Nothing

    End

End Sub

Private Sub List1_DblClick(Index As Integer)

Dim Edit    As String
    
    Text(ptxPack_No).Text = Left(List1(pLstPack).List(List1(pLstPack).ListIndex), 4)
    Text(ptxPack_No).SetFocus

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
        Case ptxPack_No                 '個装箱№
            
            
            If Len(Text(Index).Text) < Text(Index).MaxLength Then
                Call List_Disp_Proc
            End If
            
            sts = Pack_Read_Proc
            Select Case sts
                Case False              '登録済み
                
                    Text(ptxRANK_A1).Text = Format(StrConv(PACKINGREC.RANK_A1, vbUnicode), "#0")
                    Text(ptxRANK_A2).Text = Format(StrConv(PACKINGREC.RANK_A2, vbUnicode), "#0")
                    Text(ptxRANK_B1).Text = Format(StrConv(PACKINGREC.RANK_B1, vbUnicode), "#0")
                    Text(ptxRANK_B2).Text = Format(StrConv(PACKINGREC.RANK_B2, vbUnicode), "#0")
                    Text(ptxRANK_C1).Text = Format(StrConv(PACKINGREC.RANK_C1, vbUnicode), "#0")
                    Text(ptxRANK_C2).Text = Format(StrConv(PACKINGREC.RANK_C2, vbUnicode), "#0")
                
                Case True
                    
'                    Text(ptxRANK_A1).Text = ""
'                    Text(ptxRANK_A2).Text = ""
'                    Text(ptxRANK_B1).Text = ""
'                    Text(ptxRANK_B2).Text = ""
'                    Text(ptxRANK_C1).Text = ""
'                    Text(ptxRANK_C2).Text = ""
                
                Case SYS_ERR
                    Unload Me
            End Select
        Case Else
            If Not IsNumeric(Text(Index).Text) Then
                Beep
                MsgBox "入力した項目はエラーです｡ "
                Text(Index).SetFocus
                Exit Sub
            Else
                Text(Index).Text = Format(Text(Index).Text, "##0")
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
    
    List1(pLstPack).Clear

    Call UniCode_Conv(K0_PACKING.PACKING_NO, Text(ptxPack_No).Text)
    
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "個装箱マスタ")
                Exit Function
        End Select
    
        Call ListBox_Item_Set(Edit)
    
     
        List1(pLstPack).AddItem Edit
     
        com = BtOpGetNext
     Loop

    List_Disp_Proc = False

End Function

Private Function Pack_Read_Proc() As Integer

Dim sts     As Integer

    Pack_Read_Proc = True
    
    Call UniCode_Conv(K0_PACKING.PACKING_NO, Text(ptxPack_No).Text)
    
    sts = BTRV(BtOpGetEqual, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Exit Function
        Case Else
            Pack_Read_Proc = SYS_ERR
            Call File_Error(sts, BtOpGetEqual, "個装箱マスタ")
            Exit Function
    End Select

    Pack_Read_Proc = False
    
End Function

Private Function Update_Proc() As Integer

Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer
Dim com     As Integer
Dim Edit    As String
                                            
                                            
    Update_Proc = True
    
    If Err_Chk_Proc() Then
        Update_Proc = False
        Exit Function
    End If
    
    Call Input_Lock
    
    Call UniCode_Conv(K0_PACKING.PACKING_NO, Text(ptxPack_No).Text)
    
    Do
        
        sts = BTRV(BtOpGetEqual + BtSNoWait, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<PACKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "個装箱マスタ")
                Exit Function
        End Select
    
    Loop
    
    If sts = BtErrKeyNotFound Then
                            '新規データ
        Call UniCode_Conv(PACKINGREC.PACKING_NO, Text(ptxPack_No).Text)
        Call UniCode_Conv(PACKINGREC.FILLER, "")
        com = BtOpInsert
    Else
        com = BtOpUpdate
    End If
    
    Call UniCode_Conv(PACKINGREC.RANK_A1, Format(CLng(Text(ptxRANK_A1).Text), "00000000"))
    Call UniCode_Conv(PACKINGREC.RANK_A2, Format(CLng(Text(ptxRANK_A2).Text), "00000000"))
    Call UniCode_Conv(PACKINGREC.RANK_B1, Format(CLng(Text(ptxRANK_B1).Text), "00000000"))
    Call UniCode_Conv(PACKINGREC.RANK_B2, Format(CLng(Text(ptxRANK_B2).Text), "00000000"))
    Call UniCode_Conv(PACKINGREC.RANK_C1, Format(CLng(Text(ptxRANK_C1).Text), "00000000"))
    Call UniCode_Conv(PACKINGREC.RANK_C2, Format(CLng(Text(ptxRANK_C2).Text), "00000000"))
    
    Do
        sts = BTRV(com, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<PACKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "個装箱マスタ")
                Exit Function
        End Select
    Loop
    
        
    If com = BtOpInsert Then
        Call ListBox_Item_Set(Edit)
        List1(pLstPack).AddItem Edit
        
    Else
        Call ListBox_Update_Proc(0)
    End If
    
    
    Text(ptxPack_No).Text = ""
    Call Input_UnLock
    Text(ptxPack_No).SetFocus

    Update_Proc = False
End Function

Private Function Delete_Proc() As Integer

Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer
Dim com     As Integer
                                            
                                            
    Delete_Proc = True
    
    Call Input_Lock
    
    Call UniCode_Conv(K0_PACKING.PACKING_NO, Text(ptxPack_No).Text)
    
    Do
        
        sts = BTRV(BtOpGetEqual + BtSNoWait, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<PACKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Delete_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "個装箱マスタ")
                Exit Function
        End Select
    
    Loop
    
    Do
        sts = BTRV(BtOpDelete, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<PACKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Delete_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "個装箱マスタ")
                Exit Function
        End Select
    Loop
    
    Call ListBox_Update_Proc(1)
    
    
    Text(ptxPack_No).Text = ""
    Call Input_UnLock
    Text(ptxPack_No).SetFocus

    Delete_Proc = False
End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1010901.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010901)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010901)


    F1010901.MousePointer = vbDefault

End Sub


Private Function Err_Chk_Proc() As Integer

Dim i   As Integer

    Err_Chk_Proc = True

    If Len(Trim(Text(ptxPack_No).Text)) < Text(ptxPack_No).MaxLength Then
        Beep
        MsgBox "入力した項目はエラーです｡ "
        Text(ptxPack_No).SetFocus
        Exit Function
    End If

    For i = ptxRANK_A1 To ptxRANK_C2
        If Len(Trim(Text(i).Text)) = 0 Then
            Beep
            MsgBox "入力した項目はエラーです｡ "
            Text(i).SetFocus
            Exit Function
        End If
    
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "入力した項目はエラーです｡ "
            Text(i).SetFocus
            Exit Function
        End If
    
        Text(i).Text = Format(CInt(Text(i).Text), "#0")
    
    Next i


    If CInt(Text(ptxRANK_A1).Text) > CInt(Text(ptxRANK_A2).Text) And _
        CInt(Text(ptxRANK_A2).Text) > CInt(Text(ptxRANK_B1).Text) And _
        CInt(Text(ptxRANK_B1).Text) > CInt(Text(ptxRANK_B2).Text) And _
        CInt(Text(ptxRANK_B2).Text) > CInt(Text(ptxRANK_C1).Text) And _
        CInt(Text(ptxRANK_C1).Text) > CInt(Text(ptxRANK_C2).Text) Then
    Else
        Beep
        MsgBox "入力した項目はエラーです｡ "
        Text(ptxRANK_A1).SetFocus
        Exit Function
    End If
    
    
    Err_Chk_Proc = False
    
End Function

Private Sub ListBox_Update_Proc(Mode As Integer)

Dim i       As Integer
Dim Edit    As String

    For i = 0 To List1(pLstPack).ListCount - 1
    
        If Trim(StrConv(PACKINGREC.PACKING_NO, vbUnicode)) = Trim(Left(List1(pLstPack).List(i), 4)) Then
            List1(pLstPack).RemoveItem i
    
        
            If Mode = 1 Then
                Exit Sub
            End If
        
            Exit For
        End If
    
    Next i


    Call ListBox_Item_Set(Edit)
    
     
    List1(pLstPack).AddItem Edit

End Sub

Private Sub ListBox_Item_Set(Edit As String)
     
Dim Work    As String
     
    Edit = StrConv(PACKINGREC.PACKING_NO, vbUnicode) & Space(13)
    
    Work = Format(StrConv(PACKINGREC.RANK_A1, vbUnicode), "#0")
    If Len(Work) < 5 Then
        Edit = Edit & Space(5 - Len(Work)) & Work & Space(8)
    Else
        Edit = Edit & Work & Space(8)
    End If
    
    Work = Format(StrConv(PACKINGREC.RANK_A2, vbUnicode), "#0")
    If Len(Work) < 5 Then
        Edit = Edit & Space(5 - Len(Work)) & Work & Space(8)
    Else
        Edit = Edit & Work & Space(8)
    End If
    
    Work = Format(StrConv(PACKINGREC.RANK_B1, vbUnicode), "#0")
    If Len(Work) < 5 Then
        Edit = Edit & Space(5 - Len(Work)) & Work & Space(8)
    Else
        Edit = Edit & Work & Space(8)
    End If
    
    Work = Format(StrConv(PACKINGREC.RANK_B2, vbUnicode), "#0")
    If Len(Work) < 5 Then
        Edit = Edit & Space(5 - Len(Work)) & Work & Space(8)
    Else
        Edit = Edit & Work & Space(8)
    End If
    
    Work = Format(StrConv(PACKINGREC.RANK_C1, vbUnicode), "#0")
    If Len(Work) < 5 Then
        Edit = Edit & Space(5 - Len(Work)) & Work & Space(8)
    Else
        Edit = Edit & Work & Space(8)
    End If
    
    Work = Format(StrConv(PACKINGREC.RANK_C2, vbUnicode), "#0")
    If Len(Work) < 5 Then
        Edit = Edit & Space(5 - Len(Work)) & Work & Space(8)
    Else
        Edit = Edit & Work & Space(8)
    End If

End Sub
Private Function Data_Proc() As Integer

Dim FileNo          As Integer
Dim fileName        As String
Dim Ret             As Integer

Dim com             As Integer
Dim sts             As Integer

    Call Input_Lock

    FileNo = FreeFile
    fileName = PACKING_CSV
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & Right(Trim(fileName), Len(Trim(fileName)) - Ret)

    On Error GoTo Error_Proc

    Open (fileName) For Output As FileNo
    
    Write #FileNo, "個装箱№", "Ａ－１", "Ａ－２", "Ｂ－１", "Ｂ－２", "Ｃ－１", "Ｃ－２"

    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "個装箱マスタ")
                Exit Function
        End Select
    
        Write #FileNo, StrConv(PACKINGREC.PACKING_NO, vbUnicode),
        Write #FileNo, Format(StrConv(PACKINGREC.RANK_A1, vbUnicode), "#0"),
        Write #FileNo, Format(StrConv(PACKINGREC.RANK_A2, vbUnicode), "#0"),
        Write #FileNo, Format(StrConv(PACKINGREC.RANK_B1, vbUnicode), "#0"),
        Write #FileNo, Format(StrConv(PACKINGREC.RANK_B2, vbUnicode), "#0"),
        Write #FileNo, Format(StrConv(PACKINGREC.RANK_C1, vbUnicode), "#0"),
        Write #FileNo, Format(StrConv(PACKINGREC.RANK_C2, vbUnicode), "#0")
    
    
        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock
    
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"


    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        Data_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        Data_Proc = True
    End If

    Call Input_UnLock



End Function

