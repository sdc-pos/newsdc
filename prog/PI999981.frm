VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PI999981 
   Caption         =   "商品化指図票一括発行"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
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
   ScaleHeight     =   6270
   ScaleWidth      =   10185
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   13
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   12
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   0
      Left            =   2940
      TabIndex        =   9
      Top             =   5160
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   3900
      Left            =   4830
      TabIndex        =   6
      Top             =   1200
      Width           =   3585
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1680
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "読込み"
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
      Left            =   630
      TabIndex        =   4
      Top             =   720
      Width           =   1380
   End
   Begin VB.ListBox List1 
      Height          =   3900
      Left            =   630
      TabIndex        =   3
      Top             =   1200
      Width           =   3165
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
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
      Left            =   7455
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印刷"
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
      Left            =   6090
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ＮＧ件数"
      Height          =   255
      Index           =   4
      Left            =   6615
      TabIndex        =   11
      Top             =   5760
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "ＯＫ件数"
      Height          =   255
      Index           =   3
      Left            =   6615
      TabIndex        =   10
      Top             =   5280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "読込み件数"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   5280
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "印刷結果"
      Height          =   255
      Index           =   1
      Left            =   4830
      TabIndex        =   7
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "仕向け先"
      Height          =   255
      Index           =   0
      Left            =   525
      TabIndex        =   2
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "PI999981"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'テキスト用添字
Private Const ptxJGYOBU% = 0            '事業部
Private Const ptxNAIGAI% = 1            '国内外


Private Const ptxS_YMD% = 2             '開始　日付範囲
Private Const ptxE_YMD% = 3             '終了　日付範囲

Private Const ptxCOUNT% = 4             '対象個数

Private Const ptxSEL_CLASS% = 5         '選択　ｸﾗｽ
Private Const ptxSEL_BOX% = 6           '選択　ｸﾗｽ


Private Const ptxS_SOKO_No% = 7         '開始　倉庫№
Private Const ptxS_Retu% = 8            '開始　列
Private Const ptxS_Ren% = 9             '開始　連
Private Const ptxS_Dan% = 10            '開始　段

Private Const ptxe_SOKO_No% = 11        '終了　倉庫№
Private Const ptxe_Retu% = 12           '終了　列
Private Const ptxe_Ren% = 13            '終了　連
Private Const ptxe_Dan% = 14            '終了　段

Private Const pcmbSHIMUKE% = 0

Private IN_cnt  As Integer
Private OK_cnt  As Integer
Private NG_cnt  As Integer


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PI999981.MousePointer = vbHourglass

    Call Ctrl_Lock(PI999981)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI999981)


    PI999981.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer, chk As Integer, flg) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim yn          As Integer
    
    
    Error_Check_Proc = True
    
        
        
    Error_Check_Proc = False
    

End Function







Private Sub Command1_Click(Index As Integer)

Dim ans             As Integer
Dim i               As Integer

Dim f               As New PI999982

Dim rpt2            As New PI99998F2


Dim com             As Integer
Dim sts             As Integer

Dim FileNo          As Long
Dim wkText          As String

Dim Skip_F          As Boolean


    Select Case Index
        Case 0              '印刷
            
            
            
            Beep
            ans = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                List2.Clear
                
                OK_cnt = 0
                NG_cnt = 0
                Text1(1).Text = Format(OK_cnt, "#,##0")
                Text1(2).Text = Format(NG_cnt, "#,##0")
                
                For i = 0 To List1.ListCount - 1
            
            
            
                    Taget_SHIMUKE_CODE_KEY = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2)
                    Taget_JGYOBU_key = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)
                    Taget_NAIGAI_key = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1)
                    Taget_Hin_key = Trim(Left(List1.List(i), 20))
                    
                    Skip_F = False
                    
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Taget_JGYOBU_key)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Taget_NAIGAI_key)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Taget_Hin_key)
                    
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                            Skip_F = True
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Sub
                    
                    End Select
                    
                    
                    
                    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Taget_SHIMUKE_CODE_KEY)
                    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Taget_JGYOBU_key)
                    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Taget_NAIGAI_key)
                    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Taget_Hin_key)
                    
                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "0")
                    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
                
                    
                    
                    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                            Exit Sub
                    
                    End Select
                    
                    
                    
                    
                    If Skip_F Then
                        List2.AddItem Taget_Hin_key & " " & "NG"
                        NG_cnt = NG_cnt + 1
                        Text1(2).Text = Format(NG_cnt, "#,##0")
                    Else
                    
                        Set rpt2 = New PI99998F2
                        'レポートを印刷します。（true：印刷ダイアログあり false：なし）
                        rpt2.PrintReport False
                        Set rpt2 = Nothing
                    
                        List2.AddItem Taget_Hin_key & " " & "OK"
                        OK_cnt = OK_cnt + 1
                        Text1(1).Text = Format(OK_cnt, "#,##0")
                    
                    
                    End If
                Next i
            
                MsgBox "印刷が終了しました。"
            
            
            End If
        Case 1              '終了
            Unload Me
    
        Case 2
    
            List1.Clear
            
            CommonDialog1.Filter = "すべてのファイル (*.*)|*.*|"
            CommonDialog1.FilterIndex = 2
        
            On Error GoTo ErrHandler
        
            CommonDialog1.ShowOpen
    
            FileNo = FreeFile
            Open CommonDialog1.fileName For Input As #FileNo
    
            IN_cnt = 0
            Text1(0).Text = Format(IN_cnt, "#,##0")
    
            Do Until eof(FileNo)
                Line Input #FileNo, wkText
                If Trim(wkText) = "" Then
                    Exit Do
                End If
    
                List1.AddItem Trim(wkText)
                IN_cnt = IN_cnt + 1
    
                Text1(0).Text = Format(IN_cnt, "#,##0")
    
            Loop
    
            Close #FileNo
    
    End Select

ErrHandler:
    
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c           As String * 128
Dim sts         As Integer
Dim i           As Integer

Dim MUKE_CODE   As Variant


    If App.PrevInstance Then
        MsgBox "同一プログラム実行中です。"
        End
    End If

                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止します。"
        End
    End If
    LOG_F = RTrim(c)
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    
    
    Load PI999982
    
    
    
    
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    
    
    
    Doukon_Tbl_No(0) = "①"
    Doukon_Tbl_No(1) = "②"
    Doukon_Tbl_No(2) = "③"
    Doukon_Tbl_No(3) = "④"
    Doukon_Tbl_No(4) = "⑤"
    Doukon_Tbl_No(5) = "⑥"
    Doukon_Tbl_No(6) = "⑦"
    Doukon_Tbl_No(7) = "⑧"
    Doukon_Tbl_No(8) = "⑨"
    Doukon_Tbl_No(9) = "⑩"
    Doukon_Tbl_No(10) = "⑪"
    Doukon_Tbl_No(11) = "⑫"
    Doukon_Tbl_No(12) = "⑬"
    Doukon_Tbl_No(13) = "⑭"
    Doukon_Tbl_No(14) = "⑮"
    Doukon_Tbl_No(15) = "⑯"
    Doukon_Tbl_No(16) = "⑰"
    Doukon_Tbl_No(17) = "⑱"
    Doukon_Tbl_No(18) = "⑲"
    Doukon_Tbl_No(19) = "⑳"
    
    
    
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
    
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
    
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
    
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
    
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI999981 = Nothing
    Set PI999982 = Nothing

    End
End Sub


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


