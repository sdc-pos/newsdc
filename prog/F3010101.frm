VERSION 5.00
Begin VB.Form F3010101 
   BackColor       =   &H00C0C0C0&
   Caption         =   "GLICS 在庫取り込みV1.00"
   ClientHeight    =   4710
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   7320
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
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '画面の中央
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "データ集計中"
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
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "GLICS在庫取り込み処理"
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
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   5040
   End
End
Attribute VB_Name = "F3010101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type YUKO_SOKO_TBL             '有効ﾎｽﾄ倉庫取り込みテーブル
    HS_SOKO As String * 2
    NAIGAI As String * 1
End Type

Private SOKO_T() As YUKO_SOKO_TBL       '倉庫情報
Private Function HS_ZAI_INIT() As Integer


Dim In_Rec      As String
Dim In_Text     As Variant
Dim c           As String * 128
Dim FileNo      As Integer
Dim fileName    As String
    
Dim ans         As Integer
    
Dim com         As Integer
Dim sts         As Integer
    
Dim upd_com     As Integer
Dim T_Zai_Qty       As Long
Dim SAI_QTY         As Long
    
    
Dim Skip_Flg    As Boolean
    
Dim i           As Integer
Dim j           As Integer
    
    HS_ZAI_INIT = True
    
                                '初期クリアー
    com = BtOpGetFirst
    Do
        DoEvents
                
        Do
            sts = BTRV(com + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrEOF
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                HS_ZAI_INIT = True
            End Select
        Loop


        If sts = BtErrEOF Then
            Exit Do
        End If

        Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "00000000")
        Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "00000000")
        
        Do
            sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目マスタ")
                HS_ZAI_INIT = True
            End Select
        Loop
                                
        com = BtOpGetNext
    Loop
                                
                                
                                
                                
    com = BtOpGetGreater

    Call UniCode_Conv(K0_SUMZ.JGYOBU, SENTAKU)
    Call UniCode_Conv(K0_SUMZ.NAIGAI, "")
    Call UniCode_Conv(K0_SUMZ.HIN_GAI, "")

    Do
                
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(SUMZREC.JGYOBU, vbUnicode) <> SENTAKU Then
                        sts = BtErrEOF
                    End If
                    
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'これはない
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "在庫集計データ")
                    Exit Function
            End Select
        Loop
        
        If sts Then
            Exit Do
        End If
        
        '前回→前々回
        Call UniCode_Conv(SUMZREC.ZEN_HS_ZAIQTY, StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
        
        Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
        Call UniCode_Conv(SUMZREC.SAI_QTY, StrConv(SUMZREC.T_Zai_Qty, vbUnicode))
        
        Do
            sts = BTRV(BtOpUpdate, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'これはない
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "在庫集計データ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
        
        DoEvents
    Loop
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                'ファイル名取り込み
    If GetIni("FILE", "HS_NEW_ZAI", "SYS", c) Then
        Beep
        MsgBox "ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    fileName = RTrim(c)
    
    
    
    FileNo = FreeFile
        
    On Error GoTo HS_SIJ_Op_Err    'ｴﾗｰﾄﾗｯﾌﾟON
    Open fileName For Input As #FileNo
    On Error GoTo 0
        
        
        
    Do While Not EOF(FileNo)
        
        DoEvents
        Line Input #FileNo, In_Rec
        In_Text = Split(In_Rec, vbTab, -1)
    
        If CStr(In_Text(0)) = "00023100" Then
    
            Call UniCode_Conv(K0_ITEM.JGYOBU, SENTAKU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
            Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(In_Text(1)))
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    HS_ZAI_INIT = True
                End Select
            Loop
    
    
            If sts = BtNoErr Then
                Select Case In_Text(2)
                    Case "S2"
                        If IsNumeric(In_Text(3)) Then
                            Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, Format(CLng(In_Text(3)), "00000000"))
                        Else
                            Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "00000000")
                        End If
                    Case "P2"
                        If IsNumeric(In_Text(3)) Then
                            Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, Format(CLng(In_Text(3)), "00000000"))
                        Else
                            Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "00000000")
                        End If
                End Select
                Call UniCode_Conv(ITEMREC.S_TANTO, CStr(In_Text(7)))
    
    
                Do
                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "品目マスタ")
                        HS_ZAI_INIT = True
                    End Select
                Loop
    
    
            End If
    
    
            Skip_Flg = False
            If Not IsNumeric(In_Text(3)) Then
                Skip_Flg = True
            End If
        
            
            If Not Skip_Flg Then
''                Skip_Flg = True
''                For i = 0 To UBound(JGYOBU_T)
''                    If SENTAKU = JGYOBU_T(i).CODE Then
''                        For j = 0 To UBound(SOKO_T, 2)
''                            If In_Text(2) = SOKO_T(i, j).HS_SOKO Then
''                                Skip_Flg = False
''                                Exit For
''                            End If
''                        Next j
''                        Exit For
''                    End If
''                Next i
        
                If Not Skip_Flg Then
                    '対象データ
                    Call UniCode_Conv(K0_SUMZ.JGYOBU, SENTAKU)
                    Call UniCode_Conv(K0_SUMZ.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_SUMZ.HIN_GAI, CStr(In_Text(1)))
            
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                        Select Case sts
                            Case BtNoErr
                                upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                upd_com = BtOpInsert
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'これはない
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "在庫集計データ")
                                Exit Function
                        End Select
                    Loop
                
                    If upd_com = BtOpInsert Then
                    '新規追加時、品目マスタよりホスト棚番獲得
                    
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SENTAKU)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, SOKO_T(i, j).NAIGAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(In_Text(1)))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                            'ありえないがスルー
                                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                Call UniCode_Conv(ITEMREC.ST_REN, "")
                                Call UniCode_Conv(ITEMREC.ST_DAN, "")
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        End Select
            
            
                        Call UniCode_Conv(SUMZREC.JGYOBU, SENTAKU)
                        Call UniCode_Conv(SUMZREC.NAIGAI, SOKO_T(i, j).NAIGAI)
                        Call UniCode_Conv(SUMZREC.HIN_GAI, CStr(In_Text(1)))
                        Call UniCode_Conv(SUMZREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                        Call UniCode_Conv(SUMZREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                        Call UniCode_Conv(SUMZREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                        Call UniCode_Conv(SUMZREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(SUMZREC.T_Zai_Qty, "00000000")
                        Call UniCode_Conv(SUMZREC.ZEN_Zai_Qty, "00000000")
                        Call UniCode_Conv(SUMZREC.SYK_E_QTY, "00000000")
                        Call UniCode_Conv(SUMZREC.NYUKA_YQTY, "00000000")
                        Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                        Call UniCode_Conv(SUMZREC.ZEN_HS_ZAIQTY, "00000000")
                        Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")
                        Call UniCode_Conv(SUMZREC.FILLER, "")
                        
                    End If
                    
                    Call UniCode_Conv(SUMZREC.SUM_DT, Format(Date, "yyyymmdd"))
            
                    Call UniCode_Conv(SUMZREC.HS_ZAIQTY, Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) + CLng(CStr(In_Text(3))), "00000000"))
            
                    SAI_QTY = CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) - CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
                    
                    If SAI_QTY >= 0 Then
                        Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "00000000"))
                    Else
                        Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "0000000"))
                    End If
            
                    Do
                    
                        sts = BTRV(upd_com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'これはない
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, upd_com, "在庫集計データ")
                                Exit Function
                        End Select
                
                    Loop
                End If
            End If
        End If
    
    
    Loop

    Close #FileNo
    Exit Function
HS_SIJ_Op_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case ErrDiskNotReady
            Beep
            ans = MsgBox("ドライブを確認して下さい", vbYesNo + vbExclamation + vbDefaultButton1, "確認入力")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable
            Beep
            ans = MsgBox("ドライブまたはパスが見つかりません" & fileName, vbExclamation)
        Case ErrNotFound
            Beep
            ans = MsgBox("ファイルが見つかりません" & fileName, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [HS_SIJ Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select
End Function

Private Sub Form_Activate()

Dim ans As Integer

'    ans = MsgBox("「GLICS在庫取り込み処理」実行しますか？", vbYesNo, "確認処理")
'
'    If ans = vbNo Then
'        Unload Me
'    End If

        
    F3010101.MousePointer = vbHourglass
    Label1(1).Visible = True
    
    If HS_ZAI_INIT() Then           '差異チェック処理
        Unload Me
    End If
    
    Label1(1).Visible = False
    

    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()
Dim c   As String * 128
Dim sts As Integer
Dim i   As Integer
Dim j   As Integer

Dim Max_Soko    As Integer
    
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
                                                 
    If JGYOB_TB_Set(1) Then      '事業部の獲得
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                
                                '倉庫最大数を取り込み
                                
    If GetIni(App.EXEName, "MAX_SOKO", "SYS", c) Then
        Max_Soko = 1
    Else
        If Not IsNumeric(RTrim(c)) Then
            Max_Soko = 1
        Else
            Max_Soko = CInt(RTrim(c))
        End If
    End If
                                
                                '在庫取り込み用テーブル作成
    ReDim SOKO_T(0 To UBound(JGYOBU_T), 0 To Max_Soko - 1)
                                '倉庫情報取り込み
    For i = 0 To UBound(JGYOBU_T)
        j = 0
        Do
                                '有効倉庫獲得
            If GetIni(App.EXEName, "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0"), "SYS", c) Then
                Beep
                MsgBox "倉庫情報の獲得に失敗しました。処理を中止して下さい。"
                End
            End If
    
            If Trim(c) = "**" Then  '倉庫指定終了
                Exit Do
            End If
    
    
'            ReDim Preserve SOKO_T(0 To i, 0 To j)
            SOKO_T(i, j).HS_SOKO = Trim(c)
                                '国内外情報獲得
            If GetIni(App.EXEName, "NAIG" & JGYOBU_T(i).CODE & Format(j + 1, "0"), "SYS", c) Then
                Beep
                MsgBox "国内外情報の獲得に失敗しました。処理を中止して下さい。"
                End
            End If
            
            SOKO_T(i, j).NAIGAI = Trim(c)
            j = j + 1
        Loop
    
    Next i
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If

                                '品目マスタＯＰＥＮ
    If SUMZ_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F3010101 = Nothing

    End
End Sub


