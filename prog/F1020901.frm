VERSION 5.00
Begin VB.Form F1020901 
   BackColor       =   &H00FFFFFF&
   Caption         =   "複数原産国部品入庫管理リスト　(F102090 2011.07.14 12:00)"
   ClientHeight    =   3312
   ClientLeft      =   2028
   ClientTop       =   2268
   ClientWidth     =   10932
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3312
   ScaleWidth      =   10932
   StartUpPosition =   2  '画面の中央
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1020901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NormalFont      As New StdFont              '印刷フォント
Dim NormalBoldFont  As New StdFont              '印刷フォント
Dim SmallFont       As New StdFont              '印刷フォント
Dim SmallBoldFont   As New StdFont              '印刷フォント
Dim LargeFont       As New StdFont              '印刷フォント
Dim LargeUnderFont  As New StdFont              '印刷フォント

Private Const MGN_L% = 1                        '左余白（桁数：１から）
Private Const MGN_U% = 2                        '上余白（行数：１から）
Private Const LMax% = 6





Private Sub Form_Activate()

    If Print_Proc() Then
        Unload Me
    End If


    Unload Me

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
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

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


                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

                                '入荷データファイルOPEN
    If Y_NYU_Open(0) Then
        Unload Me
    End If
                                '在庫データファイルOPEN
    If ZAIKO_Open(0) Then
        Unload Me
    End If
                                '月平均データファイルOPEN
    If AVE_SYUKA_Open(0) Then
        Unload Me
    End If
    Show


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '入荷データファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷データファイル")
        End If
    End If
                                            '在庫データファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データファイル")
        End If
    End If
                                            '月平均出荷ＣＬＯＳＥ
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "月平均出荷")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020901 = Nothing

    End
End Sub


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1020901.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020901)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020901)


    F1020901.MousePointer = vbDefault

End Sub
Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   印刷処理
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim i           As Integer

Dim Lcnt        As Integer
    
Dim Print_Now   As String
    
    
Dim SUMI_QTY()      As Long
Dim MI_QTY()        As Long
Dim GENSANKOKU()    As String * 20
    
Dim svSoko      As String * 2
Dim svRetu      As String * 2
Dim svRen       As String * 2
Dim svDan       As String * 2
    
    
Dim Maru_Suuji  As String
    
Dim Read_next   As Integer
    
    
    Print_Proc = True

    Print_Now = Format(Now, "YYYY/MM/DD HH:MM")
    
    Printer.Orientation = vbPRORLandscape
    Lcnt = 99
    
    With NormalFont
        .NAME = F1020901.Font.NAME
        .Size = 11
        .Bold = False
    
    End With
    With NormalBoldFont
        .NAME = F1020901.Font.NAME
        .Size = 11
        .Bold = True
    End With
    
    With SmallFont
        .NAME = F1020901.Font.NAME
        .Size = 9
        .Bold = False
    End With
    
    With SmallBoldFont
        .NAME = F1020901.Font.NAME
        .Size = 9
        .Bold = True
    End With
    
    
    With LargeFont
        .NAME = F1020901.Font.NAME
        .Size = 14
        .Bold = True
    End With
    
    With LargeUnderFont
        .NAME = F1020901.Font.NAME
        .Size = 14
        .Bold = True
        .Underline = True
    End With
    
    
    
    com = BtOpGetGreater


    Call UniCode_Conv(K4_Y_NYU.LIST_OUT_END_F, "0")
    Call UniCode_Conv(K4_Y_NYU.JGYOBU, "")
    Call UniCode_Conv(K4_Y_NYU.NAIGAI, "")
    Call UniCode_Conv(K4_Y_NYU.HIN_NO, "")



    Do
        DoEvents
        
        sts = BTRV(com, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
        
        
        Select Case sts
            Case BtNoErr
                If StrConv(Y_NYUREC.LIST_OUT_END_F, vbUnicode) <> "0" Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
        
        
            Case Else
                Call File_Error(sts, com, "入荷予定データ", 0)
                Exit Function

        End Select
        
        
        svSoko = ""
        If StrConv(Y_NYUREC.LIST_NYU_KANRI_F, vbUnicode) <> "0" Then
        Else
        
        
    
    
            Call UniCode_Conv(K4_ZAIKO.JGYOBU, StrConv(Y_NYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K4_ZAIKO.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K4_ZAIKO.HIN_GAI, StrConv(Y_NYUREC.HIN_NO, vbUnicode))
            Call UniCode_Conv(K4_ZAIKO.Soko_No, "")
            Call UniCode_Conv(K4_ZAIKO.Retu, "")
            Call UniCode_Conv(K4_ZAIKO.Ren, "")
            Call UniCode_Conv(K4_ZAIKO.Dan, "")
    
    
    
            com = BtOpGetGreater
    
            
            
            
            
            Do
            
                DoEvents
                sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
                
                
                Select Case sts
                    Case BtNoErr
                    
                    
                        If Trim(StrConv(Y_NYUREC.JGYOBU, vbUnicode)) <> Trim(StrConv(ZAIKOREC.JGYOBU, vbUnicode)) Or _
                            Trim(StrConv(Y_NYUREC.NAIGAI, vbUnicode)) <> Trim(StrConv(ZAIKOREC.NAIGAI, vbUnicode)) Or _
                            Trim(StrConv(Y_NYUREC.HIN_NO, vbUnicode)) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                            
                            Exit Do
                        End If
                                        
                    
                    Case BtErrEOF
                        Exit Do
                
                
                    Case Else
                        Call File_Error(sts, com, "在庫データ", 0)
                        Exit Function
        
                End Select
            
                If StrConv(Y_NYUREC.ID_NO2, vbUnicode) = StrConv(ZAIKOREC.ID_NO2, vbUnicode) Then
                
                Else
            
                    If Trim(svSoko) = "" Then
                        svSoko = StrConv(ZAIKOREC.Soko_No, vbUnicode)
                        svRetu = StrConv(ZAIKOREC.Retu, vbUnicode)
                        svRen = StrConv(ZAIKOREC.Ren, vbUnicode)
                        svDan = StrConv(ZAIKOREC.Dan, vbUnicode)
                        
                        Erase SUMI_QTY
                        Erase MI_QTY
                        Erase GENSANKOKU
                        
                        ReDim Preserve SUMI_QTY(0 To 0)
                        ReDim Preserve MI_QTY(0 To 0)
                        ReDim Preserve GENSANKOKU(0 To 0)
                        GENSANKOKU(0) = StrConv(ZAIKOREC.GENSANKOKU, vbUnicode)
                        SUMI_QTY(0) = 0
                        MI_QTY(0) = 0
                    End If
                
                    If svSoko <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Or _
                        svRetu <> StrConv(ZAIKOREC.Retu, vbUnicode) Or _
                        svRen <> StrConv(ZAIKOREC.Ren, vbUnicode) Or _
                        svDan <> StrConv(ZAIKOREC.Dan, vbUnicode) Then
                
                
                
                        For i = 0 To UBound(GENSANKOKU)
                
                
                            
                            If Head_Print_Proc(Print_Now, Lcnt) Then
                                Exit Function
                            End If
        
                            Lcnt = Lcnt + 1
                            
                            Select Case Lcnt
                            
                            
                            
                                Case 1
                                    Maru_Suuji = "①"
                                Case 2
                                    Maru_Suuji = "②"
                                Case 3
                                    Maru_Suuji = "③"
                                Case 4
                                    Maru_Suuji = "④"
                                Case 5
                                    Maru_Suuji = "⑤"
                                Case 6
                                    Maru_Suuji = "⑥"
                                Case 7
                                    Maru_Suuji = "⑦"
                                Case 8
                                    Maru_Suuji = "⑧"
                                Case 9
                                    Maru_Suuji = "⑨"
                                Case 10
                                    Maru_Suuji = "⑩"
                                Case 11
                                    Maru_Suuji = "⑪"
                                Case 12
                                    Maru_Suuji = "⑫"
                                Case 13
                                    Maru_Suuji = "⑬"
                                Case 14
                                    Maru_Suuji = "⑭"
                                Case 15
                                    Maru_Suuji = "⑮"
                                Case 16
                                    Maru_Suuji = "⑯"
                                Case 17
                                    Maru_Suuji = "⑰"
                                Case 18
                                    Maru_Suuji = "⑱"
                                Case 19
                                    Maru_Suuji = "⑲"
                                Case 20
                                    Maru_Suuji = "⑳"
                                Case Else
                                    Maru_Suuji = "～"
                            End Select
                            
                            
                            
                            '------------------------------------   17行目
                            Set Printer.Font = NormalFont
                            Printer.Print Tab(MGN_L + 13);
                            Printer.Print "原産国名：";
                            Printer.Print Tab(MGN_L + 56);      '60-->56
                            Printer.Print "数量" & Maru_Suuji & "(商品化済 - 未商品)";
                            Printer.Print Tab(MGN_L + 85);
                            Printer.Print "棚番" & Maru_Suuji;
                            Printer.Print Tab(MGN_L + 105);
                            Printer.Print "現品票";
                            Set Printer.Font = SmallFont
                            Printer.Print Tab(MGN_L + 141);
                            Printer.Print "ﾗﾍﾞﾙ発行";
                            Printer.Print Tab(MGN_L + 151);
                            Printer.Print "専用ﾗﾍﾞﾙ";
                            Printer.Print Tab(MGN_L + 161);
                            Set Printer.Font = NormalFont
                            Printer.Print "　　備考"
                            '------------------------------------   18行目
                            Set Printer.Font = NormalFont
                            Printer.Print Tab(MGN_L + 5);
                            
                            If Lcnt = 1 Then
                                Printer.Print "在庫品";
                            Else
                                Printer.Print "　　　";
                            End If
                            Printer.Print Tab(MGN_L + 13);
                            Set Printer.Font = NormalBoldFont
                            Printer.Print StrConv(Trim(GENSANKOKU(i)), vbWide);  'trim
                            Printer.Print Tab(MGN_L + 49);      '53-->49
                            Printer.Print StrConv(Space(8 - Len(Format(SUMI_QTY(i) + MI_QTY(i), "#,0"))) & _
                                            Format(SUMI_QTY(i) + MI_QTY(i), "#,0") & "(" & _
                                            Format(SUMI_QTY(i), "#,0") & "-" & _
                                            Format(MI_QTY(i), "#,0") & ")", vbWide);
                            Set Printer.Font = NormalFont
                            Printer.Print Tab(MGN_L + 85);
                            Printer.Print svSoko & "-" & svRetu & "-" & svRen & "-" & svDan;
                            Printer.Print Tab(MGN_L + 107);
                            Printer.Print "－";
                            Set Printer.Font = SmallFont
                            Printer.Print Tab(MGN_L + 143);     '114-->143
                            Set Printer.Font = NormalFont
                            Printer.Print "□";
                            Set Printer.Font = SmallFont
                            Printer.Print Tab(MGN_L + 153);     '121-->153
                            Set Printer.Font = NormalFont
                            Printer.Print "□"
                            '------------------------------------   19行目
                            Set Printer.Font = NormalFont
                            Printer.Print Tab(MGN_L + 5);
                            Printer.Print String(70, "─")
        
        
        
                        Next i
                    
                    
                    
                    
                    
                    
                        svSoko = StrConv(ZAIKOREC.Soko_No, vbUnicode)
                        svRetu = StrConv(ZAIKOREC.Retu, vbUnicode)
                        svRen = StrConv(ZAIKOREC.Ren, vbUnicode)
                        svDan = StrConv(ZAIKOREC.Dan, vbUnicode)
                        
                        Erase SUMI_QTY
                        Erase MI_QTY
                        Erase GENSANKOKU
                        
                        ReDim Preserve SUMI_QTY(0 To 0)
                        ReDim Preserve MI_QTY(0 To 0)
                        ReDim Preserve GENSANKOKU(0 To 0)
                        GENSANKOKU(0) = StrConv(ZAIKOREC.GENSANKOKU, vbUnicode)
                        SUMI_QTY(0) = 0
                        MI_QTY(0) = 0
                    
                    
                    
                    
                    
                    End If
                    
                    
                    For i = 0 To UBound(GENSANKOKU)
                    
                    
                        If Trim(StrConv(ZAIKOREC.GENSANKOKU, vbUnicode)) = Trim(GENSANKOKU(i)) Then
                            Exit For
                        End If
                    
                    Next i
                    
                    
                    
                    If i > UBound(GENSANKOKU) Then
                    
                    
                        ReDim Preserve SUMI_QTY(0 To i)
                        ReDim Preserve MI_QTY(0 To i)
                        ReDim Preserve GENSANKOKU(0 To i)
                        GENSANKOKU(i) = StrConv(ZAIKOREC.GENSANKOKU, vbUnicode)
                        SUMI_QTY(i) = 0
                        MI_QTY(i) = 0
                    
                    End If
                    
                    
                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                        SUMI_QTY(i) = SUMI_QTY(i) + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    Else
                        MI_QTY(i) = MI_QTY(i) + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    End If
                
                End If
                
                
                com = BtOpGetNext
    
            Loop
    
    
            If Trim(svSoko) <> "" Then
            
                For i = 0 To UBound(GENSANKOKU)
        
        
                    If Head_Print_Proc(Print_Now, Lcnt) Then
                        Exit Function
                    End If
    
                    Lcnt = Lcnt + 1
                    
                    
                    Select Case Lcnt
                    
                    
                    
                        Case 1
                            Maru_Suuji = "①"
                        Case 2
                            Maru_Suuji = "②"
                        Case 3
                            Maru_Suuji = "③"
                        Case 4
                            Maru_Suuji = "④"
                        Case 5
                            Maru_Suuji = "⑤"
                        Case 6
                            Maru_Suuji = "⑥"
                        Case 7
                            Maru_Suuji = "⑦"
                        Case 8
                            Maru_Suuji = "⑧"
                        Case 9
                            Maru_Suuji = "⑨"
                        Case 10
                            Maru_Suuji = "⑩"
                        Case 11
                            Maru_Suuji = "⑪"
                        Case 12
                            Maru_Suuji = "⑫"
                        Case 13
                            Maru_Suuji = "⑬"
                        Case 14
                            Maru_Suuji = "⑭"
                        Case 15
                            Maru_Suuji = "⑮"
                        Case 16
                            Maru_Suuji = "⑯"
                        Case 17
                            Maru_Suuji = "⑰"
                        Case 18
                            Maru_Suuji = "⑱"
                        Case 19
                            Maru_Suuji = "⑲"
                        Case 20
                            Maru_Suuji = "⑳"
                        Case Else
                            Maru_Suuji = "～"
                    End Select
                    
                    '------------------------------------   17行目
                    Set Printer.Font = NormalFont
                    Printer.Print Tab(MGN_L + 13);
                    Printer.Print "原産国名：";
                    Printer.Print Tab(MGN_L + 56);      '60-->56
                    Printer.Print "数量" & Maru_Suuji & "(商品化済 - 未商品)";
                    Printer.Print Tab(MGN_L + 85);
                    Printer.Print "棚番" & Maru_Suuji;
                    Printer.Print Tab(MGN_L + 105);
                    Printer.Print "現品票";
                    Set Printer.Font = SmallFont
                    Printer.Print Tab(MGN_L + 141);
                    Printer.Print "ﾗﾍﾞﾙ発行";
                    Printer.Print Tab(MGN_L + 151);
                    Printer.Print "専用ﾗﾍﾞﾙ";
                    Printer.Print Tab(MGN_L + 161);
                    Set Printer.Font = NormalFont
                    Printer.Print "　　備考"
                    '------------------------------------   18行目
                    Set Printer.Font = NormalFont
                    Printer.Print Tab(MGN_L + 5);
                    
                    If Lcnt = 1 Then
                        Printer.Print "在庫品";
                    Else
                        Printer.Print "　　　";
                    End If
                    Printer.Print Tab(MGN_L + 13);
                    Set Printer.Font = NormalBoldFont
                    Printer.Print StrConv(Trim(GENSANKOKU(i)), vbWide);
                    Printer.Print Tab(MGN_L + 49);      '53-->49
                    Printer.Print StrConv(Space(8 - Len(Format(SUMI_QTY(i) + MI_QTY(i), "#,0"))) & _
                                    Format(SUMI_QTY(i) + MI_QTY(i), "#,0") & "(" & _
                                    Format(SUMI_QTY(i), "#,0") & "-" & _
                                    Format(MI_QTY(i), "#,0") & ")", vbWide);
                    Set Printer.Font = NormalFont
                    Printer.Print Tab(MGN_L + 85);
                    Printer.Print svSoko & "-" & svRetu & "-" & svRen & "-" & svDan;
                    Printer.Print Tab(MGN_L + 107);
                    Printer.Print "－";
                    Set Printer.Font = SmallFont
                    Printer.Print Tab(MGN_L + 143);     '114-->143
                    Set Printer.Font = NormalFont
                    Printer.Print "□";
                    Set Printer.Font = SmallFont
                    Printer.Print Tab(MGN_L + 153);     '121-->153
                    Set Printer.Font = NormalFont
                    Printer.Print "□"
                    '------------------------------------   19行目
                    Set Printer.Font = NormalFont
                    Printer.Print Tab(MGN_L + 5);
                    Printer.Print String(70, "─")
    
    
                Next i
    
    
    
    
    
            End If
        
        
        

        
        End If
        







        If Trim(svSoko) <> "" Then
            Lcnt = 7
        
        Else
        
        End If


        If StrConv(Y_NYUREC.LIST_NYU_KANRI_F, vbUnicode) = "0" Then
            Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "9")
        
            If Trim(svSoko) = "" Then
                If Head_Print_Proc(Print_Now, Lcnt) Then
                    Exit Function
                End If
            
                Lcnt = 7
        
        
            End If
        
        End If
        
        If StrConv(Y_NYUREC.LIST_NYU_KANRI_F, vbUnicode) = "9" And StrConv(Y_NYUREC.LIST_NYU_CHECK_F, vbUnicode) = "9" Then
            Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "9")
            
            
                        
            
            Read_next = BtOpGetGreater
        
        
        Else
            Read_next = BtOpGetNext
        End If
        
        
        
        sts = BTRV(BtOpUpdate, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
        If sts <> BtNoErr Then
        
            Call File_Error(sts, BtOpUpdate, "入荷予定データ", 0)
            Exit Function

        End If


        If Read_next = BtOpGetGreater Then

            Call UniCode_Conv(K4_Y_NYU.LIST_OUT_END_F, "0")
            Call UniCode_Conv(K4_Y_NYU.JGYOBU, "")
            Call UniCode_Conv(K4_Y_NYU.NAIGAI, "")
            Call UniCode_Conv(K4_Y_NYU.HIN_NO, "")
    
        End If
        com = Read_next
    Loop

    Printer.EndDoc

    Print_Proc = False

End Function


Private Function Head_Print_Proc(Print_Now As String, Lcnt As Integer) As Integer
'----------------------------------------------------------------------------
'                   ヘッダー処理
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

    
    Head_Print_Proc = True


    If Lcnt < 6 Then
        Head_Print_Proc = False
        Exit Function
    End If

    If Lcnt = 99 Then
    Else
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    
    '------------------------------------   1行目
    Printer.Print Tab(MGN_L + 55);
    Set Printer.Font = LargeFont

    Printer.Print "複数原産国部品入庫管理リスト";

    Printer.Print Tab(MGN_L + 100);
    Set Printer.Font = SmallFont
    Printer.Print Print_Now

    '------------------------------------   2行目
    Printer.Print
    '------------------------------------   3行目
    Printer.Print
    
    
    '------------------------------------   6行目
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "┌───────┬───────┐";
    Printer.Print
    
    
    
    '------------------------------------   6行目
    Set Printer.Font = NormalFont
    Printer.Print Tab(MGN_L + 7);
    Printer.Print "品番";
    Printer.Print Tab(MGN_L + 39);
    Printer.Print "品名";
    Printer.Print Tab(MGN_L + 90);
    Printer.Print "月平均出荷数";
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "│　承　認　印　│　作業完了印　│";
    Printer.Print
    '------------------------------------   6行目
    
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "├───────┼───────┤";
    Printer.Print
    '------------------------------------   7行目
    Set Printer.Font = LargeUnderFont
    Printer.Print Tab(MGN_L + 5);
    Printer.Print Trim(StrConv(StrConv(Y_NYUREC.HIN_NO, vbUnicode), vbWide));
    
    Set Printer.Font = LargeFont
    Printer.Print Tab(MGN_L + 30);
    Printer.Print StrConv(Y_NYUREC.HIN_NAME, vbUnicode);
    Set Printer.Font = LargeFont
    Printer.Print Tab(MGN_L + 70);
    
    
    Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(Y_NYUREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(Y_NYUREC.HIN_NO, vbUnicode))
    
    sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    
    
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
            Call UniCode_Conv(AVE_SYUKAREC.AVE_SYUKA, "0")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "月平均出荷数", 0)
            Exit Function

    End Select
    Printer.Print Space(12 - Len(Format(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode), "#,##0"))) & _
                        Format(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode), "#,##0");
    
    
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "│　　　　　　　│　　　　　　　│";
    Printer.Print
    '------------------------------------   8行目
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "│　　　　　　　│　　　　　　　│";
    Printer.Print
    '------------------------------------   9行目
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "│　　　　　　　│　　　　　　　│";
    Printer.Print
    '------------------------------------   10行目
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "└───────┴───────┘";
    Printer.Print
    '------------------------------------   11行目
    Printer.Print
    '------------------------------------   12行目
    Printer.Print
    '------------------------------------   13行目
    Set Printer.Font = NormalFont
    Printer.Print Tab(MGN_L + 13);
    Printer.Print "原産国名：";
    Printer.Print Tab(MGN_L + 56);      '60-->56
    Printer.Print "数量：";
    Printer.Print Tab(MGN_L + 85);
    Printer.Print "棚番：";
    Printer.Print Tab(MGN_L + 105);
    Printer.Print "現品票";
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 141);
    Printer.Print "ﾗﾍﾞﾙ発行";
    Printer.Print Tab(MGN_L + 151);
    Printer.Print "専用ﾗﾍﾞﾙ";
    Printer.Print Tab(MGN_L + 161);
    Set Printer.Font = NormalFont
    Printer.Print "入庫棚番"
    '------------------------------------   14行目
'    Printer.Print
    '------------------------------------   15行目
    Set Printer.Font = NormalBoldFont
    Printer.Print Tab(MGN_L + 5);
    Printer.Print "入荷品";
    Printer.Print Tab(MGN_L + 13);
    Set Printer.Font = NormalBoldFont
    Printer.Print StrConv(Trim(StrConv(Y_NYUREC.GENSANKOKU, vbUnicode)), vbWide);           'trim
    Printer.Print Tab(MGN_L + 49);      '53-->49
    Printer.Print StrConv(Space(8 - Len(Format(StrConv(Y_NYUREC.SURYO, vbUnicode), "#,##0"))) & _
                        Format(StrConv(Y_NYUREC.SURYO, vbUnicode), "#,##0"), vbWide);
    Set Printer.Font = NormalFont
    Printer.Print Tab(MGN_L + 85);
    Printer.Print Mid(StrConv(Y_NYUREC.NYUKO_TANABAN, vbUnicode), 1, 2) & "-" & _
                    Mid(StrConv(Y_NYUREC.NYUKO_TANABAN, vbUnicode), 3, 2) & "-" & _
                    Mid(StrConv(Y_NYUREC.NYUKO_TANABAN, vbUnicode), 5, 2) & "-" & _
                    Mid(StrConv(Y_NYUREC.NYUKO_TANABAN, vbUnicode), 7, 2);
    Printer.Print Tab(MGN_L + 107);
    Printer.Print "□";
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 143);     '114-->143
    Set Printer.Font = NormalFont
    Printer.Print "□";
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 153);     '121-->153
    Set Printer.Font = NormalFont
    Printer.Print "□"
    '------------------------------------   16行目
    Set Printer.Font = NormalFont
    Printer.Print Tab(MGN_L + 5);
    Printer.Print String(70, "─")
    '------------------------------------   18行目


    Lcnt = 0

    Head_Print_Proc = False

End Function
