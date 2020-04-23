Attribute VB_Name = "CompProc"
'*======================================================
' 2015.01.10 mdlProcより　移行
'*======================================================




Public Sub COMPO_Check_Cnt_Proc(KO_CNT As Integer, RD_CNT As Integer)
'-------------------------------------------------------
'
'   『指図票データの構成部品数』
'
'       2011.03.02
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim ans         As Integer

Dim i           As Integer

    
    KO_CNT = UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL) + 1
    RD_CNT = 0
    
    For i = 0 To UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL)
    
        If ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(i).RD_FLAG Then
        
        
            RD_CNT = RD_CNT + 1
        
        End If
    
    Next i
            


End Sub



Public Function COMPO_Check_Update_Proc(KO_CNT As Integer, RD_CNT As Integer, Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『指図票データの構成部品数更新』
'
'       2012.04.20
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim ans         As Integer

Dim i           As Integer

Dim DateTime    As String * 12

    
    COMPO_Check_Update_Proc = True
    
    KO_CNT = UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL) + 1
    RD_CNT = 0
    
    For i = 0 To UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL)
    
        If ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(i).RD_FLAG Then
        
        
            RD_CNT = RD_CNT + 1
        
        End If
    
    Next i
            


    '----------------------------------- データ更新処理開始 -----------
                                    'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Exit Function
    End If

    
    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, ID_KANRI_TBL(ING_No).SHIJI_No)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")

    DateTime = Format(Now, "YYYYMMDDHHSSMM")
    
    com = BtOpGetGreaterEqual


    Do
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
    
        sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
                            
                If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> ID_KANRI_TBL(ING_No).SHIJI_No Then
                    Exit Do
                End If
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "指図票ﾃﾞｰﾀ(子)", 0)
                COMPO_Check_Update_Proc = SYS_ERR
                GoTo Abort_Tran
        End Select
        
        '担当者
        Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_TANTO, ID_KANRI_TBL(ING_No).TANTO_CODE)
        '日時
        Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_YMDHS, DateTime)
        'ﾁｪｯｸ済み数
        Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_Sumi_Cnt, Format(RD_CNT, "00"))
        '構成数
        Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_ALL_Cnt, Format(KO_CNT, "00"))
        
        sts = BTRV(BtOpUpdate, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpUpdate, "指図票ﾃﾞｰﾀ(子)", 0)
                COMPO_Check_Update_Proc = SYS_ERR
                GoTo Abort_Tran
        End Select
    
        com = BtOpGetNext
    
    Loop
                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        GoTo Abort_Tran
    End If


    COMPO_Check_Update_Proc = False
    
    Exit Function
Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function





Public Function COMPO_Check_Read_Proc(SHIJI_No As String, KO_CNT As Integer) As Integer
'-------------------------------------------------------
'
'   『指図票データの構成部品数』
'
'       2011.03.02
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer


Dim i           As Integer

            
            
Dim Check_f     As Boolean  '2011.04.18
Dim j           As Integer  '2011.04.18
            
            
            
    COMPO_Check_Read_Proc = True



    
        
    
    
    
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, SHIJI_No)
    
    
    sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            COMPO_Check_Read_Proc = BtErrKeyNotFound
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
            COMPO_Check_Read_Proc = SYS_ERR
            Exit Function
    End Select
    
    
    Erase ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL
    KO_CNT = -1
    
    
    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, SHIJI_No)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
    
    com = BtOpGetGreater
            
            
    Do
    
    
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
    
        sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "指図票ﾃﾞｰﾀ(子)", 0)
                COMPO_Check_Read_Proc = SYS_ERR
                Exit Function
        End Select
    
    
    
        If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> SHIJI_No Then
            Exit Do
        End If
    
    
        If StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode) = "3" Then
    
    
''''''''''''    2011.04.18
            Check_f = False
            If Kousei_check_F Then
                If Kousei_check_Tb(0) = "*" Then
                    Check_f = True
                Else
                
                
                    For j = 0 To UBound(Kousei_check_Tb)
                        If Kousei_check_Tb(j) = Trim(StrConv(P_SSHIJI_K_REC.KO_SYUBETSU, vbUnicode)) Then
                            Check_f = True
                            Exit For
                        End If
                    
                    Next j
                
                
                End If
            End If
            If Check_f Then
''''''''''''    2011.04.18
    
                If KO_CNT = -1 Then
                    KO_CNT = KO_CNT + 1
            
                    ReDim Preserve ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(0 To KO_CNT)
                        
                    ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).KO_JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
                    ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).KO_NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode)
                    ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).KO_HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
            
            
            
                Else
            
            
                    For i = 0 To UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL)
                    
                        If Trim(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(i).KO_HIN_GAI) = Trim(StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)) Then
                            Exit For
                        End If
                    
                    Next i
            
                    If i > UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL) Then
        
                        KO_CNT = KO_CNT + 1
                
                        ReDim Preserve ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(0 To KO_CNT)
                            
                        ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).KO_JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
                        ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).KO_NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode)
                        ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).KO_HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
            
                    End If
            
            
                End If

''''''''''''    2011.04.18
            End If
''''''''''''    2011.04.18
        
        
        End If
    Loop
    
    COMPO_Check_Read_Proc = False

End Function

Public Function COMPO_CHECK_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『商品化構成ﾁｪｯｸ処理』
'       2011.03.02
'
'-------------------------------------------------------
Dim sts             As Integer



Dim SHIJI_No        As String * 8

Dim Hinban          As String * 20

Dim i               As Integer
Dim j               As Integer
Dim k               As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1



Dim MENU_NO         As String * 2

Dim KO_CNT          As Integer
Dim RD_CNT          As Integer



    COMPO_CHECK_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（指図票№）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_SHIJI_NO      '指図票№
    
    
                        SHIJI_No = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        
                        
                        
                        
                        sts = COMPO_Check_Read_Proc(SHIJI_No, KO_CNT)
                        
                        
                        Select Case sts
                            Case False          '正常
                                
                            
                            
                            Case BtErrKeyNotFound
                                    
                                
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "該当データなし", "", "")
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                COMPO_CHECK_PROC = False
                                Exit Function
                            
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        
                        End Select
                
                        
                        '子部品なし
                        If KO_CNT < 0 Then
                        
                        
                        
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "子部品なし", "", "")
                            
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            COMPO_CHECK_PROC = False
                            Exit Function
                        
                        
                        
                        
                        End If
                        
                        
                        'キャンセル済
                        If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = "1" Then
                            
                            
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "キャンセル済", "", "")
                            
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            COMPO_CHECK_PROC = False
                            Exit Function
                        End If
                                                
                        '完了済
                        If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = "1" Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "完了済", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            COMPO_CHECK_PROC = False
                            Exit Function
                        End If
                        
                        
                        
                        
                        
                        ID_KANRI_TBL(ING_No).SHIJI_No = SHIJI_No
                        ID_KANRI_TBL(ING_No).Hinban = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)
                        ID_KANRI_TBL(ING_No).Input_Line = 0
                                                
                        
                        
                        
                        
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        
                        
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.FileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                        Call COMPO_Check_Cnt_Proc(KO_CNT, RD_CNT)
                                                                                
                                                                                
                                                                                
                                                                                
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SHIJI_No & "     " & Format(RD_CNT, "##0") & "/" & Format(KO_CNT, "#"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SHIJI_No & "     " & Format(RD_CNT, "##0") & "/" & Format(KO_CNT, "#"))
                                                                                '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                                                                                'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Hinban)
                                                                                
                                                                                
                                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '入力桁数
                                                                                
                        Send_Text.Box_Type(2).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "20"
                                                                                
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = ""                    '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '入力桁数
                         Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '２回目の受信（ラベル品番）
            For i = 0 To M_Gyo - 1
            
                Select Case i
                
                                    
                    Case 0
                                    
                                    
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    
                        If Trim(Hinban) = Ent_Para Then
                        
                            '2011.06.14
                            For k = 0 To UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL)
                                If ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).RD_FLAG <> 1 Then
                                    Exit For
                                End If
                            Next k
                            
                            If k > UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL) Then
                            Else
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & ID_KANRI_TBL(ING_No).SHIJI_No, "ﾁｪｯｸ未完了", "", "")
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                COMPO_CHECK_PROC = False
                                Exit Function
                            
                            End If
                            
                            
                            
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            
                        
                        
                        
                            '-----------------------------------------------ヘッダー
                            Send_Text.sts = Sts_OK                                  'ステータス　OK
                            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                            Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                    
                            Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                    
                            Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                    
                            Send_Text.FileName = ""                                 '送信データファイル名
                            ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                    
                            Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                    
                            '-----------------------------------------------１行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '表示内容
                                                                                    
                                                                                    '数値初期表示
                            Send_Text.Box_Type(0).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(0).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                    '入力桁数
                            Send_Text.Box_Type(0).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                            '-----------------------------------------------２行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(1).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "指図票№" & ID_KANRI_TBL(ING_No).SHIJI_No)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "指図票№" & ID_KANRI_TBL(ING_No).SHIJI_No)
                                                                                    '数値初期表示
                            Send_Text.Box_Type(1).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(1).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                    '入力桁数
                            Send_Text.Box_Type(1).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                            '-----------------------------------------------３行目
                                                                                            'BOX属性
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "")
                                                                                    '数値初期表示
                            Send_Text.Box_Type(2).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(2).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                                                    '入力桁数
                                                                                    
                            Send_Text.Box_Type(2).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                                                                                    
                                                                                    
                                                                                    
                            Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                                    '-----------------------------------------------４行目
                            Send_Text.Box_Type(3).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "チェック完了")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "チェック完了")
                
                                                                                    '数値初期表示
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(3).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                                    '入力桁数
                            Send_Text.Box_Type(3).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                                    '-----------------------------------------------４行目
                                                                                            'BOX属性
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                                    '数値初期表示
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(4).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                    '入力桁数
                            Send_Text.Box_Type(4).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                
                            Sendbuf = Text_Create_Proc()
                        
                            COMPO_CHECK_PROC = False
        
                            Exit Function
                            
                            
                            
                            '2011.06.14
                                                    
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                                    
                            Exit For
                        End If
                    Case 2, 3, 4        '品番
                        
                        
                        
                        For j = 2 To 4
                        
                            If j - 2 = ID_KANRI_TBL(ING_No).Input_Line Then
                        
                                                    
                        
                                Hinban = ID_KANRI_TBL(ING_No).Recv_text(j)
                    
                    
                                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                                Select Case sts
                                    Case BtNoErr
                            
                                    Case BtErrKeyNotFound
                                
                                    Case Else
                                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                        Exit Function
                        
                                End Select
                    
                    
                    
                    
                                For k = 0 To UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL)
                                
                                    If Trim(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).KO_HIN_GAI) = Trim(Hinban) Then
                                        Exit For
                                    End If
                                
                                Next k
                    
                    
                                If k > UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL) Then
                                
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & ID_KANRI_TBL(ING_No).SHIJI_No, Hinban, "品番ｴﾗｰ", "")
                                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    COMPO_CHECK_PROC = False
                                    Exit Function
                                
                                
                                End If
                                                            
                                                            
                        
                                ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).RD_FLAG = 1
                            
                            
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.04.20
                                If COMPO_Check_Update_Proc(KO_CNT, RD_CNT, Sendbuf) Then
                                    Exit Function
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.04.20
                            
                            
                            
                            
                            
                                Exit For
                           End If
                        
                        
                        
                        Next j
                        
                        ID_KANRI_TBL(ING_No).RD_HINBAN(ID_KANRI_TBL(ING_No).Input_Line) = Hinban
                        ID_KANRI_TBL(ING_No).Input_Line = ID_KANRI_TBL(ING_No).Input_Line + 1
                        
                        If ID_KANRI_TBL(ING_No).Input_Line > 2 Then
                            ID_KANRI_TBL(ING_No).Input_Line = 0
                        
                            ID_KANRI_TBL(ING_No).RD_HINBAN(0) = ""
                            ID_KANRI_TBL(ING_No).RD_HINBAN(1) = ""
                            ID_KANRI_TBL(ING_No).RD_HINBAN(2) = ""
                        
                        End If
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                            
                            
                            
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.FileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                            
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                                                                                
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                        Call COMPO_Check_Cnt_Proc(KO_CNT, RD_CNT)
                                                                                
                                                                                
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SHIJI_No & "     " & Format(RD_CNT, "##0") & "/" & Format(KO_CNT, "#"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SHIJI_No & "     " & Format(RD_CNT, "##0") & "/" & Format(KO_CNT, "#"))
                                                                                '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                        For j = 0 To 2
                        
                            If j = ID_KANRI_TBL(ING_No).Input_Line Then
                                                                                    'BOX属性
                                Send_Text.Box_Type(j + 2).Box_Type = TYPE_BCANK
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Box_Type = TYPE_BCANK
                                                                                '表示内容
                                Call UniCode_Conv(Send_Text.Box_Type(j + 2).LCD, LCD_Hinban)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).LCD, LCD_Hinban)
                        
                        
                        
                                                                                        '数値初期表示
                                Send_Text.Box_Type(j + 2).INIT = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).INIT = ""
                                                                                        '初期カーソル位置
                                Send_Text.Box_Type(j + 2).Start_Pos = "01"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Start_Pos = "01"
                                                                                        '入力桁数
                                                                                        
                                Send_Text.Box_Type(j + 2).Max_Size = "20"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Max_Size = "20"
                                                                                        
                                                                                        
                                Send_Text.Box_Type(j + 2).MENU = ""                     'メニュ―番号
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).MENU = ""
                        
                        
                        
                        
                            Else
                                                                                    'BOX属性
                                Send_Text.Box_Type(j + 2).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Box_Type = TYPE_REF
                                Call UniCode_Conv(Send_Text.Box_Type(j + 2).LCD, ID_KANRI_TBL(ING_No).RD_HINBAN(j))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).LCD, ID_KANRI_TBL(ING_No).RD_HINBAN(j))
                        
                        
                            
                            
                                                                                        '数値初期表示
                                Send_Text.Box_Type(j + 2).INIT = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).INIT = ""
                                                                                        '初期カーソル位置
                                Send_Text.Box_Type(j + 2).Start_Pos = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Start_Pos = ""
                                                                                        '入力桁数
                                                                                        
                                Send_Text.Box_Type(j + 2).Max_Size = "00"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Max_Size = "00"
                                                                                        
                                                                                        
                                Send_Text.Box_Type(j + 2).MENU = ""                       'メニュ―番号
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).MENU = ""
                            
                            
                            End If
                        
                        Next j
                                                                                    
                                                                                    
                                                                                    
                                                                                    
                                                                                    
                                                                                    
    
                        Sendbuf = Text_Create_Proc()
                                    
                        Exit For
                           
                End Select
            
            Next i
        Case Step_Sagyo3_RES        '３回目の受信（ＥＮＴ）
        
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
            
            If Trim(MENU_NO) = "" Then
            Else
            '作業ﾛｸﾞ出力
                
                If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                    (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                                    MENU_NO, _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                     ID_KANRI_TBL(ING_No).Hinban, , , , , , , , , _
                                                     ID_KANRI_TBL(ING_No).SHIJI_No) Then
                    COMPO_CHECK_PROC = SYS_ERR
                    Exit Function
                End If
            End If
            '次の作業要求
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                    Exit Function
            End Select

            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
        
            Sendbuf = Text_Create_Proc()
        
        
    End Select
        

    COMPO_CHECK_PROC = False
    


End Function


Public Function COMPO_OSAKA_CHECK_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『大阪ＰＣ　部材ｾﾝﾀｰ商品化構成ﾁｪｯｸ処理』
'       2012.03.16
'
'-------------------------------------------------------
Dim sts             As Integer



Dim SHIJI_No        As String * 8

Dim Hinban          As String * 20

Dim i               As Integer
Dim j               As Integer
Dim k               As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1



Dim MENU_NO         As String * 2

Dim KO_CNT          As Integer
Dim RD_CNT          As Integer


Dim Found_F         As Boolean
Dim wkJgoybu        As String * 1

    COMPO_OSAKA_CHECK_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（指図票№）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_SHIJI_NO      '指図票№
    
    
                        SHIJI_No = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        
                        
                        
                        
                        sts = COMPO_OSAKA_Check_Read_Proc(SHIJI_No, KO_CNT)
                        
                        
                        Select Case sts
                            Case False          '正常
                                
                            
                            
                            Case BtErrKeyNotFound
                                    
                                
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "該当データなし", "", "")
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                COMPO_OSAKA_CHECK_PROC = False
                                Exit Function
                            
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        
                        End Select
                
                        
                        '子部品なし
                        If KO_CNT < 0 Then
                        
                        
                        
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "子部品なし", "", "")
                            
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            COMPO_OSAKA_CHECK_PROC = False
                            Exit Function
                        
                        
                        
                        
                        End If
                        
                        
                        'キャンセル済
                        If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = "1" Then
                            
                            
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "キャンセル済", "", "")
                            
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            COMPO_OSAKA_CHECK_PROC = False
                            Exit Function
                        End If
                                                
                        '完了済
                        If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = "1" Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "完了済", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            COMPO_OSAKA_CHECK_PROC = False
                            Exit Function
                        End If
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 子部品チェック 20012.04.13
                        If StrConv(P_SSHIJI_O_REC.COMPO_END_F, vbUnicode) = "9" Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "子部品ﾁｪｯｸ", "完了済", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            COMPO_OSAKA_CHECK_PROC = False
                            Exit Function
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 子部品チェック 20012.04.13
                        
                        
                        
                        
                        
                        ID_KANRI_TBL(ING_No).SHIJI_No = SHIJI_No
                        ID_KANRI_TBL(ING_No).Hinban = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)
                        ID_KANRI_TBL(ING_No).Input_Line = 0
                                                
                        
                        
                        
                        
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        
                        
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.FileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                        Call COMPO_Check_Cnt_Proc(KO_CNT, RD_CNT)
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SHIJI_No & "     " & Format(RD_CNT, "##0") & "/" & Format(KO_CNT, "#"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SHIJI_No & "     " & Format(RD_CNT, "##0") & "/" & Format(KO_CNT, "#"))
                                                                                '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                                                                                'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Hinban)
                                                                                
                                                                                
                                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '入力桁数
                                                                                
                        Send_Text.Box_Type(2).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "20"
                                                                                
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = ""                    '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '入力桁数
                         Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '２回目の受信（ラベル品番）
            For i = 0 To M_Gyo - 1
            
                Select Case i
                
                                    
                    Case 0
                                    
                                    
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    
                        If Trim(Hinban) = Ent_Para Then
                        
                            '2011.06.14
                            For k = 0 To UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL)
                                If ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).RD_FLAG <> 1 Then
                                    Exit For
                                End If
                            Next k
                            
                            If k > UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL) Then
                            Else
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & ID_KANRI_TBL(ING_No).SHIJI_No, "ﾁｪｯｸ未完了", "", "")
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                COMPO_OSAKA_CHECK_PROC = False
                                Exit Function
                            
                            End If
                            
                            
                            
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            
                        
                        
                        
                            '-----------------------------------------------ヘッダー
                            Send_Text.sts = Sts_OK                                  'ステータス　OK
                            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                            Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                    
                            Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                    
                            Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                    
                            Send_Text.FileName = ""                                 '送信データファイル名
                            ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                    
                            Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                    
                            '-----------------------------------------------１行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '表示内容
                                                                                    
                                                                                    '数値初期表示
                            Send_Text.Box_Type(0).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(0).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                    '入力桁数
                            Send_Text.Box_Type(0).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                            '-----------------------------------------------２行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(1).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "指図票№" & ID_KANRI_TBL(ING_No).SHIJI_No)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "指図票№" & ID_KANRI_TBL(ING_No).SHIJI_No)
                                                                                    '数値初期表示
                            Send_Text.Box_Type(1).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(1).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                    '入力桁数
                            Send_Text.Box_Type(1).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                            '-----------------------------------------------３行目
                                                                                            'BOX属性
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "")
                                                                                    '数値初期表示
                            Send_Text.Box_Type(2).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(2).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                                                    '入力桁数
                                                                                    
                            Send_Text.Box_Type(2).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                                                                                    
                                                                                    
                                                                                    
                            Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                                    '-----------------------------------------------４行目
                            Send_Text.Box_Type(3).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "チェック完了")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "チェック完了")
                
                                                                                    '数値初期表示
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(3).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                                    '入力桁数
                            Send_Text.Box_Type(3).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                                    '-----------------------------------------------４行目
                                                                                            'BOX属性
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                                    '数値初期表示
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(4).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                    '入力桁数
                            Send_Text.Box_Type(4).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                
                            Sendbuf = Text_Create_Proc()
                        
                            COMPO_OSAKA_CHECK_PROC = False
        
                            Exit Function
                            
                            
                            
                            '2011.06.14
                            Exit For
                        End If
                    Case 2, 3, 4        '品番
                        
                        
                        
                        For j = 2 To 4
                        
                            If j - 2 = ID_KANRI_TBL(ING_No).Input_Line Then
                        
                                                    
                        
                                Hinban = ID_KANRI_TBL(ING_No).Recv_text(j)
                    
                    
                                For k = 0 To UBound(JGYOBU_T)
                                
                                
                                    If JGYOBU_T(k).CODE = SHIZAI Then
                                        sts = Item_Read_Proc(JGYOBU_T(k).CODE, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI, , BUZAI)
                                    Else
                                        sts = Item_Read_Proc(JGYOBU_T(k).CODE, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                                            Select Case sts
                                            Case BtNoErr
                                                Exit For
                                            Case BtErrKeyNotFound
                                        
                                            Case Else
                                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                                Sendbuf = Text_Create_Proc()
                                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                                Exit Function
                                
                                        End Select
                                    End If
                                Next k
                    
                    
                                Found_F = False
                                For k = 0 To UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL)
                                
                                    If Trim(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).KO_HIN_GAI) = Trim(Hinban) Then
                                    
                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.04.13
                                        If ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).RD_FLAG = 1 Then
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & ID_KANRI_TBL(ING_No).SHIJI_No, Hinban, "チェック済", "")
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            COMPO_OSAKA_CHECK_PROC = False
                                            Exit Function
                                        End If
                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.04.13
                                    
                                        Found_F = True
                                        If ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).RD_FLAG = 0 Then
                                            ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).RD_FLAG = 1
                                            Exit For
                                        End If
                                    
                                    End If
                                
                                
                                Next k
                    
                    
                                If Not Found_F Then
                                
                                
                                
                                    For k = 0 To UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL)
                                
                                        '品番読み替えでチェック
                                        
                                        If ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).KO_JGYOBU = SHIZAI Then
                                            Call UniCode_Conv(K1_FURIKAE.JGYOBU_GO, BUZAI)
                                            Call UniCode_Conv(K1_FURIKAE.NAIGAI_GO, NAIGAI_NAI)
                                            Call UniCode_Conv(K1_FURIKAE.HIN_GO, Hinban)
                                        
                                            Call UniCode_Conv(K1_FURIKAE.JGYOBU_MAE, BUZAI)
                                            Call UniCode_Conv(K1_FURIKAE.NAIGAI_MAE, NAIGAI_NAI)
                                            Call UniCode_Conv(K1_FURIKAE.HIN_MAE, ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).KO_HIN_GAI)
                                        
                                    
                                            sts = BTRV(BtOpGetEqual, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K1_FURIKAE, Len(K1_FURIKAE), 1)
                                            Select Case sts
                                                Case BtNoErr
                                                    
                                                                                                        
                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.04.13
                                                    If ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).RD_FLAG = 1 Then
                                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & ID_KANRI_TBL(ING_No).SHIJI_No, Hinban, "チェック済", "")
                                                        Sendbuf = Text_Create_Proc()
                                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                                        COMPO_OSAKA_CHECK_PROC = False
                                                        Exit Function
                                                    End If
                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.04.13
                                                    
                                                    
                                                    ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(k).RD_FLAG = 1
                                                    Exit For
                                                    
                                                Case BtErrKeyNotFound
                                                
'                                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & ID_KANRI_TBL(ING_No).SHIJI_No, Hinban, "品番ｴﾗｰ", "")
'
'                                                    Sendbuf = Text_Create_Proc()
'                                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                                    COMPO_OSAKA_CHECK_PROC = False
'                                                    Exit Function
                                                
                                                Case Else
                                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                                    Sendbuf = Text_Create_Proc()
                                                    Call File_Error(sts, BtOpGetEqual, "品番振替マスタ", 0)
                                                    Exit Function
                                            End Select
                                    
                                    
                                    
                                        End If
                                    Next k
                                
                                    If k > UBound(ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL) Then
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & ID_KANRI_TBL(ING_No).SHIJI_No, Hinban, "品番ｴﾗｰ", "")
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        COMPO_OSAKA_CHECK_PROC = False
                                        Exit Function
                                    End If
                                        
                                                            
                                End If
                        
                            
                           End If
                        
                        
                        
                        Next j
                        
                        ID_KANRI_TBL(ING_No).RD_HINBAN(ID_KANRI_TBL(ING_No).Input_Line) = Hinban
                        ID_KANRI_TBL(ING_No).Input_Line = ID_KANRI_TBL(ING_No).Input_Line + 1
                        
                        If ID_KANRI_TBL(ING_No).Input_Line > 2 Then
                            ID_KANRI_TBL(ING_No).Input_Line = 0
                        
                            ID_KANRI_TBL(ING_No).RD_HINBAN(0) = ""
                            ID_KANRI_TBL(ING_No).RD_HINBAN(1) = ""
                            ID_KANRI_TBL(ING_No).RD_HINBAN(2) = ""
                        
                        End If
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                            
                            
                            
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.FileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                            
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                                                                                
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                        Call COMPO_Check_Cnt_Proc(KO_CNT, RD_CNT)
                                                                                
                                                                                
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SHIJI_No & "     " & Format(RD_CNT, "##0") & "/" & Format(KO_CNT, "#"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SHIJI_No & "     " & Format(RD_CNT, "##0") & "/" & Format(KO_CNT, "#"))
                                                                                '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                        For j = 0 To 2
                        
                            If j = ID_KANRI_TBL(ING_No).Input_Line Then
                                                                                    'BOX属性
                                Send_Text.Box_Type(j + 2).Box_Type = TYPE_BCANK
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Box_Type = TYPE_BCANK
                                                                                '表示内容
                                Call UniCode_Conv(Send_Text.Box_Type(j + 2).LCD, LCD_Hinban)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).LCD, LCD_Hinban)
                        
                        
                        
                                                                                        '数値初期表示
                                Send_Text.Box_Type(j + 2).INIT = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).INIT = ""
                                                                                        '初期カーソル位置
                                Send_Text.Box_Type(j + 2).Start_Pos = "01"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Start_Pos = "01"
                                                                                        '入力桁数
                                                                                        
                                Send_Text.Box_Type(j + 2).Max_Size = "20"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Max_Size = "20"
                                                                                        
                                                                                        
                                Send_Text.Box_Type(j + 2).MENU = ""                     'メニュ―番号
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).MENU = ""
                        
                        
                        
                        
                            Else
                                                                                    'BOX属性
                                Send_Text.Box_Type(j + 2).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Box_Type = TYPE_REF
                                Call UniCode_Conv(Send_Text.Box_Type(j + 2).LCD, ID_KANRI_TBL(ING_No).RD_HINBAN(j))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).LCD, ID_KANRI_TBL(ING_No).RD_HINBAN(j))
                        
                        
                            
                            
                                                                                        '数値初期表示
                                Send_Text.Box_Type(j + 2).INIT = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).INIT = ""
                                                                                        '初期カーソル位置
                                Send_Text.Box_Type(j + 2).Start_Pos = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Start_Pos = ""
                                                                                        '入力桁数
                                                                                        
                                Send_Text.Box_Type(j + 2).Max_Size = "00"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).Max_Size = "00"
                                                                                        
                                                                                        
                                Send_Text.Box_Type(j + 2).MENU = ""                       'メニュ―番号
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j + 2).MENU = ""
                            
                            
                            End If
                        
                        Next j
                                                                                                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   全品目チェック済みは　ブザー音を替える  2012.04.13
                        If (KO_CNT = RD_CNT) Then
                            Send_Text.Buzzer = Buzzer_DOUBLE
                            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DOUBLE
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   全品目チェック済みは　ブザー音を替える  2012.04.13
                                                                                    
                                                                                    
                                                                                    
    
                        Sendbuf = Text_Create_Proc()
                                    
                        Exit For
                           
                End Select
            
            Next i
        Case Step_Sagyo3_RES        '３回目の受信（ＥＮＴ）
        
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
            
            If Trim(MENU_NO) = "" Then
            Else
            '作業ﾛｸﾞ出力
                
                If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                    (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                                    MENU_NO, _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                     ID_KANRI_TBL(ING_No).Hinban, , , , , , , , , _
                                                     ID_KANRI_TBL(ING_No).SHIJI_No) Then
                    COMPO_OSAKA_CHECK_PROC = SYS_ERR
                    Exit Function
                End If
            End If
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 指図表(親)更新  2012.04.13
            Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, ID_KANRI_TBL(ING_No).SHIJI_No)
            sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case BtErrKeyNotFound
                    Call Err_Send_Proc("未登録異常", "指図表(親)", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    COMPO_OSAKA_CHECK_PROC = False
                    Exit Function
                Case Else
                '重要な要因なので未登録はシステム停止とする
                    Call Err_Send_Proc("システム異常発生", "指図表(親)", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual, "指図表(親)", 0)
                    Exit Function
            End Select
                        
            Call UniCode_Conv(P_SSHIJI_O_REC.COMPO_END_F, "9")
            sts = BTRV(BtOpUpdate, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                    Call Err_Send_Proc("システム異常発生", "指図表(親)", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpUpdate, "指図表(親)", 0)
                    Exit Function
            End Select
                   
            
            
            
            
            
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 指図表(親)更新  2012.04.13
            
            
            
            
            '次の作業要求
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                    Exit Function
            End Select

            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
        
            Sendbuf = Text_Create_Proc()
        
        
    End Select
        

    COMPO_OSAKA_CHECK_PROC = False
    


End Function


Public Function COMPO_OSAKA_Check_Read_Proc(SHIJI_No As String, KO_CNT As Integer) As Integer
'-------------------------------------------------------
'
'   『大阪ＰＣ　部材ｾﾝﾀｰ指図票データの構成部品数』
'
'       2012.03.16
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer


Dim i           As Integer

            
            
Dim Check_f     As Boolean
Dim j           As Integer
            
            
            
    COMPO_OSAKA_Check_Read_Proc = True



    
        
    
    
    
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, SHIJI_No)
    
    
    sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            COMPO_OSAKA_Check_Read_Proc = BtErrKeyNotFound
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "指図票データ(親)", 0)
            COMPO_OSAKA_Check_Read_Proc = SYS_ERR
            Exit Function
    End Select
    
    
    Erase ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL
    KO_CNT = -1
    
    
    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, SHIJI_No)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
    
    com = BtOpGetGreater
            
            
    Do
    
    
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
    
        sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                COMPO_OSAKA_Check_Read_Proc = SYS_ERR
                Exit Function
        End Select
    
    
    
        If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> SHIJI_No Then
            Exit Do
        End If
    
    
        If StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode) = "3" Then
    
    
            Check_f = False
            If Kousei_check_F Then
                If Kousei_check_Tb(0) = "*" Then
                    Check_f = True
                Else
                
                
                    For j = 0 To UBound(Kousei_check_Tb)
                        If Kousei_check_Tb(j) = Trim(StrConv(P_SSHIJI_K_REC.KO_SYUBETSU, vbUnicode)) Then
                            Check_f = True
                            Exit For
                        End If
                    
                    Next j
                
                
                End If
            End If
            If Check_f Then
                KO_CNT = KO_CNT + 1
                
                ReDim Preserve ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(0 To KO_CNT)
                            
                ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).KO_JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
                ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).KO_NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode)
                ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).KO_HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
            
                ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).RD_FLAG = False
            
                ID_KANRI_TBL(ING_No).CHK_HINBAN_TBL(KO_CNT).KO_QTY = Val(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
            
            End If
        
        
        End If
    Loop
    
    COMPO_OSAKA_Check_Read_Proc = False

End Function

Public Function COMPO_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『構成表示処理』    2006.10.10
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

'2010.12.07
'Dim Hinban      As String * 13
Dim Hinban      As String * 20
'2010.12.07

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim FileNo      As Integer

Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag

Dim i           As Integer
Dim j           As Integer

Dim In_cnt      As Integer

    
    COMPO_Proc = True

    If Right(F1100101.CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = F1100101.CtrsWsk1.SendFolder & "\" & BA_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    Else
        FullPath = F1100101.CtrsWsk1.SendFolder & BA_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    End If


    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '品番
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    
                                    
                '最初に出荷予定で
                Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Hinban)
                sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Hinban = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                    Case BtErrKeyNotFound
                    Case Else
                        '重要な要因なので未登録はシステム停止とする
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "出荷予定ﾃﾞｰﾀ", 0)
                        Exit Function
                End Select
                                    
                                    
                                    
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        COMPO_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                
                End Select
                
                For j = 0 To UBound(SHIMUKE_TBL)
                    If ID_KANRI_TBL(ING_No).JGYOBU = SHIMUKE_TBL(j).JGYOBU And _
                        ID_KANRI_TBL(ING_No).NAIGAI = SHIMUKE_TBL(j).NAIGAI Then
                        Exit For
                    End If
                Next j
                If j > UBound(SHIMUKE_TBL) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "仕向け先設定エラー", "", "")
                
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    COMPO_Proc = False
                    Exit Function
                End If
                
                
                '-件数獲得
                
                Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_TBL(j).SHIMUKE_CODE)
                Call UniCode_Conv(K0_P_COMPO.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                Call UniCode_Conv(K0_P_COMPO.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Hinban)
                Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
                Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
                
                com = BtOpGetGreater
                In_cnt = 0
                            
                Do
                    
                    sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    Select Case sts
                        Case BtNoErr
                            '仕向け先／事業部／内外／品番ブレーク？
                            If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_TBL(j).SHIMUKE_CODE Or _
                                StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                                StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Or _
                                Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Or _
                                StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                                
                                Exit Do
                            
                            End If
                        
                        Case BtErrEOF
                            
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "構成マスタ", 0)
                            Exit Function
                    End Select
                    
                    
                                
                    In_cnt = In_cnt + 1
                    
                    com = BtOpGetNext
                Loop
                
                
                
                On Error Resume Next
                Kill (FullPath)             '送信用ファイル削除
                On Error GoTo 0
        
                FileNo = FreeFile           '送信用ファイルＯＰＥＮ
                Open FullPath For Binary As #FileNo
        
                SendFileRec.Title = "0"     'タイトル行
                Call UniCode_Conv(SendFileRec.LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
                    
                SendFileRec.Title = "0"     '品番
                Call UniCode_Conv(SendFileRec.LCD, Hinban & Format(In_cnt, "#0") & "件")
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
                    
                            
                Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_TBL(j).SHIMUKE_CODE)
                Call UniCode_Conv(K0_P_COMPO.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                Call UniCode_Conv(K0_P_COMPO.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Hinban)
                Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
                Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
                
                com = BtOpGetGreater
                
                Do
                    
                    sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    Select Case sts
                        Case BtNoErr
                            '仕向け先／事業部／内外／品番ブレーク？
                            If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_TBL(j).SHIMUKE_CODE Or _
                                StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                                StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Or _
                                Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Or _
                                StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                                
                                Exit Do
                            
                            End If
                        
                        Case BtErrEOF
                            
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "構成マスタ", 0)
                            Exit Function
                    End Select
                    
                    
                                
                    SendFileRec.Title = "1"
                            
                            
                    'コードマスタ読み込み
                    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
                    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))

                    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                            
                            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                        
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "コードマスタ ", 0)
                            Exit Function
                    End Select
                            
                    Call UniCode_Conv(SendFileRec.LCD, Left(StrConv(P_CODEREC.C_RNAME, vbUnicode), 2) & _
                                        ":" & _
                                        Left(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode), 13) & _
                                        Format(CLng(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0"))
                    SendFileRec.CRLF = vbCrLf
                    Put #FileNo, , SendFileRec
                        
                    
                    com = BtOpGetNext
                Loop
        
                Close #FileNo
        
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '送信メッセージを作成する
    Send_Text.sts = Sts_OK                                  'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_REF                     '表示画面フラグ 参照画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_REF
    
    Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '送信データファイル名
    Send_Text.FileName = BA_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    ID_KANRI_TBL(ING_No).Send_Text.FileName = BA_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    
    Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------１～５行目
                                                            
    For i = 0 To M_Gyo - 1
                                                            'BOX属性
        Send_Text.Box_Type(i).Box_Type = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
                                                            '表示内容
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
                                                            '数値初期表示
        Send_Text.Box_Type(i).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                            '初期カーソル位置
        Send_Text.Box_Type(i).Start_Pos = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                            '入力桁数
        Send_Text.Box_Type(i).Max_Size = "00"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                            'メニュ―番号
        Send_Text.Box_Type(i).MENU = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
    Next i

    Sendbuf = Text_Create_Proc()
    
    
    
    COMPO_Proc = False
    

End Function

