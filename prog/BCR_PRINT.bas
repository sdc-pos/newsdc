Attribute VB_Name = "BCR_PRINT"
Option Explicit

Public Function BCR_DAKUTO_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『ダクトバーコード印字』
'
'       2017.04.10
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Hinban      As String * 20
Dim Maisu       As Long



Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1



Dim FileNo      As Integer
Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag


Dim MENU_NO     As String * 2

Dim FileName    As String


    BCR_DAKUTO_Proc = True


    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES                        '品番／枚数

            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                            
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "品目コード", "未登録です", "")
                            
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                BCR_DAKUTO_Proc = False
                                Exit Function
                            
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        
                        End Select
                                            
                                            
                    Case LCD_MAISU                  '枚数
                                            
                        If Not IsNumeric(ID_KANRI_TBL(ING_No).Recv_text(i)) Then
                            '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "数字を入力して下さい", "", "")
                        
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            BCR_DAKUTO_Proc = False
                            Exit Function
                                            
                        End If
                                            
                                                                    
                        Maisu = Val(ID_KANRI_TBL(ING_No).Recv_text(i))
                        If Maisu < 1 Or Maisu > 9999 Then
                        
                            '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Format(Maisu), "枚数は、1～9999の", "範囲で入力して下さい", "")
                        
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            BCR_DAKUTO_Proc = False
                            Exit Function
                                                
                        End If
                                            
                        FileName = R1_SendFile
                        
                        If BCR_Print_File_Make_Proc(1, FileName, Hinban, Maisu) Then
                        End If
                                
                                
                                    
                        'データ送信
                        ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                                    
                                                    
                        ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                    
                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                    
                        
                        
                        ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU
                        ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI

                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).SURYO = Maisu
                    
                    
                    
                    
                    
                    
                        '-----------------------------------------------ヘッダー
                    
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                        Send_Text.Display_Flg = Display_LABEL                   '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_LABEL
                    
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                    
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                    
                        Send_Text.FileName = FileName                           '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = FileName
                    
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                    
                        Sendbuf = Text_Create_Proc()
                        Exit For
                    
                    
                End Select
            Next i
        Case Step_PRINT_RES         '２回目の受信（印刷終了）
    
                '作業ﾛｸﾞ出力    2008.08.08
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
                                                        ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                        ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                                        MENU_NO, _
                                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                        ID_KANRI_TBL(ING_No).Hinban, _
                                                        , _
                                                        ID_KANRI_TBL(ING_No).SURYO) Then
                        BCR_DAKUTO_Proc = SYS_ERR
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

    
    
    
    
    BCR_DAKUTO_Proc = False
    


End Function
Public Function BCR_JAN_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『JANバーコード印字』
'
'       2017.04.10
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Hinban      As String * 20
Dim Maisu       As Long



Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1



Dim FileNo      As Integer
Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag


Dim MENU_NO     As String * 2

Dim FileName    As String


    BCR_JAN_Proc = True


    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES                        '品番／枚数

            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                            
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "品目コード", "未登録です", "")
                            
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                BCR_JAN_Proc = False
                                Exit Function
                            
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        
                        End Select
                                            
                                            
                    Case LCD_MAISU                  '枚数
                                            
                        If Not IsNumeric(ID_KANRI_TBL(ING_No).Recv_text(i)) Then
                            '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "数字を入力して下さい", "", "")
                        
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            BCR_JAN_Proc = False
                            Exit Function
                                            
                        End If
                                            
                                                                    
                        Maisu = Val(ID_KANRI_TBL(ING_No).Recv_text(i))
                        If Maisu < 1 Or Maisu > 9999 Then
                        
                            '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Format(Maisu), "枚数は、1～9999の", "範囲で入力して下さい", "")
                        
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            BCR_JAN_Proc = False
                            Exit Function
                                                
                        End If
                                            
                        FileName = R2_SendFile
                        
                        If BCR_Print_File_Make_Proc(2, FileName, Hinban, Maisu) Then
                        End If
                                
                                
                                    
                        'データ送信
                        ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                                    
                                                    
                        ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                    
                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                    
                        
                        
                        ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU
                        ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI

                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).SURYO = Maisu
                    
                    
                    
                    
                    
                    
                        '-----------------------------------------------ヘッダー
                    
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                        Send_Text.Display_Flg = Display_LABEL                   '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_LABEL
                    
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                    
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                    
                        Send_Text.FileName = FileName                           '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = FileName
                    
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                    
                        Sendbuf = Text_Create_Proc()
                        Exit For
                    
                    
                End Select
            Next i
        Case Step_PRINT_RES         '２回目の受信（印刷終了）
    
                '作業ﾛｸﾞ出力    2008.08.08
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
                                                        ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                        ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                                        MENU_NO, _
                                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                        ID_KANRI_TBL(ING_No).Hinban, _
                                                        , _
                                                        ID_KANRI_TBL(ING_No).SURYO) Then
                        BCR_JAN_Proc = SYS_ERR
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

    
    
    
    
    BCR_JAN_Proc = False

End Function
                    
Public Function BCR_Inspe_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『検品ラベルバーコード印字』
'
'       2017.04.10
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim ID_NO       As String * 7


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1



Dim FileNo      As Integer
Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag


Dim MENU_NO     As String * 2

Dim FileName    As String


    BCR_Inspe_Proc = True


    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES                        'ID

            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_ID                     'ID_NO
                        
                        
                        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.06.14
'                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) <> 7 Then
'
'                            '   -------------------------------- エラーメッセージ作成
'                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "ＩＤは７桁で", "入力して下さい", "")
'
'                            Sendbuf = Text_Create_Proc()
'                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                            BCR_Inspe_Proc = False
'                            Exit Function
'
'                        End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.06.14
                        
                        
                        
                        
                        'ID_NO = ID_KANRI_TBL(ING_No).Recv_text(i)          2017.06.14
                        ID_NO = ID_KANRI_TBL(ING_No).Recv_text(i)           '2017.06.14
                        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.06.14
'                        If IsNumeric(ID_KANRI_TBL(ING_No).Recv_text(i)) Then                    '2017.06.14
'                            ID_NO = Format(Val(ID_KANRI_TBL(ING_No).Recv_text(i)), "0000000")   '2017.06.14
'                        Else                                                                    '2017.06.14
'                            ID_NO = ID_KANRI_TBL(ING_No).Recv_text(i)                           '2017.06.14
'                        End If                                                                  '2017.06.14
'
'
'                        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, ID_NO)
'                        sts = BTRV(BtOpGetGreater, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
'                        Select Case sts
'                            Case BtNoErr
'                                If ID_NO <> Mid(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 1, 7) Then
'
'                                    '   -------------------------------- エラーメッセージ作成
'                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "ＩＤ未登録です", "", "")7
'
'                                    Sendbuf = Text_Create_Proc()
'                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                    BCR_Inspe_Proc = False
'                                    Exit Function
'
'                                End If
'
'
'                            Case BtErrEOF
'
'                                '   -------------------------------- エラーメッセージ作成
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "ＩＤ未登録です", "", "")
'
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                BCR_Inspe_Proc = False
'                                Exit Function
'
'                            Case Else
'                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                Call File_Error(sts, BtOpGetGreater, "出荷予定(H)", 0)
'                                Exit Function
'                        End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2017.06.14
                        
                                            
                        FileName = R3_SendFile
                        
                        If BCR_Print_File_Make_Proc(3, FileName, ID_NO, 1) Then
                        End If
                                
                                
                                    
                        'データ送信
                        ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                                    
                                                    
                        ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                    
                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                    
                        
                        
                        ID_KANRI_TBL(ING_No).S_JGYOBU = ID_KANRI_TBL(ING_No).JGYOBU
                        ID_KANRI_TBL(ING_No).S_NAIGAI = ID_KANRI_TBL(ING_No).NAIGAI

                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                    
                    
                    
                    
                    
                    
                        '-----------------------------------------------ヘッダー
                    
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                        Send_Text.Display_Flg = Display_LABEL                   '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_LABEL
                    
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                    
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                    
                        Send_Text.FileName = FileName                           '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = FileName
                    
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                    
                        Sendbuf = Text_Create_Proc()
                        Exit For
                    
                    
                End Select
            Next i
        Case Step_PRINT_RES         '２回目の受信（印刷終了）
    
                '作業ﾛｸﾞ出力    2008.08.08
                If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                    MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                            
                Else
                            
                    MENU_NO = ""
                End If
                If Trim(MENU_NO) = "" Then
                Else
                '作業ﾛｸﾞ出力
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.06.14
'                    If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
'                                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
'                                                        ID_KANRI_TBL(ING_No).S_JGYOBU, _
'                                                        ID_KANRI_TBL(ING_No).S_NAIGAI, _
'                                                        MENU_NO, _
'                                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
'                                                        , _
'                                                        , _
'                                                        , _
'                                                        , _
'                                                        , _
'                                                        ID_KANRI_TBL(ING_No).ID_NO) Then
'                        BCR_Inspe_Proc = SYS_ERR
'                        Exit Function
'                    End If
                
                
                
                    If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                        ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                        ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                                        MENU_NO, _
                                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                        , _
                                                        , _
                                                        1, _
                                                        , _
                                                        , _
                                                        ID_KANRI_TBL(ING_No).ID_NO) Then
                        BCR_Inspe_Proc = SYS_ERR
                        Exit Function
                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.06.14
                
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

    
    
    
    
    BCR_Inspe_Proc = False


End Function
                    
                    
Public Function BCR_TANA_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『棚番バーコード印字』
'
'       2017.04.10
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Tanaban       As String * 9


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1



Dim FileNo      As Integer
Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag


Dim MENU_NO     As String * 2

Dim FileName    As String


    BCR_TANA_Proc = True


    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES                        '棚番

            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban                '棚番
                        
                        
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) <> 9 Then
                        
                            '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "棚番は９桁で", "入力して下さい", "")
                        
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            BCR_TANA_Proc = False
                            Exit Function
                                                
                        End If
                        Tanaban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                            
                        If Mid(Tanaban, 1, 1) <> "/" Then
                                            
                            '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Tanaban, "棚番の先頭一桁は、", "/を入力して下さい。", "")
                        
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            BCR_TANA_Proc = False
                            Exit Function
                                            
                        End If
                                            
                                            
                        FileName = R4_SendFile
                        
                        If BCR_Print_File_Make_Proc(4, FileName, Tanaban, 1) Then
                        End If
                                
                                
                                    
                        'データ送信
                        ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                                    
                                                    
                        ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                    
                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                    
                        
                        
                        ID_KANRI_TBL(ING_No).S_JGYOBU = ID_KANRI_TBL(ING_No).JGYOBU
                        ID_KANRI_TBL(ING_No).S_NAIGAI = ID_KANRI_TBL(ING_No).NAIGAI

                        ID_KANRI_TBL(ING_No).Tanaban = Right(Tanaban, 8)
                    
                    
                    
                    
                    
                    
                        '-----------------------------------------------ヘッダー
                    
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                        Send_Text.Display_Flg = Display_LABEL                   '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_LABEL
                    
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                    
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                    
                        Send_Text.FileName = FileName                           '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = FileName
                    
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                    
                        Sendbuf = Text_Create_Proc()
                        Exit For
                    
                    
                End Select
            Next i
        Case Step_PRINT_RES         '２回目の受信（印刷終了）
    
                '作業ﾛｸﾞ出力    2008.08.08
                If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                    MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                            
                Else
                            
                    MENU_NO = ""
                End If
                If Trim(MENU_NO) = "" Then
                Else
                '作業ﾛｸﾞ出力
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.06.14
'                    If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
'                                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
'                                                        ID_KANRI_TBL(ING_No).S_JGYOBU, _
'                                                        ID_KANRI_TBL(ING_No).S_NAIGAI, _
'                                                        MENU_NO, _
'                                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
'                                                        , _
'                                                        , _
'                                                        , _
'                                                        ID_KANRI_TBL(ING_No).Tanaban) Then
'                        BCR_TANA_Proc = SYS_ERR
'                        Exit Function
'                    End If
                
                
                
                    If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                        ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                        ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                                        MENU_NO, _
                                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                        , _
                                                        , _
                                                        1, _
                                                        ID_KANRI_TBL(ING_No).Tanaban) Then
                        BCR_TANA_Proc = SYS_ERR
                        Exit Function
                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.06.14
                
                
                
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

    
    
    
    
    BCR_TANA_Proc = False

End Function


Private Function BCR_Print_File_Make_Proc(Mode As Integer, FileName As String, strPrint As String, Optional Maisu As Long = 1) As Integer
'-------------------------------------------------------
'
'   『ﾊﾞｰｺｰﾄﾞ印刷用ﾌｧｲﾙ作成』
'
'       2017.04.10
'-------------------------------------------------------
    
Dim sts         As Integer


Dim FileNo      As Long

Dim FullPath    As String

Dim wkHinban    As String * 14
Dim wklen       As Long
    
Dim wkTANABAN   As String * 8
    
Dim Den_SU      As Long
    
    
    
    BCR_Print_File_Make_Proc = True
    If Right(F1100101.CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = F1100101.CtrsWsk1.SendFolder & "\" & FileName & Format(ID_KANRI_TBL(ING_No).ID, "000") & ".txt"
    Else
        FullPath = F1100101.CtrsWsk1.SendFolder & FileName & Format(ID_KANRI_TBL(ING_No).ID, "000") & ".txt"
    End If




    On Error Resume Next
    Kill (FullPath)             '送信用ファイル削除
    On Error GoTo 0
        
    FileNo = FreeFile           '送信用ファイルＯＰＥＮ
    Open FullPath For Output As #FileNo





    Select Case Mode
        Case 1              'ダクトラベル
            
            Print #FileNo, "#"
            Print #FileNo, "JOB"
            Print #FileNo, LABEL01_DEF
            Print #FileNo, "START"
            Print #FileNo, LABEL01_HIN_F
            Print #FileNo, LABEL01_HIN_T
            wkHinban = Trim(strPrint)
'            If Len(Trim(wkHinban)) < 14 Then
'                wklen = 14 - Len(Trim(wkHinban))
'
'
'                wklen = ToRoundDown(CCur(wklen / 2), 0)
'                If wklen < 1 Then
'                Else
'                    wkHinban = Space(wklen) & Trim(wkHinban)
'                End If
'            End If
            Print #FileNo, wkHinban
            Print #FileNo, LABEL01_HIN_B
            Print #FileNo, Left(strPrint, 14)
    
            Print #FileNo, LABEL01_BIK_F
            Print #FileNo, LABEL01_BIK_T
            Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
            Call UniCode_Conv(K0_ITEM.HIN_GAI, strPrint)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.BIKOU20, "")
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                
                    Call UniCode_Conv(ITEMREC.IRI_QTY, "")
                                    
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                    BCR_Print_File_Make_Proc = SYS_ERR
                    Exit Function
            End Select
            Print #FileNo, Trim(StrConv(ITEMREC.BIKOU20, vbUnicode)) & "ｍ"
    
    
            Print #FileNo, LABEL01_IRI_F
            Print #FileNo, LABEL01_IRI_T
            If IsNumeric(Trim(StrConv(ITEMREC.IRI_QTY, vbUnicode))) Then
'                Print #FileNo, Format(Val(Trim(StrConv(ITEMREC.IRI_QTY, vbUnicode))), "#0") & "本入り"
                Print #FileNo, Format(Val(Trim(StrConv(ITEMREC.IRI_QTY, vbUnicode))), "#0") & "本入"
            Else
'                Print #FileNo, "本入り"
                Print #FileNo, "本入"
            End If
    
    
            Print #FileNo, LABEL01_LOC_F
            Print #FileNo, LABEL01_LOC_T
'            Print #FileNo, StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)
            Print #FileNo, "ﾛｹｰｼｮﾝ"
            
            Print #FileNo, LABEL01_LOC_B
            Print #FileNo, "/" & StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>> 001/[枚数] DELETE 2017.04.24
'            Print #FileNo, LABEL01_QTY_F
'            Print #FileNo, LABEL01_QTY_T
'            Print #FileNo, "001/" & Maisu
'>>>>>>>>>>>>>>>>>>>>>>>>>>>> 001/[枚数] DELETE 2017.04.24
                        
            Print #FileNo, "QTY P=" & Maisu
            Print #FileNo, "END"
            Print #FileNo, "JOBE"
    
    
        Case 2              'Janラベル
    
            Print #FileNo, "#"
            Print #FileNo, "JOB"
            Print #FileNo, LABEL02_DEF
            Print #FileNo, "START"
            Print #FileNo, LABEL02_HIN_F
            Print #FileNo, LABEL02_HIN_T
    
            wkHinban = Trim(strPrint)
            If Len(Trim(wkHinban)) < 14 Then
                wklen = 14 - Len(Trim(wkHinban))
                
                
                wklen = ToRoundDown(CCur(wklen / 2), 0)
                If wklen < 1 Then
                Else
                    wkHinban = Space(wklen) & Trim(wkHinban)
                End If
            End If
            Print #FileNo, wkHinban
            Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
            Call UniCode_Conv(K0_ITEM.HIN_GAI, strPrint)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.JAN_CODE, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                    BCR_Print_File_Make_Proc = SYS_ERR
                    Exit Function
            End Select
            Print #FileNo, LABEL02_HIN_B
            Print #FileNo, StrConv(ITEMREC.JAN_CODE, vbUnicode)
            
            Print #FileNo, "QTY P=" & Maisu
            Print #FileNo, "END"
            Print #FileNo, "JOBE"
    
    
        Case 3              '検品ラベル
    
    
    
            Call UniCode_Conv(K4_Y_SYU_H.ID_NO, strPrint & "zz")
            sts = BTRV(BtOpGetLess, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
            Select Case sts
                Case BtNoErr
                    If strPrint <> Mid(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 1, 7) Then
                
                    
                        Den_SU = 0
                    
                    Else
                    
                        Den_SU = Val(Mid(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 8, 2))
                    
                    End If
                
                
                Case BtErrEOF
                        
                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "")
                    Call UniCode_Conv(Y_SYU_HREC.OKURISAKI, "")
                
                Case Else
                    
                    Call File_Error(sts, BtOpGetGreater, "出荷予定(H)", 0)
                    BCR_Print_File_Make_Proc = SYS_ERR
                    Exit Function
            
            End Select
    
    
    
            Call UniCode_Conv(K4_Y_SYU_H.ID_NO, strPrint)
            sts = BTRV(BtOpGetGreater, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
            Select Case sts
                Case BtNoErr
                    If strPrint <> Mid(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 1, 7) Then
                
                    
                        Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "")
                        Call UniCode_Conv(Y_SYU_HREC.OKURISAKI, "")
                    
                    End If
                
                
                Case BtErrEOF
                        
                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "")
                    Call UniCode_Conv(Y_SYU_HREC.OKURISAKI, "")
                
                Case Else
                    
                    Call File_Error(sts, BtOpGetGreater, "出荷予定(H)", 0)
                    BCR_Print_File_Make_Proc = SYS_ERR
                    Exit Function
            
            End Select
    
    
    
    
    
    
    
            Print #FileNo, "#"
            Print #FileNo, "JOB"
            Print #FileNo, LABEL03_DEF
            Print #FileNo, "START"
            Print #FileNo, LABEL03_ID_F
            Print #FileNo, LABEL03_ID_T
            Print #FileNo, strPrint
            
            Print #FileNo, LABEL03_UN_F
            Print #FileNo, LABEL03_UN_T
            Print #FileNo, StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)
            
            
            
            Print #FileNo, LABEL03_ID_B
            Print #FileNo, strPrint
                
                
            Print #FileNo, LABEL03_OKURI_F
            Print #FileNo, LABEL03_OKURI_T
            Print #FileNo, StrConv(Trim(Left(StrConv(Y_SYU_HREC.OKURISAKI, vbUnicode), 15)), vbWide)
                
                
            If Den_SU <> 0 Then                     '2017.06.14
                Print #FileNo, LABEL03_DEN_F
                Print #FileNo, LABEL03_DEN_T
                Print #FileNo, Format(Den_SU, "#0点")
            End If                                  '2017.06.14
                
                
                
            Print #FileNo, "QTY P=1"
            Print #FileNo, "END"
            Print #FileNo, "JOBE"
    
        Case 4              '棚番ラベル
    
            Print #FileNo, "#"
            Print #FileNo, "JOB"
            Print #FileNo, LABEL04_DEF
            Print #FileNo, "START"
            Print #FileNo, LABEL04_LOC_F
            Print #FileNo, LABEL04_LOC_T
            
            Print #FileNo, strPrint
            Print #FileNo, LABEL04_LOC_B
            Print #FileNo, strPrint
                
            Print #FileNo, "QTY P=1"
            Print #FileNo, "END"
            Print #FileNo, "JOBE"
    
    
    End Select




    Close #FileNo

    FileName = FileName & Format(ID_KANRI_TBL(ING_No).ID, "000") & ".txt"


    BCR_Print_File_Make_Proc = False



End Function
