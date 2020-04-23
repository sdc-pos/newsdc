Attribute VB_Name = "SYOHIN_CHECK_PROC"
Option Explicit
Public Function SHOUHINKA_CHECK_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『商品化ﾁｪｯｸ処理』
'       2010.09.03
'
'   mdlProcより移動 2015.11.07
'-------------------------------------------------------
Dim sts             As Integer



Dim SHIJI_No        As String * 8


Dim SHIJI_QTY       As String * 11

'2010.12.07
'Dim Hinban          As String * 13
Dim Hinban          As String * 20
'2010.12.07

Dim i               As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1



Dim MENU_NO         As String * 2

Dim HIN_NAI         As String * 20
Dim HIN_NAI_READ    As Integer



    SHOUHINKA_CHECK_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（指図票№）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_SHIJI_NO      '指図票№
    
    
                        SHIJI_No = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        
                        
                        
                        Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, SHIJI_No)
                        sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
                        Select Case sts
                            Case False          '正常
                                
                            
                            
                            Case BtErrKeyNotFound
                                    
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "指図票№:" & SHIJI_No, "該当データなし", "", "")
                                
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "該当データなし", "", "")
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SHOUHINKA_CHECK_PROC = False
                                Exit Function
                            
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        
                        End Select
                
                        
                        
                        'キャンセル済
                        If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = "1" Then
'                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "指図票№:" & SHIJI_No, "キャンセル済", "", "")
                            
                            
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "キャンセル済", "", "")
                            
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SHOUHINKA_CHECK_PROC = False
                            Exit Function
                        End If
                                                
                        '完了済
                        If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = "1" Then
'                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "指図票№:" & SHIJI_No, "完了済", "", "")
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "完了済", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SHOUHINKA_CHECK_PROC = False
                            Exit Function
                        End If
                        
                        
                        
                        
                        SHIJI_QTY = Format(Val(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)) - Val(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "00000000.00")
                        If Val(SHIJI_QTY) < 0 Then
                            SHIJI_QTY = "00000000.00"
                        End If
                        
                        ID_KANRI_TBL(ING_No).SHIJI_No = SHIJI_No
                        ID_KANRI_TBL(ING_No).Hinban = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)
                        ID_KANRI_TBL(ING_No).SHIJI_QTY = SHIJI_QTY
                        ID_KANRI_TBL(ING_No).LABEL_CNT = 0
                        ID_KANRI_TBL(ING_No).GENPIN_CNT = 0
                        
                        
                        
                        
                        
                        
                        
                        
                        
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
                
                        Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                                                                                
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
                        '2010.12.07
'                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                                                                                
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                        '2010.12.07
                                                                                
                                                                                
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
                                                                                'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_L_HIN_CNT)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_L_HIN_CNT)
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                '入力桁数
                        '2010.12.07
'                        Send_Text.Box_Type(3).Max_Size = "13"
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                        Send_Text.Box_Type(3).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                        '2010.12.07
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
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
            
                
'                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                Select Case i
                
                    Case 0
                
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        If Trim(Hinban) = Ent_Para Then
                        
                        
                        
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
                    
                            Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                            
                            '-----------------------------------------------１行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '表示内容
'                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                    
                                                                                    
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
                            '2010.12.07
'                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                                                                                    
                            '2010.12.07
                                                                                    
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
                                                                                    'BOX属性
                            Send_Text.Box_Type(3).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                    '表示内容
                            
                            
                            '2010.12.07
'                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
'                                                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & _
'                                                                                    Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & _
                                                                                    Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            
                            '2010.12.07
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
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                                                                                    '数値初期表示
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(4).Start_Pos = "01"                  '数値は５桁固定
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    '入力桁数
'2011.04.11                            Send_Text.Box_Type(4).Max_Size = "13"
'2011.04.11                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "13"
                                                                                    
                                                                                    
                            Send_Text.Box_Type(4).Max_Size = "20"                           '2011.04.11 13-->20
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"      '2011.04.11 13-->20
                                                                                    
                                                                                    
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
                            Sendbuf = Text_Create_Proc()
                        
                            SHOUHINKA_CHECK_PROC = False
    
                            Exit Function
            
                        End If
                
                
                
                
                
                
                    '品番
                    Case 3
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                                    
                        'sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)                                 '2018.09.19
                        sts = Item_Read4_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI, , , 1, HIN_NAI_READ, HIN_NAI)    '2018.09.19
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                        
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                
                        End Select
                        
                        
                        '2018.09.19
                        If Trim(HIN_NAI) = "" Then
                            ID_KANRI_TBL(ING_No).IN_HINBAN_L = Hinban
                        Else
                            ID_KANRI_TBL(ING_No).IN_HINBAN_L = HIN_NAI
                        End If
                        '2018.09.19
                        
                        
                        If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                            If Split(Trim(Hinban), " ")(0) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then '2019/10/28 広島 商品化チェック品番バーコード空白対応
                            
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
                                Send_Text.buzzer = Buzzer_DOUBLE                        'ブザー音　標準
                                ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                                '-----------------------------------------------１行目
                                                                                        'BOX属性
                                Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                        '表示内容
    '                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
    '                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                        
                                                                                        
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
                                '2010.12.07
    '                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
    '                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                                                                                        
                                                    
                                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                                
                                
                                '2010.12.07
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
                                                                                        
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
                                                                                        'BOX属性
                                Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                        '表示内容
                                '2010.12.07
    '                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Hinban & "    ｴﾗｰ")
    '                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Hinban & "    ｴﾗｰ")
                                                                                        
                                '2018.09.19
                                'Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(Hinban, 16) & " ｴﾗｰ")
                                'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(Hinban, 16) & " ｴﾗｰ")
                                
                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_L, 16) & " ｴﾗｰ")
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_L, 16) & " ｴﾗｰ")
                                '2018.09.19
                                
                                '2010.12.07
                                                                                        
                                                                                        '数値初期表示
                                Send_Text.Box_Type(3).INIT = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                        '初期カーソル位置
                                Send_Text.Box_Type(3).Start_Pos = "01"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                        '入力桁数
                                '2010.12.07
    '                            Send_Text.Box_Type(3).Max_Size = "13"
    '                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                Send_Text.Box_Type(3).Max_Size = "20"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                '2010.12.07
                                                                                        
                                Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                                '-----------------------------------------------４行目
                                                                                        'BOX属性
                                Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                        '表示内容
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
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
                                SHOUHINKA_CHECK_PROC = False
        
                                Exit Function
                        
                            End If
                        
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
                
                        
                        
                        '2018.09.19
                        If HIN_NAI_READ = 1 Then
                        
                            Send_Text.buzzer = Buzzer_DEF                        'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        Else
                        '2018.09.19
                        
                            Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
                        End If                              '2018.09.19

                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                                                                                
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
                        '2010.12.07
'                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                        
                        
                        '2010.12.07
                                                                                
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
                                                                                'BOX属性
                        
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        
                        
                        ID_KANRI_TBL(ING_No).LABEL_CNT = ID_KANRI_TBL(ING_No).LABEL_CNT + 1
                        '2010.12.07
'                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
'                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Hinban & _
'                                                                                Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        
                        '2018.09.19
                        'Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                        '                                                        Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(Hinban, 16) & _
                        '                                                        Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        
                        
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_L, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_L, 16) & _
                                                                                Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        
                        
                        '2018.09.19
                        
                        '2010.12.07
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                '入力桁数
                        '2010.12.07
                        Send_Text.Box_Type(3).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                                                                
                        Send_Text.Box_Type(3).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                        '2010.12.07
                                                                                
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
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
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        Case Step_Sagyo3_RES        '３回目の受信（現品票品番）
            For i = 0 To M_Gyo - 1
            
                
'                Select Case Trim(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD)
                Select Case i
                
                                    
                    Case 0
                                    
                                    
                                    
                                    
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    
                        If Trim(Hinban) = Ent_Para Then
                        
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ
                            
                        
                        
                        
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
                    
                            Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                                    
                            '-----------------------------------------------１行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '表示内容
'                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                    
                                                                                    
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
                            Send_Text.Box_Type(4).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    '入力桁数
                            Send_Text.Box_Type(4).Max_Size = "20"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                
                            Sendbuf = Text_Create_Proc()
                        
                            SHOUHINKA_CHECK_PROC = False
        
                            Exit Function

                                                            
                    End If
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                    Case 4          '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        
                        
                        
                        
                        'sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)                                 '2018.09.19
                        ' Item_Read4_Proc->Item_Read3_Proc 2018.10.17
                        sts = Item_Read4_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI, , , 1, HIN_NAI_READ, HIN_NAI)    '2018.09.19
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                        
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                
                        End Select
                        
                        
                            '2018.09.19
                            If Trim(HIN_NAI) = "" Then
                                ID_KANRI_TBL(ING_No).IN_HINBAN_G = Hinban
                            Else
                                ID_KANRI_TBL(ING_No).IN_HINBAN_G = HIN_NAI
                            End If
                            '2018.09.19
                        
                        
                        
                        If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                            If Split(Trim(Hinban), " ")(0) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then '2019/10/28 広島 商品化チェック現品票バーコード空白対応
                        
                        
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
                    
                            Send_Text.buzzer = Buzzer_DOUBLE                        'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                            
                            '-----------------------------------------------１行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '表示内容
'                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                    
                                                                                    
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
                                                        
                            '2010.12.07
'                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))                            '2010.12.08
                            
                            '2010.12.07
                                                                                    
                                                                                    
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
                                                                                    'BOX属性
                            Send_Text.Box_Type(3).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                    '表示内容
                            '2010.12.07
'                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, ID_KANRI_TBL(ING_No).Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
'                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, ID_KANRI_TBL(ING_No).Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
'                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))

                            
                            '2018.09.19
                            
                            'Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                            '                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                            '                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_L, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_L, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            
                            
                            '2018.09.19
                            
                            '2010.12.07


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
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                                                    '表示内容
                            '2010.12.07
'                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Hinban & "    ｴﾗｰ")
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Hinban & "    ｴﾗｰ")
                                                                                    
                            '2018.09.19
                            'Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Left(Hinban, 16) & " ｴﾗｰ")
                            'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Left(Hinban, 16) & " ｴﾗｰ")
                            
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_G, 16) & " ｴﾗｰ")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_G, 16) & " ｴﾗｰ")
                            '2018.09.19
                            
                            
                            '2010.12.07
                                                                                    
                                                                                    '数値初期表示
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(4).Start_Pos = "01"                  '数値は５桁固定
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    '入力桁数
''                            Send_Text.Box_Type(4).Max_Size = "13"
''                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "13"
                                                                                    
                            Send_Text.Box_Type(4).Max_Size = "20"                               '2011.04.11 13-->20
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"          '2011.04.11 13-->20
                                                                                    
                                                                                    
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
                            Sendbuf = Text_Create_Proc()
                            SHOUHINKA_CHECK_PROC = False
        
                            Exit Function
                                                    
                            End If
                        
                        
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
                
                        
                        
                        If HIN_NAI_READ = 1 Then        '2018.09.19
                            Send_Text.buzzer = Buzzer_DEF                        '2018.09.19
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF   '2018.09.19
                        Else                                                        '2018.09.19
                            Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        End If                                                      '2018.09.19
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
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
                        '2010.12.07
'                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                                                                                
                        
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                        '2010.12.07
                                                                                
                                                                                
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
                                                                                
                        '2010.12.07
'                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
'                                                                            Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
'                                                                            Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))


                        '2018.09.19
                        'Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                        '                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                        '                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_L, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                            Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_L, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                            Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        
                        '2018.09.19
                        '2010.12.07


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
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        ID_KANRI_TBL(ING_No).GENPIN_CNT = ID_KANRI_TBL(ING_No).GENPIN_CNT + 1
                        '2010.12.07
'                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
'                                                                                Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Hinban & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
'                                                                                Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                                                                                
                        '2018.09.19
                        'Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Left(Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                        '                                                        Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Left(Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                        '                                                        Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                        
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_G, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).IN_HINBAN_G, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                        
                        '2018.09.19
                        
                        
                        '2010.12.07
                                                                                
                                                                                
                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                '入力桁数
                        '2010.12.07
'                        Send_Text.Box_Type(4).Max_Size = "13"
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(4).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"
                        '2010.12.07
                                                                                
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
                        Sendbuf = Text_Create_Proc()
                
                
                
                End Select
            
            Next i
        
        
        
        
        
        Case Step_Sagyo4_RES        '５回目の受信（Any Key）
            
            '----------------------------------- データ更新処理開始 -----------
                                            'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
            
            
            Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, ID_KANRI_TBL(ING_No).SHIJI_No)
            sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case False          '正常
                    
                
                
                Case BtErrKeyNotFound
                        
'                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "指図票№:" & SHIJI_No, "該当データなし", "", "")
                    
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "該当データなし", "", "")
                    
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    GoTo Abort_Tran
                
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    GoTo Abort_Tran
            
            End Select
            
            Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, ID_KANRI_TBL(ING_No).TANTO_CODE)
            Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
            Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_LABEL_CNT, Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "000"))
            Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GENPIN_CNT, Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "000"))
            sts = BTRV(BtOpUpdate, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case False          '正常
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    GoTo Abort_Tran
            End Select
            
            
            
            
            
            
            
            
            
            
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
                                                     ID_KANRI_TBL(ING_No).Hinban, , CLng(ID_KANRI_TBL(ING_No).SHIJI_QTY), , , , , , , _
                                                     ID_KANRI_TBL(ING_No).SHIJI_No, Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "000"), Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "000")) Then
                    SHOUHINKA_CHECK_PROC = SYS_ERR
                    GoTo Abort_Tran
                End If
            End If
            '2004.07.16 ↑
                                        
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
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

    SHOUHINKA_CHECK_PROC = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function


Public Function SHOUHINKA_CHECK_GAI_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『商品化ﾁｪｯｸ処理(外装品番有り)』
'       2015.11.07
'
'-------------------------------------------------------
Dim sts             As Integer



Dim SHIJI_No        As String * 8


Dim SHIJI_QTY       As String * 11

Dim Hinban          As String * 20

Dim i               As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1



Dim MENU_NO         As String * 2



    SHOUHINKA_CHECK_GAI_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（指図票№）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_SHIJI_NO      '指図票№
    
    
                        SHIJI_No = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        
                        
                        
                        Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, SHIJI_No)
                        sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
                        Select Case sts
                            Case False          '正常
                                
                            
                            
                            Case BtErrKeyNotFound
                                    
                                
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "該当データなし", "", "")
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SHOUHINKA_CHECK_GAI_PROC = False
                                Exit Function
                            
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        
                        End Select
                
                        
                        
                        'キャンセル済
                        If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = "1" Then
                            
                            
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "キャンセル済", "", "")
                            
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SHOUHINKA_CHECK_GAI_PROC = False
                            Exit Function
                        End If
                                                
                        '完了済
                        If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) = "1" Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "完了済", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SHOUHINKA_CHECK_GAI_PROC = False
                            Exit Function
                        End If
                        
                        
                        
                        
                        SHIJI_QTY = Format(Val(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)) - Val(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "00000000.00")
                        If Val(SHIJI_QTY) < 0 Then
                            SHIJI_QTY = "00000000.00"
                        End If
                        
                        ID_KANRI_TBL(ING_No).SHIJI_No = SHIJI_No
                        ID_KANRI_TBL(ING_No).Hinban = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)
                        ID_KANRI_TBL(ING_No).SHIJI_QTY = SHIJI_QTY
                        ID_KANRI_TBL(ING_No).LABEL_CNT = 0
                        ID_KANRI_TBL(ING_No).GENPIN_CNT = 0
                        ID_KANRI_TBL(ING_No).GAISOU_CNT = 0
                        
                        
                        
                        
                        
                        
                        
                        
                        
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
                
                        Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
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
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_L_HIN_CNT)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_L_HIN_CNT)
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
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_GAI_HIN_CNT)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_GAI_HIN_CNT)
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
                        '-----------------------------------------------５行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
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
                        SHOUHINKA_CHECK_GAI_PROC = False
    
                        Exit Function
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '２回目の受信（ラベル品番）
            For i = 0 To M_Gyo - 1
            
                
                Select Case i
                
                    Case 0
                
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        If Trim(Hinban) = Ent_Para Then
                        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  未チェックの確認    2016.04.15
                            If ID_KANRI_TBL(ING_No).LABEL_CNT = 0 Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & ID_KANRI_TBL(ING_No).SHIJI_No, "商品化未チェックです", "ラベル品番", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SHOUHINKA_CHECK_GAI_PROC = False
                                Exit Function
                            
                            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  未チェックの確認    2016.04.15
                                            
                            
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
                    
                            Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                            
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
                            
                            
                            
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & _
                                                                                    Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            
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
                                                                                    'BOX属性
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_GAI_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_GAI_HIN_CNT)
                                                                                    '数値初期表示
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(3).Start_Pos = "01"                  '数値は５桁固定
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                    
                            Send_Text.Box_Type(3).Max_Size = "20"                           '2011.04.11 13-->20
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"      '2011.04.11 13-->20
                            
                            
                            '-----------------------------------------------５行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                                                                                    '数値初期表示
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(4).Start_Pos = "01"                  '数値は５桁固定
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    
                            Send_Text.Box_Type(4).Max_Size = "20"                           '2011.04.11 13-->20
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"      '2011.04.11 13-->20
                                                                                    
                                                                                    
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
                            Sendbuf = Text_Create_Proc()
                        
                            SHOUHINKA_CHECK_GAI_PROC = False
    
                            Exit Function
            
                        End If
                
                
                
                
                
                
                    '品番
                    Case 2
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                                    
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
                        
                        
                        
                        
                        If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                        
                        
                        
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
                    
                            Send_Text.buzzer = Buzzer_DOUBLE                        'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                            
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
                            Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(Hinban, 16) & " ｴﾗｰ")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(Hinban, 16) & " ｴﾗｰ")
                                                                                    
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
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_G_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_G_HIN_CNT)
                                                                                    '数値初期表示
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(3).Start_Pos = ""                    '数値は５桁固定
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                                    '入力桁数
                             Send_Text.Box_Type(3).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------５行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
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
                            SHOUHINKA_CHECK_GAI_PROC = False
    
                            Exit Function
                        
                        
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
                
                        Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
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
                        
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        
                        
                        ID_KANRI_TBL(ING_No).LABEL_CNT = ID_KANRI_TBL(ING_No).LABEL_CNT + 1
                        
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(Hinban, 16) & _
                                                                                Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
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
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_GAI_HIN_CNT)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_GAI_HIN_CNT)
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = ""                    '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        
                        '-----------------------------------------------５行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
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
                    
                        SHOUHINKA_CHECK_GAI_PROC = False

                        Exit Function
                
                
                
                End Select
            
            Next i
        
        
                
        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   外装ラベル処理
        
        
        Case Step_Sagyo3_RES        '３回目の受信（外装ラベル品番）
            For i = 0 To M_Gyo - 1
            
                
                Select Case i
                
                    Case 0
                
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        If Trim(Hinban) = Ent_Para Then
                        
                        
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ
                            
                            
                            
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
                    
                            Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                            
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
                            
                            
                            
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & _
                                                                                    Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            
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
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))) & _
                                                                                    Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & _
                                                                                    Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))
                            
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
                            
                            
                            '-----------------------------------------------５行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                                                                                    '数値初期表示
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(4).Start_Pos = "01"                  '数値は５桁固定
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    
                            Send_Text.Box_Type(4).Max_Size = "20"                           '2011.04.11 13-->20
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"      '2011.04.11 13-->20
                                                                                    
                                                                                    
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
                            Sendbuf = Text_Create_Proc()
                        
                            SHOUHINKA_CHECK_GAI_PROC = False
    
                            Exit Function
            
                        End If
                
                
                
                
                
                
                    '品番
                    Case 3
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                                    
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
                        
                        
                        
                        
                        If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                        
                        
                        
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
                    
                            Send_Text.buzzer = Buzzer_DOUBLE                        'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                            
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
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                                                                                    
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
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(Hinban, 16) & " ｴﾗｰ")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(Hinban, 16) & " ｴﾗｰ")
                                                                                    '数値初期表示
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(3).Start_Pos = "01"                  '数値は５桁固定
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                    '入力桁数
                             Send_Text.Box_Type(3).Max_Size = "20"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------５行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
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
                            SHOUHINKA_CHECK_GAI_PROC = False
    
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
                
                        Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
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
                        
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(Hinban, 16) & _
                                                                                Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
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
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                
                        ID_KANRI_TBL(ING_No).GAISOU_CNT = ID_KANRI_TBL(ING_No).GAISOU_CNT + 1
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(Hinban, 16) & _
                                                                                Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = "01"                  '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                '入力桁数
                         Send_Text.Box_Type(3).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
    
                        
                        '-----------------------------------------------５行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
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
                    
                        SHOUHINKA_CHECK_GAI_PROC = False

                        Exit Function
                
                
                
                End Select
            
            Next i
        
        
        
        
        
        
        
        
        
        
        
        Case Step_Sagyo4_RES        '４回目の受信（現品票品番）
            For i = 0 To M_Gyo - 1
            
                
                Select Case i
                
                                    
                    Case 0
                                    
                                    
                                    
                                    
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    
                        If Trim(Hinban) = Ent_Para Then
                        
                        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  未チェックの確認    2016.04.15
                            If ID_KANRI_TBL(ING_No).GENPIN_CNT = 0 Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & ID_KANRI_TBL(ING_No).SHIJI_No, "商品化未チェックです", "現品票品番", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SHOUHINKA_CHECK_GAI_PROC = False
                                Exit Function
                            
                            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  未チェックの確認    2016.04.15
                        
                        
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                            
                        
                        
                        
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
                    
                            Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                                    
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
                            Send_Text.Box_Type(4).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    '入力桁数
                            Send_Text.Box_Type(4).Max_Size = "20"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                
                            Sendbuf = Text_Create_Proc()
                        
                            SHOUHINKA_CHECK_GAI_PROC = False
        
                            Exit Function

                        End If
                                    
                    Case 4          '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        
                        
                        
                        
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
                        
                        
                        
                        
                        
                        If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                        
                        
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ
                            
                            
                            
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
                    
                            Send_Text.buzzer = Buzzer_DOUBLE                        'ブザー音　標準
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                            
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
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))


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
                                                                                    'BOX属性
                            Send_Text.Box_Type(3).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))


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
                            '-----------------------------------------------５行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                                                    '表示内容
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Left(Hinban, 16) & " ｴﾗｰ")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Left(Hinban, 16) & " ｴﾗｰ")
                                                                                    
                                                                                    '数値初期表示
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(4).Start_Pos = "01"                  '数値は５桁固定
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    '入力桁数
                                                                                    
                            Send_Text.Box_Type(4).Max_Size = "20"                               '2011.04.11 13-->20
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"          '2011.04.11 13-->20
                                                                                    
                                                                                    
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
                            Sendbuf = Text_Create_Proc()
                            SHOUHINKA_CHECK_GAI_PROC = False
        
                            Exit Function
                                                    
                        
                        
                        
                        End If
                
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ
                        
                        
                        
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
                
                        Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
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
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                '表示内容
                                                                                


                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                            Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                            Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))


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
                                                                                


                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))) & _
                                                                            Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))) & _
                                                                            Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "#0"))


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
                        '-----------------------------------------------５行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        ID_KANRI_TBL(ING_No).GENPIN_CNT = ID_KANRI_TBL(ING_No).GENPIN_CNT + 1
                                                                                
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Left(Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Left(Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                                                                                
                                                                                
                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(4).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"
                                                                                
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
                        Sendbuf = Text_Create_Proc()
                
                        SHOUHINKA_CHECK_GAI_PROC = False

                        Exit Function
                
                
                End Select
            
            Next i
        
        
        
        
        
        Case Step_Sagyo5_RES        '６回目の受信（Any Key）
            
            '----------------------------------- データ更新処理開始 -----------
                                            'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
            
            
            Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, ID_KANRI_TBL(ING_No).SHIJI_No)
            sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case False          '正常
                    
                
                
                Case BtErrKeyNotFound
                        
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "指図票№:" & SHIJI_No, "該当データなし", "", "")
                    
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    GoTo Abort_Tran
                
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    GoTo Abort_Tran
            
            End Select
            
            Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, ID_KANRI_TBL(ING_No).TANTO_CODE)
            Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
            Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_LABEL_CNT, Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "000"))
            Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GENPIN_CNT, Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "000"))
            Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GAISOU_CNT, Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "000"))              '2015.11.07
            sts = BTRV(BtOpUpdate, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case False          '正常
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    GoTo Abort_Tran
            End Select
            
            
            
            
            
            
            
            
            
            
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
                                                     ID_KANRI_TBL(ING_No).Hinban, , CLng(ID_KANRI_TBL(ING_No).SHIJI_QTY), , , , , , , _
                                                     ID_KANRI_TBL(ING_No).SHIJI_No, Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "000"), Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "000"), , , , Format(ID_KANRI_TBL(ING_No).GAISOU_CNT, "000")) Then
                    SHOUHINKA_CHECK_GAI_PROC = SYS_ERR
                    GoTo Abort_Tran
                End If
            End If
            '2004.07.16 ↑
                                        
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
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

    SHOUHINKA_CHECK_GAI_PROC = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function

