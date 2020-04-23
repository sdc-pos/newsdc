Attribute VB_Name = "mdlMODULE"
Option Explicit


Public Function MODULE_INSPE_CHECK_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『モジュール検品処理』
'       2014.06.24
'2016.05.14 Private--> Public
'-------------------------------------------------------
Dim sts             As Integer



Dim Hinban          As String * 20

Dim Location        As String * 8

Dim i               As Integer
Dim j               As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


Dim MI_QTY          As Long
Dim SUMI_QTY        As Long
Dim Zaiko_QTY       As Long

Dim Use_QTY         As Long

Dim HANTEI_MARK     As String

Dim wkDate          As String * 10

Dim MENU_NO         As String * 2


    MODULE_INSPE_CHECK_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（品番）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                            '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "品番エラー", "", "")
                    
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                MODULE_INSPE_CHECK_PROC = False
                                Exit Function
                        
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                
                        End Select
    
    
    
    
                        '---------------------  <品番（モジュール）> --------------------------------------------
                        Call UniCode_Conv(K0_M_ITEM.JGYOBU, RET_JGYOBU)
                        Call UniCode_Conv(K0_M_ITEM.NAIGAI, RET_NAIGAI)
                        Call UniCode_Conv(K0_M_ITEM.HIN_GAI, Hinban)
                        sts = BTRV(BtOpGetEqual, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            Case BtErrKeyNotFound
                            '   -------------------------------- エラーメッセージ作成
                            
                                
                                '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "モジュール対象外", "×廃棄候補", "")
                                '
                                '
                                'Sendbuf = Text_Create_Proc()
                                'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                'MODULE_INSPE_CHECK_PROC = False
                                'Exit Function
                                
                                                            
                                ID_KANRI_TBL(ING_No).Hinban = Hinban
                                
                                ID_KANRI_TBL(ING_No).MEMO = "モジュール品目未登録"
                                
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                            
                                Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "モジュール品目未登録", "", "", Buzzer_DEF)
                                Sendbuf = Text_Create_Proc()
                            
                                MODULE_INSPE_CHECK_PROC = False
                                
                                Exit Function
                            
                                '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                            
                            
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ(モジュール)", 0)
                                Exit Function
                        End Select
    
                        '---------------------  <モジュール対象チェック>
                        If StrConv(M_ITEM_REC.MODULE_KBN, vbUnicode) <> "1" Then
                        
                            '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "モジュール対象外", "×廃棄候補", "")
                            '
                            '
                            'Sendbuf = Text_Create_Proc()
                            'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            'MODULE_INSPE_CHECK_PROC = False
                            'Exit Function
                            
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "モジュール対象外"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "モジュール対象外", "×廃棄候補", "", Buzzer_DEF)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK_PROC = False
                            
                            Exit Function
                            
                            '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
         
                        End If
                        '---------------------  <供給打ち切りチェック>
                        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "0" Then
                            
                            '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "0非対象", "×廃棄候補", "")
                            '
                            '
                            'Sendbuf = Text_Create_Proc()
                            'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            'MODULE_INSPE_CHECK_PROC = False
                            'Exit Function
                        
                        
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "0非対象"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "0非対象", "×廃棄候補", "", Buzzer_DEF)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK_PROC = False
                            
                            Exit Function
                            '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                        End If
                            
                        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "3" Then
                            
                            '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "3打切り", "×廃棄候補", "")
                            '
                            '
                            '
                            'Sendbuf = Text_Create_Proc()
                            'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            'MODULE_INSPE_CHECK_PROC = False
                            'Exit Function
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "3打切り"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "3打切り", "×廃棄候補", "", Buzzer_DEF)
                            
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK_PROC = False
                            Exit Function
                            '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                        
                        End If
                            
                        '---------------------  <ユニットチェック>      2014.07.03 DELETE
                        'If StrConv(M_ITEM_REC.MODULE_UNIT_KBN, vbUnicode) = "2" Then
                        '
                        '    '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                        '    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "2ユニット子", "担当者確認", "")
                        '    '
                        '    '
                        '    '
                        '    'Sendbuf = Text_Create_Proc()
                        '    'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        '    'MODULE_INSPE_CHECK_PROC = False
                        '    'Exit Function
                        '
                        '
                        '    ID_KANRI_TBL(ING_No).Hinban = Hinban
                        '
                        '    ID_KANRI_TBL(ING_No).MEMO = "2ユニット子"
                        '
                        '    ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        '
                        '    Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "2ユニット子", "担当者確認", "")
                        '    Sendbuf = Text_Create_Proc()
                        '
                        '    MODULE_INSPE_CHECK_PROC = False
                        '    Exit Function
                        '
                        '    '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                        'End If
                        '---------------------  <治具チェック>
                        If StrConv(M_ITEM_REC.KENSA_JIGU, vbUnicode) = "1" Then
                            
                            '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "治具なし", "10番へ移動候補", "")
                            '
                            '
                            '
                            'Sendbuf = Text_Create_Proc()
                            'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            'MODULE_INSPE_CHECK_PROC = False
                            'Exit Function
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "治具なし"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "治具なし", "10番へ移動候補", "", Buzzer_DEF)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK_PROC = False
                            Exit Function
                            
                            '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                        
                        End If
                        
                        '---------------------  <設変チェック>  2014.07.02
                        If StrConv(M_ITEM_REC.SETUHEN_KBN, vbUnicode) = "1" Then
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "×設変有り"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "×設変有り", "", "", Buzzer_DEF)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK_PROC = False
                            Exit Function
                            
                            '--------------------   エラーメッセージを確認メッセージに変更  2014.07.01
                        
                        End If
                        
                        
                        '---------------------  <判定>
                        '現在庫
                        
                        Zaiko_QTY = 0
                        For j = 0 To UBound(Nara_Soko_T)
                        
                            Location = Nara_Soko_T(j)
                        
                            'If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, RET_JGYOBU, RET_NAIGAI, Hinban, Location) Then     '2018.09.18
                            If NEW_SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, RET_JGYOBU, RET_NAIGAI, Hinban, Location) Then  '2018.09.18
                                Exit Function
                            End If
                            Zaiko_QTY = Zaiko_QTY + (SUMI_QTY + MI_QTY)
                        Next j
                    
                    
                        '4ヵ月在庫
                        Use_QTY = Val(StrConv(M_ITEM_REC.HITUYO_SU, vbUnicode)) * Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))
                        
                        
                        HANTEI_MARK = "D 再生候補"
                        If Use_QTY >= 200 Then
                            If Zaiko_QTY >= Use_QTY Then
                                HANTEI_MARK = "A 鮮度確認"
                            Else
                                HANTEI_MARK = "B 再生候補"
                            End If
                        Else
                            If Zaiko_QTY >= Use_QTY Then
                                HANTEI_MARK = "C 鮮度確認"
                            Else
                                HANTEI_MARK = "D 再生候補"
                            End If
                        End If
        
    
    
                        'wkDate = Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 1, 4) & _
                        '        "/" & Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 5, 2) & _
                        '        "/" & Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 7, 2)
    
    
                        
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        
                        ID_KANRI_TBL(ING_No).MEMO = HANTEI_MARK
                        

'               --- 2014.07.17
'                        Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
'                                                Hinban, _
'                                                "判定:" & HANTEI_MARK, _
'                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode)), "#") & "ヵ月在庫:" & Format(Use_QTY), _
'                                                "現在庫:" & Format(Zaiko_QTY))
                        
                        
                        Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                Hinban, _
                                                "判定:" & HANTEI_MARK, _
                                                "上限在庫:" & Format(Use_QTY), _
                                                "現 在 庫:" & Format(Zaiko_QTY), Buzzer_DEF)
                        
'               --- 2014.07.17
                        
                        Sendbuf = Text_Create_Proc()
                        
                        
                        '-----------------------------------------------ヘッダー
                        'Send_Text.sts = Sts_OK                                  'ステータス　OK
                        'ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                        '
                        'Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        'ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                        '
                        'Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        'ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                        '
                        'Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        'ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                        '
                        'Send_Text.FileName = ""                                 '送信データファイル名
                        'ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                        '
                        'Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        'ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        '
                        '-----------------------------------------------１行目
                        '                                                        'BOX属性
                        'Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        '                                                        '表示内容
                        'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, "判定:" & HANTEI_MARK)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, "判定:" & HANTEI_MARK)
                        '                                                        '数値初期表示
                        'Send_Text.Box_Type(0).INIT = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                        '                                                        '初期カーソル位置
                        'Send_Text.Box_Type(0).Start_Pos = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                        '                                                        '入力桁数
                        'Send_Text.Box_Type(0).Max_Size = "00"
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                        '
                        'Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                        '                                                        'BOX属性
                        'Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        '                                                        '表示内容
                        'Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "○再生候補")
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "○再生候補")
                        '                                                        '数値初期表示
                        'Send_Text.Box_Type(1).INIT = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                        '                                                        '初期カーソル位置
                        'Send_Text.Box_Type(1).Start_Pos = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                        '                                                        '入力桁数
                        'Send_Text.Box_Type(1).Max_Size = "00"
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                        '
                        'Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                        '                                                        'BOX属性
                        'Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        '                                                        '表示内容
                        '
                        '
                        'wkDate = Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 1, 4) & _
                        '        "/" & Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 5, 2) & _
                        '        "/" & Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 7, 2)
                        '
                        'Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "設変:" & wkDate)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "設変:" & wkDate)
                        '                                                        '数値初期表示
                        'Send_Text.Box_Type(2).INIT = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                        '                                                        '初期カーソル位置
                        'Send_Text.Box_Type(2).Start_Pos = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                        '                                                        '入力桁数
                        '
                        'Send_Text.Box_Type(2).Max_Size = "00"
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                        '
                        '
                        '
                        'Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                        '                                                        'BOX属性
                        'Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        '                                                        '表示内容
                        'Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode)), "#") & "ヵ月在庫:" & Format(Use_QTY))
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode)), "#") & "ヵ月在庫:" & Format(Use_QTY))
                        '                                                        '数値初期表示
                        'Send_Text.Box_Type(3).INIT = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                        '                                                        '初期カーソル位置
                        'Send_Text.Box_Type(3).Start_Pos = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                        '                                                        '入力桁数
                        '
                        'Send_Text.Box_Type(3).Max_Size = "00"
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                        '
                        '
                        '
                        'Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------５行目
                        '
                        '                                                        'BOX属性
                        'Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        '                                                        '表示内容
                        'Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "現在庫:" & Format(Zaiko_QTY))
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "現在庫:" & Format(Zaiko_QTY))
                        '                                                        '数値初期表示
                        'Send_Text.Box_Type(4).INIT = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                        '                                                        '初期カーソル位置
                        'Send_Text.Box_Type(4).Start_Pos = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                        '                                                        '入力桁数
                        '
                        'Send_Text.Box_Type(4).Max_Size = "00"
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                        '
                        'Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                        '
                        '
                        'Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '２回目の受信（ENT）
                        
                        
            '----------------------------------- データ更新処理開始 -----------
                                            'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                            
                            
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
                                                     ID_KANRI_TBL(ING_No).Hinban, , , , , , , , , , , , , , _
                                                     ID_KANRI_TBL(ING_No).MEMO) Then
                    MODULE_INSPE_CHECK_PROC = SYS_ERR
                    GoTo Abort_Tran
                End If
            End If
                                
                                                        
'-----------------------------<移動歴出力>













'-----------------------------<移動歴出力>
                                                        
                                                        
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
        
        
            MODULE_INSPE_CHECK_PROC = False

            Exit Function
                        
        
    End Select
            
    MODULE_INSPE_CHECK_PROC = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


Public Function MODULE_INSPE_CHECK2_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『モジュール検品処理2』
'       2015.02.19
'2016.05.14 Private--> Public
'
'-------------------------------------------------------
Dim sts             As Integer



Dim Hinban          As String * 20

Dim Location        As String * 8

Dim i               As Integer
Dim j               As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


Dim MI_QTY          As Long
Dim SUMI_QTY        As Long
Dim Zaiko_QTY       As Long

Dim Use_QTY         As Long

Dim HANTEI_MARK     As String

Dim wkDate          As String * 10

Dim MENU_NO         As String * 2


    MODULE_INSPE_CHECK2_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（品番）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                            '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "品番エラー", "", "")
                    
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                MODULE_INSPE_CHECK2_PROC = False
                                Exit Function
                        
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                
                        End Select
    
    
    
    
                        '---------------------  <品番（モジュール）> --------------------------------------------
                        Call UniCode_Conv(K0_M_ITEM.JGYOBU, RET_JGYOBU)
                        Call UniCode_Conv(K0_M_ITEM.NAIGAI, RET_NAIGAI)
                        Call UniCode_Conv(K0_M_ITEM.HIN_GAI, Hinban)
                        sts = BTRV(BtOpGetEqual, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            Case BtErrKeyNotFound
                            '   -------------------------------- エラーメッセージ作成
                                
                                                            
                                ID_KANRI_TBL(ING_No).Hinban = Hinban
                                
                                ID_KANRI_TBL(ING_No).MEMO = "モジュール品目未登録"
                                
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                            
                                Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "モジュール品目未登録", "", "", Buzzer_DEF)
                                Sendbuf = Text_Create_Proc()
                            
                                MODULE_INSPE_CHECK2_PROC = False
                                
                                Exit Function
                            
                            
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ(モジュール)", 0)
                                Exit Function
                        End Select
    
                        '---------------------  <在庫集計>
    
                        '現在庫
                        
                        Zaiko_QTY = 0
                        For j = 0 To UBound(Nara_Soko_T)
                        
                            Location = Nara_Soko_T(j)
                        
                            'If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, RET_JGYOBU, RET_NAIGAI, Hinban, Location) Then     '2018.09.18
                            If NEW_SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, RET_JGYOBU, RET_NAIGAI, Hinban, Location) Then  '2018.09.18
                                Exit Function
                            End If
                            Zaiko_QTY = Zaiko_QTY + (SUMI_QTY + MI_QTY)
                        Next j
                    
                    
                        '4ヵ月在庫
                        Use_QTY = Val(StrConv(M_ITEM_REC.HITUYO_SU, vbUnicode)) * Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))
    
    
    
                        '---------------------  <モジュール対象チェック>
                        If StrConv(M_ITEM_REC.MODULE_KBN, vbUnicode) = "0" Then
                            
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "モジュール対象外"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "モジュール対象外", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "ヶ月在庫:" & Format(Use_QTY), _
                                                                                "現在庫:" & Format(Zaiko_QTY), Buzzer_DEF)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            
                            Exit Function
                            
                        End If
                        
                        
                        If StrConv(M_ITEM_REC.MODULE_KBN, vbUnicode) = "9" Then
                            
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "生涯発注"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "生涯発注", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "ヶ月在庫:" & Format(Use_QTY), _
                                                                                "現在庫:" & Format(Zaiko_QTY), Buzzer_DOUBLE)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            
                            Exit Function
                            
                        End If
                        
                        '>>>>>>>>>  2017.11.24
                        If StrConv(M_ITEM_REC.MODULE_KBN, vbUnicode) = "8" Then
                            
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "全数残し"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "全数残し", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "ヶ月在庫:" & Format(Use_QTY), _
                                                                                "現在庫:" & Format(Zaiko_QTY), Buzzer_DOUBLE)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            
                            Exit Function
                            
                        End If
                        
                        
                        '---------------------  <供給打ち切りチェック>
                        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "0" Then
                        
                        
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "0非対象"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "0非対象", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "ヶ月在庫:" & Format(Use_QTY), _
                                                                                "現在庫:" & Format(Zaiko_QTY), Buzzer_DEF)
                            
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            
                            Exit Function
                        End If
                            
                        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "3" Then
                            
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "3打切り"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "3打切り", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "ヶ月在庫:" & Format(Use_QTY), _
                                                                                "現在庫:" & Format(Zaiko_QTY), Buzzer_DOUBLE)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            Exit Function
                        
                        End If
                            
                        '---------------------  <治具チェック>
                        If StrConv(M_ITEM_REC.KENSA_JIGU, vbUnicode) = "1" Then
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "治具なし"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "治具なし", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "ヶ月在庫:" & Format(Use_QTY), _
                                                                                "現在庫:" & Format(Zaiko_QTY), Buzzer_DEF)
                        
                            Sendbuf = Text_Create_Proc()
                        
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            Exit Function
                            
                        
                        End If
                        
                        '---------------------  <設変チェック>  2014.07.02
                        If StrConv(M_ITEM_REC.SETUHEN_KBN, vbUnicode) = "1" Then
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "×設変有り"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "×設変有り", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "ヶ月在庫:" & Format(Use_QTY), _
                                                                                "現在庫:" & Format(Zaiko_QTY), Buzzer_DOUBLE)
                                                        
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            Exit Function
                            
                        
                        End If
                        
                        
                        '---------------------  <判定>
                        
                        
                        HANTEI_MARK = "D 再生候補"
                        If Use_QTY >= 200 Then
                            If Zaiko_QTY >= Use_QTY Then
                                HANTEI_MARK = "A 鮮度確認"
                            Else
                                HANTEI_MARK = "B 再生候補"
                            End If
                        Else
                            If Zaiko_QTY >= Use_QTY Then
                                HANTEI_MARK = "C 鮮度確認"
                            Else
                                HANTEI_MARK = "D 再生候補"
                            End If
                        End If
        
                        
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        
                        ID_KANRI_TBL(ING_No).MEMO = HANTEI_MARK
                        

                        
                        
                        Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                Hinban, _
                                                "判定:" & HANTEI_MARK, _
                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "ヶ月在庫:" & Format(Use_QTY), _
                                                "現在庫:" & Format(Zaiko_QTY), Buzzer_DEF)
                        
                        
                        Sendbuf = Text_Create_Proc()
                        
                        
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '２回目の受信（ENT）
                        
                        
            '----------------------------------- データ更新処理開始 -----------
                                            'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                            
                            
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
                                                     ID_KANRI_TBL(ING_No).Hinban, , , , , , , , , , , , , , _
                                                     ID_KANRI_TBL(ING_No).MEMO) Then
                    MODULE_INSPE_CHECK2_PROC = SYS_ERR
                    GoTo Abort_Tran
                End If
            End If
                                
                                                        

'-----------------------------<移動歴出力>
                                                        
                                                        
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
        
        
            MODULE_INSPE_CHECK2_PROC = False

            Exit Function
                        
        
    End Select
            
    MODULE_INSPE_CHECK2_PROC = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


Private Sub MODULE_TEXT_PROC(Line1 As String, Line2 As String, Line3 As String, Line4 As String, LINE5 As String, buzzer As String)
'-------------------------------------------------------
'
'   『モジュール検品処理　確認ﾃｷｽﾄ作成』
'       2014.06.24
'
'-------------------------------------------------------
                        
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

'    Send_Text.buzzer = Buzzer_DEF                           'ブザー音　標準
    Send_Text.buzzer = buzzer                               'ブザー音　標準
'    ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
    ID_KANRI_TBL(ING_No).Send_Text.buzzer = buzzer
                        
    '-----------------------------------------------１行目
                                                            'BOX属性
    Send_Text.Box_Type(0).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, Line1)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, Line1)
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
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Line2)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Line2)
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

    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Line3)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Line3)
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
    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Line4)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Line4)
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
    Send_Text.Box_Type(4).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LINE5)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LINE5)
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

End Sub


Public Function Module_In_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『モジュール品番入庫処理のチェック＆更新処理』
'
'
'       2018.10.03
'-------------------------------------------------------
Dim i           As Integer

Dim Hinban      As String * 20


Dim Tanaban     As String * 8
Dim sts         As Integer

Dim QTY         As Long
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim RET_JGYOBU  As String * 1
Dim RET_NAIGAI  As String * 1

Dim MENU_NO     As String * 2

Dim WK_CODE     As String * 5       '2007.05.28
Dim WK_TANKA    As String * 11      '2007.05.28


    Module_In_Proc = True
    
    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（棚番／品番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                            
                            
                    Case LCD_Tanaban        '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        
                        If Trim(Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                        '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.SOKO_NO, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                        
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")    '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                    Module_In_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 混載チェック
                            If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                                If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                                    StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")  '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Module_In_Proc = False
                                    Exit Function
                                End If
                            End If
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.SOKO_NO, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                        
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")    '2017.09.22
                            
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                    Module_In_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
                    
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                                    
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")    '2017.09.22
                        
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                
                                Module_In_Proc = False
                                Exit Function
                            End If
                    
                    
                        End If
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        '------------------ 品目マスタ読込み
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                            '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.SOKO_NO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.SOKO_NO, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")    '2017.09.22
                            
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                            Module_In_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "品番エラー", "", "")  '2017.09.22
                            
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                Module_In_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
                    
                    
                    
                    
                        '在庫集計
'                        sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, RET_JGYOBU, RET_NAIGAI, Hinban, SUMI_QTY, MI_QTY)
'                        Select Case sts
'                            Case False
'                            Case True           'ここでは発生しない
'                                Exit Function
'                            Case SYS_ERR
'                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                Exit Function
'                            Case SYS_CANCEL
'                                Call Err_Send_Proc("在庫使用中", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                Module_In_Proc = False
'                                Exit Function
'                        End Select
        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '棚番をセーブ
                        ID_KANRI_TBL(ING_No).Hinban = Hinban            '品番をセーブ
                        ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '送信する商品化済み数量
                        ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '送信する未商品の数量
                                                                        
                        ID_KANRI_TBL(ING_No).RET_JGYOBU = RET_JGYOBU      '資材対応の事業部
                        ID_KANRI_TBL(ING_No).RET_NAIGAI = RET_NAIGAI      '資材対応の国内外
            
            
            
            
                        '数量付きの送信メッセージを作成する
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
                                                                                            
                        Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                                                                        'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                        '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                                Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
            
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                                Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                                                                        '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                        '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                                        '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "08"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                            
                        Send_Text.Box_Type(1).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                                                                        'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                        '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                                         '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                       '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                        '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "20"
                                                                                            
                        Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                        'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                                        '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_MI_Suryo)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_MI_Suryo)
                                                                        '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                        '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                                        '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                            
                        Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------５行目
                                                                        'BOX属性
                    
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                        '表示内容
'                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "在庫数：" & Format(MI_QTY + SUMI_QTY, "#0"))
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "在庫数：" & Format(MI_QTY + SUMI_QTY, "#0"))
                                                                        
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
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                 'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                    
                        Sendbuf = Text_Create_Proc()
                    
                End Select
            
            Next i
            
        Case Step_Sagyo2_RES        '２回目の受信（数量）
            
            For i = 0 To M_Gyo - 1
                
                Select Case i
            
                    
                    Case 3
            
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Module_In_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Module_In_Proc = False
                            Exit Function
                        End If
            
                '----------------------------------- データ更新処理開始 -----------
                                                    
                                                            'トランザクション開始
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                                        
                                        
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                                    
                            MENU_NO = ""
                        End If
                                        
                                        
                       If RET_JGYOBU = SHIZAI Then
                           Call UniCode_Conv(K0_ITEM.JGYOBU, RET_JGYOBU)
                           Call UniCode_Conv(K0_ITEM.NAIGAI, RET_NAIGAI)
                           Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
                           sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                           Select Case sts
                               Case BtNoErr
                               Case Else
                                   Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                   Sendbuf = Text_Create_Proc()
                                   Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                                   Exit Function
                           End Select
                       
                           '品目ﾏｽﾀの最新仕入先／単価が設定されていた時は、こちらの項目を使用  2007.05.28
                           If Not IsNumeric(StrConv(ITEMREC.LAST_TANKA, vbUnicode)) Then
                               
                               WK_CODE = StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode)
                               WK_TANKA = StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)
                           Else
                               WK_CODE = StrConv(ITEMREC.LAST_CODE, vbUnicode)
                               WK_TANKA = StrConv(ITEMREC.LAST_TANKA, vbUnicode)
                           
                           End If
                       
                       
                       
                           sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).RET_JGYOBU, _
                                                   ID_KANRI_TBL(ING_No).RET_NAIGAI, _
                                                   ID_KANRI_TBL(ING_No).Hinban, _
                                                   Format(Now, "YYYYMMDD"), _
                                                   ID_KANRI_TBL(ING_No).Tanaban, _
                                                   (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                   0, _
                                                   QTY, _
                                                   Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                                   ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                   FILE_RETRY, , _
                                                   WK_CODE, _
                                                   WK_TANKA, , _
                                                   MENU_NO)
                    
                           Select Case sts
                               Case False
                               Case True           '入庫時は発生しない
                               Case SYS_CANCEL
                                   'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")        '2017.09.22
                                   Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "処理中断", "", "", "")    '2017.09.22
                                   Sendbuf = Text_Create_Proc()
                                   ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                   
                                   Module_In_Proc = False
                                   GoTo Abort_Tran
                               Case SYS_ERR
                                   Sendbuf = Text_Create_Proc()
                                   Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                   Module_In_Proc = SYS_ERR    'システム異常発生
                                   
                                   GoTo Abort_Tran
                           End Select
                       
                       
                       
                       
                       
                       
                       Else
                                                           
                                                               
                                                               '入庫更新
                           sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).RET_JGYOBU, _
                                                   ID_KANRI_TBL(ING_No).RET_NAIGAI, _
                                                   ID_KANRI_TBL(ING_No).Hinban, _
                                                   Format(Now, "YYYYMMDD"), _
                                                   ID_KANRI_TBL(ING_No).Tanaban, _
                                                   (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                   0, _
                                                   QTY, _
                                                   Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                                   ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                   FILE_RETRY, , , , , MENU_NO)
                           Select Case sts
                               Case False
                               Case True           '入庫時は発生しない
                               Case SYS_CANCEL
                                   'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")        '2017.09.22
                                   Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "処理中断", "", "", "")    '2017.09.22
                                   Sendbuf = Text_Create_Proc()
                                   ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                   
                                   Module_In_Proc = False
                                   GoTo Abort_Tran
                               Case SYS_ERR
                                   Sendbuf = Text_Create_Proc()
                                   Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                   Module_In_Proc = SYS_ERR    'システム異常発生
                                   
                                   GoTo Abort_Tran
                           End Select
                       End If
                
                                                            'トランザクション終了
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                                
                                
                                
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                        
                        ID_KANRI_TBL(ING_No).Inp_QTY = QTY          '入力数量
                                
                        '在庫付きの送信メッセージを作成する
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
                                
                                
                        '在庫集計
                        sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, ID_KANRI_TBL(ING_No).Tanaban, ID_KANRI_TBL(ING_No).RET_JGYOBU, ID_KANRI_TBL(ING_No).RET_NAIGAI, ID_KANRI_TBL(ING_No).Hinban, SUMI_QTY, MI_QTY)
                        Select Case sts
                            Case False
                            Case True           'ここでは発生しない
                                Exit Function
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc("在庫使用中", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Module_In_Proc = False
                                Exit Function
                        End Select
                                
                                
                                
                                
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
                                                                                            
                        Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                                                                        'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                        '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                                Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2))
            
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                                Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2))
                                                                        '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                        '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                                        '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "08"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                            
                        Send_Text.Box_Type(1).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                                                                        'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                        '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban)
                                                                         '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                       '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                        '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "20"
                                                                                            
                        Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                        'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                        '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "未商品：" & QTY)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "未商品：" & QTY)
                                                                        '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                        '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                                        '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                            
                        Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------５行目
                                                                        'BOX属性
                    
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                        '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "在庫数：" & Format(MI_QTY + SUMI_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "在庫数：" & Format(MI_QTY + SUMI_QTY, "#0"))
                                                                        '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                        '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                        '入力桁数
                        Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                 'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                        
                        
                        
                        Sendbuf = Text_Create_Proc()
                    
                    
                    
                    End Select
    
    
                Next i
    
    
    
            Case Step_Sagyo3_RES        '３回目の受信（ENT）
                    
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
            
            
                Module_In_Proc = False
    
                Exit Function
    
    
    End Select
    
    Module_In_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


