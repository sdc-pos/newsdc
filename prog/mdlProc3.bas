Attribute VB_Name = "mdlProc3"
Option Explicit


Public Function NYUKO_KENPIN_OSAKA_S_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『大阪ＰＣ　資材検収入庫』
'
'   2012.03.06
'
'-------------------------------------------------------
Dim i               As Integer


Dim Hinban          As String * 20
Dim Tanaban         As String * 8
Dim QTY             As Long

Dim sts             As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2


Dim wkORDER_QTY     As Long
Dim wkNYUKO_QTY     As Long
Dim wkQty           As Long

Dim ST_TANABAN      As String * 11


    NYUKO_KENPIN_OSAKA_S_Proc = True


    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（品番）

            For i = 0 To M_Gyo - 1
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI, , BUZAI)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")      '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "品番エラー", "", "")  '2017.09.22
                            
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                NYUKO_KENPIN_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
                        
                        
                        ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU
                        ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            ST_TANABAN = ""
                        Else
                            ST_TANABAN = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                        
                        
                        
                        '当月仕入予定残集計
                        If ORDER_ZAN_Proc(RET_JGYOBU, NAIGAI_NAI, Hinban, wkORDER_QTY, wkNYUKO_QTY) Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
                        
                        ID_KANRI_TBL(ING_No).ORDER_QTY = wkORDER_QTY
                        ID_KANRI_TBL(ING_No).NYUKO_QTY = wkNYUKO_QTY
                        
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
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                        '>>>>>>>>>>>    2017.09.22
                        'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>>>>>    2017.09.22
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Tanaban & " " & ST_TANABAN)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Tanaban & " " & ST_TANABAN)
                                                                                
                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "16"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "16"
                                                                                
                                                                                
                                                                                
                                                                                
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
                        Send_Text.Box_Type(3).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                        '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                            
                        Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------５行目
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
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                    
                    
                    

                        Sendbuf = Text_Create_Proc()
                        
                        Exit Function
                
                End Select
            Next i
    
    
        Case Step_Sagyo2_RES        '２回目の受信（棚番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    Case LCD_Tanaban        '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        
                        '------------------ 倉庫マスタ読込み
                        Call UniCode_Conv(K0_SOKO.SOKO_NO, Left(Tanaban, 2))
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                    
                            '   -------------------------------- エラーメッセージ作成
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")    '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                NYUKO_KENPIN_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                Exit Function
                        End Select
                        '------------------ 混載チェック
                        If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                            If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).S_JGYOBU Or _
                                StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")          '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")      '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                NYUKO_KENPIN_OSAKA_S_Proc = False
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
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")            '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")        '2017.09.22
                        
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                NYUKO_KENPIN_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                Exit Function
                        End Select
                    
                        '------------------ 禁止棚のチェック
                        If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                                
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")            '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")        '2017.09.22
                    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            Exit Function
                        End If
                        ID_KANRI_TBL(ING_No).Tanaban = Tanaban
    
    
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))

                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
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
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "注残(" & USE_YM & "):" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "注残(" & USE_YM & "):" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
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
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                        '表示内容
                        
                        wkQty = ID_KANRI_TBL(ING_No).ORDER_QTY - ID_KANRI_TBL(ING_No).NYUKO_QTY
                        If wkQty < 0 Then
                            wkQty = 0
                        End If
                        
                        
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Suryo)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Suryo)
                                                                        '数値初期表示
                        Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(wkQty, "#0"))) & Format(wkQty, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(wkQty, "#0"))) & Format(wkQty, "#0")
                                                                        '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                                        '入力桁数
                        Send_Text.Box_Type(4).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
                        
                        NYUKO_KENPIN_OSAKA_S_Proc = False
                        Exit Function
                End Select
            Next i
        Case Step_Sagyo3_RES        '３回目の受信（数量）
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
            
            
                    Case LCD_Suryo          '数量
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                
                        If QTY <= (ID_KANRI_TBL(ING_No).ORDER_QTY - ID_KANRI_TBL(ING_No).NYUKO_QTY) Then
                        
                        
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
                                                                '入庫更新
                            sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
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
                                    
                                    NYUKO_KENPIN_OSAKA_S_Proc = False
                                    GoTo Abort_Tran
                                Case SYS_ERR
                                    Sendbuf = Text_Create_Proc()
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    NYUKO_KENPIN_OSAKA_S_Proc = SYS_ERR    'システム異常発生
                                    
                                    GoTo Abort_Tran
                            End Select
                                        
                            GoTo End_Tran
                        
                                                    
                        Else
                        
                            ID_KANRI_TBL(ING_No).INPUT_QTY = QTY
                        
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
                    
                            Send_Text.Buzzer = Buzzer_DOUBLE                        'ブザー音　二重警告音
                            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DOUBLE
                            
                            '-----------------------------------------------１行目
                                                                                    'BOX属性
'                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                                                                                    '表示内容
'                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
'                                                                                    '数値初期表示
'                            Send_Text.Box_Type(0).INIT = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
'                                                                                    '初期カーソル位置
'                            Send_Text.Box_Type(0).Start_Pos = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
'                                                                                    '入力桁数
'                            Send_Text.Box_Type(0).Max_Size = "00"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
'
'                            Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                            '-----------------------------------------------２行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).Hinban)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "注文残:" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "注文残:" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
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
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "総入庫数:" & Format(ID_KANRI_TBL(ING_No).NYUKO_QTY + QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "総入庫数:" & Format(ID_KANRI_TBL(ING_No).NYUKO_QTY + QTY, "#0"))
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
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "棚入れを行いますか？")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "棚入れを行いますか？")
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
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_CAN_ANS)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_CAN_ANS)
                                                                                    '数値初期表示
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(4).Start_Pos = "20"                  '数値は５桁固定
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "20"
                                                                                    '入力桁数
                             Send_Text.Box_Type(4).Max_Size = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "01"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
                            Sendbuf = Text_Create_Proc()
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            
                            Exit Function
                        End If
                
                End Select
            Next i
    
    
        Case Step_Sagyo4_RES        '４回目の受信（処理継続(Y/N)）
            
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
            
            
                    Case LCD_CAN_ANS          '数量
                        If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) <> "1" And Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) <> "9" Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "1 OR 9 を入力", "して下さい。", "")          '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "1 OR 9 を入力", "して下さい。", "")      '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                
                
                        If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) = "1" Then
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
                                                                '入庫更新
                            sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                                    ID_KANRI_TBL(ING_No).Hinban, _
                                                    Format(Now, "YYYYMMDD"), _
                                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                    0, _
                                                    ID_KANRI_TBL(ING_No).INPUT_QTY, _
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
                                    
                                    NYUKO_KENPIN_OSAKA_S_Proc = False
                                    GoTo Abort_Tran
                                Case SYS_ERR
                                    Sendbuf = Text_Create_Proc()
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    NYUKO_KENPIN_OSAKA_S_Proc = SYS_ERR    'システム異常発生
                                    
                                    GoTo Abort_Tran
                            End Select
                                        
                            GoTo End_Tran
                        Else
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
                            
                            
                            
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            
                            Exit Function
                        
                        End If
            
                End Select
            Next i
    End Select

End_Tran:
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
    
    
    
    NYUKO_KENPIN_OSAKA_S_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Public Function ORDER_ZAN_Proc(JGYOBU As String, NAIGAI As String, Hinban As String, ORDER_QTY As Long, NYUKO_QTY As Long) As Integer
'-------------------------------------------------------
'
'   『大阪ＰＣ　資材注文残集計』
'
'   2012.03.06
'
'-------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
    
    
Dim i           As Integer
    
Dim wkJgyobu    As String * 1
    
    
    ORDER_ZAN_Proc = True
    
    If JGYOBU = BUZAI Then
        wkJgyobu = SHIZAI
    Else
        wkJgyobu = JGYOBU
    End If
    
    Call UniCode_Conv(K7_P_SHORDER.USE_YM, USE_YM)
    Call UniCode_Conv(K7_P_SHORDER.JGYOBU, wkJgyobu)
    Call UniCode_Conv(K7_P_SHORDER.NAIGAI, NAIGAI)
    Call UniCode_Conv(K7_P_SHORDER.HIN_GAI, Hinban)
    Call UniCode_Conv(K7_P_SHORDER.CANCEL_F, P_CANCEL_OFF)
    
    ORDER_QTY = 0
    com = BtOpGetGreaterEqual
    
        
    Do
        
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K7_P_SHORDER, Len(K7_P_SHORDER), 7)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_SHORDER_REC.USE_YM, vbUnicode) <> USE_YM Or _
                    StrConv(P_SHORDER_REC.JGYOBU, vbUnicode) <> wkJgyobu Or _
                    StrConv(P_SHORDER_REC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    RTrim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) <> RTrim(Hinban) Or _
                    StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) <> P_CANCEL_OFF Then
                       
                    Exit Do
                    
                End If
            
            Case BtErrEOF
                Exit Do
            
            '   -------------------------------- エラーメッセージ作成
            Case Else
                '重要な要因なので未登録はシステム停止とする
                Call File_Error(sts, BtOpGetEqual, "資材注文ﾃﾞｰﾀ", 0)
                Exit Function
        End Select
    
        ORDER_QTY = ORDER_QTY + (Val(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - Val(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)))
    
        com = BtOpGetNext
    
    Loop
    
    Call UniCode_Conv(K1_IDO.JGYOBU, JGYOBU)
    Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K1_IDO.HIN_GAI, Hinban)
    Call UniCode_Conv(K1_IDO.JITU_DT, BUZAI_DATE_S)
    Call UniCode_Conv(K1_IDO.JITU_TM, "")
    NYUKO_QTY = 0
    com = BtOpGetGreaterEqual
    
    Do
        
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        Select Case sts
            Case BtNoErr
            
                If StrConv(IDOREC.JGYOBU, vbUnicode) <> JGYOBU Or _
                    StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    RTrim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> RTrim(Hinban) Or _
                    StrConv(IDOREC.JITU_DT, vbUnicode) > BUZAI_DATE_E Then
                       
                    Exit Do
                    
                End If
            
            Case BtErrEOF
                Exit Do
            
            '   -------------------------------- エラーメッセージ作成
            Case Else
                '重要な要因なので未登録はシステム停止とする
                Call File_Error(sts, BtOpGetEqual, "在庫移動歴", 0)
                Exit Function
        End Select
    
    
    
            
    
        For i = 0 To UBound(IN_TANA_S_OSAKA)
        
            If StrConv(IDOREC.RIRK_ID, vbUnicode) = IN_TANA_S_OSAKA(i) Then
            
                NYUKO_QTY = NYUKO_QTY + (Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)))
                    
            
            End If
        
        Next i
        
        com = BtOpGetNext
    
    Loop
    
    
    
    
    ORDER_ZAN_Proc = False

End Function


Public Function NYUKO_MAEGARI_OSAKA_S_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『大阪ＰＣ　資材検収入庫』
'
'   2012.03.06
'
'-------------------------------------------------------
Dim i               As Integer


Dim Hinban          As String * 20
Dim Tanaban         As String * 8
Dim QTY             As Long

Dim sts             As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2


Dim wkORDER_QTY     As Long
Dim wkNYUKO_QTY     As Long
Dim wkQty           As Long

Dim ST_TANABAN      As String * 11


    NYUKO_MAEGARI_OSAKA_S_Proc = True


    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（品番）

            For i = 0 To M_Gyo - 1
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI, , SHIZAI)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")      '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "品番エラー", "", "")  '2017.09.22
                            
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                NYUKO_MAEGARI_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
                        
                        
                        ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU
                        ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            ST_TANABAN = ""
                        Else
                            ST_TANABAN = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                        
                        
                        
                        '当月仕入予定残集計
                        If ORDER_ZAN_Proc(RET_JGYOBU, NAIGAI_NAI, Hinban, wkORDER_QTY, wkNYUKO_QTY) Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
                        
                        ID_KANRI_TBL(ING_No).ORDER_QTY = wkORDER_QTY
                        ID_KANRI_TBL(ING_No).NYUKO_QTY = wkNYUKO_QTY
                        
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
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                        '>>>>>>>>   2017.09.22
                        'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>>   2017.09.22
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Tanaban & " " & ST_TANABAN)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Tanaban & " " & ST_TANABAN)
                                                                                
                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "16"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "16"
                                                                                
                                                                                
                                                                                
                                                                                
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
                        Send_Text.Box_Type(3).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                        '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                            
                        Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------５行目
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
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                    
                    
                    

                        Sendbuf = Text_Create_Proc()
                        
                        Exit Function
                
                End Select
            Next i
    
    
        Case Step_Sagyo2_RES        '２回目の受信（棚番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    Case LCD_Tanaban        '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        
                        '------------------ 倉庫マスタ読込み
                        Call UniCode_Conv(K0_SOKO.SOKO_NO, Left(Tanaban, 2))
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                    
                            '   -------------------------------- エラーメッセージ作成
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")    '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                NYUKO_MAEGARI_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                Exit Function
                        End Select
                        '------------------ 混載チェック
                        If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                            If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).S_JGYOBU Or _
                                StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")          '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")      '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                NYUKO_MAEGARI_OSAKA_S_Proc = False
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
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")    '2017.09.22
                        
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                NYUKO_MAEGARI_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                Exit Function
                        End Select
                    
                        '------------------ 禁止棚のチェック
                        If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                                
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")            '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")        '2017.09.22
                            
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            Exit Function
                        End If
                        ID_KANRI_TBL(ING_No).Tanaban = Tanaban
    
    
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))

                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
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
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "注残(" & USE_YM & "):" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "注残(" & USE_YM & "):" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
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
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                        '表示内容
                        
                        wkQty = ID_KANRI_TBL(ING_No).ORDER_QTY - ID_KANRI_TBL(ING_No).NYUKO_QTY
                        If wkQty < 0 Then
                            wkQty = 0
                        End If
                        
                        
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Suryo)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Suryo)
                                                                        '数値初期表示
                        Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(wkQty, "#0"))) & Format(wkQty, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(wkQty, "#0"))) & Format(wkQty, "#0")
                                                                        '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                                        '入力桁数
                        Send_Text.Box_Type(4).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
                        
                        NYUKO_MAEGARI_OSAKA_S_Proc = False
                        Exit Function
                End Select
            Next i
        Case Step_Sagyo3_RES        '３回目の受信（数量）
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
            
            
                    Case LCD_Suryo          '数量
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                
                        If QTY <= (ID_KANRI_TBL(ING_No).ORDER_QTY - ID_KANRI_TBL(ING_No).NYUKO_QTY) Then
                        
                        
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
                                                                '入庫更新
                            sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
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
                                    
                                    NYUKO_MAEGARI_OSAKA_S_Proc = False
                                    GoTo Abort_Tran
                                Case SYS_ERR
                                    Sendbuf = Text_Create_Proc()
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    NYUKO_MAEGARI_OSAKA_S_Proc = SYS_ERR    'システム異常発生
                                    
                                    GoTo Abort_Tran
                            End Select
                                        
                            GoTo End_Tran
                        
                                                    
                        Else
                        
                            ID_KANRI_TBL(ING_No).INPUT_QTY = QTY
                        
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
                    
                            Send_Text.Buzzer = Buzzer_DOUBLE                        'ブザー音　二重警告音
                            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DOUBLE
                            
                            '-----------------------------------------------１行目
                                                                                    'BOX属性
'                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                                                                                    '表示内容
'                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
'                                                                                    '数値初期表示
'                            Send_Text.Box_Type(0).INIT = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
'                                                                                    '初期カーソル位置
'                            Send_Text.Box_Type(0).Start_Pos = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
'                                                                                    '入力桁数
'                            Send_Text.Box_Type(0).Max_Size = "00"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
'
'                            Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                            '-----------------------------------------------２行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).Hinban)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "注文残:" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "注文残:" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
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
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "総入庫数:" & Format(ID_KANRI_TBL(ING_No).NYUKO_QTY + QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "総入庫数:" & Format(ID_KANRI_TBL(ING_No).NYUKO_QTY + QTY, "#0"))
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
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "棚入れを行いますか？")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "棚入れを行いますか？")
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
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                    '表示内容
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_CAN_ANS)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_CAN_ANS)
                                                                                    '数値初期表示
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(4).Start_Pos = "20"                  '数値は５桁固定
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "20"
                                                                                    '入力桁数
                             Send_Text.Box_Type(4).Max_Size = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "01"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
                            Sendbuf = Text_Create_Proc()
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            
                            Exit Function
                        End If
                
                End Select
            Next i
    
    
        Case Step_Sagyo4_RES        '４回目の受信（処理継続(Y/N)）
            
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
            
            
                    Case LCD_CAN_ANS          '数量
                        If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) <> "1" And Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) <> "9" Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "1 OR 9 を入力", "して下さい。", "")          '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "1 OR 9 を入力", "して下さい。", "")      '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                
                
                        If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) = "1" Then
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
                                                                '入庫更新
                            sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                                    ID_KANRI_TBL(ING_No).Hinban, _
                                                    Format(Now, "YYYYMMDD"), _
                                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                    0, _
                                                    ID_KANRI_TBL(ING_No).INPUT_QTY, _
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
                                    
                                    NYUKO_MAEGARI_OSAKA_S_Proc = False
                                    GoTo Abort_Tran
                                Case SYS_ERR
                                    Sendbuf = Text_Create_Proc()
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    NYUKO_MAEGARI_OSAKA_S_Proc = SYS_ERR    'システム異常発生
                                    
                                    GoTo Abort_Tran
                            End Select
                                        
                            GoTo End_Tran
                        Else
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
                            
                            
                            
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            
                            Exit Function
                        
                        End If
            
                End Select
            Next i
    End Select

End_Tran:
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
    
    
    
    NYUKO_MAEGARI_OSAKA_S_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


