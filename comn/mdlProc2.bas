Attribute VB_Name = "mdlProc2"
Option Explicit

'[2014/02/10 - M.MATSUYAMA 移動(Ver2.0.0)] F1100101から移動

Public Function Tanto_Check_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『担当者コードのチェック』
'
'-------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Tanto_Check_Proc = True

    For i = 0 To M_Gyo
        
        If ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = TYPE_REF Then
        Else
                                '担当者マスタ読み込み
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, ID_KANRI_TBL(ING_No).Recv_text(i))
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    ID_KANRI_TBL(ING_No).TANTO_CODE = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                    Exit For
                Case BtErrKeyNotFound
                    
                    '   -------------------------------- エラーメッセージ作成
                    Call Err_Send_Proc("担当者未登録", "", "", "", "")
                    
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                    
                    Tanto_Check_Proc = False
                    Exit Function
                Case Else
                    Sendbuf = Text_Create_Proc()
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ", 0)
                    Exit Function
            End Select
        
        End If
    
    Next i

    If i > M_Gyo Then                       '実際はありえない（担当者が未入力）
        ID_KANRI_TBL(ING_No).Step = Step_Start
        '   -------------------------------- エラーメッセージ作成
        Call Err_Send_Proc("担当者未登録", "", "", "", "")
                    
        Sendbuf = Text_Create_Proc()
        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                    
        Tanto_Check_Proc = False
        Exit Function
    End If

'----------------------------------------------- '専用メニュー獲得
    If Menu_Type = 1 Then
                        '共通メニュー
    Else
        Call UniCode_Conv(K0_P_TMENU.TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
        sts = BTRV(BtOpGetEqual, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
        Select Case sts
            Case BtNoErr
'                ID_KANRI_TBL(ING_No).MENU_GRP = StrConv(P_TMENUREC.TANTO_CODE, vbUnicode)
            Case BtErrKeyNotFound
                    
'                ID_KANRI_TBL(ING_No).MENU_GRP = ""
                '   -------------------------------- エラーメッセージ作成
                Call Err_Send_Proc("担当者メニュー", "未登録", "", "", "")
                    
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
            
                Tanto_Check_Proc = False
                Exit Function
            Case Else
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "担当者メニュー", 0)
                Exit Function
        End Select
            
    
    
    End If
'----------------------------------------------- 'メニュー情報＆作業情報の初期化
    
    If UBound(JGYOBU_T) = 0 Then
                                                '１事業部固定
        ID_KANRI_TBL(ING_No).JGYOBU = JGYOBU_T(0).CODE
    Else
        ID_KANRI_TBL(ING_No).JGYOBU = ""
    End If
    
    If UBound(NAIGAI) = 0 Then
                                                '国内外固定
        ID_KANRI_TBL(ING_No).NAIGAI = NAIGAI(0).CODE
    Else
        ID_KANRI_TBL(ING_No).NAIGAI = ""
    End If
    
    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'    ID_KANRI_TBL(ING_No).MENU_LV3 = ""

    ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
    ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'    ID_KANRI_TBL(ING_No).PageNo_LV3 = 0

    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ""
    ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = ""
    ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = ""

'---------------------------------------------- 'メニュー送信
    If Menu_Send_Proc(Sendbuf) Then
        Exit Function
    End If

    Tanto_Check_Proc = False

End Function


'[2016/05/14 -  mdlProcから移動

Public Function Menu_Send_Proc(Optional Sendbuf As String) As Integer

'-------------------------------------------------------
'
'   『メニューテキスト作成』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Menu_Tbl()  As Menu_Tbl_tag
Dim Menu_Cnt    As Integer
Dim Max_Page    As Integer


Dim PageNo      As Integer

Dim Gyo_Suu     As Integer
Dim Start_Gyo   As Integer
Dim End_Gyo     As Integer


Dim WK_LV1      As String * 3
Dim WK_LV2      As String * 3



Dim wkHex       As String   '2017.09.07


    Menu_Send_Proc = True
'----------------------------------------------- '事業部選択あり
    If Trim(ID_KANRI_TBL(ING_No).JGYOBU) = "" Then
        Call JGYOBU_MENU_SET

        Sendbuf = Text_Create_Proc()


        Menu_Send_Proc = False
        Exit Function
    End If
'----------------------------------------------- '国内外選択あり
    If ID_KANRI_TBL(ING_No).NAIGAI = " " Then

        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        ID_KANRI_TBL(ING_No).MENU_LV2 = ""

        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0


        Call NAIGAI_MENU_SET
        Sendbuf = Text_Create_Proc
        Menu_Send_Proc = False
        Exit Function
    Else
        '2010.03.30
        If Trim(ID_KANRI_TBL(ING_No).MENU_LV1) = "" Then
            If ID_KANRI_TBL(ING_No).Step < Step_MENU1_REQ Then
                ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
            End If
'            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
        
'            ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'            ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
        
'            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ""
'            ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = ""
'            ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = ""
        End If
        '2010.03.30
    
    
    End If
    '   -------------------------------- レベル１　トップメニューの管理
    If Trim(ID_KANRI_TBL(ING_No).MENU_LV1) = "" Then
'        ST_LOG_OUT_F = True '2008.08.08
        
        
        MENU_UP_F = False   '2008.08.08
        
        'ﾒﾆｭｰｸﾞﾙｰﾌﾟ
        Call UniCode_Conv(K0_P_TMENU.TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
        Erase Menu_Tbl

        sts = BTRV(BtOpGetEqual, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
        Select Case sts
            
            Case BtNoErr
            Case BtErrKeyNotFound
            
                        
                Call Err_Send_Proc("メニュー未登録", "", "", "", "")
                Sendbuf = Text_Create_Proc()
'                If UBound(NAIGAI) = 0 Then
'                    ID_KANRI_TBL(ING_No).Step = Step_Start
'                Else
'                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'                End If
                
                
                
                
                
                
                
                If UBound(NAIGAI) = 0 Then
                    
                    
                    
                    If UBound(JGYOBU_T) = 0 Then
                    
                    
                        ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                
    '                    Call Menu_Send_Proc(Sendbuf)
                
                
                    Else
                
                        '事業部要求でループする
                        ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
                        ID_KANRI_TBL(ING_No).JGYOBU = ""
                        
                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
            
    '                    Call Menu_Send_Proc(Sendbuf)
                    End If
                
                Else
                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                    ID_KANRI_TBL(ING_No).NAIGAI = ""
    '                Call Menu_Send_Proc(Sendbuf)
                End If
                
                Menu_Send_Proc = False
                
                Exit Function
            
            
            Case Else

                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, com, "担当者別ﾒﾆｭｰ", 0)
                Exit Function
        End Select


        Menu_Cnt = -1
        For i = 0 To 179     '29--->179 2006.10.11
            If Trim(StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)) = "" Then
                Exit For
            End If
        
            If StrConv(P_TMENUREC.MENU_T(i).JGYOBU, vbUnicode) = ID_KANRI_TBL(ING_No).JGYOBU And _
                StrConv(P_TMENUREC.MENU_T(i).NAIGAI, vbUnicode) = ID_KANRI_TBL(ING_No).NAIGAI Then
        
                Menu_Cnt = Menu_Cnt + 1
                ReDim Preserve Menu_Tbl(Menu_Cnt)
            
                Menu_Tbl(Menu_Cnt).MENU_NO = StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)
                
                Call UniCode_Conv(K0_P_MENU.JGYOBU, StrConv(P_TMENUREC.MENU_T(i).JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_MENU.NAIGAI, StrConv(P_TMENUREC.MENU_T(i).NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_MENU.MENU_NO, StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode))
                
                sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
                Select Case sts
                    
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                                
                        Call Err_Send_Proc("メニュー異常", "", "", "", "")
'                        Sendbuf = Text_Create_Proc()
'                        If UBound(NAIGAI) = 0 Then
'                            ID_KANRI_TBL(ING_No).Step = Step_Start
'                        Else
'                            ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'                        End If
'                        Menu_Send_Proc = False
'                        Exit Function
                    
                    
                        If UBound(NAIGAI) = 0 Then
                            
                            
                            
                            If UBound(JGYOBU_T) = 0 Then
                            
                            
                                ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                        
            '                    Call Menu_Send_Proc(Sendbuf)
                        
                        
                            Else
                        
                                '事業部要求でループする
                                ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
                                ID_KANRI_TBL(ING_No).JGYOBU = ""
                                
                                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                                ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                    
            '                    Call Menu_Send_Proc(Sendbuf)
                            End If
                        
                        Else
                            ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                            ID_KANRI_TBL(ING_No).NAIGAI = ""
            '                Call Menu_Send_Proc(Sendbuf)
                        End If
                        
                        Menu_Send_Proc = False
                        
                        Exit Function
                    
                    
                    
                    
                    Case Else
        
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, com, "メニュー管理マスタ", 0)
                        Exit Function
                End Select
                
                
                
                Menu_Tbl(Menu_Cnt).MENU_NO = StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)
                Menu_Tbl(Menu_Cnt).Disp = StrConv(P_MENUREC.MENU_DSP, vbUnicode)
        
        
            End If
        
        Next i


        If Menu_Cnt = -1 Then
        '   -------------------------------- エラーメッセージ作成
            Call Err_Send_Proc("メニュー未登録", "", "", "", "")
            Sendbuf = Text_Create_Proc()
'            If UBound(NAIGAI) = 0 Then
'                ID_KANRI_TBL(ING_No).Step = Step_Start
'            Else
'                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'            End If
            
            
            
            
            
            If UBound(NAIGAI) = 0 Then
                
                
                
                If UBound(JGYOBU_T) = 0 Then
                
                
                    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
            
'                    Call Menu_Send_Proc(Sendbuf)
            
            
                Else
            
                    '事業部要求でループする
                    ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
                    ID_KANRI_TBL(ING_No).JGYOBU = ""
                    
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
        
'                    Call Menu_Send_Proc(Sendbuf)
                End If
            
            Else
                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                ID_KANRI_TBL(ING_No).NAIGAI = ""
'                Call Menu_Send_Proc(Sendbuf)
            End If
            
'''''            Menu_Send_Proc = False
            
            Exit Function
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        End If


        Max_Page = Int(CDbl((Menu_Cnt + 1) / M_Gyo + 0.9))
        PageNo = ID_KANRI_TBL(ING_No).PageNo_LV1

        Start_Gyo = PageNo * M_Gyo
        End_Gyo = (PageNo * M_Gyo) + (M_Gyo - 1)


        Send_Text.sts = Sts_OK                                      'ステータス　OK
        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK

        Send_Text.Display_Flg = Display_MENU                        '表示画面フラグ メニュー画面
        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
        If Max_Page = 1 Then
            Send_Text.End_Menu = Menu_Only          '１画面のみ
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
        Else
            If (Max_Page - 1) = PageNo Then
                Send_Text.End_Menu = MENU_END       '最終ページ
                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = MENU_END
            Else
                If PageNo = 0 Then
                    Send_Text.End_Menu = Menu_Head  '先頭ページ
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Head
                Else
                    Send_Text.End_Menu = Menu_Mid   '途中ページ
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Mid
                End If
            End If
        End If
        Send_Text.FileName = ""                                         '送信データファイル名
        ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
        Send_Text.buzzer = Buzzer_DEF                                   'ブザー音　標準
        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
'---------------------------------------------------------------
        Gyo_Suu = 0
        j = -1
        For i = Start_Gyo To End_Gyo
            j = j + 1
            If i > UBound(Menu_Tbl) Then
                Send_Text.Box_Type(j).Box_Type = ""                 'BOX属性
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = ""
                Call UniCode_Conv(Send_Text.Box_Type(j).LCD, "")    '表示内容
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, "")
                Send_Text.Box_Type(j).INIT = ""                     '数値初期値
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                Send_Text.Box_Type(j).MENU = ""                     'メニュ―番号
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = ""
            Else
                Gyo_Suu = Gyo_Suu + 1
                Send_Text.Box_Type(j).Box_Type = TYPE_MENU          'BOX属性
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = TYPE_MENU
                Call UniCode_Conv(Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                Send_Text.Box_Type(j).INIT = ""                     '数値初期値
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO
                
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO
            End If
        Next i
        
        Send_Text.Menu_Suu = Format(Gyo_Suu, "00")      'メニュー項目数
        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(Gyo_Suu, "00")
        ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信
        Sendbuf = Text_Create_Proc()
        
    Else
        '   -------------------------------- レベル２　作業メニューの管理
        If Trim(ID_KANRI_TBL(ING_No).MENU_LV2) = "" Then
            
            
'2008.08.12            If ST_LOG_OUT_F Then            '2008.08.08

'                ST_LOG_OUT_F = False        '2008.08.08

                
'                If Not MENU_UP_F Then       '2008.08.08
'
'                    If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
'                                                Format(ID_KANRI_TBL(ING_No).ID, "000"), _
'                                                ID_KANRI_TBL(ING_No).JGYOBU, _
'                                                ID_KANRI_TBL(ING_No).NAIGAI, _
'                                                ID_KANRI_TBL(ING_No).MENU_LV1, _
'                                                "ST", , , , , , , , , FILE_RETRY) Then
'
'                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
'                        Sendbuf = Text_Create_Proc()
'                        Exit Function
'
'                    End If
'
'                End If


 '2008.08.12           End If
            
            
            
            
            
            
            
            
            
            'ﾒﾆｭｰｸﾞﾙｰﾌﾟ
            Call UniCode_Conv(K0_P_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
            Call UniCode_Conv(K0_P_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
            Call UniCode_Conv(K0_P_MENU.MENU_NO, ID_KANRI_TBL(ING_No).MENU_LV1)
            
            
            Erase Menu_Tbl
    
            sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
            Select Case sts
                
                Case BtNoErr
                Case BtErrKeyNotFound
                
                            
                    Call Err_Send_Proc("メニュー未登録", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    If UBound(NAIGAI) = 0 Then
                        ID_KANRI_TBL(ING_No).Step = Step_Start
                    Else
                        ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                    End If
                    Menu_Send_Proc = False
                    Exit Function
                
                
                Case Else
    
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, com, "担当者別ﾒﾆｭｰ", 0)
                    Exit Function
            End Select
    
    
            Menu_Cnt = -1
            For i = 0 To 19
                If Trim(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)) = "" Then
                    Exit For
                End If
            
            
                Menu_Cnt = Menu_Cnt + 1
                ReDim Preserve Menu_Tbl(Menu_Cnt)
                
                    
                    
                Menu_Tbl(Menu_Cnt).MENU_NO = StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)
                Menu_Tbl(Menu_Cnt).PARAM = StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode)
                Menu_Tbl(Menu_Cnt).Disp = StrConv(P_MENUREC.SAGYO(i).Disp, vbUnicode)
            
                Menu_Tbl(Menu_Cnt).LOG_OUT = StrConv(P_MENUREC.SAGYO(i).LOG_OUT, vbUnicode)
            
            Next i
    
    
            If Menu_Cnt = -1 Then
            '   -------------------------------- エラーメッセージ作成
                Call Err_Send_Proc("メニュー未登録", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                If UBound(NAIGAI) = 0 Then
                    ID_KANRI_TBL(ING_No).Step = Step_Start
                Else
                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                End If
                Menu_Send_Proc = False
                Exit Function
            End If
    
    
            Max_Page = Int(CDbl((Menu_Cnt + 1) / M_Gyo + 0.9))
            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV2
    
            Start_Gyo = PageNo * M_Gyo
            End_Gyo = (PageNo * M_Gyo) + (M_Gyo - 1)
    
    
            Send_Text.sts = Sts_OK                                      'ステータス　OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_MENU                        '表示画面フラグ メニュー画面
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
            If Max_Page = 1 Then
                Send_Text.End_Menu = Menu_Only          '１画面のみ
                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
            Else
                If (Max_Page - 1) = PageNo Then
                    Send_Text.End_Menu = MENU_END       '最終ページ
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = MENU_END
                Else
                    If PageNo = 0 Then
                        Send_Text.End_Menu = Menu_Head  '先頭ページ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Head
                    Else
                        Send_Text.End_Menu = Menu_Mid   '途中ページ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Mid
                    End If
                End If
            End If
            Send_Text.FileName = ""                                         '送信データファイル名
            ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
            Send_Text.buzzer = Buzzer_DEF                                   'ブザー音　標準
            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
    '---------------------------------------------------------------
            Gyo_Suu = 0
            j = -1
            For i = Start_Gyo To End_Gyo
                j = j + 1
                If i > UBound(Menu_Tbl) Then
                    Send_Text.Box_Type(j).Box_Type = ""                 'BOX属性
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = ""
                    Call UniCode_Conv(Send_Text.Box_Type(j).LCD, "")    '表示内容
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, "")
                    Send_Text.Box_Type(j).INIT = ""                     '数値初期値
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                    Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                    Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                    Send_Text.Box_Type(j).MENU = ""                     'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = ""
                
                    Send_Text.Box_Type(j).MENU18 = ""                     'メニュ―番号 2017.09.07
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU18 = ""              '2017.09.07
                
                
                Else
                    Gyo_Suu = Gyo_Suu + 1
                    Send_Text.Box_Type(j).Box_Type = TYPE_MENU          'BOX属性
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = TYPE_MENU
                    Call UniCode_Conv(Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                    Send_Text.Box_Type(j).INIT = ""                     '数値初期値
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                    Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                    Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                                                                        
                                                                        
'>>>>>>>>>>>>>>>>>>>    2017.09.07
                                                                        'メニュ―番号 & ﾊﾟﾗﾒｰﾀ
'                     Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM                              '2017.09.07
'                     ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM         '2017.09.07
                
                     
                     
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MTS_FLG = ""         '2017.10.27
                     
                    wkHex = f10sinTo16sin(Menu_Tbl(i).PARAM)
                     
                     
                     
                    If Trim(wkHex) = "" Then
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MTS_FLG = "1"    '2017.10.27
                    
                        wkHex = Menu_Tbl(i).PARAM
                    End If
                    Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & wkHex                                           '2017.09.07
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & wkHex                      '2017.09.07

                    
'>>>>>>>>>>>>>>>>>>>    2017.09.07
                
                
                    Send_Text.Box_Type(j).MENU18 = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM                          'メニュ―番号 2017.09.07
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU18 = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM     '2017.09.07
                
                End If
            Next i
            
            Send_Text.Menu_Suu = Format(Gyo_Suu, "00")      'メニュー項目数
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(Gyo_Suu, "00")
            ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信
            Sendbuf = Text_Create_Proc()
            
        
        End If
    End If
    
    Menu_Send_Proc = False




End Function

'[2016/05/14 -  mdlProcから移動

Public Function Menu_Recv_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『２階層以上のメニュー送信』
'
'-------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

    
Dim MTS     As String * 8
Dim SS      As String * 8
    
    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_MENU1_RES
    '   -------------------------------- 次ﾚﾍﾞﾙへ
            If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> BEF_Page And Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> NEXT_Page Then

                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
            End If
            
            If Menu_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
        Case Else
    
            If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) = BEF_Page Or Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) = NEXT_Page Then
                If Menu_Send_Proc() Then
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                End If
            Else
    
    '   -------------------------------- メニュー管理マスタ読込み
                Call UniCode_Conv(K0_P_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                Call UniCode_Conv(K0_P_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                Call UniCode_Conv(K0_P_MENU.MENU_NO, ID_KANRI_TBL(ING_No).MENU_LV1)
        
                sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
                Select Case sts
                    Case BtNoErr
    
                        For i = 0 To 19
                        
                            If Trim(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)) = Trim(ID_KANRI_TBL(ING_No).MENU_LV2) And _
                               ((StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode) = (ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)) Or _
                                (Trim(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode)) = (Trim(Format(ID_KANRI_TBL(ING_No).MTS_CODE, "#0") & ID_KANRI_TBL(ING_No).SS_CODE)))) Then
                                                                
                                '>>>>>>>>>  2017.11.30
                                If (Trim(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode)) = (Trim(Format(ID_KANRI_TBL(ING_No).MTS_CODE, "#0") & ID_KANRI_TBL(ING_No).SS_CODE))) Then
                                    ID_KANRI_TBL(ING_No).MTS_CODE = Trim(Format(ID_KANRI_TBL(ING_No).MTS_CODE, "#0"))
                                End If
                                '>>>>>>>>>  2017.11.30
                                
                                Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode), 1))
                                Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode), 1))
                                
                                sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                                Select Case sts
                                
                                    Case BtNoErr
                                        'ｽｷｬﾅ表示名称
                                        ID_KANRI_TBL(ING_No).YOIN_DNAME = StrConv(P_MENUREC.SAGYO(i).Disp, vbUnicode)
                                
                                
                                        '2010.09.15
                                        ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME = StrConv(P_MENUREC.SAGYO(i).Disp, vbUnicode)
                                
                                        
                                        ID_KANRI_TBL(ING_No).SAGYO_LOG = StrConv(P_MENUREC.SAGYO(i).LOG_OUT, vbUnicode)
                                                                    
                                        If StrConv(YOINREC.PARAM_F, vbUnicode) = "1" Then   '向け先なら（出荷）
                                            ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(YOINREC.YOIN_CODE, vbUnicode)
'2006.01.30                                            ID_KANRI_TBL(ING_No).MTS_CODE = Left(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode), 8)
'2006.01.30                                            ID_KANRI_TBL(ING_No).SS_CODE = Right(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode), 8)
                                        End If
                                                                                            '検品（向け先指定）なら
                                        If StrConv(YOINREC.CODE_TYPE, vbUnicode) = ACT_KENPIN_MTS Then
                                            ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                        End If
                                        
                                                                                            '検品（直送指定）なら   2016.10.14
                                        If StrConv(YOINREC.CODE_TYPE, vbUnicode) = ACT_KENPIN_Drct Then
                                            ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                        End If
                                        
                                        
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = StrConv(YOINREC.CODE_TYPE, vbUnicode)
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = StrConv(YOINREC.SOKO_NO, vbUnicode)
                                
                                    Case BtErrKeyNotFound
    
                                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                                        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                                        '   -------------------------------- エラーメッセージ作成
                                        Call Err_Send_Proc("要因マスタ", "未登録", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                                        Menu_Recv_Proc = False
                                        Exit Function
                                    Case Else
                                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Call File_Error(sts, BtOpGetEqual, "メニュー管理", 0)
                                        Exit Function
                                End Select
                                
                                Exit For
                            End If
                        
                        Next i
    
                        If i > 19 Then
                
    
                            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                            '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc("メニュー管理マスタ", "設定ミス", Trim(ID_KANRI_TBL(ING_No).MENU_LV1), "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                            Menu_Recv_Proc = False
                            Exit Function
                        
                        End If
                
                        If Sagyo_Send_Proc() Then
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
        
        
                    Case BtErrKeyNotFound
        
                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc("メニュー管理マスタ", "未登録", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                        Menu_Recv_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "メニュー管理", 0)
                        Exit Function
        
                End Select
            End If
        End Select
        Sendbuf = Text_Create_Proc()
    
    Menu_Recv_Proc = False

End Function
'2006.01.30Private Function Menu_Recv_Proc(Sendbuf As String) As Integer
'2006.01.30'-------------------------------------------------------
'2006.01.30'
'2006.01.30'   『２階層以上のメニュー送信』
'2006.01.30'
'2006.01.30'-------------------------------------------------------
'2006.01.30Dim sts As Integer
'2006.01.30
'2006.01.30    Menu_Recv_Proc = True
'2006.01.30                                        'メニュ管理マスタの読み込み
'2006.01.30    Call UniCode_Conv(K1_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'2006.01.30    Call UniCode_Conv(K1_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'2006.01.30    Call UniCode_Conv(K1_MENU.MENU_GRP_NO, ID_KANRI_TBL(ING_No).MENU_GRP)
'2006.01.30
'2006.01.30    If ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ Then
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV1, "000")
'2006.01.30    Else
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV1, ID_KANRI_TBL(ING_No).MENU_LV1)
'2006.01.30    End If
'2006.01.30
'2006.01.30    Call UniCode_Conv(K1_MENU.MENU_LV2, ID_KANRI_TBL(ING_No).MENU_LV2)
'2006.01.30    Call UniCode_Conv(K1_MENU.MENU_LV3, ID_KANRI_TBL(ING_No).MENU_LV3)
'2006.01.30    sts = BTRV(BtOpGetEqual, MENU_POS, MENUREC, Len(MENUREC), K1_MENU, Len(K1_MENU), 1)
'2006.01.30    Select Case sts
'2006.01.30        Case BtNoErr
'2006.01.30        Case BtErrKeyNotFound
'2006.01.30
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_GRP = ""
'2006.01.30
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30            '   -------------------------------- エラーメッセージ作成
'2006.01.30            Call Err_Send_Proc("メニュー管理", "未登録", "", "", "")
'2006.01.30
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'2006.01.30
'2006.01.30            Menu_Recv_Proc = False
'2006.01.30            Exit Function
'2006.01.30        Case Else
'2006.01.30            Call Err_Send_Proc("システム異常発生", "", "", "", "")
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            Call File_Error(sts, BtOpGetEqual, "メニュー管理", 0)
'2006.01.30            Exit Function
'2006.01.30    End Select
'2006.01.30
'2006.01.30    If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> BEF_Page And Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> NEXT_Page And _
'2006.01.30     StrConv(MENUREC.MENU_KBN, vbUnicode) = "1" Then
'2006.01.30
'2006.01.30
'2006.01.30                                            '要因の読込み
'2006.01.30        Call UniCode_Conv(K0_YOIN.CODE_TYPE, StrConv(MENUREC.CODE_TYPE, vbUnicode))
'2006.01.30        Call UniCode_Conv(K0_YOIN.YOIN_CODE, StrConv(MENUREC.YOIN_CODE, vbUnicode))
'2006.01.30        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
'2006.01.30        Select Case sts
'2006.01.30            Case BtNoErr
'2006.01.30                ID_KANRI_TBL(ING_No).YOIN_DNAME = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
'2006.01.30
'2006.01.30                If StrConv(YOINREC.PARAM_F, vbUnicode) = "1" Then   '向け先なら（出荷）
'2006.01.30                    ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(MENUREC.YOIN_CODE, vbUnicode)
'2006.01.30                    ID_KANRI_TBL(ING_No).MTS_CODE = Left(StrConv(MENUREC.PARAM, vbUnicode), 8)
'2006.01.30                    ID_KANRI_TBL(ING_No).SS_CODE = Right(StrConv(MENUREC.PARAM, vbUnicode), 8)
'2006.01.30                End If
'2006.01.30                                                                    '検品（向け先指定）なら
'2006.01.30                If StrConv(YOINREC.CODE_TYPE, vbUnicode) = ACT_KENPIN_MTS Then
'2006.01.30                    ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(MENUREC.YOIN_CODE, vbUnicode)
'2006.01.30                End If
'2006.01.30
'2006.01.30                ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = StrConv(YOINREC.CODE_TYPE, vbUnicode)
'2006.01.30                ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = StrConv(YOINREC.YOIN_CODE, vbUnicode)
'2006.01.30                ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = StrConv(MENUREC.PARAM, vbUnicode)
'2006.01.30
'2006.01.30            Case BtErrKeyNotFound
'2006.01.30
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_GRP = ""
'2006.01.30
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30                '   -------------------------------- エラーメッセージ作成
'2006.01.30                Call Err_Send_Proc("要因マスタ", "未登録", "", "", "")
'2006.01.30
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'2006.01.30
'2006.01.30                Menu_Recv_Proc = False
'2006.01.30                Exit Function
'2006.01.30            Case Else
'2006.01.30                Call Err_Send_Proc("システム異常発生", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                Call File_Error(sts, BtOpGetEqual, "メニュー管理", 0)
'2006.01.30                Exit Function
'2006.01.30        End Select
'2006.01.30
'2006.01.30        If Sagyo_Send_Proc() Then
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            Exit Function
'2006.01.30        End If
'2006.01.30    Else
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
'2006.01.30
'2006.01.30        If Menu_Send_Proc() Then
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            Exit Function
'2006.01.30        End If
'2006.01.30    End If
'2006.01.30
'2006.01.30    Sendbuf = Text_Create_Proc()
'2006.01.30
'2006.01.30
'2006.01.30    Menu_Recv_Proc = False
'2006.01.30
'2006.01.30End Function


'2006.01.30Private Function Menu_Send_Proc(Optional Sendbuf As String) As Integer
'2006.01.30'-------------------------------------------------------
'2006.01.30'
'2006.01.30'   『メニューテキスト作成』
'2006.01.30'
'2006.01.30'-------------------------------------------------------
'2006.01.30Dim sts         As Integer
'2006.01.30Dim com         As Integer
'2006.01.30
'2006.01.30Dim i           As Integer
'2006.01.30Dim j           As Integer
'2006.01.30
'2006.01.30Dim Menu_Tbl()  As Menu_Tbl_tag
'2006.01.30Dim Menu_Cnt    As Integer
'2006.01.30Dim Max_Page    As Integer
'2006.01.30
'2006.01.30
'2006.01.30Dim PageNo      As Integer
'2006.01.30
'2006.01.30Dim Gyo_Suu     As Integer
'2006.01.30Dim Start_Gyo   As Integer
'2006.01.30Dim End_Gyo     As Integer
'2006.01.30
'2006.01.30
'2006.01.30Dim WK_LV1      As String * 3
'2006.01.30Dim WK_LV2      As String * 3
'2006.01.30Dim WK_LV3      As String * 3
'2006.01.30
'2006.01.30
'2006.01.30    Menu_Send_Proc = True
'2006.01.30'----------------------------------------------- '事業部選択あり
'2006.01.30    If ID_KANRI_TBL(ING_No).JGYOBU = " " Then
'2006.01.30        Call JGYOBU_MENU_SET
'2006.01.30
'2006.01.30        Sendbuf = Text_Create_Proc()
'2006.01.30
'2006.01.30
'2006.01.30        Menu_Send_Proc = False
'2006.01.30        Exit Function
'2006.01.30    End If
'2006.01.30'----------------------------------------------- '国内外選択あり
'2006.01.30    If ID_KANRI_TBL(ING_No).NAIGAI = " " Then
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_GRP = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30        Call NAIGAI_MENU_SET
'2006.01.30        Sendbuf = Text_Create_Proc
'2006.01.30        Menu_Send_Proc = False
'2006.01.30        Exit Function
'2006.01.30    End If
'2006.01.30'----------------------------------------------- 'メニュー管理の読み込み
'2006.01.30    If Len(Trim(ID_KANRI_TBL(ING_No).MENU_GRP)) = 0 Then
'2006.01.30                                    'ここで未確定なのは共通メニュー時だけ
'2006.01.30        Call UniCode_Conv(K1_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'2006.01.30        Call UniCode_Conv(K1_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_GRP_NO, ALL_MENU_GRP)
'2006.01.30
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV1, "")
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV2, "")
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV3, "")
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30        sts = BTRV(BtOpGetGreaterEqual, MENU_POS, MENUREC, Len(MENUREC), K1_MENU, Len(K1_MENU), 1)
'2006.01.30        Select Case sts
'2006.01.30            Case BtNoErr
'2006.01.30            Case BtErrEOF
'2006.01.30
'2006.01.30            '   -------------------------------- エラーメッセージ作成
'2006.01.30                Call Err_Send_Proc("メニュー未登録", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                If UBound(NAIGAI) = 0 Then
'2006.01.30                    ID_KANRI_TBL(ING_No).Step = Step_Start
'2006.01.30                Else
'2006.01.30                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'2006.01.30                End If
'2006.01.30                Menu_Send_Proc = False
'2006.01.30                Exit Function
'2006.01.30            Case Else
'2006.01.30                Call Err_Send_Proc("システム異常発生", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ", 0)
'2006.01.30                Exit Function
'2006.01.30        End Select
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_GRP = StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
'2006.01.30
'2006.01.30
'2006.01.30    End If
'2006.01.30    '   -------------------------------- メニュー管理マスタ読込み
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_GRP_NO, ID_KANRI_TBL(ING_No).MENU_GRP)
'2006.01.30    Call UniCode_Conv(K0_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'2006.01.30    Call UniCode_Conv(K0_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_LV1, ID_KANRI_TBL(ING_No).MENU_LV1)
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_LV2, ID_KANRI_TBL(ING_No).MENU_LV2)
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_LV3, ID_KANRI_TBL(ING_No).MENU_LV3)
'2006.01.30
'2006.01.30    Erase Menu_Tbl
'2006.01.30
'2006.01.30    com = BtOpGetGreater
'2006.01.30
'2006.01.30    Menu_Cnt = -1
'2006.01.30    Do
'2006.01.30        sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
'2006.01.30        Select Case sts
'2006.01.30            Case BtNoErr
'2006.01.30                If ID_KANRI_TBL(ING_No).MENU_GRP <> StrConv(MENUREC.MENU_GRP_NO, vbUnicode) Or _
'2006.01.30                    ID_KANRI_TBL(ING_No).JGYOBU <> StrConv(MENUREC.JGYOBU, vbUnicode) Or _
'2006.01.30                    ID_KANRI_TBL(ING_No).NAIGAI <> StrConv(MENUREC.NAIGAI, vbUnicode) Then
'2006.01.30                    Exit Do
'2006.01.30                End If
'2006.01.30
'2006.01.30                WK_LV1 = ID_KANRI_TBL(ING_No).MENU_LV1
'2006.01.30                WK_LV2 = ID_KANRI_TBL(ING_No).MENU_LV2
'2006.01.30                WK_LV3 = ID_KANRI_TBL(ING_No).MENU_LV3
'2006.01.30
'2006.01.30
'2006.01.30                If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV1)) = 0 Then
'2006.01.30                    WK_LV1 = StrConv(MENUREC.MENU_LV1, vbUnicode)
'2006.01.30                    WK_LV2 = ""
'2006.01.30                    WK_LV3 = ""
'2006.01.30
'2006.01.30
'2006.01.30'                    ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30                    ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30                    ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30                Else
'2006.01.30                    If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV2)) = 0 Then
'2006.01.30                        WK_LV2 = StrConv(MENUREC.MENU_LV2, vbUnicode)
'2006.01.30                        WK_LV3 = ""
'2006.01.30
'2006.01.30'                        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30                        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30                    Else
'2006.01.30                        If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV3)) = 0 Then
'2006.01.30                            WK_LV3 = StrConv(MENUREC.MENU_LV3, vbUnicode)
'2006.01.30
'2006.01.30'                            ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30                        End If
'2006.01.30                    End If
'2006.01.30                End If
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30            Case BtErrEOF
'2006.01.30                Exit Do
'2006.01.30            Case Else
'2006.01.30
'2006.01.30                Call Err_Send_Proc("システム異常発生", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                Call File_Error(sts, com, "メニュー管理マスタ", 0)
'2006.01.30                Exit Function
'2006.01.30        End Select
'2006.01.30
'2006.01.30
'2006.01.30    '   -------------------------------- メニュー情報を保存
'2006.01.30        If WK_LV1 <> StrConv(MENUREC.MENU_LV1, vbUnicode) Or _
'2006.01.30            WK_LV2 <> StrConv(MENUREC.MENU_LV2, vbUnicode) Or _
'2006.01.30            WK_LV3 <> StrConv(MENUREC.MENU_LV3, vbUnicode) Then
'2006.01.30        Else
'2006.01.30
'2006.01.30
'2006.01.30            Menu_Cnt = Menu_Cnt + 1
'2006.01.30
'2006.01.30            ReDim Preserve Menu_Tbl(Menu_Cnt)
'2006.01.30
'2006.01.30            If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV1)) = 0 Then
'2006.01.30                Menu_Tbl(Menu_Cnt).CODE = StrConv(MENUREC.MENU_LV1, vbUnicode)
'2006.01.30            Else
'2006.01.30                If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV2)) = 0 Then
'2006.01.30                    Menu_Tbl(Menu_Cnt).CODE = StrConv(MENUREC.MENU_LV2, vbUnicode)
'2006.01.30                Else
'2006.01.30                    If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV3)) = 0 Then
'2006.01.30                        Menu_Tbl(Menu_Cnt).CODE = StrConv(MENUREC.MENU_LV3, vbUnicode)
'2006.01.30                    End If
'2006.01.30                End If
'2006.01.30            End If
'2006.01.30
'2006.01.30
'2006.01.30            Menu_Tbl(Menu_Cnt).Disp = StrConv(MENUREC.DISPLAY_ITEM, vbUnicode)
'2006.01.30        End If
'2006.01.30
'2006.01.30        com = BtOpGetNext
'2006.01.30
'2006.01.30    Loop
'2006.01.30
'2006.01.30    If Menu_Cnt = -1 Then
'2006.01.30            '   -------------------------------- エラーメッセージ作成
'2006.01.30        Call Err_Send_Proc("メニュー未登録", "", "", "", "")
'2006.01.30        Sendbuf = Text_Create_Proc()
'2006.01.30        If UBound(NAIGAI) = 0 Then
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_Start
'2006.01.30        Else
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'2006.01.30        End If
'2006.01.30        Menu_Send_Proc = False
'2006.01.30        Exit Function
'2006.01.30    End If
'2006.01.30
'2006.01.30'----------------------------------------------- 'メニュー送信テキスト作成
'2006.01.30'''    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1   'メニュー送信
'2006.01.30    '---------------------------------------------------------------
'2006.01.30    Max_Page = Int(CDbl((Menu_Cnt + 1) / M_Gyo + 0.9))
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30'    Start_Gyo = ID_KANRI_TBL(ING_No).PageNo * M_Gyo
'2006.01.30'    End_Gyo = (ID_KANRI_TBL(ING_No).PageNo * M_Gyo) + (M_Gyo - 1)
'2006.01.30
'2006.01.30
'2006.01.30    Select Case ID_KANRI_TBL(ING_No).Step
'2006.01.30        Case Step_MENU1_REQ, Step_MENU1_RES
'2006.01.30            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV1
'2006.01.30        Case Step_MENU2_REQ, Step_MENU2_RES
'2006.01.30            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV2
'2006.01.30        Case Step_MENU3_REQ, Step_MENU3_RES
'2006.01.30            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV3
'2006.01.30    End Select
'2006.01.30
'2006.01.30
'2006.01.30    Start_Gyo = PageNo * M_Gyo
'2006.01.30    End_Gyo = (PageNo * M_Gyo) + (M_Gyo - 1)
'2006.01.30
'2006.01.30
'2006.01.30    Send_Text.sts = Sts_OK                                      'ステータス　OK
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
'2006.01.30
'2006.01.30    Send_Text.Display_Flg = Display_MENU                        '表示画面フラグ メニュー画面
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
'2006.01.30                                                                '最終メニューフラグ
'2006.01.30    If Max_Page = 1 Then
'2006.01.30        Send_Text.End_Menu = Menu_Only          '１画面のみ
'2006.01.30        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
'2006.01.30    Else
'2006.01.30        If (Max_Page - 1) = PageNo Then
'2006.01.30            Send_Text.End_Menu = Menu_End       '最終ページ
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_End
'2006.01.30        Else
'2006.01.30            If PageNo = 0 Then
'2006.01.30                Send_Text.End_Menu = Menu_Head  '先頭ページ
'2006.01.30                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Head
'2006.01.30            Else
'2006.01.30                Send_Text.End_Menu = Menu_Mid   '途中ページ
'2006.01.30                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Mid
'2006.01.30            End If
'2006.01.30        End If
'2006.01.30    End If
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30    Send_Text.fileName = ""                                         '送信データファイル名
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
'2006.01.30
'2006.01.30    Send_Text.Buzzer = Buzzer_DEF                                   'ブザー音　標準
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
'2006.01.30    '---------------------------------------------------------------
'2006.01.30    Gyo_Suu = 0
'2006.01.30    j = -1
'2006.01.30    For i = Start_Gyo To End_Gyo
'2006.01.30        j = j + 1
'2006.01.30        If i > UBound(Menu_Tbl) Then
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Box_Type = ""                 'BOX属性
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = ""
'2006.01.30
'2006.01.30            Call UniCode_Conv(Send_Text.Box_Type(j).LCD, "")    '表示内容
'2006.01.30            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, "")
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).INIT = ""                     '数値初期値
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).MENU = ""                     'メニュ―番号
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = ""
'2006.01.30
'2006.01.30
'2006.01.30        Else
'2006.01.30            Gyo_Suu = Gyo_Suu + 1
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Box_Type = TYPE_MENU          'BOX属性
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = TYPE_MENU
'2006.01.30                                                                '表示内容
'2006.01.30            Call UniCode_Conv(Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
'2006.01.30            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).INIT = ""                     '数値初期値
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).MENU = Menu_Tbl(i).CODE       'メニュ―番号
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).CODE
'2006.01.30
'2006.01.30        End If
'2006.01.30
'2006.01.30    Next i
'2006.01.30
'2006.01.30    Send_Text.Menu_Suu = Format(Gyo_Suu, "00")      'メニュー項目数
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(Gyo_Suu, "00")
'2006.01.30
'2006.01.30
'2006.01.30    ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信
'2006.01.30
'2006.01.30    Sendbuf = Text_Create_Proc()
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30    Menu_Send_Proc = False
'2006.01.30
'2006.01.30End Function


'[2016/05/14 -  mdlProcから移動
Public Function Sagyo_Main_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『作業受信時のメイン処理』
'
'-------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim Found_Flg   As Boolean
    
    
    Sagyo_Main_Proc = True
    
    
    Found_Flg = False
    For i = 0 To UBound(WEL_Para_Tbl, 1)
        For j = 0 To UBound(WEL_Para_Tbl, 2)
                                        '最初は２桁で検索
            If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = _
                WEL_Para_Tbl(i, j).Action Then
                Found_Flg = True
                Exit For
            End If
        
        Next j
            
        If Found_Flg Then
            Exit For
        End If
    
    Next i
    
    If Not Found_Flg Then
        
        For i = 0 To UBound(WEL_Para_Tbl, 1)
            For j = 0 To UBound(WEL_Para_Tbl, 2)
                                        '最後は１桁で検索
               If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = Left(WEL_Para_Tbl(i, j).Action, 1) Then
                    Found_Flg = True
                    Exit For
                End If
        
            Next j
            
            If Found_Flg Then
                Exit For
            End If
        
        
        Next i
            
    End If


    If Not Found_Flg Then
                        
                        'ありえない異常（該当作業パラメータなし）
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_GRP = ""
                
        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
                
        '   -------------------------------- エラーメッセージ作成
        Call Err_Send_Proc("作業パラメータ（INI）", "未登録", "", "", "")
                    
        Sendbuf = Text_Create_Proc()
        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
            
        Sagyo_Main_Proc = False
        Exit Function
    
    End If

    Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE
        Case ACT_ZAITEI_IN          '在訂＋
        
        
            If Zaitei_In_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        
        Case ACT_ZAITEI_OUT         '在訂－
            
            '2007.10.02
            If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE) = _
                Wel_S_SHOUHI Then
                        
                '2007.10.02 資材消費専用
                If S_SHOUHI_Out_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
                        
                        
            Else
                
                '2015.02.21 資材消費(新)
                If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE) = _
                    Wel_S_SHOUHI2 Then
                    If S_SHOUHI_Out2_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Else
                
                
                
                    If Zaitei_Out_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                End If
            End If
        
        
        Case ACT_NYUKA              '入荷
    
        Case ACT_SYUKA_KEI          '出荷(出荷予定有り)→向け先宣言
        
            
            If MTS_Dec_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        
        Case ACT_SYUKA_HYO          '出荷(出庫表)
        
            If SYUKO_HYO_Dec_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        Case ACT_SYUKA_GAI          '出荷(出荷予定無し)
        
            If Out_Plan_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        
        Case ACT_IDO_IN             '移動入庫
        
'-----------------------------------------  2012.03.06
'            If Ido_In_Proc(Sendbuf, i, j) Then
'                Exit Function
'            End If
        
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
        
                Case Wel_IDO_IN_OSAKA          '2012.03.15
            
                    If Ido_In_OSAKA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Else
            
                    If Ido_In_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            End Select
        
        
        
        Case ACT_IDO_OUT            '移動出庫
        
            
            
'-----------------------------------------  2012.03.06
'            '2011.06.01
'            If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE) = _
'                Wel_HIN_FURIKAE_MAINA Then
'                '2011.06.01
'                If Ido_Out_Hin_Furikae_Proc(Sendbuf, i, j) Then
'                    Exit Function
'                End If
'
'            Else
'                If Ido_Out_Proc(Sendbuf, i, j) Then
'                    Exit Function
'                End If
'
'            End If
        
        Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
            Case Wel_HIN_FURIKAE_MAINA
        
                If Ido_Out_Hin_Furikae_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
        
        
            Case Wel_IDO_OUT_OSAKA          '2012.03.10
            
                If Ido_Out_OSAKA_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
            
            Case Wel_IDO_OUT_OSAKA2         '2014.11.07
            
                If Ido_Out_OSAKA_NEW_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
            
            
            Case Wel_IDO_OUT_OSAKA3         '2016.05.11
            
                If Ido_Out_OSAKA_NEW2_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
            
            
            
            
            Case Else
                If Ido_Out_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
        
        End Select
'-----------------------------------------  2012.03.06
        
        Case ACT_DENPYO_ID          '伝票ＩＤ
        
            If DEN_ID_Dec_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_DENPYO_ID2          '*伝票ＩＤ
        
            If DEN_ID_Dec2_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_KENPIN             '検品
        
            If Inspe_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_WEL_ETC            'WEL専用（照会）
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
            
                Case Wel_TANAOROSI      '「WEL 棚卸し」の要因
                
                    If Tanaorosi_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_TANAHYOJI      '「WEL 棚番表示」の要因
                
                    If Tanahyoji_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                
                Case Wel_HIN_SHOGO      '「WEL 品番別照合」の要因
                    
                    If Hin_Shogo_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_AVE_SYUKA      '「WEL 月平均出荷数」の要因
                
                    If Ave_Syuka_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_HOST_ZAIKO     '「WEL ホスト在庫照会」の要因
                    
                    If Host_Zaiko_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                            
                Case Wel_ST_TANABAN     '「WEL 標準棚番設定」の要因
                    
                    If St_Tanaban_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                            
            
                Case Wel_RIREKI         '「WEL 当日出庫履歴」の要因
                
                    If Rireki_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_SUII           '「WEL 出荷推移」の要因

                    If Suii_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If

                Case Wel_TANA_HIN_SHOGO '「WEL 棚番・品番別照合」の要因

                    If Tana_Hin_Shogo_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_TANAHYOJI_KASO '「WEL 棚番表示(仮想優先)」の要因
                
                    If Tanahyoji_Kaso_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_COMPO          '「WEL 構成表示」の要因 2006.10.15
                
                    If COMPO_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_KEN_RIREKI     '「WEL 検品実績」の要因 2006.10.15
                
                    If KEN_Rireki_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_RIREKI2        '「WEL 当日出庫履歴」の要因 2009.01.09
                
                    If Rireki2_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_AVE_SYUKA_ID   '「WEL 月平均　ＩＤ読み込み」の要因 2009.03.19
                
                    If Ave_Syuka_ID_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_KEN_ZAN_ID   '「WEL 集約梱包残」の要因 2010.02.15
            
            
                    If KEN_ZAN_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                '2010.12.13
                Case Wel_AKI_LOC                '「WEL 空きロケーションの検索」の要因
            
            
                        
                    If AKI_LOC_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                '2011.07.05
                Case Wel_S_AVE_SYUKA
            
                    If S_Ave_Syuka_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            End Select
    
        Case ACT_KENPIN_MTS             '検品（ＭＴＳ読み込みあり）
        
            If Inspe_Proc_MTS(Sendbuf, i, j) Then
                Exit Function
            End If
    
    
        Case ACT_GOODS_ONFF             '商品／未商品切り替え
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
    
                Case Wel_GOODS_ONOFF_ONO        '「WEL 商品/未商品切り替え　小野」の要因
                
                    If GOODS_ONOFF_Ono_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_GOODS_ONOFF_SIGA       '「WEL 商品/未商品切り替え　滋賀」の要因
            
                    If GOODS_ONOFF_Siga_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            End Select
    
    
        Case ACT_SPECIAL_PROCESS    '特殊処理
        
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
    
                Case Wel_RETURNED_GOODS         '「良品返品」の要因
                
                    If RETURNED_GOODS_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                
                Case Wel_LOCATION_MOVE         '「棚移動」の要因
                
                
                    If Location_Move_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
            
                Case Wel_RETURNED_GOODS_OSAKA   '「大阪ＰＣ　良品返品」の要因   2007.09.12
                
                    If RETURNED_GOODS_OSAKA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_SYUKA_CENCEL           '「出荷ｷｬﾝｾﾙ」の要因    2007.11.02
            
                    If SYUKA_CANCEL_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_ZAIKO_SEISA            '「在庫精査」の要因    2008.11.20
            
                    If Zaiko_Seisa_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_SYUKO_CANCEL           '「出庫CANCEL」の要因    2008.12.05
            
                    If SYUKO_CANCEL_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                                                '「移管処理」の要因    2009.02.26
                Case Wel_IKAN_1, Wel_IKAN_2, Wel_IKAN_3
            
                    
                    Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                        
                        Case Wel_IKAN_1
                    
                            If LSBU_IKAN_Proc(Sendbuf, i, j, 1) Then
                                Exit Function
                            End If
            
                        Case Wel_IKAN_2
                    
                            If LSBU_IKAN_Proc(Sendbuf, i, j, 2) Then
                                Exit Function
                            End If
            
                        Case Wel_IKAN_3
                    
                            If LSBU_IKAN_Proc(Sendbuf, i, j, 3) Then
                                Exit Function
                            End If
            
            
                    End Select
            
            
            
            
                                                '「移管削除処理」の要因    2009.03.09
                Case Wel_IKAN_DEL
            
                    
                
                    If LSBU_IKAN_DEL_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                '「移管表示処理」の要因    2009.03.09
                Case Wel_IKAN_DSP
            
                    
                
                    If LSBU_IKAN_DSP_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                '「棚番追加処理」の要因    2009.03.17
                Case Wel_TANA_INS
            
                    
                
                    If TANA_INS_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                '「出荷削除処理」の要因    2009.03.17
                Case Wel_SYUKA_DEL
            
                    
                
                    If SYUKA_DEL_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
                                                '「才数／口数処理」の要因    2010.03.09
                Case Wel_SAI_SU
            
                    
                
                    If SAI_SU_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
            
                                                '2010.12.13
                Case Wel_TANA_USE               '「WEL 棚使用状況」の要因
                                                
                                                
                                                
                    If TANA_USE_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                
                                                
                                                    
                                                
                                                '2011.03.05
                Case Wel_LABEL_PRINT            '「WEL ﾗﾍﾞﾙ発行」の要因
            
            
                    
                    
            
            
            
                    If LABEL_Print_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                                                '2011.08.05
                Case Wel_JAN_SET                '「WEL 品番(JAN)登録」の要因
            
                    If JAN_SET_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
            
                Case Wel_T_back                 '2015.01.22 引取処理(広島)
            
            
                    If T_back_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
                                                '2015.10.06
                Case Wel_LABEL_PRINT_CNT    '「WEL ﾗﾍﾞﾙ発行 枚数指定」の要因
            
            
                    
                    
            
            
            
                    If LABEL_Print_Cnt_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
            End Select
        
        Case ACT_KENPIN_DEN             '検品（大阪ＰＣ向け）   2006.12.07
        
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
'                Case Wel_Inspe_DEN, Wel_Inspe_DEN2              '2009.06.03
                Case Wel_Inspe_DEN                              '2009.06.03
        
        
        
                    If Inspe_Proc_DEN(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
        
                Case Wel_Inspe_E_BAG     '2010.01.21
                    
                    If Inspe_Proc_E_BAG(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
                Case Wel_KYOSEI_END
    
                    If KYOSEI_END_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
    
                Case Wel_LABEL_REPRINT  '2010.01.21
    
                    If LABEL_RePrint_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_Inspe_LOGISTIC '2010.01.25
                    
                    If Inspe_Proc_LOGISTIC(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
    
                Case Wel_SEK_PACKING    '2011.04.25
                    
                    If SEK_PACKING_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
    
    
                Case Wel_SEK_Inspe      '2011.05.09
            
                    If Inspe_Proc_SEK(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_SEK_SYUGO_PACKING  '2011.05.12
                    
                    If SEK_SYUGO_PACKING_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_SEK_KYOSEI_END     '2011.06.28
            
                    If SEK_KYOSEI_END_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
            
'------------------------------------------------------------------ 全数検品対応    2012.03.21
                Case Wel_Inspe_DEN_ALL                  '「WEL 大阪検品」の要因
                    If Inspe_Proc_DEN_ALL(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case Wel_Inspe_LOGISTIC_ALL             '「WEL ﾛｼﾞｽﾃｯｸｽ」の要因
                    If Inspe_Proc_LOGISTIC_ALL(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case Wel_SEK_PACKING_ALL                '「WEL 積水邸別梱包処理」の要因
                    If SEK_PACKING_ALL_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
'------------------------------------------------------------------ 全数検品対応    2012.03.21
            
            
                Case Wel_Inspe_E_BAG_ALL                                            '2012.06.20
                    If Inspe_Proc_E_BAG_ALL(Sendbuf, i, j) Then                     '2012.06.20
                        Exit Function                                               '2012.06.20
                    End If                                                          '2012.06.20
            
            
            End Select
    
        Case ACT_SYUKA_HYO_OSAKA        '出庫表出庫（大阪ＰＣ向け）   2007.03.16
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                Case Wel_SYUKO_HYO_OSAKA
        
                    If SYUKO_HYO_Dec_OSAKA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
    
            End Select
    
    
    
    
        Case ACT_IN_KENPIN_OSAKA        '入庫検品（大阪ＰＣ向け）   2007.06.07
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                Case Wel_IN_KENPIN_OSAKA
        
                    If NYUKO_KENPIN_OSAKA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
    
                Case WEL_IN_TANA_S_OSAKA            '資材検収入庫 2012.03.01
    
                    If NYUKO_KENPIN_OSAKA_S_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
                Case WEL_MAEGARI_TANA_S_OSAKA       '資材前借入庫 2016.05.30
    
                    If NYUKO_MAEGARI_OSAKA_S_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
    
    
            End Select
    
        Case ACT_IN_TANA_OSAKA        '入庫検品（大阪ＰＣ向け）   2007.06.07
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                Case Wel_IN_TANA_OSAKA
        
                    If NYUKO_TANA_OSAKA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
    
            End Select
    
    
        Case ACT_FURIKAE                '資材振替処理   2007.06.28
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                Case Wel_FURIKAE
        
                    If Furikae_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
        
        
                Case Wel_HIN_FURIKAE_PLUS   '品番振替処理   2011.06.01
        
                    If Hin_Furikae_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
        
                Case Wel_HIN_FURIKA_S    '部材センター振替出庫
                    
                    If Hin_Furikae_S_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
    
            End Select
    
    
        Case ACT_BINNO                  '便№処理(移管用) 2009.03.11
    
            If LSBU_IKAN_BinNo_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
    
    
        Case ACT_KENPIN_GAI             '海外向け検品   2009.08.07
        
            
            
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE   '2014.03.05
                Case WEL_KENPIN_GAI                                                                             '2014.03.05
                    If Inspe_Proc_GAI(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case WEL_KENPIN_Su                          '数量検品   2014.03.05
                    If Inspe_Proc_Su(Sendbuf, i, j) Then    '           2014.03.05
                        Exit Function                       '           2014.03.05
                    End If                                  '           2014.03.05
    
            End Select                                                                                          '2014.03.05
    
'        Case ACT_SAI_SU                 '才数／口数   2010.01.21
'
'            If SAI_SU_Proc(Sendbuf, i, j) Then
'                Exit Function
'            End If
    
    
        
        Case ACT_SHOUHINKA              '商品化ﾁｪｯｸ     2010.09.03
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                
                Case Wel_SHOUHINKA_CHECK
                
                    If SHOUHINKA_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_SHOUHINKA_GAI_CHECK                           '2015.11.07
                                                                        '2015.11.07
                    If SHOUHINKA_CHECK_GAI_PROC(Sendbuf, i, j) Then     '2015.11.07
                        Exit Function                                   '2015.11.07
                    End If                                              '2015.11.07
            
            
            
            
                Case Wel_HINBAN_CHECK    '品番ﾁｪｯｸ   2010.09.10
                
                    If HINBAN_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_COMPO_CHECK    '構成ﾁｪｯｸ   2011.03.02
                
                    If COMPO_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_COMPO_OSAKA_CHECK  '大阪ＰＣ　部材ｾﾝﾀｰ構成ﾁｪｯｸ   2012.03.16
                
                    If COMPO_OSAKA_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_KENPIN_OSAKA       '大阪ＰＣ　検品   2012.03.16
                
                    If KENPIN_OSAKA_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_KENPIN_OSAKA_NEW   '大阪ＰＣ　検品   2016.05.20
                
                    If KENPIN_OSAKA_NEW_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_KENPIN_OSAKA_NEW2  '大阪ＰＣ　検品(エラー表示有り)   2016.06.27
                
                    If KENPIN_OSAKA_NEW2_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            

            
            End Select


'-----------------------------------------------    床暖房　製造№  2013.06.06
        Case ACT_LotNo
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                
                Case Wel_LotNo_IN_CHECK
                
                    If LOTNO_IN_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If

                Case Wel_LotNo_OUT_CHECK
                
                    If LOTNO_OUT_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If

                Case Wel_LotNo_OUT_CANCEL
                
                    If LOTNO_OUT_CANCEL_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If

                Case Wel_LotNo_LABEL_PRINT
                
                    If LOTNO_LABEL_PRINT_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
                Case Wel_InvNo_OUT_CHECK                                '2014.07.01
                                                                        '2014.07.01
                    If INVNO_OUT_CHECK_PROC(Sendbuf, i, j) Then         '2014.07.01
                        Exit Function                                   '2014.07.01
                    End If                                              '2014.07.01
                                    
                Case Wel_InvNo_OUT_CANCEL                               '2014.07.01
                                                                        '2014.07.01
                    If INVNO_OUT_CANCEL_PROC(Sendbuf, i, j) Then        '2014.07.01
                        Exit Function                                   '2014.07.01
                    End If                                              '2014.07.01
            
            
            End Select
'-----------------------------------------------    床暖房　製造№  2013.06.06


'-----------------------------------------------    モジュール 2014.06.24
        Case ACT_MODULE
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                
                Case Wel_MODULE_INSPE
                
                    If MODULE_INSPE_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_MODULE_INSPE2                                  '2015.02.19
                
                    If MODULE_INSPE_CHECK2_PROC(Sendbuf, i, j) Then     '2015.02.19
                        Exit Function                                   '2015.02.19
                    End If                                              '2015.02.19
            
            End Select
'-----------------------------------------------    モジュール 2014.06.24



'-----------------------------------------------    (新)直送検品    2016.10.14
        Case ACT_KENPIN_Drct
        
            If Inspe_Proc_Drct(Sendbuf, i, j) Then
                Exit Function
            End If
'-----------------------------------------------    (新)直送検品    2016.10.14

'-----------------------------------------------   ﾊﾞｰｺｰﾄﾞ印字  2017.04.10
        Case ACT_BCR_PRINT

            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                Case Wel_BCR_DAKUTO
                    If BCR_DAKUTO_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case Wel_BCR_JAN
                    If BCR_JAN_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case Wel_BCR_Inspe
                    If BCR_Inspe_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case Wel_BCR_TANA
                    If BCR_TANA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            End Select
            

            






'-----------------------------------------------   ﾊﾞｰｺｰﾄﾞ印字  2017.04.10






    End Select

    Sagyo_Main_Proc = False

End Function

'[2016/05/14 -  mdlProcから移動

Public Function Cancel_Proc(Sendbuf As String, Optional Mode As Integer = 0, Optional Para As String = "  ") As Integer
'-------------------------------------------------------
'
'   『キャンセル処理（前画面検索）』
'
'-------------------------------------------------------
    
    Cancel_Proc = True
        
    
    Select Case ID_KANRI_TBL(ING_No).Step
    
        Case Step_Start         '子機電源ＯＮ
        Case Step_TANTO_REQ     '担当者要求
            
            Call Re_Send_Proc(Sendbuf)
                        
        Case Step_JGYOBU_REQ    '事業部要求

'            ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'            ID_KANRI_TBL(ING_No).JGYOBU = ""
'            Call Start_Proc(Sendbuf)

            '事業部要求でループする
            ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
            ID_KANRI_TBL(ING_No).JGYOBU = ""
            
            
            If UBound(NAIGAI) = 0 Then
            Else
                ID_KANRI_TBL(ING_No).NAIGAI = ""
            End If
            
'            ID_KANRI_TBL(ING_No).MENU_GRP = ""
            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
            
            
            
            '2010.04.23
            ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
            ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
            '2010.04.23
            
            

            Call Menu_Send_Proc(Sendbuf)


        Case Step_NAIGAI_REQ    '国内外要求
            
            
            If UBound(JGYOBU_T) = 0 Then
            
                        
                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                ID_KANRI_TBL(ING_No).NAIGAI = ""
                Call Menu_Send_Proc(Sendbuf)
            
            Else
            
                ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
                ID_KANRI_TBL(ING_No).JGYOBU = ""
                ID_KANRI_TBL(ING_No).NAIGAI = ""
            
'               ID_KANRI_TBL(ING_No).MENU_GRP = ""
                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'               ID_KANRI_TBL(ING_No).MENU_LV3 = ""
            
            
                '2010.04.23
                ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
                ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
                '2010.04.23
            
            
            
                Call Menu_Send_Proc(Sendbuf)
            End If

        Case Step_MENU1_REQ     'メニュー１要求
        
            
            If Mode = 0 Then            '2008.08.08
'                ST_LOG_OUT_F = False    '2008.08.08
            End If                      '2008.08.08
            
            If UBound(NAIGAI) = 0 Then
                
                
                
                If UBound(JGYOBU_T) = 0 Then
                
                
                
                '国内外の切り分けなし
    '                ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
    '                ID_KANRI_TBL(ING_No).JGYOBU = ""
    '                Call Start_Proc(Sendbuf)
                
                    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                    
                    '2010.04.23
                    ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
                    '2010.04.23
            
                
                    Call Menu_Send_Proc(Sendbuf)
            
            
                Else
            
                    '事業部要求でループする
                    ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
                    ID_KANRI_TBL(ING_No).JGYOBU = ""
        '            ID_KANRI_TBL(ING_No).NAIGAI = ""
                    
        '            ID_KANRI_TBL(ING_No).MENU_GRP = ""
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
        '            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
                    
                    
                    '2010.04.23
                    ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
                    ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
                    '2010.04.23
                    
                    
        
                    Call Menu_Send_Proc(Sendbuf)
                End If
            
            Else
                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                ID_KANRI_TBL(ING_No).NAIGAI = ""
                Call Menu_Send_Proc(Sendbuf)
            End If
        
        
        
        Case Step_MENU2_REQ     'メニュー２要求
        
            
            If Not CANCEL_OPE Then      '2008.09.01
                
                If Para <> "EN" Then
                
                
                
                    '前回がエラー送信
                    Call Re_Send_Proc(Sendbuf)
                    
                    
                    
                    Cancel_Proc = False
                    Exit Function
            
                End If
            End If
            
            
            
            
            
            ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        
            If Mode = 0 Then            '2008.08.08
'                ST_LOG_OUT_F = False    '2008.08.08
            End If                      '2008.08.08
        
        
            Call Menu_Send_Proc(Sendbuf)
        
        
'2006.01.30        Case Step_MENU3_REQ     'メニュー３要求
'2006.01.30
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_MENU2_REQ
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30
'2006.01.30            Call Menu_Send_Proc(Sendbuf)

        Case Step_Sagyo1_REQ    '作業１要求
'2006.01.30            If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV3)) <> 0 Then
'2006.01.30                ID_KANRI_TBL(ING_No).Step = Step_MENU3_REQ
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30                Call Menu_Send_Proc(Sendbuf)
'2006.01.30            Else
                
                
            If Mode = 0 Then            '2008.08.08
'                ST_LOG_OUT_F = False    '2008.08.08
            End If                      '2008.08.08
                
                If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV2)) <> 0 Then
                    MENU_UP_F = True   '2008.08.08
                    ID_KANRI_TBL(ING_No).Step = Step_MENU2_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                    Call Menu_Send_Proc(Sendbuf)
                Else
                    MENU_UP_F = True   '2008.08.08
                    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                    Call Menu_Send_Proc(Sendbuf)
                End If
'2006.01.30            End If
                                                    
                                                    
                                                    
                                                    '作業２／作業３／作業４／作業５要求
        Case Step_Sagyo2_REQ, Step_Sagyo3_REQ, Step_Sagyo4_REQ, Step_Sagyo5_REQ, Step_PRINT_REQ
        
        
        
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
        
                Case Wel_SHOUHINKA_CHECK
        
        
                    Select Case ID_KANRI_TBL(ING_No).Step
                    
                        Case Step_Sagyo2_REQ
                    
                    
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                            If Sagyo_Send_Proc() Then
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            End If
                            
                            Sendbuf = Text_Create_Proc()
                    
                    
                        Case Step_Sagyo3_REQ
                    
                    
                    
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
                            
                            '>>>>>>>>>  2017.09.22
                            'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                            'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                            
                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                            
                            '>>>>>>>>>  2017.09.22
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
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
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
                            
                            
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & _
                                                                                    Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                                                                                    '数値初期表示
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(3).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                    '入力桁数
                            Send_Text.Box_Type(3).Max_Size = "13"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------４行目
                                                                                    'BOX属性
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '表示内容
                            
                            If ID_KANRI_TBL(ING_No).GENPIN_CNT = 0 Then
                            
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                                                                                        Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & _
                                                                                        Space(7 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                            End If
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
                    
                    
                                            
                    
                    
                        Case Step_Sagyo4_REQ
                    
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
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
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
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
    
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
                            If ID_KANRI_TBL(ING_No).GENPIN_CNT = 0 Then
                            
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                                                                                        Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                                                                                        Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                            End If
                                                                                    '数値初期表示
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '初期カーソル位置
                            Send_Text.Box_Type(4).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    '入力桁数
                            Send_Text.Box_Type(4).Max_Size = "13"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "13"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
                            Sendbuf = Text_Create_Proc()
                    
                    
                    End Select
        
        
        
        
                Case Wel_SHOUHINKA_GAI_CHECK        '2016.04.05
        
                    Select Case ID_KANRI_TBL(ING_No).Step
                    
                        Case Step_Sagyo1_REQ
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                            If Sagyo_Send_Proc() Then
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            End If
                            
                            Sendbuf = Text_Create_Proc()
                        Case Step_Sagyo2_REQ
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                            If Sagyo_Send_Proc() Then
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            End If
                            
                            Sendbuf = Text_Create_Proc()
                        
                        
                        Case Step_Sagyo3_REQ
                        
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
        
                            Exit Function
                        
                        
                        Case Step_Sagyo4_REQ
                    
                    
                    
                    
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
    
                            Exit Function
                    
                    
                    End Select
                Case Else
        
            
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                    If Sagyo_Send_Proc() Then
                        Sendbuf = Text_Create_Proc()
                        Exit Function
                    End If
                    
                    Sendbuf = Text_Create_Proc()
                
                    If Mode = 0 Then            '2008.08.08
        '                ST_LOG_OUT_F = False    '2008.08.08
                    End If                      '2008.08.08
    
            End Select
    
    End Select
    
    Cancel_Proc = False


End Function

'[2016/05/14 -  mdlProcから移動
Public Function tmpZaiko_Clear_Proc() As Integer
'-------------------------------------------------------
'
'   『在庫データ（一時データ）の消去』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim RETRY_CNT   As Integer
    
    
    tmpZaiko_Clear_Proc = True
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        tmpZaiko_Clear_Proc = SYS_ERR
        Exit Function
    End If
    
    com = BtOpGetFirst

    Do
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        RETRY_CNT = 0
        Do
            sts = BTRV(com + BtSNoWait, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
            Select Case sts
                                
                Case BtNoErr
                    
                    Exit Do
                
                Case BtErrEOF
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ（一時データ）")
                        tmpZaiko_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
'                    DoEvents
                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                        DoEvents                                                    '2016.01.26
                    End If                                                          '2016.01.26
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ（一時データ）")
                    tmpZaiko_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
    
        
        RETRY_CNT = 0
        
        Do
        
            sts = BTRV(BtOpDelete, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ（一時データ）")
                        tmpZaiko_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
'                    DoEvents
                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                        DoEvents                                                    '2016.01.26
                    End If                                                          '2016.01.26
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ（一時データ）")
                    tmpZaiko_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
        Loop
    
        com = BtOpGetNext
    
    Loop

End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        tmpZaiko_Clear_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    tmpZaiko_Clear_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function


