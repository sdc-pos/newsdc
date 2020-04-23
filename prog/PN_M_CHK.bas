Attribute VB_Name = "PN_M_CHK"


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MT_2009.06.01
Function PN_CHK(PN_CODE As String, NAIGAI As String, InsTanto As String, Optional Inv_Mode As Integer = 0, Optional Message = 0)

'   引数：PN_CODE           チェック対象品番
'         NaiGai            G（外部品番）、N（内部品番）
'         InsTanto          ＩＴＥＭ追加担当者（もしくはプログラム名）

Dim yn          As Integer
Dim W_Msg       As String
Dim W_STR       As String
    
Dim sts         As Integer
    
Dim sts1        As Integer
    
    
    PN_CHK = True
    
    
    
    
    
    Select Case NAIGAI
        Case "G"
            sts = PN_M_GET(Last_JGYOBU, PN_CODE, 0)
            If sts Then
                If sts = BtErrKeyNotFound Then
                    If Inv_Mode = 1 Then
                        MsgBox "入力した項目はエラーです。（外部品番）"
                        Exit Function
                    Else
                    End If
                Else
                    MsgBox "入力した項目はエラーです。（外部品番）"
                    Exit Function
                End If
            End If
        Case Else
            sts = PN_M_GET2(Last_JGYOBU, PN_CODE, 0)
            If sts Then
                
                If sts = BtErrKeyNotFound Then
                    If Inv_Mode = 1 Then
                        MsgBox "入力した項目はエラーです。（内部品番）"
                        Exit Function
                    Else
                    End If
                Else
                    MsgBox "入力した項目はエラーです。（内部品番）"
                    Exit Function
                End If
            
            Else
            
            
                If Message = 1 Then
            
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(PN_MREC.PN, vbUnicode))
                    sts1 = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts1
                        
                        Case BtNoErr
                        
                            W_Msg = "対内品番を変更しますか？" & Chr(13) & Chr(10)
                            W_Msg = W_Msg & "対外品番：" & Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) & Chr(13) & Chr(10)
                            W_Msg = W_Msg & "対内品番：" & Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode)) & "→" & Trim(PN_CODE) & Chr(13) & Chr(10)

                        
                        
                            yn = MsgBox(W_Msg, vbYesNo + vbDefaultButton2, "対内品番変更確認")
                            If yn = vbYes Then
                    
                    
                    
                                Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(PN_MREC.SPn, vbUnicode))     '品番（内部）
            
    
    
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, InsTanto)        '追加　担当者
        
                                W_Date = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))  '追加　日時
                            
                            
                                Do
                                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                            If ans = vbCancel Then
                                                PN_CHK = False
                                                Exit Function
                                            End If
                                            
                                        Case Else
                                            Call File_Error(sts, BtOpInsert, "品目マスタ")
                                            MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                            Exit Function
                                    End Select
                                Loop
                    
                            End If
                                
                            PN_CHK = False
                            Exit Function
                        
                        Case BtErrKeyNotFound
                        Case Else
                    
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            MsgBox "システム異常が発生しました。処理を中止して下さい。"
                            Exit Function
                    End Select
                End If
            End If
    
    End Select
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   原産国対応  2012.01.23  削除 2012.02.06
'    Call UniCode_Conv(K0_Country.CountryCode, StrConv(PN_MREC.MadeInCode, vbUnicode))
'    sts = BTRV(BtOpGetEqual, Country_POS, CountryREC, Len(CountryREC), K0_Country, Len(K0_Country), 0)
'    Select Case sts
'        Case BtNoErr
'
'        Case BtErrKeyNotFound
'            Call UniCode_Conv(CountryREC.CountryName, "")
'            Call UniCode_Conv(CountryREC.CountryName2, "")
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "Countryマスタ")
'            MsgBox "システム異常が発生しました。処理を中止して下さい。"
'            Exit Function
'    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   原産国対応  2012.01.23
    
    W_Msg = ""
                                
    W_Msg = W_Msg & "該当する品番がありません。" & Chr(13) & Chr(10)
    W_Msg = W_Msg & "下記の品番を新規登録しますか？" & Chr(13) & Chr(10)
    W_Msg = W_Msg & " " & Chr(13) & Chr(10)
    W_Msg = W_Msg & "事業部　：" & Last_JGYOBU & Chr(13) & Chr(10)
    W_Msg = W_Msg & "国内外　：1 国内" & Chr(13) & Chr(10)
    W_Msg = W_Msg & "品　番　：" & RTrim(StrConv(PN_MREC.PN, vbUnicode)) & Chr(13) & Chr(10)
    W_Msg = W_Msg & "対内品番：" & RTrim(StrConv(PN_MREC.SPn, vbUnicode)) & Chr(13) & Chr(10)
    W_Msg = W_Msg & "品　名　：" & RTrim(StrConv(PN_MREC.PName, vbUnicode)) & Chr(13) & Chr(10)
    
    
    
    
    
    
    
    If IsNumeric(StrConv(PN_MREC.Tanka2, vbUnicode)) Then
        W_STR = Format(CDbl(Trim(StrConv(PN_MREC.Tanka2, vbUnicode))), "###,##0.00")
    Else
        W_STR = "0"
    End If
    W_Msg = W_Msg & "単価１　：" & W_STR & Chr(13) & Chr(10)
    
    If IsNumeric(StrConv(PN_MREC.Tanka3, vbUnicode)) Then
        W_STR = Format(CDbl(Trim(StrConv(PN_MREC.Tanka3, vbUnicode))), "###,##0.00")
    Else
        W_STR = "0"
    End If
    W_Msg = W_Msg & "単価２　：" & W_STR & Chr(13) & Chr(10)
    
    If IsNumeric(StrConv(PN_MREC.Tanka4, vbUnicode)) Then
        W_STR = Format(CDbl(Trim(StrConv(PN_MREC.Tanka4, vbUnicode))), "###,##0.00")
    Else
        W_STR = "0"
    End If
    W_Msg = W_Msg & "単価３　：" & W_STR & Chr(13) & Chr(10)
    W_Msg = W_Msg & "現物表示原産国 ：" & RTrim(StrConv(PN_MREC.MadeIn, vbUnicode)) & Chr(13) & Chr(10)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   原産国対応  2012.01.23  --> 原産国  2012.02.06
'    W_Msg = W_Msg & "MadeInCode：" & RTrim(StrConv(PN_MREC.MadeInCode, vbUnicode)) & Chr(13) & Chr(10)
    W_Msg = W_Msg & "原産国：" & RTrim(StrConv(PN_MREC.GENSANKOKU, vbUnicode)) & Chr(13) & Chr(10)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   原産国対応  2012.01.23
                                
    yn = MsgBox(W_Msg, vbYesNo + vbExclamation, "確認入力")
    If yn = vbNo Then
        
        Exit Function
    End If

    If Item_PUT_Proc(InsTanto) Then
        W_STR = "追加に失敗しました。" & Chr(13) & Chr(10) & "再試行願います。"
        MsgBox W_STR
        
        Exit Function
    End If
    
    
    PN_CHK = False
    
    
End Function


'                                           MT_2009.06.01
Function Item_PUT_Proc(InsTanto As String) As Integer

Dim sts         As Integer
Dim ans         As Integer
Dim W_Date      As String

    Item_PUT_Proc = True
    
    
    
    
    
    
    
    Call Rclr_ITEMREC

    Call UniCode_Conv(ITEMREC.JGYOBU, Last_JGYOBU)                          '事業部区分
    Call UniCode_Conv(ITEMREC.NAIGAI, "1")                                  '国内外

    Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(PN_MREC.PN, vbUnicode))      '品番（外部）
    Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(PN_MREC.PName, vbUnicode))  '品名
    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(PN_MREC.SPn, vbUnicode))     '品番（内部）

    Call UniCode_Conv(ITEMREC.ST_SOKO, StrConv(PN_MREC.SOKO, vbUnicode))    '標準入庫倉庫 倉庫
    
    Call UniCode_Conv(ITEMREC.ST_RETU, "")          '             列
    Call UniCode_Conv(ITEMREC.ST_REN, "")           '             連
    Call UniCode_Conv(ITEMREC.ST_DAN, "")           '             段
    
    Call UniCode_Conv(ITEMREC.JAN_CODE, "")         'Janコード
    Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")      'グリックス棚番１
    Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")      'グリックス棚番２
    Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")      'グリックス棚番３


''*------------------------------------------ 2005.11.15 追加(商品ﾗﾍﾞﾙ項目) ▽
'    Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, StrConv(PN_MREC.PName, vbUnicode))      '商品ﾗﾍﾞﾙ   品名

                                                    '           価格(1)
    
        
    
    
    If IsNumeric(StrConv(PN_MREC.Tanka2, vbUnicode)) Then
        Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(Val(StrConv(PN_MREC.Tanka2, vbUnicode)), "0000000000"))
    Else
        Call UniCode_Conv(ITEMREC.L_URIKIN1, "0000000000")
    End If
                                                    
                                                    
                                                    '           価格(2)
    If IsNumeric(StrConv(PN_MREC.Tanka3, vbUnicode)) Then
        Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(Val(StrConv(PN_MREC.Tanka3, vbUnicode)), "0000000000"))
    Else
        Call UniCode_Conv(ITEMREC.L_URIKIN2, "0000000000")
    End If
                                                    
                                                    
                                                    '           価格(3)
    If IsNumeric(StrConv(PN_MREC.Tanka4, vbUnicode)) Then
        Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(Val(StrConv(PN_MREC.Tanka4, vbUnicode)), "0000000000"))
    Else
        Call UniCode_Conv(ITEMREC.L_URIKIN3, "0000000000")
    End If
    
    Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")    '           適用機種備考(→機種（３）)
    Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")    '           作業指示
    Call UniCode_Conv(ITEMREC.L_BIKOU3, "")         '           備考３
                                                    
                                                    
                                                    
'Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, Last_JGYOBU)
                                                    '           入り数
    Call UniCode_Conv(ITEMREC.L_IRI_QTY, String(UBound(ITEMREC.L_IRI_QTY) + 1, "0"))
''*------------------------------------------ 2005.11.15 追加(商品ﾗﾍﾞﾙ項目) △
    Call UniCode_Conv(ITEMREC.S_TANTO, "")          '収単／担当者コード


    
    
    Call UniCode_Conv(ITEMREC.GENSANKOKU, StrConv(PN_MREC.MadeIn, vbUnicode))      '原産国
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.01.23 原産国名-->PNよりに変更  2012.02.06
'    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, StrConv(CountryREC.CountryName2, vbUnicode))
    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, StrConv(PN_MREC.GENSANKOKU, vbUnicode))
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.01.23 原産国名
    


    Call UniCode_Conv(ITEMREC.INS_TANTO, InsTanto)          '追加　担当者
    
    W_Date = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
    Call UniCode_Conv(ITEMREC.Ins_DateTime, W_Date)         '追加　日時





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.12.14
    Call UniCode_Conv(ITEMREC.D_MODEL, StrConv(PN_MREC.DModel, vbUnicode))          '代表機種品目ｺｰﾄﾞ
    Call UniCode_Conv(ITEMREC.HINMOKU, StrConv(PN_MREC.HINMOKU, vbUnicode))         '品目ｺｰﾄﾞ
    Call UniCode_Conv(ITEMREC.HYO_TANKA, StrConv(PN_MREC.HyoTan, vbUnicode))        '標準単価
    Call UniCode_Conv(ITEMREC.K_KEITAI, StrConv(PN_MREC.KKeitai, vbUnicode))        '個装形態
    Call UniCode_Conv(ITEMREC.UNIT_BUHIN, StrConv(PN_MREC.UnitKbn, vbUnicode))      'ﾕﾆｯﾄ区分
    Call UniCode_Conv(ITEMREC.NAI_BUHIN, StrConv(PN_MREC.NaiKbn, vbUnicode))        '国内供給区分
    Call UniCode_Conv(ITEMREC.GAI_BUHIN, StrConv(PN_MREC.GaiKbn, vbUnicode))        '国外供給区分
    Call UniCode_Conv(ITEMREC.GLICS1_TANA, StrConv(PN_MREC.Loc1, vbUnicode))        '棚番1
    Call UniCode_Conv(ITEMREC.GLICS2_TANA, StrConv(PN_MREC.Loc2, vbUnicode))        '棚番2
    Call UniCode_Conv(ITEMREC.GLICS3_TANA, StrConv(PN_MREC.Loc3, vbUnicode))        '棚番3
    Call UniCode_Conv(ITEMREC.L_KISHU1, StrConv(PN_MREC.NaiModel, vbUnicode))       '代表機種1
    Call UniCode_Conv(ITEMREC.L_KISHU2, StrConv(PN_MREC.GaiModel, vbUnicode))       '代表機種2
    Call UniCode_Conv(ITEMREC.CS_TANTO_CD, StrConv(PN_MREC.KobaiTanto, vbUnicode))  '購買担当者
    Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, StrConv(PN_MREC.PNameEngA, vbUnicode))  '英語品名
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.12.14
    




'---------------------------------------------------------------------------------------------
    Do
        sts = BTRV(BtOpInsert, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Item_PUT_Proc = False
                    Exit Function
                End If
            Case BtErrDuplicates
                MsgBox "すでに追加されています。"
                Exit Function
                
            Case Else
                Call File_Error(sts, BtOpInsert, "品目マスタ")
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Exit Function
        End Select
    Loop
    
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)                      '事業部区分
    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")                              '国内外
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(PN_MREC.PN, vbUnicode))  '品番（外部）
    Do
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
                
            Case BtErrKeyNotFound
                MsgBox "品目マスタ　追加失敗！"
                Exit Function
            
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Item_PUT_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Exit Function
        End Select
    Loop
    
    Item_PUT_Proc = False
    
    
End Function

