Attribute VB_Name = "NYUKO_UPDATE"
Option Explicit

Public Function Nyuko_Update_Proc(JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    TO_LOCATION As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional MEMO As String = "          ", _
                                    Optional SHIIRE_CODE As String = "     ", _
                                    Optional SHIIRE_TANKA As String = "           ", _
                                    Optional KEIJYO_YM As String = "      ", _
                                    Optional MENU_NO As String = "  ", _
                                    Optional LOG_NON As Integer = 0, _
                                    Optional KAMOKU_FURIKAE As String = "  ", _
                                    Optional GENSANKOKU As String = "                    ", _
                                    Optional SHIIRE_WORK_CENTER As String = "        ", _
                                    Optional ID_NO2 As String = "            ", _
                                    Optional YOSAN_FROM As String = "     ", _
                                    Optional YOSAN_TO As String = "     ", _
                                    Optional wkMTS As String = "        ") As Integer
'****************************************************
'*      「入荷／入庫処理」在庫データ更新
'*
'*  在庫データの更新を行う。
'*  (引数の設定ミスはこちらではチェックしない)
'*  ※トランザクション処理が必要な場合は呼び元で行う事
'*  使用ﾌｧｲﾙ    :   在庫データ
'*                  品目マスタ
'*                  要因マスタ
'*                  在庫移動歴
'*                  入荷実績
'*                  倉庫マスタ
'*
'*  引数：  事業部（省略不可）
'*          国内外（省略不可）
'*          品番外部(省略不可)
'*          TO列（XXXXXXXX(倉庫№+列+連+段)省略不可）
'*          入荷日(YYYYMMDD 省略不可)
'*          要因(省略不可)
'*          商品化済み実績数（何れか一方必須）
'*          未商品実績数　　（　　〃　　　　）
'*          ID(省略不可)
'*          担当者（省略不可）
'*          リトライ(省略可 １桁目:1=画面メッセージ有 0:無，２桁目:リトライ回数(0～9 0:無限))
'*          メモ(省略可 履歴に出力するメモ内容)
'*          仕入先ｺｰﾄﾞ（資材用省略可　2006.01.05）
'*          仕入単価（資材用省略可　2006.01.05）
'*          計上年月（資材用省略可　2006.01.05）
'*          ﾒﾆｭｰｸﾞﾙｰﾌﾟ（原価管理項目）  2006.01.30
'*  戻り値: false       :正常
'*          true        :継続可能な異常
'*          SYS_ERR     :継続できない異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'****************************************************
Dim sts             As Integer
Dim com             As Integer

Dim RETRY_CNT       As Integer
Dim MESG_FLG        As Integer
Dim RETRY_SU        As Integer
    
Dim ans             As Integer
    
Dim Ins_DateTime    As String * 14                  '2004.12.09
    
    
    Nyuko_Update_Proc = True
                                                                      
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                                        
    Ins_DateTime = Format(Now, "YYYYMMDDHHMMSS")    '2004.12.09
    '*------------------------------------------------------'品目ﾏｽﾀの確保
    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)               '事業部
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)               '内外
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)             '品番（外部）
        
    RETRY_CNT = 0
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                If MESG_FLG = 1 Then
                    Beep
                    MsgBox "品目コードが存在しません。更新処理を中止します。", vbOKOnly, "確認入力"
                End If
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ", 0)
                Exit Function
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                If RETRY_SU <> 0 Then
                    
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ", 0)
                        Nyuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                
                End If
                
                If MESG_FLG = 0 Then
'                    DoEvents
                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                        DoEvents                                                    '2016.01.26
                    End If                                                          '2016.01.26
                Else
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Nyuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            
            
            Case BtErrDEAD_LOCK
                Nyuko_Update_Proc = SYS_CANCEL
                Exit Function
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Nyuko_Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
    
'============================================================
'************************************************************ 商品化済み更新
    If SUMI_JITU_QTY <> 0 Then
    '*------------------------------------------------------'在庫データ読込み
        Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(TO_LOCATION, 1, 2)) '倉庫№
        Call UniCode_Conv(K0_ZAIKO.Retu, Mid(TO_LOCATION, 3, 2))    '列
        Call UniCode_Conv(K0_ZAIKO.Ren, Mid(TO_LOCATION, 5, 2))     '連
        Call UniCode_Conv(K0_ZAIKO.Dan, Mid(TO_LOCATION, 7, 2))     '段
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                  '事業部
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                  '内外
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                '品番（外部）
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "0")                   '商品／未商品
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, NYUKA_DT)              '入荷日
    
        RETRY_CNT = 0
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    com = BtOpInsert
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                    If RETRY_SU <> 0 Then
                    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫ﾃﾞｰﾀ", 0)
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                    
                        End If
                
                    End If
                
                    If MESG_FLG = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case BtErrDEAD_LOCK
                    Nyuko_Update_Proc = SYS_CANCEL
                    Exit Function
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                    Nyuko_Update_Proc = SYS_ERR
                    Exit Function
            End Select
            
        Loop
                                                                                
        If com = BtOpInsert Then
                                        '新規追加
            Call UniCode_Conv(ZAIKOREC.SOKO_NO, Mid(TO_LOCATION, 1, 2)) '倉庫№
            Call UniCode_Conv(ZAIKOREC.Retu, Mid(TO_LOCATION, 3, 2))    '列
            Call UniCode_Conv(ZAIKOREC.Ren, Mid(TO_LOCATION, 5, 2))     '連
            Call UniCode_Conv(ZAIKOREC.Dan, Mid(TO_LOCATION, 7, 2))     '段
            Call UniCode_Conv(ZAIKOREC.JGYOBU, JGYOBU)                  '事業部
            Call UniCode_Conv(ZAIKOREC.NAIGAI, NAIGAI)                  '内外
            Call UniCode_Conv(ZAIKOREC.HIN_GAI, HIN_GAI)                '品番（外部）
            Call UniCode_Conv(ZAIKOREC.GOODS_ON, "0")                   '商品／未商品
            Call UniCode_Conv(ZAIKOREC.NYUKA_DT, NYUKA_DT)              '入荷日
                                                                        '入庫日
            Call UniCode_Conv(ZAIKOREC.NYUKO_DT, Format(Date, "yyyymmdd"))
                                                                        '品番（内部）
            Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                                                        '有効在庫数
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(SUMI_JITU_QTY, "00000000"))
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '排他フラグ
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '使用中子機ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '使用中プログラム
                                                                        '商品化日付
            Call UniCode_Conv(ZAIKOREC.GOODS_YMD, Format(Now, "YYYYMMDD"))
            
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, SHIIRE_CODE)        '仕入先ｺｰﾄﾞ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, SHIIRE_TANKA)      '仕入単価
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, KEIJYO_YM)            '計上年月
            
            
            
            
            '------------   2010.07.08 ▽
            Call UniCode_Conv(ZAIKOREC.GENSANKOKU, GENSANKOKU)          '原産国
                                                                        '資材仕入先ﾜｰｸｾﾝﾀｰ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
            Call UniCode_Conv(ZAIKOREC.ID_NO2, ID_NO2)                  'ID_NO
            Call UniCode_Conv(ZAIKOREC.YOSAN_FROM, YOSAN_FROM)          '予算　元
            Call UniCode_Conv(ZAIKOREC.YOSAN_TO, YOSAN_TO)              '予算　先
            '------------   2010.07.08 ▽
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            Call UniCode_Conv(ZAIKOREC.FILLER, "")
        Else
                                        '在庫数更新
                                                                        '在庫数
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) + SUMI_JITU_QTY, "00000000"))
                                                                    
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, SHIIRE_CODE)        '仕入先ｺｰﾄﾞ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, SHIIRE_TANKA)      '仕入単価
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, KEIJYO_YM)            '計上年月
        
        
        
        
            '------------   2010.07.08 ▽
            Call UniCode_Conv(ZAIKOREC.GENSANKOKU, GENSANKOKU)          '原産国
                                                                        '資材仕入先ﾜｰｸｾﾝﾀｰ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
            Call UniCode_Conv(ZAIKOREC.ID_NO2, ID_NO2)                  'ID_NO
            Call UniCode_Conv(ZAIKOREC.YOSAN_FROM, YOSAN_FROM)          '予算　元
            Call UniCode_Conv(ZAIKOREC.YOSAN_TO, YOSAN_TO)              '予算　先
            '------------   2010.07.08 ▽
        
        
        End If
    
        RETRY_CNT = 0
    '*------------------------------------------------------'在庫データ出力
        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                    If RETRY_SU <> 0 Then
                    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                            Call File_Error(sts, com, "在庫ﾃﾞｰﾀ", 0)
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                    
                        End If
                
                    End If
                
                    If MESG_FLG = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                
                Case BtErrDEAD_LOCK
                    Nyuko_Update_Proc = SYS_CANCEL
                    Exit Function
                
                Case Else
                    Call File_Error(sts, com, "在庫データ")
                    Nyuko_Update_Proc = SYS_ERR
                    Exit Function
                
            End Select
        Loop
'============================================================
    '*------------------------------------------------------'在庫移動歴出力
'        sts = IDOREKI_OUTPUT_PROC(Space(8), _
'                                    TO_LOCATION, _
'                                    JGYOBU, _
'                                    NAIGAI, _
'                                    HIN_GAI, _
'                                    NYUKA_DT, _
'                                    YOIN, _
'                                    SUMI_JITU_QTY, _
'                                    0, _
'                                    ID, _
'                                    TANTO_CODE, _
'                                    RETRY, , _
'                                    MEMO, _
'                                    Ins_DateTime, _
'                                    SHIIRE_CODE, SHIIRE_TANKA, KEIJYO_YM, MENU_NO)
'        If sts Then
'            Nyuko_Update_Proc = sts
'            Exit Function
'        End If
    End If
'************************************************************ 未商品更新
    If MI_JITU_QTY <> 0 Then
    '*------------------------------------------------------'在庫データ読込み
        Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(TO_LOCATION, 1, 2)) '倉庫№
        Call UniCode_Conv(K0_ZAIKO.Retu, Mid(TO_LOCATION, 3, 2))    '列
        Call UniCode_Conv(K0_ZAIKO.Ren, Mid(TO_LOCATION, 5, 2))     '連
        Call UniCode_Conv(K0_ZAIKO.Dan, Mid(TO_LOCATION, 7, 2))     '段
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                  '事業部
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                  '内外
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                '品番（外部）
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "1")                   '商品／未商品
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, NYUKA_DT)              '入荷日
    
        RETRY_CNT = 0
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    com = BtOpInsert
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                    If RETRY_SU <> 0 Then
                    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫ﾃﾞｰﾀ", 0)
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                    
                        End If
                
                    End If
                
                    If MESG_FLG = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                
                
                Case BtErrDEAD_LOCK
                    Nyuko_Update_Proc = SYS_CANCEL
                    Exit Function
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                    Nyuko_Update_Proc = SYS_ERR
                    Exit Function
            End Select
            
        Loop
                                                                                
        If com = BtOpInsert Then
                                        '新規追加
            Call UniCode_Conv(ZAIKOREC.SOKO_NO, Mid(TO_LOCATION, 1, 2)) '倉庫№
            Call UniCode_Conv(ZAIKOREC.Retu, Mid(TO_LOCATION, 3, 2))    '列
            Call UniCode_Conv(ZAIKOREC.Ren, Mid(TO_LOCATION, 5, 2))     '連
            Call UniCode_Conv(ZAIKOREC.Dan, Mid(TO_LOCATION, 7, 2))     '段
            Call UniCode_Conv(ZAIKOREC.JGYOBU, JGYOBU)                  '事業部
            Call UniCode_Conv(ZAIKOREC.NAIGAI, NAIGAI)                  '内外
            Call UniCode_Conv(ZAIKOREC.HIN_GAI, HIN_GAI)                '品番（外部）
            Call UniCode_Conv(ZAIKOREC.GOODS_ON, "1")                   '商品／未商品
            Call UniCode_Conv(ZAIKOREC.NYUKA_DT, NYUKA_DT)              '入荷日
                                                                        '入庫日
            Call UniCode_Conv(ZAIKOREC.NYUKO_DT, Format(Date, "yyyymmdd"))
                                                                        '品番（内部）
            Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                                                        '有効在庫数
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(MI_JITU_QTY, "00000000"))
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '排他フラグ
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '使用中子機ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '使用中プログラム
            
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, SHIIRE_CODE)        '仕入先ｺｰﾄﾞ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, SHIIRE_TANKA)      '仕入単価
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, KEIJYO_YM)            '計上年月
            
            '------------   2010.07.08 ▽
            Call UniCode_Conv(ZAIKOREC.GENSANKOKU, GENSANKOKU)          '原産国
                                                                        '資材仕入先ﾜｰｸｾﾝﾀｰ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
            Call UniCode_Conv(ZAIKOREC.ID_NO2, ID_NO2)                  'ID_NO
            Call UniCode_Conv(ZAIKOREC.YOSAN_FROM, YOSAN_FROM)          '予算　元
            Call UniCode_Conv(ZAIKOREC.YOSAN_TO, YOSAN_TO)              '予算　先
            '------------   2010.07.08 ▽
            
            Call UniCode_Conv(ZAIKOREC.FILLER, "")
        Else
                                        '在庫数更新
                                                                    '在庫数
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) + MI_JITU_QTY, "00000000"))
        
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, SHIIRE_CODE)        '仕入先ｺｰﾄﾞ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, SHIIRE_TANKA)      '仕入単価
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, KEIJYO_YM)            '計上年月
            '------------   2010.07.08 ▽
            Call UniCode_Conv(ZAIKOREC.GENSANKOKU, GENSANKOKU)          '原産国
                                                                        '資材仕入先ﾜｰｸｾﾝﾀｰ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
            Call UniCode_Conv(ZAIKOREC.ID_NO2, ID_NO2)                  'ID_NO
            Call UniCode_Conv(ZAIKOREC.YOSAN_FROM, YOSAN_FROM)          '予算　元
            Call UniCode_Conv(ZAIKOREC.YOSAN_TO, YOSAN_TO)              '予算　先
            '------------   2010.07.08 ▽
        
        
        End If
    
        RETRY_CNT = 0
    '*------------------------------------------------------'在庫データ出力
        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                    If RETRY_SU <> 0 Then
                    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                            Call File_Error(sts, com, "在庫ﾃﾞｰﾀ", 0)
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                    
                        End If
                
                    End If
                
                    If MESG_FLG = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                
                
                Case BtErrDEAD_LOCK
                    Nyuko_Update_Proc = SYS_CANCEL
                    Exit Function
                
                Case Else
                    Call File_Error(sts, com, "在庫データ")
                    Nyuko_Update_Proc = SYS_ERR
                    Exit Function
                
            End Select
        Loop
'============================================================
'    '*------------------------------------------------------'在庫移動歴出力
'        sts = IDOREKI_OUTPUT_PROC(Space(8), _
'                                    TO_LOCATION, _
'                                    JGYOBU, _
'                                    NAIGAI, _
'                                    HIN_GAI, _
'                                    NYUKA_DT, _
'                                    YOIN, _
'                                    0, _
'                                    MI_JITU_QTY, _
'                                    ID, _
'                                    TANTO_CODE, _
'                                    RETRY, , _
'                                    MEMO, _
'                                    Ins_DateTime, _
'                                    SHIIRE_CODE, SHIIRE_TANKA, KEIJYO_YM, MENU_NO)
'        If sts Then
'            Nyuko_Update_Proc = sts
'            Exit Function
'        End If
    End If
'============================================================
    '*------------------------------------------------------'在庫移動歴出力 2008.08.08 出力箇所移動
        
    If Trim(KAMOKU_FURIKAE) = "" Then       '2009.06.26
    
        sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                    TO_LOCATION, _
                                    JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    NYUKA_DT, _
                                    YOIN, _
                                    SUMI_JITU_QTY, _
                                    MI_JITU_QTY, _
                                    ID, _
                                    TANTO_CODE, _
                                    RETRY, , _
                                    MEMO, _
                                    Ins_DateTime, _
                                    SHIIRE_CODE, SHIIRE_TANKA, KEIJYO_YM, MENU_NO, , , , , , , , GENSANKOKU, SHIIRE_WORK_CENTER, ID_NO2, YOSAN_FROM, YOSAN_TO, wkMTS)
    
    Else
    
        sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                    TO_LOCATION, _
                                    JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    NYUKA_DT, _
                                    KAMOKU_FURIKAE, _
                                    SUMI_JITU_QTY, _
                                    MI_JITU_QTY, _
                                    ID, _
                                    TANTO_CODE, _
                                    RETRY, , _
                                    MEMO, _
                                    Ins_DateTime, _
                                    SHIIRE_CODE, SHIIRE_TANKA, KEIJYO_YM, MENU_NO, , , , , , , , GENSANKOKU, SHIIRE_WORK_CENTER, ID_NO2, YOSAN_FROM, YOSAN_TO, wkMTS)
    
    
    End If
    
    
    If sts Then
        Nyuko_Update_Proc = sts
        Exit Function
    End If
    
    
    
'    If YOIN = YOIN_MAEGARI Then        2016.05.30
    If YOIN = YOIN_MAEGARI Or YOIN = WEL_MAEGARI_TANA_S_OSAKA Then  '2016.05.30
    '*------------------------------------------------------'前借りデータ読込み
        Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
            Case SHIZAI
                '資材前借処理
                Call UniCode_Conv(K0_P_NYU.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_NYU.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_NYU.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_P_NYU.NYUKA_DT, Format(Now, "YYYYMMDD"))
                
                
                RETRY_CNT = 0
                Do
                                                '前借りﾃﾞｰﾀ読込み
                    sts = BTRV(BtOpGetEqual + BtSNoWait, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            com = BtOpUpdate
                            Exit Do
                        Case BtErrKeyNotFound
                            com = BtOpInsert
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                'リトライ回数チェック
                            If RETRY_SU <> 0 Then
                            
                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                                '回数オーバー
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "資材前借", 0)
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                            
                                End If
                        
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                        
                        Case BtErrDEAD_LOCK
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "資材前借データ")
                            Nyuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                    
                Loop
        
                If com = BtOpInsert Then
                                            '新規追加
                                                            '事業部
                    Call UniCode_Conv(P_NYUREC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                                                            '国内外
                    Call UniCode_Conv(P_NYUREC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                                                            '品目（外部）
                    Call UniCode_Conv(P_NYUREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                            '入荷日
                    Call UniCode_Conv(P_NYUREC.NYUKA_DT, Format(Now, "YYYYMMDD"))
                                                            '実績数量
                    Call UniCode_Conv(P_NYUREC.NYUKA_QTY, Format(SUMI_JITU_QTY + MI_JITU_QTY, "00000000"))
                                                            '相殺日付
                    Call UniCode_Conv(P_NYUREC.SOUSAI_DT, "")
                                                            '相殺数
                    Call UniCode_Conv(P_NYUREC.SOUSAI_DT, "00000000")
                                                            '登録端末
                    Call UniCode_Conv(P_NYUREC.WS_ID, ID)
                                        
                                                            '仕入先
                    Call UniCode_Conv(P_NYUREC.SHIIRE_CODE, SHIIRE_CODE)
                                            
                    Call UniCode_Conv(P_NYUREC.SHIIRE_TANKA, SHIIRE_TANKA)
                    
                    
                    Call UniCode_Conv(P_NYUREC.FILLER, "")
                
                                                            '登録日時
                    Call UniCode_Conv(P_NYUREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                Else
                                                            '実績数量
                    SUMI_JITU_QTY = SUMI_JITU_QTY + MI_JITU_QTY + CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode))
                    Call UniCode_Conv(P_NYUREC.NYUKA_QTY, Format(SUMI_JITU_QTY, "00000000"))
                End If
            '*------------------------------------------------------'前借りデータ出力
                RETRY_CNT = 0
                Do
                    sts = BTRV(com, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                'リトライ回数チェック
                            If RETRY_SU <> 0 Then
                                
                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                                '回数オーバー
                                    Call File_Error(sts, com, "資材前借", 0)
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                            
                                End If
                            
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                        
                        Case BtErrDEAD_LOCK
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        
                        
                        Case Else
                            Call File_Error(sts, com, "資材前借")
                            Nyuko_Update_Proc = SYS_ERR
                            Exit Function
                                
                    End Select
                Loop
            
            
            
            
            
            
            
            Case Else
                '部品前借処理
                Call UniCode_Conv(K0_J_NYU.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_J_NYU.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_J_NYU.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            
                RETRY_CNT = 0
                Do
                                                '前借りﾃﾞｰﾀ読込み
                    sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            com = BtOpUpdate
                            Exit Do
                        Case BtErrKeyNotFound
                            com = BtOpInsert
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                'リトライ回数チェック
                            If RETRY_SU <> 0 Then
                            
                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                                '回数オーバー
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "入荷実績", 0)
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                            
                                End If
                        
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                        
                        Case BtErrDEAD_LOCK
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "入荷実績データ")
                            Nyuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                    
                Loop
        
                If com = BtOpInsert Then
                                            '新規追加
                                                            '事業部
                    Call UniCode_Conv(J_NYUREC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                                                            '国内外
                    Call UniCode_Conv(J_NYUREC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                                                            '品目（外部）
                    Call UniCode_Conv(J_NYUREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                            '実績数量
                    Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(SUMI_JITU_QTY + MI_JITU_QTY, "00000000"))
                    Call UniCode_Conv(J_NYUREC.FILLER, "")
                
                
                    Call UniCode_Conv(J_NYUREC.INS_DATE, Format(Now, "YYYYMMDD"))            '2019.01.12
                Else
                                                            '実績数量
                    SUMI_JITU_QTY = SUMI_JITU_QTY + MI_JITU_QTY + CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
                    Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(SUMI_JITU_QTY, "00000000"))
                End If
            
            
            
                          
            '*------------------------------------------------------'前借りデータ出力
                RETRY_CNT = 0
                Do
                    sts = BTRV(com, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                'リトライ回数チェック
                            If RETRY_SU <> 0 Then
                                
                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                                '回数オーバー
                                    Call File_Error(sts, com, "入荷実績", 0)
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                            
                                End If
                            
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                        
                        Case BtErrDEAD_LOCK
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        
                        Case Else
                            Call File_Error(sts, com, "入荷実績データ")
                            Nyuko_Update_Proc = SYS_ERR
                            Exit Function
                                
                    End Select
                Loop
        End Select
    End If
'============================================================
                                        
                                        '倉庫ﾏｽﾀ読込み
    Call UniCode_Conv(K0_SOKO.SOKO_NO, Mid(TO_LOCATION, 1, 2))
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound           '有るとまずいがエラーにしない
            Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_KASO)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "倉庫ﾏｽﾀ")
            Nyuko_Update_Proc = SYS_ERR
            Exit Function
    End Select
    
    If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_JITU Then
                                        '標準棚番
'        If Last_JGYOBU = SOJIKI Or _
'            Last_JGYOBU = SENTAKU Then
                                        '掃除機は設定済みの上書きをしない
            If StrConv(ITEMREC.ST_SET_DT, vbUnicode) = Space(8) Then
                Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Date, "yyyymmdd"))
                Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(TO_LOCATION, 1, 2))
                Call UniCode_Conv(ITEMREC.ST_RETU, Mid(TO_LOCATION, 3, 2))
                Call UniCode_Conv(ITEMREC.ST_REN, Mid(TO_LOCATION, 5, 2))
                Call UniCode_Conv(ITEMREC.ST_DAN, Mid(TO_LOCATION, 7, 2))
            End If
'        Else
'            Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Date, "yyyymmdd"))
'            Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(TO_LOCATION, 1, 2))
'            Call UniCode_Conv(ITEMREC.ST_RETU, Mid(TO_LOCATION, 3, 2))
'            Call UniCode_Conv(ITEMREC.ST_REN, Mid(TO_LOCATION, 5, 2))
'            Call UniCode_Conv(ITEMREC.ST_DAN, Mid(TO_LOCATION, 7, 2))
'        End If
                                        '前回入庫棚
        Call UniCode_Conv(ITEMREC.BEF_SOKO, Mid(TO_LOCATION, 1, 2))
        Call UniCode_Conv(ITEMREC.BEF_RETU, Mid(TO_LOCATION, 3, 2))
        Call UniCode_Conv(ITEMREC.BEF_REN, Mid(TO_LOCATION, 5, 2))
        Call UniCode_Conv(ITEMREC.BEF_DAN, Mid(TO_LOCATION, 7, 2))
    End If
                                        '最終入庫日
    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, Format(Date, "yyyymmdd"))
                                        '最終入荷日付
    If StrConv(ITEMREC.LAST_INP_DT, vbUnicode) < NYUKA_DT Then
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, NYUKA_DT)
    End If
    '*------------------------------------------------------'品目マスタ出力
    RETRY_CNT = 0
    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                If RETRY_SU <> 0 Then
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                        Call File_Error(sts, BtOpUpdate, "品目マスタ", 0)
                        Nyuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    
                    End If
                    
                End If
                
                If MESG_FLG = 0 Then
'                    DoEvents
                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                        DoEvents                                                    '2016.01.26
                    End If                                                          '2016.01.26
                Else
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Nyuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case BtErrDEAD_LOCK
                Nyuko_Update_Proc = SYS_CANCEL
                Exit Function
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目マスタ")
                Nyuko_Update_Proc = SYS_ERR
                Exit Function
                        
        End Select
    Loop
'============================================================

    
    Nyuko_Update_Proc = False
    
    
End Function
