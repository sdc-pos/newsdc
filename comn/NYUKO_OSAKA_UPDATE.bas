Attribute VB_Name = "NYUKO_OSAKA_UPDATE"
Option Explicit

Public DAITO_SOKO_NO       As String * 2


Public Function Nyuko_OSAKA_Update_Proc(JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    TO_LOCATION As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    GYO_INS As String, _
                                    DEN_NO As String, _
                                    SEQ_NO As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional MEMO As String = "          ", _
                                    Optional MENU_NO As String = "  ") As Integer
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
'*          入庫作成(省略不可 0:更新 1:追加)
'*          伝票№(省略不可)
'*          SEQNO(省略不可　ﾚｺｰﾄﾞ№)
'*          リトライ(省略可 １桁目:1=画面メッセージ有 0:無，２桁目:リトライ回数(0～9 0:無限))
'*          メモ(省略可 履歴に出力するメモ内容)
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
    
    
    Nyuko_OSAKA_Update_Proc = True
                                                                      
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
                        Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                
                End If
                
                If MESG_FLG = 0 Then
'                    DoEvents                                                       '2016.01.26
                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                        DoEvents                                                    '2016.01.26
                    End If                                                          '2016.01.26
                Else
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Nyuko_OSAKA_Update_Proc = SYS_ERR
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
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                    
                        End If
                
                    End If
                
                    If MESG_FLG = 0 Then
'                        DoEvents                                                       '2016.01.26
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
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
            
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, "")                 '仕入先ｺｰﾄﾞ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "")                '仕入単価
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, "")                   '計上年月
            
            
            Call UniCode_Conv(ZAIKOREC.FILLER, "")
        Else
                                        '在庫数更新
                                                                        '在庫数
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) + SUMI_JITU_QTY, "00000000"))
                                                                    
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, "")                 '仕入先ｺｰﾄﾞ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "")                '仕入単価
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, "")                   '計上年月
        
        
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
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
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
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com, "在庫データ")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
                    Exit Function
                
            End Select
        Loop
'============================================================
    '*------------------------------------------------------'在庫移動歴出力
        sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                    TO_LOCATION, _
                                    JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    NYUKA_DT, _
                                    YOIN, _
                                    SUMI_JITU_QTY, _
                                    0, _
                                    ID, _
                                    TANTO_CODE, _
                                    RETRY, , _
                                    MEMO, _
                                    Ins_DateTime, _
                                    , , , MENU_NO, , , , , DEN_NO, Format(Now, "YYYYMMDD"))
        If sts Then
            Nyuko_OSAKA_Update_Proc = sts
            Exit Function
        End If
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
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
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
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
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
            
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, "")                 '仕入先ｺｰﾄﾞ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "")                '仕入単価
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, "")                   '計上年月
            
            
            Call UniCode_Conv(ZAIKOREC.FILLER, "")
        Else
                                        '在庫数更新
                                                                        '在庫数
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) + MI_JITU_QTY, "00000000"))
        
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, "")                 '仕入先ｺｰﾄﾞ
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "")                '仕入単価
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, "")                   '計上年月
        
        
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
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
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
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com, "在庫データ")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
                    Exit Function
                
            End Select
        Loop
'============================================================
    '*------------------------------------------------------'在庫移動歴出力
        sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                    TO_LOCATION, _
                                    JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    NYUKA_DT, _
                                    YOIN, _
                                    0, _
                                    MI_JITU_QTY, _
                                    ID, _
                                    TANTO_CODE, _
                                    RETRY, , _
                                    MEMO, _
                                    Ins_DateTime, _
                                    , , , MENU_NO, , , , , DEN_NO, Format(Now, "YYYYMMDD"))
        If sts Then
            Nyuko_OSAKA_Update_Proc = sts
            Exit Function
        End If
    End If
'============================================================
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
            Nyuko_OSAKA_Update_Proc = SYS_ERR
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
                        Nyuko_OSAKA_Update_Proc = SYS_CANCEL
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
                        Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目マスタ")
                Nyuko_OSAKA_Update_Proc = SYS_ERR
                Exit Function
                        
        End Select
    Loop
'============================================================

'============================================================   入荷データ更新／作成
    If GYO_INS = "9" Then '2007.09.12
    Else
    
        If GYO_INS = "0" Then
            com = BtOpUpdate
        
            Call UniCode_Conv(K0_Y_NYU_O.SEQ_NO, SEQ_NO)
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                        If MESG_FLG = 1 Then
                            Beep
                            MsgBox "入荷予定が存在しません。更新処理を中止します。", vbOKOnly, "確認入力"
                        End If
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "入荷予定", 0)
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                            
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                                '回数オーバー
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "入荷予定", 0)
                                Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        
                        End If
                        
                        If MESG_FLG = 0 Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<Y_NYU_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "入荷予定")
                        Nyuko_OSAKA_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            Loop
        
        
        
        
        Else
            com = BtOpInsert
            sts = BTRV(BtOpGetLast, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
            Select Case sts
                Case BtNoErr
                    SEQ_NO = StrConv(Y_NYU_O_REC.SEQ_NO, vbUnicode)
                Case BtErrEOF
                    SEQ_NO = "000"
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "入荷予定ﾃﾞｰﾀ")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        
            SEQ_NO = Format(CInt(SEQ_NO) + 1, "000")
                
            Call UniCode_Conv(Y_NYU_O_REC.JGYOBU, JGYOBU)
            Call UniCode_Conv(Y_NYU_O_REC.SOKO_NO, DAITO_SOKO_NO)
            Call UniCode_Conv(Y_NYU_O_REC.SEQ_NO, SEQ_NO)
            Call UniCode_Conv(Y_NYU_O_REC.NYUKO_YMD, Format(Now, "YYYYMMDD"))
            Call UniCode_Conv(Y_NYU_O_REC.DEN_NO, "000000")
        
            Call UniCode_Conv(Y_NYU_O_REC.MAKER_CODE, StrConv(ITEMREC.MAKER_CODE, vbUnicode))
            Call UniCode_Conv(Y_NYU_O_REC.NAIGAI, NAIGAI)
        
            Call UniCode_Conv(Y_NYU_O_REC.HIN_NO, HIN_GAI)
        
            Call UniCode_Conv(Y_NYU_O_REC.Y_SURYO, "00000000")
        
            Call UniCode_Conv(Y_NYU_O_REC.ORDER_NO, "")
        
            Call UniCode_Conv(Y_NYU_O_REC.FILLER, "")
        End If
    
        If Trim(DEN_NO) <> "" And DEN_NO <> "000000" Then
            Call UniCode_Conv(Y_NYU_O_REC.DEN_NO, DEN_NO)
        End If
        Call UniCode_Conv(Y_NYU_O_REC.J_SURYO, Format(MI_JITU_QTY + SUMI_JITU_QTY, "00000000"))
    
        Call UniCode_Conv(Y_NYU_O_REC.TANTO_CODE, TANTO_CODE)
        
        Call UniCode_Conv(Y_NYU_O_REC.KENPIN_F, "1")
        
        Call UniCode_Conv(Y_NYU_O_REC.WEL_ID, "")
        Call UniCode_Conv(Y_NYU_O_REC.PRG_ID, "")
        
            
        RETRY_CNT = 0
        Do
            sts = BTRV(com, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'リトライ回数チェック
                    If RETRY_SU <> 0 Then
                        
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                            '回数オーバー
                            Call File_Error(sts, com, "入荷予定", 0)
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
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
                        ans = MsgBox("他端末でデータ使用中です。<Y_NYU_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com, "入荷予定")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        Loop
    
    End If
    
    
    Nyuko_OSAKA_Update_Proc = False
    
    
End Function
