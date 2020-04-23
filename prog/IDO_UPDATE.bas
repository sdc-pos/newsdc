Attribute VB_Name = "IDO_UPDATE"
Option Explicit
'---------------------------------------------- *更新用在庫ワーク
'ポジショニング
Public wZAIKO_POS   As POSBLK
'データ・バッファ
Public wZAIKOREC    As ZAIKOREC_Tag
'キー・データ
Public K0_wZAIKO    As KEY0_ZAIKO
Public K1_wZAIKO    As KEY1_ZAIKO
Public K2_wZAIKO    As KEY2_ZAIKO

Public Function IDO_Update_Proc(JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    FROM_LOCATION As String, _
                                    TO_LOCATION As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional MEMO As String = "          ", _
                                    Optional MENU_NO As String = "  ") As Integer
'****************************************************
'*      「移動処理」在庫データ更新
'*
'*  在庫データの更新を行う。
'*  (引数の設定ミスはこちらではチェックしない)
'*  使用ﾌｧｲﾙ    :   在庫データ
'*                  品目マスタ
'*                  要因マスタ
'*                  在庫移動歴
'*                  倉庫マスタ
'*  引数：  事業部（省略不可）
'*          国内外（省略不可）
'*          品番外部(省略不可)
'*          FROM列（XXXXXXXX(倉庫№+列+連+段)省略不可）
'*          TO列（XXXXXXXX(倉庫№+列+連+段)省略不可）
'*          入荷日(YYYYMMDD 省略可 省略時FIFO)
'*          要因(省略不可)
'*          商品化済み実績数（何れか一方必須）
'*          未商品実績数　　（　　〃　　　　）
'*          ID(省略不可)
'*          担当者（省略不可）
'*          リトライ(省略可 １桁目:1=画面メッセージ有 0:無，２桁目:リトライ回数(0～9 0:無限))
'*          メモ(省略可 履歴に出力するメモ内容)
'*          ﾒﾆｭｰｸﾞﾙｰﾌﾟ（原価管理項目）  2006.01.30
'*  戻り値: false       :正常
'*          true        :継続可能な異常
'*          SYS_ERR     :継続できない異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'****************************************************
Dim sts         As Integer
Dim com         As Integer
Dim Upd_com     As Integer


Dim RETRY_CNT   As Integer
Dim MESG_FLG    As Integer
Dim RETRY_SU    As Integer
    
Dim ans         As Integer
    
Dim Zan_Qty     As Long
Dim WK_Qty      As Long
    
Dim DEN_NO      As String * 6
Dim TO_NAIGAI   As String * 1
    
Dim IDO_GOODS_ON_F  As String * 1
Dim IDO_GOODS_YMD   As String * 8
    
Dim Ins_DateTime    As String * 14              '2004.12.09


    IDO_Update_Proc = True
                                                                      
                                                                      
                                                                      
                                                                      
                                                                      
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                                        
    Ins_DateTime = Format(Now, "YYYYMMDDHHMMSS")    '2004.12.09
                                        
    '*------------------------------------------------------'品目ﾏｽﾀ（FROM側）の確保
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
                        IDO_Update_Proc = SYS_CANCEL
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
                        IDO_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                IDO_Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop


                                        
'*------------------------------------------------------'倉庫ﾏｽﾀ読込み
    Call UniCode_Conv(K0_SOKO.SOKO_NO, Mid(TO_LOCATION, 1, 2))
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound           '有るとまずいがエラーにしない
            Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_KASO)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "倉庫ﾏｽﾀ")
            IDO_Update_Proc = SYS_ERR
            Exit Function
    End Select

    IDO_GOODS_ON_F = "1"
    IDO_GOODS_YMD = ""
'    If Left(YOIN, 1) = ACT_IDO_OUT Then
    If JGYOBU <> SHIZAI Then
    '資材品は振替しない2006.01.10
        If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = "0" Then
            IDO_GOODS_ON_F = "0"
            IDO_GOODS_YMD = Format(Now, "YYYYMMDD")
        End If

    End If
'    End If
'============================================================
    If YOIN = YOIN_FURIKAE Then     '国内外振替は内外を反転
        If NAIGAI = NAIGAI_NAI Then
            TO_NAIGAI = NAIGAI_GAI
        Else
            TO_NAIGAI = NAIGAI_NAI
        End If
    Else
        TO_NAIGAI = NAIGAI
    End If
    
    
    If Len(Trim(NYUKA_DT)) = 0 Then
    '*------------------------------------------------------'入荷日指定無し 在庫データ読込み（FROM側の処理）
        '作業ﾛｸﾞ出力    '2008.08.06
        If P_SAGYO_LOG_OUTPUT_PROC(TANTO_CODE, _
                                    ID, _
                                    JGYOBU, _
                                    NAIGAI, _
                                    MENU_NO, _
                                    YOIN, _
                                    HIN_GAI, _
                                    SUMI_JITU_QTY, _
                                    MI_JITU_QTY, _
                                    FROM_LOCATION, _
                                    TO_LOCATION) Then
            Exit Function
        End If
        '*
        '*--------------------  商品化済みの処理
        If SUMI_JITU_QTY <> 0 Then
        
            Zan_Qty = SUMI_JITU_QTY

            Do

                Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   'FROM倉庫№
                Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      'FROM列
                Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       'FROM連
                Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       'FROM段
                Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '事業部
                Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '内外
                Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '品番（外部）
                Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "0")                       '商品／未商品
                Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")                        '入荷日

                RETRY_CNT = 0
                Do
                    sts = BTRV(BtOpGetGreater + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                                                '棚＋品ブレーク
                            If FROM_LOCATION <> (StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)) Or _
                                JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                                NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                                Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Or _
                                StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> "0" Then
                                If MESG_FLG = 1 Then
                                    Beep
                                    MsgBox "在庫数が不足しています。更新処理を中止します。", vbOKOnly, "確認入力"
                                End If
                                Exit Function
                            End If
                        
                            If Zan_Qty < CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                                Upd_com = BtOpUpdate
                                WK_Qty = Zan_Qty
                            Else
                                Upd_com = BtOpDelete
                                WK_Qty = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                            End If

                            Exit Do
                        Case BtErrEOF

                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "在庫数が不足しています。更新処理を中止します。", vbOKOnly, "確認入力"
                            End If
                            Exit Function

                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                    Call File_Error(sts, com + BtSNoWait, "在庫データ", 0)
                                    IDO_Update_Proc = SYS_CANCEL
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
                                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    IDO_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case Else
                            Call File_Error(sts, com + BtSNoWait, "在庫データ")
                            IDO_Update_Proc = SYS_ERR
                            Exit Function
                    End Select

                Loop

                If Upd_com = BtOpUpdate Then
                                                                                '有効在庫数
                    Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - Zan_Qty, "00000000"))
                    Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '排他フラグ
                    Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '使用中子機ID
                    Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '使用中プログラム
                
                End If


                RETRY_CNT = 0
    '*------------------------------------------------------'在庫データ出力
                Do
                    sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                    Call File_Error(sts, Upd_com, "在庫データ", 0)
                                    IDO_Update_Proc = SYS_CANCEL
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
                                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    IDO_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case Else
                            Call File_Error(sts, Upd_com, "在庫データ")
                            IDO_Update_Proc = SYS_ERR
                            Exit Function

                    End Select
                Loop
'============================================================
        '*------------------------------------------------------'入荷日指定無し 在庫データ読込み（TO側の処理）
                Call UniCode_Conv(K0_wZAIKO.SOKO_NO, Mid(TO_LOCATION, 1, 2))    'TO倉庫№
                Call UniCode_Conv(K0_wZAIKO.Retu, Mid(TO_LOCATION, 3, 2))       'TO列
                Call UniCode_Conv(K0_wZAIKO.Ren, Mid(TO_LOCATION, 5, 2))        'TO連
                Call UniCode_Conv(K0_wZAIKO.Dan, Mid(TO_LOCATION, 7, 2))        'TO段
                Call UniCode_Conv(K0_wZAIKO.JGYOBU, JGYOBU)                     '事業部
                Call UniCode_Conv(K0_wZAIKO.NAIGAI, TO_NAIGAI)                  '内外
                Call UniCode_Conv(K0_wZAIKO.HIN_GAI, HIN_GAI)                   '品番（外部）
                Call UniCode_Conv(K0_wZAIKO.GOODS_ON, "0")                      '商品／未商品
                                                                                '入荷日
                Call UniCode_Conv(K0_wZAIKO.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))

                RETRY_CNT = 0
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                    Select Case sts
                        Case BtNoErr

                            Upd_com = BtOpUpdate
                            Exit Do
                        Case BtErrKeyNotFound
                            Upd_com = BtOpInsert
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ", 0)
                                    IDO_Update_Proc = SYS_CANCEL
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
                                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    IDO_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                            IDO_Update_Proc = SYS_ERR
                            Exit Function
                    End Select

                Loop

                If Upd_com = BtOpInsert Then
                                                    '新規追加
                    Call UniCode_Conv(wZAIKOREC.SOKO_NO, Mid(TO_LOCATION, 1, 2))    '倉庫№
                    Call UniCode_Conv(wZAIKOREC.Retu, Mid(TO_LOCATION, 3, 2))       '列
                    Call UniCode_Conv(wZAIKOREC.Ren, Mid(TO_LOCATION, 5, 2))        '連
                    Call UniCode_Conv(wZAIKOREC.Dan, Mid(TO_LOCATION, 7, 2))        '段
                    Call UniCode_Conv(wZAIKOREC.JGYOBU, JGYOBU)                     '事業部
                    Call UniCode_Conv(wZAIKOREC.NAIGAI, TO_NAIGAI)                  '内外
                    Call UniCode_Conv(wZAIKOREC.HIN_GAI, HIN_GAI)                   '品番（外部）
                    Call UniCode_Conv(wZAIKOREC.GOODS_ON, "0")                      '商品／未商品
                                                                                    '入荷日
                    Call UniCode_Conv(wZAIKOREC.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))
                    Call UniCode_Conv(wZAIKOREC.NYUKO_DT, Format(Date, "YYYYMMDD")) '入庫日
                                                                                    '品番（内部）
                    Call UniCode_Conv(wZAIKOREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                                                                    '有効在庫数
                    Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(WK_Qty, "00000000"))
                    Call UniCode_Conv(wZAIKOREC.LOCK_F, LOCK_OFF)                   '排他フラグ
                    Call UniCode_Conv(wZAIKOREC.WEL_ID, "")                         '使用中子機ID
                    Call UniCode_Conv(wZAIKOREC.PRG_ID, "")                         '使用中ﾌﾟﾛｸﾞﾗﾑ
                                                                                    '仕入先ｺｰﾄﾞ2006.01.08
                    Call UniCode_Conv(wZAIKOREC.SHIIRE_CODE, StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
                                                                                    '仕入先単価2006.01.08
                    Call UniCode_Conv(wZAIKOREC.SHIIRE_TANKA, StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))
                                                                                    '計上年月2006.01.08
                    Call UniCode_Conv(wZAIKOREC.KEIJYO_YM, StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode))
                    
                    
                    
                    
                    
                    
    '----------------   2010.07.08 ▽
                    Call UniCode_Conv(wZAIKOREC.GENSANKOKU, StrConv(ZAIKOREC.GENSANKOKU, vbUnicode))
                    Call UniCode_Conv(wZAIKOREC.SHIIRE_WORK_CENTER, StrConv(ZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode))
                    Call UniCode_Conv(wZAIKOREC.ID_NO2, StrConv(ZAIKOREC.ID_NO2, vbUnicode))
                    Call UniCode_Conv(wZAIKOREC.YOSAN_FROM, StrConv(ZAIKOREC.YOSAN_FROM, vbUnicode))
                    Call UniCode_Conv(wZAIKOREC.YOSAN_TO, StrConv(ZAIKOREC.YOSAN_TO, vbUnicode))
    '----------------   2010.07.08 △
                    
                    
                    
                    Call UniCode_Conv(wZAIKOREC.FILLER, StrConv(ZAIKOREC.FILLER, vbUnicode))
                Else
                                                '在庫数更新
                    Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(wZAIKOREC.YUKO_Z_QTY, vbUnicode)) + WK_Qty, "00000000"))
                End If

                RETRY_CNT = 0
    '*------------------------------------------------------'在庫データ出力
                Do
                    sts = BTRV(Upd_com, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                    Call File_Error(sts, Upd_com, "在庫データ", 0)
                                    IDO_Update_Proc = SYS_CANCEL
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
                                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    IDO_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case Else
                            Call File_Error(sts, Upd_com, "在庫データ")
                            IDO_Update_Proc = SYS_ERR
                            Exit Function

                    End Select
                Loop
            '*------------------------------------------------------'在庫移動歴出力
                If YOIN = YOIN_FURIKAE Then
'2004.06.11                    sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, Space(8), JGYOBU, NAIGAI, HIN_GAI, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), YOIN_FURIKAE_OUT, SUMI_JITU_QTY, 0, ID, TANTO_CODE, RETRY, , MEMO)
                    sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                                    Space(8), _
                                                    JGYOBU, _
                                                    NAIGAI, _
                                                    HIN_GAI, _
                                                    StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                                    YOIN_FURIKAE_OUT, _
                                                    WK_Qty, 0, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1, StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                    If sts Then
                        IDO_Update_Proc = sts
                        Exit Function
                    End If
'2004.06.11                    sts = IDOREKI_OUTPUT_PROC(Space(8), TO_LOCATION, JGYOBU, TO_NAIGAI, HIN_GAI, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), YOIN_FURIKAE_IN, SUMI_JITU_QTY, 0, ID, TANTO_CODE, RETRY, , MEMO)
                    sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                                    TO_LOCATION, _
                                                    JGYOBU, _
                                                    TO_NAIGAI, _
                                                    HIN_GAI, _
                                                    StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                                    YOIN_FURIKAE_IN, _
                                                    WK_Qty, 0, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1, StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                    If sts Then
                        IDO_Update_Proc = sts
                        Exit Function
                    End If
                Else
'2004.06.11                    sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, TO_LOCATION, JGYOBU, NAIGAI, HIN_GAI, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), YOIN, SUMI_JITU_QTY, 0, ID, TANTO_CODE, RETRY, , MEMO)
                    sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                                    TO_LOCATION, _
                                                    JGYOBU, _
                                                    NAIGAI, _
                                                    HIN_GAI, _
                                                    StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                                    YOIN, _
                                                    WK_Qty, 0, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1, StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                    If sts Then
                        IDO_Update_Proc = sts
                        Exit Function
                    End If
                End If

                Zan_Qty = Zan_Qty - WK_Qty

                If Zan_Qty <= 0 Then
                    Exit Do                     '引き落とし終了
                End If

            Loop
                    
        End If
'================================================================================
        '*
        '*--------------------  未商品化の処理
        If MI_JITU_QTY <> 0 Then
        
            Zan_Qty = MI_JITU_QTY

            Do

                Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   'FROM倉庫№
                Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      'FROM列
                Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       'FROM連
                Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       'FROM段
                Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '事業部
                Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '内外
                Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '品番（外部）
                Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "1")                       '商品／未商品
                Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")                        '入荷日

                RETRY_CNT = 0
                Do
                    sts = BTRV(BtOpGetGreater + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                                                '棚＋品ブレーク
                            If FROM_LOCATION <> (StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)) Or _
                                JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                                NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                                Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Or _
                                StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> "1" Then
                                If MESG_FLG = 1 Then
                                    Beep
                                    MsgBox "在庫数が不足しています。更新処理を中止します。", vbOKOnly, "確認入力"
                                End If
                                Exit Function
                            End If
                        
                            If Zan_Qty < CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                                Upd_com = BtOpUpdate
                                WK_Qty = Zan_Qty
                            Else
                                Upd_com = BtOpDelete
                                WK_Qty = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                            End If

                            Exit Do
                        Case BtErrEOF

                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "在庫数が不足しています。更新処理を中止します。", vbOKOnly, "確認入力"
                            End If
                            Exit Function

                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                    Call File_Error(sts, com + BtSNoWait, "在庫データ", 0)
                                    IDO_Update_Proc = SYS_CANCEL
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
                                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    IDO_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case Else
                            Call File_Error(sts, com + BtSNoWait, "在庫データ")
                            IDO_Update_Proc = SYS_ERR
                            Exit Function
                    End Select

                Loop

                If Upd_com = BtOpUpdate Then
                                                                                '有効在庫数
                    Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - Zan_Qty, "00000000"))
                    Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '排他フラグ
                    Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '使用中子機ID
                    Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '使用中プログラム
                
                End If


                RETRY_CNT = 0
    '*------------------------------------------------------'在庫データ出力
                Do
                    sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                    Call File_Error(sts, Upd_com, "在庫データ", 0)
                                    IDO_Update_Proc = SYS_CANCEL
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
                                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    IDO_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case Else
                            Call File_Error(sts, Upd_com, "在庫データ")
                            IDO_Update_Proc = SYS_ERR
                            Exit Function

                    End Select
                Loop
'============================================================
        '*------------------------------------------------------'入荷日指定無し 在庫データ読込み（TO側の処理）
                Call UniCode_Conv(K0_wZAIKO.SOKO_NO, Mid(TO_LOCATION, 1, 2))    'TO倉庫№
                Call UniCode_Conv(K0_wZAIKO.Retu, Mid(TO_LOCATION, 3, 2))       'TO列
                Call UniCode_Conv(K0_wZAIKO.Ren, Mid(TO_LOCATION, 5, 2))        'TO連
                Call UniCode_Conv(K0_wZAIKO.Dan, Mid(TO_LOCATION, 7, 2))        'TO段
                Call UniCode_Conv(K0_wZAIKO.JGYOBU, JGYOBU)                     '事業部
                Call UniCode_Conv(K0_wZAIKO.NAIGAI, TO_NAIGAI)                  '内外
                Call UniCode_Conv(K0_wZAIKO.HIN_GAI, HIN_GAI)                   '品番（外部）
                Call UniCode_Conv(K0_wZAIKO.GOODS_ON, IDO_GOODS_ON_F)           '商品／未商品
                                                                                '入荷日
                Call UniCode_Conv(K0_wZAIKO.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))

                RETRY_CNT = 0
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                    Select Case sts
                        Case BtNoErr

                            Upd_com = BtOpUpdate
                            Exit Do
                        Case BtErrKeyNotFound
                            Upd_com = BtOpInsert
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ", 0)
                                    IDO_Update_Proc = SYS_CANCEL
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
                                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    IDO_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                            IDO_Update_Proc = SYS_ERR
                            Exit Function
                    End Select

                Loop

                If Upd_com = BtOpInsert Then
                                                    '新規追加
                    Call UniCode_Conv(wZAIKOREC.SOKO_NO, Mid(TO_LOCATION, 1, 2))    '倉庫№
                    Call UniCode_Conv(wZAIKOREC.Retu, Mid(TO_LOCATION, 3, 2))       '列
                    Call UniCode_Conv(wZAIKOREC.Ren, Mid(TO_LOCATION, 5, 2))        '連
                    Call UniCode_Conv(wZAIKOREC.Dan, Mid(TO_LOCATION, 7, 2))        '段
                    Call UniCode_Conv(wZAIKOREC.JGYOBU, JGYOBU)                     '事業部
                    Call UniCode_Conv(wZAIKOREC.NAIGAI, TO_NAIGAI)                  '内外
                    Call UniCode_Conv(wZAIKOREC.HIN_GAI, HIN_GAI)                   '品番（外部）
                    Call UniCode_Conv(wZAIKOREC.GOODS_ON, IDO_GOODS_ON_F)           '商品／未商品
                                                                                    '入荷日
                    Call UniCode_Conv(wZAIKOREC.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))
                    Call UniCode_Conv(wZAIKOREC.NYUKO_DT, Format(Date, "YYYYMMDD")) '入庫日
                                                                                    '品番（内部）
                    Call UniCode_Conv(wZAIKOREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                                                                    '有効在庫数
                    Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(WK_Qty, "00000000"))
                    Call UniCode_Conv(wZAIKOREC.LOCK_F, LOCK_OFF)                   '排他フラグ
                    Call UniCode_Conv(wZAIKOREC.WEL_ID, "")                         '使用中子機ID
                    Call UniCode_Conv(wZAIKOREC.PRG_ID, "")                         '使用中ﾌﾟﾛｸﾞﾗﾑ

                    Call UniCode_Conv(wZAIKOREC.GOODS_YMD, IDO_GOODS_YMD)           '商品化日
                    
                                                                                    '仕入先ｺｰﾄﾞ2006.01.08
                    Call UniCode_Conv(wZAIKOREC.SHIIRE_CODE, StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
                                                                                    '仕入先単価2006.01.08
                    Call UniCode_Conv(wZAIKOREC.SHIIRE_TANKA, StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))
                                                                                    '計上年月2006.01.08
                    Call UniCode_Conv(wZAIKOREC.KEIJYO_YM, StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode))
                    
                    
    '----------------   2010.07.08 ▽
                    Call UniCode_Conv(wZAIKOREC.GENSANKOKU, StrConv(ZAIKOREC.GENSANKOKU, vbUnicode))
                    Call UniCode_Conv(wZAIKOREC.SHIIRE_WORK_CENTER, StrConv(ZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode))
                    Call UniCode_Conv(wZAIKOREC.ID_NO2, StrConv(ZAIKOREC.ID_NO2, vbUnicode))
                    Call UniCode_Conv(wZAIKOREC.YOSAN_FROM, StrConv(ZAIKOREC.YOSAN_FROM, vbUnicode))
                    Call UniCode_Conv(wZAIKOREC.YOSAN_TO, StrConv(ZAIKOREC.YOSAN_TO, vbUnicode))
    '----------------   2010.07.08 △
                    
                    
                    
                    Call UniCode_Conv(wZAIKOREC.FILLER, StrConv(ZAIKOREC.FILLER, vbUnicode))
                Else
                                                '在庫数更新
                    Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(wZAIKOREC.YUKO_Z_QTY, vbUnicode)) + WK_Qty, "00000000"))
                End If

                RETRY_CNT = 0
    '*------------------------------------------------------'在庫データ出力
                Do
                    sts = BTRV(Upd_com, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                    Call File_Error(sts, Upd_com, "在庫データ", 0)
                                    IDO_Update_Proc = SYS_CANCEL
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
                                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    IDO_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case Else
                            Call File_Error(sts, Upd_com, "在庫データ")
                            IDO_Update_Proc = SYS_ERR
                            Exit Function

                    End Select
                Loop
            '*------------------------------------------------------'在庫移動歴出力
                If YOIN = YOIN_FURIKAE Then
'2004.06.11                    sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, Space(8), JGYOBU, NAIGAI, HIN_GAI, NYUKA_DT, YOIN_FURIKAE_OUT, 0, MI_JITU_QTY, ID, TANTO_CODE, RETRY, , MEMO)
                    sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                                    Space(8), _
                                                    JGYOBU, _
                                                    NAIGAI, _
                                                    HIN_GAI, _
                                                    StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                                    YOIN_FURIKAE_OUT, _
                                                    0, WK_Qty, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1, StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                    If sts Then
                        IDO_Update_Proc = sts
                        Exit Function
                    End If
                    
'2004.06.11                    sts = IDOREKI_OUTPUT_PROC(Space(8), TO_LOCATION, JGYOBU, TO_NAIGAI, HIN_GAI, NYUKA_DT, YOIN_FURIKAE_IN, 0, MI_JITU_QTY, ID, TANTO_CODE, RETRY, , MEMO)
                    sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                                    TO_LOCATION, _
                                                    JGYOBU, _
                                                    TO_NAIGAI, _
                                                    HIN_GAI, _
                                                    StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                                    YOIN_FURIKAE_IN, _
                                                    0, WK_Qty, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1, StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                    If sts Then
                        IDO_Update_Proc = sts
                        Exit Function
                    End If
                Else
                    If IDO_GOODS_ON_F = "0" Then
'2004.06.11                        sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, TO_LOCATION, JGYOBU, NAIGAI, HIN_GAI, NYUKA_DT, YOIN, MI_JITU_QTY, 0, ID, TANTO_CODE, RETRY, , MEMO & "商品振替")
                        sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                                    TO_LOCATION, _
                                                    JGYOBU, _
                                                    NAIGAI, _
                                                    HIN_GAI, _
                                                    StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                                    YOIN, _
                                                    WK_Qty, 0, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO & "商品振替", _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1, StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                    Else
'2004.06.11                        sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, TO_LOCATION, JGYOBU, NAIGAI, HIN_GAI, NYUKA_DT, YOIN, 0, MI_JITU_QTY, ID, TANTO_CODE, RETRY, , MEMO)
                        sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                                    TO_LOCATION, _
                                                    JGYOBU, _
                                                    NAIGAI, _
                                                    HIN_GAI, _
                                                    StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                                    YOIN, _
                                                    0, WK_Qty, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1, StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                    End If
                    
                    If sts Then
                        IDO_Update_Proc = sts
                        Exit Function
                    End If
                End If

                Zan_Qty = Zan_Qty - WK_Qty

                If Zan_Qty <= 0 Then
                    Exit Do                     '引き落とし終了
                End If

            Loop
                    
        End If
    
    Else
    '*------------------------------------------------------'入荷日指定有り 在庫データ読込み（FROM側の処理）
        '
        '----------------------------------- 商品化済み
        If SUMI_JITU_QTY <> 0 Then
        
            Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   'FROM倉庫№
            Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      'FROM列
            Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       'FROM連
            Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       'FROM段
            Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '事業部
            Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '内外
            Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '品番（外部）
            Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "0")                       '商品／未商品
            Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, NYUKA_DT)                  '入荷日
                                                                    
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        If SUMI_JITU_QTY > CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                            If MESG_FLG = 0 Then
                                Beep
                                MsgBox "在庫数が不足しています。更新処理を中止します。", vbOKOnly, "確認入力"
                            End If
                            Exit Function
                        Else
                            If SUMI_JITU_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                                Upd_com = BtOpDelete
                            Else
                                Upd_com = BtOpUpdate
                            End If
                        End If
                    
                        Exit Do
                    Case BtErrKeyNotFound
                        If MESG_FLG = 0 Then
                            Beep
                            MsgBox "在庫データが存在しません。更新処理を中止します。", vbOKOnly, "確認入力"
                        End If
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ", 0)
                                IDO_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                IDO_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                        IDO_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            
            Loop
        
            If Upd_com = BtOpUpdate Then
                                                                '有効在庫数
                Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - SUMI_JITU_QTY, "00000000"))
                Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)        '排他フラグ
                Call UniCode_Conv(ZAIKOREC.WEL_ID, "")              '使用中子機ID
                Call UniCode_Conv(ZAIKOREC.PRG_ID, "")              '使用中プログラム
            
            End If
        
            RETRY_CNT = 0
    '*------------------------------------------------------'在庫データ出力
            Do
                sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, Upd_com, "在庫データ", 0)
                                IDO_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                           If ans = vbCancel Then
                                IDO_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, Upd_com, "在庫データ")
                        IDO_Update_Proc = SYS_ERR
                        Exit Function
                           
                End Select
            Loop
'============================================================
    '*------------------------------------------------------'入荷日指定有り 在庫データ読込み（TO側の処理）
            Call UniCode_Conv(K0_wZAIKO.SOKO_NO, Mid(TO_LOCATION, 1, 2))    'TO倉庫№
            Call UniCode_Conv(K0_wZAIKO.Retu, Mid(TO_LOCATION, 3, 2))       'TO列
            Call UniCode_Conv(K0_wZAIKO.Ren, Mid(TO_LOCATION, 5, 2))        'TO連
            Call UniCode_Conv(K0_wZAIKO.Dan, Mid(TO_LOCATION, 7, 2))        'TO段
            Call UniCode_Conv(K0_wZAIKO.JGYOBU, JGYOBU)                     '事業部
            Call UniCode_Conv(K0_wZAIKO.NAIGAI, TO_NAIGAI)                  '内外
            Call UniCode_Conv(K0_wZAIKO.HIN_GAI, HIN_GAI)                   '品番（外部）
            Call UniCode_Conv(K0_wZAIKO.GOODS_ON, "0")                      '商品／未商品
            Call UniCode_Conv(K0_wZAIKO.NYUKA_DT, NYUKA_DT)                 '入荷日
                                                                    
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Upd_com = BtOpUpdate
                        Exit Do
                    Case BtErrKeyNotFound
                        Upd_com = BtOpInsert
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, BtOpGetEqual + 200, "在庫データ", 0)
                                IDO_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                               IDO_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                        IDO_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            
            Loop
    
            If Upd_com = BtOpInsert Then
                                        '新規追加
                Call UniCode_Conv(wZAIKOREC.SOKO_NO, Mid(TO_LOCATION, 1, 2))    '倉庫№
                Call UniCode_Conv(wZAIKOREC.Retu, Mid(TO_LOCATION, 3, 2))       '列
                Call UniCode_Conv(wZAIKOREC.Ren, Mid(TO_LOCATION, 5, 2))        '連
                Call UniCode_Conv(wZAIKOREC.Dan, Mid(TO_LOCATION, 7, 2))        '段
                Call UniCode_Conv(wZAIKOREC.JGYOBU, JGYOBU)                     '事業部
                Call UniCode_Conv(wZAIKOREC.NAIGAI, TO_NAIGAI)                  '内外
                Call UniCode_Conv(wZAIKOREC.HIN_GAI, HIN_GAI)                   '品番（外部）
                Call UniCode_Conv(wZAIKOREC.GOODS_ON, "0")                      '商品／未商品
                Call UniCode_Conv(wZAIKOREC.NYUKA_DT, NYUKA_DT)                 '入荷日
                                                                                '入庫日
                Call UniCode_Conv(wZAIKOREC.NYUKO_DT, Format(Date, "YYYYMMDD"))
                                                                                '品番（内部）
                Call UniCode_Conv(wZAIKOREC.HIN_NAI, StrConv(ZAIKOREC.HIN_NAI, vbUnicode))
                                                                                '有効在庫数クリアー
                Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(SUMI_JITU_QTY, "00000000"))
            
                Call UniCode_Conv(wZAIKOREC.LOCK_F, LOCK_OFF)                   '排他フラグ
                Call UniCode_Conv(wZAIKOREC.WEL_ID, "")                         '使用中子機ID
                Call UniCode_Conv(wZAIKOREC.PRG_ID, "")                         '使用中ﾌﾟﾛｸﾞﾗﾑ
                        
                Call UniCode_Conv(wZAIKOREC.GOODS_YMD, "")                      '商品化日
                
                                                                                '仕入先ｺｰﾄﾞ2006.01.08
                Call UniCode_Conv(wZAIKOREC.SHIIRE_CODE, StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
                                                                                '仕入先単価2006.01.08
                Call UniCode_Conv(wZAIKOREC.SHIIRE_TANKA, StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))
                                                                                '計上年月2006.01.08
                Call UniCode_Conv(wZAIKOREC.KEIJYO_YM, StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode))
                
                
                
    '----------------   2010.07.08 ▽
                Call UniCode_Conv(wZAIKOREC.GENSANKOKU, StrConv(ZAIKOREC.GENSANKOKU, vbUnicode))
                Call UniCode_Conv(wZAIKOREC.SHIIRE_WORK_CENTER, StrConv(ZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode))
                Call UniCode_Conv(wZAIKOREC.ID_NO2, StrConv(ZAIKOREC.ID_NO2, vbUnicode))
                Call UniCode_Conv(wZAIKOREC.YOSAN_FROM, StrConv(ZAIKOREC.YOSAN_FROM, vbUnicode))
                Call UniCode_Conv(wZAIKOREC.YOSAN_TO, StrConv(ZAIKOREC.YOSAN_TO, vbUnicode))
    '----------------   2010.07.08 △
                    
                    
                    
                Call UniCode_Conv(wZAIKOREC.FILLER, StrConv(ZAIKOREC.FILLER, vbUnicode))
            Else
                                        '在庫数更新
                Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(wZAIKOREC.YUKO_Z_QTY, vbUnicode)) + SUMI_JITU_QTY, "00000000"))
            End If
        
            RETRY_CNT = 0
    '*------------------------------------------------------'在庫データ出力
            Do
                sts = BTRV(Upd_com, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, Upd_com, "在庫データ", 0)
                                IDO_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                IDO_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, Upd_com, "在庫データ")
                        IDO_Update_Proc = SYS_ERR
                        Exit Function
                        
                End Select
            Loop
'============================================================
    '*------------------------------------------------------'在庫移動歴出力
            If YOIN = YOIN_FURIKAE Then
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                                    Space(8), _
                                                    JGYOBU, _
                                                    NAIGAI, _
                                                    HIN_GAI, _
                                                    NYUKA_DT, _
                                                    YOIN_FURIKAE_OUT, _
                                                    SUMI_JITU_QTY, 0, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, "", MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , , StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                If sts Then
                    IDO_Update_Proc = sts
                    Exit Function
                End If
                sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                                    TO_LOCATION, _
                                                    JGYOBU, _
                                                    TO_NAIGAI, _
                                                    HIN_GAI, _
                                                    NYUKA_DT, _
                                                    YOIN_FURIKAE_IN, _
                                                    SUMI_JITU_QTY, 0, _
                                                    ID, _
                                                    TANTO_CODE, RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , , StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                If sts Then
                    IDO_Update_Proc = sts
                    Exit Function
                End If
            Else
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                                    TO_LOCATION, _
                                                    JGYOBU, _
                                                    NAIGAI, _
                                                    HIN_GAI, _
                                                    NYUKA_DT, _
                                                    YOIN, _
                                                    SUMI_JITU_QTY, 0, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , , StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                If sts Then
                    IDO_Update_Proc = sts
                    Exit Function
                End If
            End If
        End If
        '===================================================================
        '
        '----------------------------------- 未商品
        If MI_JITU_QTY <> 0 Then
        
            Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   'FROM倉庫№
            Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      'FROM列
            Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       'FROM連
            Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       'FROM段
            Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '事業部
            Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '内外
            Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '品番（外部）
            Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "1")                       '商品／未商品
            Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, NYUKA_DT)                  '入荷日
                                                                    
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        If MI_JITU_QTY > CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                            If MESG_FLG = 0 Then
                                Beep
                                MsgBox "在庫数が不足しています。更新処理を中止します。", vbOKOnly, "確認入力"
                            End If
                            Exit Function
                        Else
                            If MI_JITU_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                                Upd_com = BtOpDelete
                            Else
                                Upd_com = BtOpUpdate
                            End If
                        End If
                    
                        Exit Do
                    Case BtErrKeyNotFound
                        If MESG_FLG = 0 Then
                            Beep
                            MsgBox "在庫データが存在しません。更新処理を中止します。", vbOKOnly, "確認入力"
                        End If
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ", 0)
                                IDO_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                IDO_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                        IDO_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            
            Loop
        
            If Upd_com = BtOpUpdate Then
                                                                '有効在庫数
                Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - MI_JITU_QTY, "00000000"))
                Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)        '排他フラグ
                Call UniCode_Conv(ZAIKOREC.WEL_ID, "")              '使用中子機ID
                Call UniCode_Conv(ZAIKOREC.PRG_ID, "")              '使用中プログラム
            
            End If
        
            RETRY_CNT = 0
    '*------------------------------------------------------'在庫データ出力
            Do
                sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, Upd_com, "在庫データ", 0)
                                IDO_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                           If ans = vbCancel Then
                                IDO_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, Upd_com, "在庫データ")
                        IDO_Update_Proc = SYS_ERR
                        Exit Function
                           
                End Select
            Loop
'============================================================
    '*------------------------------------------------------'入荷日指定有り 在庫データ読込み（TO側の処理）
            Call UniCode_Conv(K0_wZAIKO.SOKO_NO, Mid(TO_LOCATION, 1, 2))    'TO倉庫№
            Call UniCode_Conv(K0_wZAIKO.Retu, Mid(TO_LOCATION, 3, 2))       'TO列
            Call UniCode_Conv(K0_wZAIKO.Ren, Mid(TO_LOCATION, 5, 2))        'TO連
            Call UniCode_Conv(K0_wZAIKO.Dan, Mid(TO_LOCATION, 7, 2))        'TO段
            Call UniCode_Conv(K0_wZAIKO.JGYOBU, JGYOBU)                     '事業部
            Call UniCode_Conv(K0_wZAIKO.NAIGAI, TO_NAIGAI)                  '内外
            Call UniCode_Conv(K0_wZAIKO.HIN_GAI, HIN_GAI)                   '品番（外部）
            Call UniCode_Conv(K0_wZAIKO.GOODS_ON, IDO_GOODS_ON_F)           '商品／未商品
            Call UniCode_Conv(K0_wZAIKO.NYUKA_DT, NYUKA_DT)                 '入荷日
                                                                    
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Upd_com = BtOpUpdate
                        Exit Do
                    Case BtErrKeyNotFound
                        Upd_com = BtOpInsert
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, BtOpGetEqual + 200, "在庫データ", 0)
                                IDO_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                               IDO_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                        IDO_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            
            Loop
    
            If Upd_com = BtOpInsert Then
                                        '新規追加
                Call UniCode_Conv(wZAIKOREC.SOKO_NO, Mid(TO_LOCATION, 1, 2))    '倉庫№
                Call UniCode_Conv(wZAIKOREC.Retu, Mid(TO_LOCATION, 3, 2))       '列
                Call UniCode_Conv(wZAIKOREC.Ren, Mid(TO_LOCATION, 5, 2))        '連
                Call UniCode_Conv(wZAIKOREC.Dan, Mid(TO_LOCATION, 7, 2))        '段
                Call UniCode_Conv(wZAIKOREC.JGYOBU, JGYOBU)                     '事業部
                Call UniCode_Conv(wZAIKOREC.NAIGAI, TO_NAIGAI)                  '内外
                Call UniCode_Conv(wZAIKOREC.HIN_GAI, HIN_GAI)                   '品番（外部）
                Call UniCode_Conv(wZAIKOREC.GOODS_ON, IDO_GOODS_ON_F)           '商品／未商品
                Call UniCode_Conv(wZAIKOREC.NYUKA_DT, NYUKA_DT)                 '入荷日
                                                                                '入庫日
                Call UniCode_Conv(wZAIKOREC.NYUKO_DT, Format(Date, "YYYYMMDD"))
                                                                                '品番（内部）
                Call UniCode_Conv(wZAIKOREC.HIN_NAI, StrConv(ZAIKOREC.HIN_NAI, vbUnicode))
                                                                                '有効在庫数クリアー
                Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(MI_JITU_QTY, "00000000"))
            
                Call UniCode_Conv(wZAIKOREC.LOCK_F, LOCK_OFF)                   '排他フラグ
                Call UniCode_Conv(wZAIKOREC.WEL_ID, "")                         '使用中子機ID
                Call UniCode_Conv(wZAIKOREC.PRG_ID, "")                         '使用中ﾌﾟﾛｸﾞﾗﾑ
                        
                Call UniCode_Conv(wZAIKOREC.GOODS_YMD, IDO_GOODS_YMD)           '商品化日
                
                                                                                '仕入先ｺｰﾄﾞ2006.01.08
                Call UniCode_Conv(wZAIKOREC.SHIIRE_CODE, StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
                                                                                '仕入先単価2006.01.08
                Call UniCode_Conv(wZAIKOREC.SHIIRE_TANKA, StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))
                                                                                '計上年月2006.01.08
                Call UniCode_Conv(wZAIKOREC.KEIJYO_YM, StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode))
    '----------------   2010.07.08 ▽
                Call UniCode_Conv(wZAIKOREC.GENSANKOKU, StrConv(ZAIKOREC.GENSANKOKU, vbUnicode))
                Call UniCode_Conv(wZAIKOREC.SHIIRE_WORK_CENTER, StrConv(ZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode))
                Call UniCode_Conv(wZAIKOREC.ID_NO2, StrConv(ZAIKOREC.ID_NO2, vbUnicode))
                Call UniCode_Conv(wZAIKOREC.YOSAN_FROM, StrConv(ZAIKOREC.YOSAN_FROM, vbUnicode))
                Call UniCode_Conv(wZAIKOREC.YOSAN_TO, StrConv(ZAIKOREC.YOSAN_TO, vbUnicode))
    '----------------   2010.07.08 △
                    
                    
                    
                Call UniCode_Conv(wZAIKOREC.FILLER, StrConv(ZAIKOREC.FILLER, vbUnicode))
                
                
            Else
                                        '在庫数更新
                Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(wZAIKOREC.YUKO_Z_QTY, vbUnicode)) + MI_JITU_QTY, "00000000"))
            End If
        
            RETRY_CNT = 0
    '*------------------------------------------------------'在庫データ出力
            Do
                sts = BTRV(Upd_com, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, Upd_com, "在庫データ", 0)
                                IDO_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                IDO_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, Upd_com, "在庫データ")
                        IDO_Update_Proc = SYS_ERR
                        Exit Function
                        
                End Select
            Loop
'============================================================
    '*------------------------------------------------------'在庫移動歴出力
            If YOIN = YOIN_FURIKAE Then
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                                    Space(8), _
                                                    JGYOBU, _
                                                    NAIGAI, _
                                                    HIN_GAI, _
                                                    NYUKA_DT, _
                                                    YOIN_FURIKAE_OUT, _
                                                    0, MI_JITU_QTY, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , , StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                If sts Then
                    IDO_Update_Proc = sts
                    Exit Function
                End If
                sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                                    TO_LOCATION, _
                                                    JGYOBU, _
                                                    TO_NAIGAI, _
                                                    HIN_GAI, _
                                                    NYUKA_DT, _
                                                    YOIN_FURIKAE_IN, _
                                                    0, MI_JITU_QTY, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , , StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                If sts Then
                    IDO_Update_Proc = sts
                    Exit Function
                End If
            Else
                If IDO_GOODS_ON_F = "0" Then
                    sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                                    TO_LOCATION, _
                                                    JGYOBU, _
                                                    NAIGAI, _
                                                    HIN_GAI, _
                                                    NYUKA_DT, _
                                                    YOIN, _
                                                    MI_JITU_QTY, 0, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO & "商品振替", _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , , StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                Else
                    sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                                    TO_LOCATION, _
                                                    JGYOBU, _
                                                    NAIGAI, _
                                                    HIN_GAI, _
                                                    NYUKA_DT, _
                                                    YOIN, _
                                                    0, MI_JITU_QTY, _
                                                    ID, _
                                                    TANTO_CODE, _
                                                    RETRY, , MEMO, _
                                                    Ins_DateTime, _
                                                    StrConv(wZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                                    StrConv(wZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                                    StrConv(wZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , , StrConv(wZAIKOREC.GENSANKOKU, vbUnicode), StrConv(wZAIKOREC.SHIIRE_WORK_CENTER, vbUnicode), StrConv(wZAIKOREC.ID_NO2, vbUnicode), StrConv(wZAIKOREC.YOSAN_FROM, vbUnicode), StrConv(wZAIKOREC.YOSAN_TO, vbUnicode))
                End If
                If sts Then
                    IDO_Update_Proc = sts
                    Exit Function
                End If
            End If
        End If
    End If
    
'============================================================
    
    If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_JITU Then
                                        '標準棚番
'        If Last_JGYOBU = SOJIKI Or _
'            Last_JGYOBU = SENTAKU Then
                                        '掃除機は設定済みの上書きをしない
'''全センター設定済標準棚番は変更しない。2004.04.10
            If StrConv(ITEMREC.ST_SET_DT, vbUnicode) = Space(8) Then
                Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Date, "YYYYMMDD"))
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
                        IDO_Update_Proc = SYS_CANCEL
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
                        IDO_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目マスタ")
                IDO_Update_Proc = SYS_ERR
                Exit Function
                        
        End Select
    Loop
'============================================================
    
    IDO_Update_Proc = False
    
End Function
Public Function wZAIKO_Open(Mode As Integer) As Integer
'****************************************************
'*      「移動処理」    在庫ＯＰＥＮ処理
'*
'*  在庫ファイルを別ポインタでＯＰＥＮする
'*  (呼び元で起動時に１度だけ呼び出す)

'*  戻り値: false       :正常
'*          true        :異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    wZAIKO_Open = True
                                '在庫データ　フルパス取込み
    sts = GetIni("FILE", ZAIKO_ID, "SYS", c)
    
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
'-------------- ＯＰＥＮ処理での使用中は、立ち上げ時に１回だけのはずなので、常に画面入力とし、
'               ｷｬﾝｾﾙは、処理の起動ｷｬﾝｾﾙとする。
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    wZAIKO_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "在庫データ")
                Exit Function
        End Select
    Loop

    wZAIKO_Open = False

End Function

Public Function wZAIKO_CLOSE() As Integer

'****************************************************
'*      「移動処理」    在庫ＣＬＯＳＥ処理
'*
'*  在庫ファイルを別ポインタでＣＬＯＳＥする
'*  (呼び元で終了時に１度だけ呼び出す)
'*  戻り値: false       :正常
'*          true        :異常
'****************************************************
Dim sts As Integer
    
    wZAIKO_CLOSE = True
    
    sts = BTRV(BtOpClose, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
    
    Select Case sts
        Case BtNoErr, BtErrNoOpen
        Case Else
            Call File_Error(sts, BtOpClose, "在庫データ")
            Exit Function
    End Select

    wZAIKO_CLOSE = False

End Function
