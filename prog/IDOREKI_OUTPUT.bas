Attribute VB_Name = "IDOREKI_OUTPUT"
Option Explicit
'---------------------------------------------- *更新用品目ワーク
'ポジショニング
Public wITEM_POS    As POSBLK
'データ・バッファ
Public wITEMREC     As ITEMREC_Tag
'キー・データ
Public K0_wITEM     As KEY0_ITEM

Public Wel_S_SHOUHI             As String * 2       '「WEL 資材消費」の要因 2007.06.28          2015.03.03 移動
Public Wel_S_SHOUHI2            As String * 2      '「WEL 資材消費(新)」の要因 2015.02.21                  移動


Public Function IDOREKI_OUTPUT_PROC(FROM_LOCATION As String, _
                                    TO_LOCATION As String, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional CYU_KBN As String = " ", _
                                    Optional MEMO As String = "          ", _
                                    Optional Ins_DateTime As String = "              ", _
                                    Optional SHIIRE_CODE As String = "     ", _
                                    Optional SHIIRE_TANKA As String = "           ", _
                                    Optional KEIJYO_YM As String = "      ", _
                                    Optional MENU_NO As String = "  ", _
                                    Optional MUKE_CODE As String = "        ", _
                                    Optional SS_CODE As String = "        ", _
                                    Optional ID_NO As String = "        ", _
                                    Optional BIN_NO As String = "  ", _
                                    Optional DEN_NO As String = "      ", _
                                    Optional DEN_YMD As String = "        ", Optional LOG_MODE As Integer = 0, Optional GENSANKOKU As String = "                    ", Optional SHIIRE_WORK_CENTER As String = "        ", Optional ID_NO2 As String = "            ", Optional YOSAN_FROM As String = "     ", Optional YOSAN_TO As String = "     ", Optional wkMTS As String = "        ", Optional SEK_TEI_LABELID As String = "             ", Optional HINBAN_DAMMY As String = "                    ") As Integer
'****************************************************
'*      在庫移動歴データ更新
'*
'*  在庫移動歴の更新を行う。
'*  (引数の設定ミスはこちらではチェックしない)
'*  引数：  FROM棚（XXXXXXXX(倉庫№+列+連+段)省略可）  ※FROM/TO何れか必須
'*          TO棚（XXXXXXXX(倉庫№+列+連+段)省略可）    ※FROM/TO何れか必須
'*          事業部（省略不可）
'*          国内外（省略不可）
'*          外部品番（省略不可）
'*          入荷日(YYYYMMDD 省略不可)
'*          要因(省略不可)
'*          商品化済み実績数（=0を可とする、履歴のみ出力）
'*          未商品実績数（=0を可とする、履歴のみ出力）
'*          ID(必須)
'*          担当者
'*          リトライ(省略可 １桁目:1=画面メッセージ有 0:無，２桁目:リトライ回数(0～9 0:無限))
'*          注文区分（省略可）
'*          メモ(省略可 履歴に出力するメモ内容)
'*          ﾃﾞｰﾀ追加日時（省略可　ﾃﾞｰﾀ作成日時を一元可）
'*          仕入先ｺｰﾄﾞ（資材用省略可　2006.01.05）
'*          仕入単価（資材用省略可　2006.01.05）
'*          計上年月（資材用省略可　2006.01.05）
'*          TOPﾒﾆｭｰ(原価用 2006.01.30)
'*          向け先(原価用 2006.01.30)
'*          直送先(原価用 2006.01.30)
'*          伝票ID(原価用 2006.01.30)
'*          便№   (大阪PC出荷 2007.05.16)
'*
'*
'*          ログ出力モード (0:ここで出力 1:ここでは出力しない)
'*  戻り値: false       :正常
'*          true        :継続可能な異常
'*          SYS_ERR     :継続できない異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'*
'*  ※出荷予定／在庫データ／向け先管理マスタは呼び元で読込み済みの事
'****************************************************
Dim sts                 As Integer
Dim Sumi_Zaiko_Qty      As Long
Dim Mi_Zaiko_Qty        As Long

Dim RETRY_CNT           As Integer
Dim MESG_FLG            As Integer
Dim RETRY_SU            As Integer
    
Dim ans                 As Integer
                                            
    IDOREKI_OUTPUT_PROC = True
                                            
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                            '出庫表／伝票ID出荷時の要因を置きかえる
    If Left(YOIN, 1) = ACT_DENPYO_ID Or _
        Left(YOIN, 1) = ACT_SYUKA_HYO Or _
        Left(YOIN, 1) = ACT_DENPYO_ID2 Then 'ACT_DENPYO_ID2 追加　2015.02.21
        YOIN = ACT_SYUKA_KEI & CYU_KBN
    End If
                            
                            '品目マスタ読み込み
    Call UniCode_Conv(K0_wITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_wITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_wITEM.HIN_GAI, HIN_GAI)
    sts = BTRV(BtOpGetEqual, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(wITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
            IDOREKI_OUTPUT_PROC = SYS_ERR
            Exit Function
    End Select
    
'2008.08.06    Call UniCode_Conv(IDOREC.JITU_DT, Format(Now, "yyyymmdd"))              '実績日付
'2008.08.06    Call UniCode_Conv(IDOREC.JITU_TM, Format(Now, "HHmmss"))                '実績時刻
    
    
    If YOIN = Wel_S_SHOUHI2 Then                                            '新資材消費の要因置き換え　2015.02.21
        YOIN = Wel_S_SHOUHI
    End If
    
    
    
    If Trim(Ins_DateTime) = "" Then     '2008.09.01
    
        Call UniCode_Conv(IDOREC.JITU_DT, Format(Now, "yyyymmdd"))              '実績日付
        Call UniCode_Conv(IDOREC.JITU_TM, Format(Now, "HHmmss"))                '実績時刻
    
    Else
    
    
        Call UniCode_Conv(IDOREC.JITU_DT, Left(Ins_DateTime, 8))                '実績日付   2008.08.06
        Call UniCode_Conv(IDOREC.JITU_TM, Right(Ins_DateTime, 6))               '実績時刻   2008.08.06
    
    End If
    
    
    
    Call UniCode_Conv(IDOREC.JGYOBU, JGYOBU)                                '事業部
    Call UniCode_Conv(IDOREC.NAIGAI, NAIGAI)                                '国内外
    Call UniCode_Conv(IDOREC.HIN_GAI, HIN_GAI)                              '品番（外部）
    
    Call UniCode_Conv(IDOREC.RIRK_ID, YOIN)                                 '履歴種別
    
                                                                            
                                                                            
                                                                            '実績数量(商品化済み)
    Call UniCode_Conv(IDOREC.SUMI_JITU_QTY, Format(SUMI_JITU_QTY, "00000000"))
                                                                            '実績数量(未商品)
    Call UniCode_Conv(IDOREC.MI_JITU_QTY, Format(MI_JITU_QTY, "00000000"))
    Call UniCode_Conv(IDOREC.FROM_SOKO, Mid(FROM_LOCATION, 1, 2))           'FROM 倉庫№
    Call UniCode_Conv(IDOREC.FROM_RETU, Mid(FROM_LOCATION, 3, 2))           'FROM 列
    Call UniCode_Conv(IDOREC.FROM_REN, Mid(FROM_LOCATION, 5, 2))            'FROM 連
    Call UniCode_Conv(IDOREC.FROM_DAN, Mid(FROM_LOCATION, 7, 2))            'FROM 段
    Call UniCode_Conv(IDOREC.TO_SOKO, Mid(TO_LOCATION, 1, 2))               'TO 倉庫№
    Call UniCode_Conv(IDOREC.TO_RETU, Mid(TO_LOCATION, 3, 2))               'TO 列
    Call UniCode_Conv(IDOREC.TO_REN, Mid(TO_LOCATION, 5, 2))                'TO 連
    Call UniCode_Conv(IDOREC.TO_DAN, Mid(TO_LOCATION, 7, 2))                'TO 段
    Call UniCode_Conv(IDOREC.PRG_ID, StrConv(App.EXEName, vbUpperCase))     '出力元プログラム
            
            
            
            
            
'''2011.02.03    If YOIN = YOIN_TANASHOGO Or YOIN = YOIN_TANAHINSHOGO Then
'品番別照合を追加   2011.02.03
    If YOIN = YOIN_TANASHOGO Or YOIN = YOIN_TANAHINSHOGO Or YOIN = YOIN_HIN_SHOGO Then
       '要因＝棚照合の時は不定となる
                                                    '品番（内部）
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(wITEMREC.HIN_NAI, vbUnicode))
                                                    '入庫日
        Call UniCode_Conv(IDOREC.NYUKO_DT, "")
    Else
                                                        '品番（内部）
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(ZAIKOREC.HIN_NAI, vbUnicode))
                                                            '入庫日
        Call UniCode_Conv(IDOREC.NYUKO_DT, StrConv(ZAIKOREC.NYUKO_DT, vbUnicode))
   End If
    
    
    
    If Trim(FROM_LOCATION) = "" And Trim(TO_LOCATION) = "" Then                     '2014.03.05
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(wITEMREC.HIN_NAI, vbUnicode))     '2014.03.05
                                                    '入庫日
        Call UniCode_Conv(IDOREC.NYUKO_DT, "")                                      '2014.03.05
    End If
    
    
    
    Call UniCode_Conv(IDOREC.NYUKA_DT, NYUKA_DT)
    Call UniCode_Conv(IDOREC.WEL_ID, ID)                                    '端末ID
    
    
    
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, Mid(YOIN, 1, 1))                   '履歴名称
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, Mid(YOIN, 2, 1))
                                            '要因ﾏｽﾀ読込み
    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    Select Case sts
        Case BtNoErr
            Call UniCode_Conv(IDOREC.RIRK_NAME, StrConv(YOINREC.YOIN_DNAME, vbUnicode))
            Call UniCode_Conv(IDOREC.SUM_KBN, StrConv(YOINREC.SUM_KBN, vbUnicode))
        Case BtErrKeyNotFound
            Call UniCode_Conv(IDOREC.RIRK_NAME, "")
                                            '不明のときは在訂扱い
            Call UniCode_Conv(IDOREC.SUM_KBN, SUM_KBN_ZT)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "要因ﾏｽﾀ")
            IDOREKI_OUTPUT_PROC = SYS_ERR
            Exit Function
    End Select
                                                                            '品目名称
    Call UniCode_Conv(IDOREC.HIN_NAME, StrConv(wITEMREC.HIN_NAME, vbUnicode))
                                                                            '品目別在庫数
    If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, Mi_Zaiko_Qty, JGYOBU, NAIGAI, HIN_GAI) Then
        IDOREKI_OUTPUT_PROC = SYS_ERR
        Exit Function
    End If
    Call UniCode_Conv(IDOREC.SUMI_HIN_Zaiko_Qty, Format(Sumi_Zaiko_Qty, "00000000"))
    Call UniCode_Conv(IDOREC.MI_HIN_Zaiko_Qty, Format(Mi_Zaiko_Qty, "00000000"))
                                                                            
                                                                            'FROM棚別品目別在庫数
    If Len(Trim(FROM_LOCATION)) <> 0 Then
        If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, Mi_Zaiko_Qty, JGYOBU, NAIGAI, HIN_GAI, FROM_LOCATION) Then
            IDOREKI_OUTPUT_PROC = SYS_ERR
            Exit Function
        End If
        Call UniCode_Conv(IDOREC.SUMI_FROM_TANA_Zaiko_Qty, Format(Sumi_Zaiko_Qty, "00000000"))
        Call UniCode_Conv(IDOREC.MI_FROM_TANA_Zaiko_Qty, Format(Mi_Zaiko_Qty, "00000000"))
    Else
        Call UniCode_Conv(IDOREC.SUMI_FROM_TANA_Zaiko_Qty, "00000000")
        Call UniCode_Conv(IDOREC.MI_FROM_TANA_Zaiko_Qty, "00000000")
    End If
                                                                            'TO棚別品目別在庫数
    If Len(Trim(TO_LOCATION)) <> 0 Then
        If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, Mi_Zaiko_Qty, JGYOBU, NAIGAI, HIN_GAI, TO_LOCATION) Then
            IDOREKI_OUTPUT_PROC = SYS_ERR
            Exit Function
        End If
        Call UniCode_Conv(IDOREC.SUMI_TO_TANA_Zaiko_Qty, Format(Sumi_Zaiko_Qty, "00000000"))
        Call UniCode_Conv(IDOREC.MI_TO_TANA_Zaiko_Qty, Format(Mi_Zaiko_Qty, "00000000"))
    Else
        Call UniCode_Conv(IDOREC.SUMI_TO_TANA_Zaiko_Qty, "00000000")
        Call UniCode_Conv(IDOREC.MI_TO_TANA_Zaiko_Qty, "00000000")
    End If
                                        
    If CYU_KBN = " " Then
        Call UniCode_Conv(IDOREC.DEN_DT, "")
        Call UniCode_Conv(IDOREC.DEN_NO, "")
        Call UniCode_Conv(IDOREC.TOKU_MARK, "")
        Call UniCode_Conv(IDOREC.MUKE_CODE, "")
        Call UniCode_Conv(IDOREC.MUKE_CHG_CD, "")
        Call UniCode_Conv(IDOREC.MUKE_DNAME, "")
        Call UniCode_Conv(IDOREC.SS_CODE, "")
        Call UniCode_Conv(IDOREC.SS_NAME, "")
        Call UniCode_Conv(IDOREC.ID_NO, "")
    
        If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_TU_NYUKA Then
            Call UniCode_Conv(IDOREC.ID_NO, ID_NO2)
        End If
    
    Else
        Call UniCode_Conv(IDOREC.DEN_DT, StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode))
        Call UniCode_Conv(IDOREC.DEN_NO, StrConv(Y_SYUREC.DEN_NO, vbUnicode))
        If CYU_KBN = CYU_KBN_SPO And _
            StrConv(Y_SYUREC.TOK_KBN, vbUnicode) = "1" Then
            Call UniCode_Conv(IDOREC.TOKU_MARK, "*")                        '特売りマーク
        Else
            Call UniCode_Conv(IDOREC.TOKU_MARK, "")
        End If
        Call UniCode_Conv(IDOREC.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
    
    
                                                                                '向け先コード
        Call UniCode_Conv(IDOREC.MUKE_CODE, StrConv(MTSREC.MUKE_CODE, vbUnicode))
                                                                                '向け先名称
        Call UniCode_Conv(IDOREC.MUKE_DNAME, StrConv(MTSREC.MUKE_DNAME, vbUnicode))
                                                                                'ＳＳコード
        Call UniCode_Conv(IDOREC.SS_CODE, StrConv(MTSREC.SS_CODE, vbUnicode))
                                                                                'ＳＳ名称
        Call UniCode_Conv(IDOREC.SS_NAME, StrConv(MTSREC.SS_NAME, vbUnicode))
                                                                                '得意先略称
        Call UniCode_Conv(IDOREC.MUKE_DNAME, StrConv(MTSREC.MUKE_DNAME, vbUnicode))
    
    
    End If
                                                                        
                                                                        
    '2009.03.18
    If Left(YOIN, 1) = ACT_BINNO Then
        Call UniCode_Conv(IDOREC.SS_CODE, SS_CODE)
        Call UniCode_Conv(IDOREC.SS_NAME, SS_CODE)
    End If
                                                                        
                                                                        
    Call UniCode_Conv(IDOREC.MEMO, MEMO)                                    'メモ
                                                                            
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, TANTO_CODE)                      '担当者
                                            '担当者ﾏｽﾀ読込み
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者ﾏｽﾀ")
            IDOREKI_OUTPUT_PROC = SYS_ERR
            Exit Function
    End Select
                                                                            '担当者コード
    Call UniCode_Conv(IDOREC.TANTO_CODE, StrConv(TANTOREC.TANTO_CODE, vbUnicode))
                                                                            '担当者名称
    Call UniCode_Conv(IDOREC.TANTO_NAME, StrConv(TANTOREC.TANTO_NAME, vbUnicode))
                                                                            
                                                                            '挿入日時
    If Len(Trim(Ins_DateTime)) = 0 Then
        Ins_DateTime = StrConv(IDOREC.JITU_DT, vbUnicode) & StrConv(IDOREC.JITU_TM, vbUnicode)
    End If
    
    Call UniCode_Conv(IDOREC.Ins_DateTime, Ins_DateTime)
                                            
                                            
    Call UniCode_Conv(IDOREC.SHIIRE_CODE, SHIIRE_CODE)          '仕入ｺｰﾄﾞ2006.01.06
    Call UniCode_Conv(IDOREC.SHIIRE_TANKA, SHIIRE_TANKA)        '仕入単価2006.01.06
    Call UniCode_Conv(IDOREC.KEIJYO_YM, KEIJYO_YM)              '計上年月2006.01.06
                                            
    Call UniCode_Conv(IDOREC.BIN_NO, BIN_NO)                    '便№ 2007.05.16
                                            
                                            
    If Trim(DEN_NO) <> "" Then                                  '大阪PC　入荷伝票№ 2007.06.07
        Call UniCode_Conv(IDOREC.DEN_NO, DEN_NO)
    End If
    If Trim(DEN_YMD) <> "" Then                                 '大阪PC　入荷伝票日付 2007.06.07
        Call UniCode_Conv(IDOREC.DEN_DT, DEN_YMD)
    End If
                                            
                                            
                                            
                                            
                                            
                                            
    '----------------   2010.07.08 ▽
    Call UniCode_Conv(IDOREC.GENSANKOKU, GENSANKOKU)            '原産国名
                                                                '資材仕入先ﾜｰｸｾﾝﾀｰ
    Call UniCode_Conv(IDOREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
    Call UniCode_Conv(IDOREC.ID_NO2, ID_NO2)                    'ID_NO
    Call UniCode_Conv(IDOREC.YOSAN_FROM, YOSAN_FROM)            '予算単位（元）
    Call UniCode_Conv(IDOREC.YOSAN_TO, YOSAN_TO)                '予算単位（先）
    '----------------   2010.07.08 △
                                            
    
    '----------------   2011.04.29 ▽
    If Trim(SEK_TEI_LABELID) <> "" Then
        Call UniCode_Conv(IDOREC.ID_NO, SEK_TEI_LABELID)
    End If
    '----------------   2011.04.29 △
                                            
                                            
                                            
                                            
                                            
                                            
                                            
    Call UniCode_Conv(IDOREC.FILLER, "")
                                            
                                        '在庫移動歴出力
    
    '要因=商品完了 & 棚番TO <>"" 2019/12/25 商品化完了登録在庫計上時 移動履歴に入庫数を表示
    If YOIN = "M8" And Trim(TO_LOCATION) <> "" Then '2020/03/16 "M8" 商品化完了登録要因 決め打ちに修正
         Call UniCode_Conv(IDOREC.SUM_KBN, SUM_KBN_IN)
    End If
    
    RETRY_CNT = 0
    Do
        
        sts = BTRV(BtOpInsert, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                If RETRY_SU <> 0 Then
                    
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                        Call File_Error(sts, BtOpInsert, "在庫移動歴", 0)
                        IDOREKI_OUTPUT_PROC = SYS_CANCEL
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
                    ans = MsgBox("他端末でデータ使用中です。<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        IDOREKI_OUTPUT_PROC = SYS_CANCEL
                        Exit Function
                    End If
                End If
            
            Case BtErrDEAD_LOCK
                IDOREKI_OUTPUT_PROC = SYS_CANCEL
                Exit Function
            
            Case Else
                Call File_Error(sts, BtOpInsert, "在庫移動歴")
                IDOREKI_OUTPUT_PROC = SYS_ERR
                Exit Function
        End Select
    Loop
                            

    
    
    If LOG_MODE = 0 Then
        If Trim(MENU_NO) = "" Then
        Else
        '作業ﾛｸﾞ出力
            
            If Trim(MUKE_CODE) <> "" Then
                Call UniCode_Conv(IDOREC.MUKE_CODE, MUKE_CODE)
            End If
            
            If Trim(SS_CODE) <> "" Then
                Call UniCode_Conv(IDOREC.SS_CODE, SS_CODE)
            End If
            
            If Trim(ID_NO) <> "" Then
                Call UniCode_Conv(IDOREC.ID_NO, ID_NO)
            End If
            
            
            If App.EXEName = "F102015" Then
                Call UniCode_Conv(IDOREC.ID_NO, ID_NO2)
            End If
            
            If YOIN = RYOHEN Then
                Call UniCode_Conv(IDOREC.MUKE_CODE, Left(StrConv(IDOREC.MEMO, vbUnicode), 4))
            End If
            
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
                                                TO_LOCATION, _
                                                StrConv(IDOREC.ID_NO, vbUnicode), _
                                                StrConv(IDOREC.MUKE_CODE, vbUnicode), _
                                                StrConv(IDOREC.SS_CODE, vbUnicode), _
                                                RETRY, , , , wkMTS, , , , HINBAN_DAMMY) Then
                IDOREKI_OUTPUT_PROC = SYS_ERR
                Exit Function
            End If
        End If
    End If
                            
                            '正常終了
    IDOREKI_OUTPUT_PROC = False

End Function
                    
Public Function wITEM_Open(Mode As Integer) As Integer
'****************************************************
'*      「移動歴出力処理」    品目ＯＰＥＮ処理
'*
'*  品目マスタを別ポインタでＯＰＥＮする
'*  (呼び元で起動時に１度だけ呼び出す)
'*
'*  戻り値: false       :正常
'*          true        :異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    wITEM_Open = True
                                '品目マスタ　フルパス取込み
    sts = GetIni("FILE", ITEM_ID, "SYS", c)
    
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wITEM_POS, wITEMREC, Len(wITEMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
'-------------- ＯＰＥＮ処理での使用中は、立ち上げ時に１回だけのはずなので、常に画面入力とし、
'               ｷｬﾝｾﾙは、処理の起動ｷｬﾝｾﾙとする。
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    wITEM_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "品目マスタ")
                Exit Function
        End Select
    Loop

    wITEM_Open = False

End Function
