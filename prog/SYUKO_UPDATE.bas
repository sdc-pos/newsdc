Attribute VB_Name = "SYUKO_UPDATE"
Option Explicit
'---------------------------------------------- *更新用出荷予定ワーク
'ポジショニング
Public wY_SYU_POS   As POSBLK
'データ・バッファ
Public wY_SYUREC    As Y_SYUREC_Tag
'キー・データ
Public K3_wY_SYU    As KEY3_Y_SYU


Public Function Syuko_Update_Proc(JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    FROM_LOCATION As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    SYUKA_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional MEMO As String = "          ", _
                                    Optional CYU_KBN As String = " ", _
                                    Optional MUKE_CODE As String = "                ", _
                                    Optional SYUKA_YMD As String = "        ", _
                                    Optional DEN_NO As String = "          ", _
                                    Optional ID_NO As String = "        ", _
                                    Optional MENU_NO As String = "  ", _
                                    Optional LOG_NON As Integer = 0) As Integer
'****************************************************
'*      「出荷／出庫処理」在庫データ更新
'*
'*  在庫データの更新を行う。
'*  (引数の設定ミスはこちらではチェックしない)
'*  ※トランザクション処理が必要な場合は呼び元で行う事
'*  使用ﾌｧｲﾙ    :   在庫データ
'*                  品目マスタ
'*                  要因マスタ
'*                  向け先マスタ
'*                  出荷予定データ
'*                  在庫移動歴
'*  引数：  事業部（省略不可）
'*          国内外（省略不可）
'*          品番外部(省略不可)
'*          FROM棚番（XXXXXXXX(倉庫№+列+連+段)省略不可）
'*          入荷日(YYYYMMDD 省略可 省略時FIFO)
'*          要因(省略不可)
'*          商品化済み実績数（いずれか必須）
'*          未商品実績数    （　　　〃　　）
'*          出荷数量        （            ）
'*          ID(省略不可)
'*          担当者（省略不可）
'*          リトライ(省略可 １桁目:1=画面メッセージ有 0:無，２桁目:リトライ回数(0～9 0:無限))
'*          メモ(省略可 履歴に出力するメモ内容)
'*          注文区分（出荷時必須）
'*          伝票ＩＤ（出荷時必須）
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
    
Dim GET_DEN_NO  As String * 6
Dim GET_ID_NO   As String * 12
    
Dim JITU_QTY    As Long
Dim GOODS_F     As String * 1
    
Dim Wk_SUMI_JITU_QTY As Long
Dim Wk_MI_JITU_QTY As Long
Dim Wk_SYUKA_QTY As Long

Dim Ins_DateTime    As String * 14              '2004.12.09
    
Dim wkYOIN      As String * 2
    
    
'''''   2011.04.04
Dim Total_SUMI_JITU_QTY     As Long
Dim Total_MI_JITU_QTY       As Long
'''''   2011.04.04
    
    
    Syuko_Update_Proc = True
                                                                      
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                                        
                                        
    Wk_SUMI_JITU_QTY = SUMI_JITU_QTY
    Wk_MI_JITU_QTY = MI_JITU_QTY
    Wk_SYUKA_QTY = SYUKA_QTY
                                        
    Ins_DateTime = Format(Now, "YYYYMMDDHHMMSS")            '2004.12.09
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
                Exit Function
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                If RETRY_SU <> 0 Then
                    
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ", 0)
                        Syuko_Update_Proc = SYS_CANCEL
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
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case BtErrDEAD_LOCK
                Syuko_Update_Proc = SYS_CANCEL
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Syuko_Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
    '*------------------------------------------------------'出荷時出荷予定の確保
    If CYU_KBN <> " " Then       '出荷時向け先ﾏｽﾀ読込み
        Call UniCode_Conv(K0_MTS.MUKE_CODE, Left(MUKE_CODE, 8))
        Call UniCode_Conv(K0_MTS.SS_CODE, Right(MUKE_CODE, 8))
                                    '向け先マスタを読み込み向け先ｺｰﾄﾞをｾｯﾄする
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound           '有るとまずいがエラーにしない
                Call UniCode_Conv(MTSREC.MUKE_DNAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先管理ﾏｽﾀ")
                Syuko_Update_Proc = SYS_ERR
                Exit Function
        End Select
    
    End If
    
    Select Case CYU_KBN
        Case " "
        Case CYU_KBN_KIN                '緊急時、出荷予定を起票する
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                      '使用子機ＩＤ
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                      '使用中プログラム
            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)             '完了区分
            Call UniCode_Conv(Y_SYUREC.DT_SYU, "R")                     'データ種別
            Call UniCode_Conv(Y_SYUREC.JGYOBU, Last_JGYOBU)             '事業部区分
            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_KIN)        '注文区分
            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_KIN)
            If Len(Trim(ID_NO)) <> 0 Then                               'ＩＤ№
                Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
                Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)
            Else
                sts = Den_No_Set_Proc(21, Last_JGYOBU, GET_ID_NO)
                If sts Then
                    Syuko_Update_Proc = SYS_ERR
                    Exit Function
                End If
                Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, GET_ID_NO)
                Call UniCode_Conv(Y_SYUREC.ID_NO, GET_ID_NO)
            End If
                                                                            
            Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI)                  '国内外
                                                                    
            Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)             '品目番号
            Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)                 '品目番号
                                                                        '得意先コード
            Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, Left(MUKE_CODE, 8))
            Call UniCode_Conv(Y_SYUREC.MUKE_CODE, Left(MUKE_CODE, 8))
                                                                        '直送先コード
            Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, Right(MUKE_CODE, 8))
            Call UniCode_Conv(Y_SYUREC.SS_CODE, Right(MUKE_CODE, 8))
                                                                        '出荷日
            Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, SYUKA_YMD)
            Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, SYUKA_YMD)
            Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, SYUKA_YMD)
    
    
    
            Call UniCode_Conv(Y_SYUREC.JGYOBA, "")                      '事業場ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")                    'データ区分
            Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")                    '取引区分

                
            Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")               '会計用事業場ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")               '資産管理用事業場ｺｰﾄﾞ

            If Len(Trim(DEN_NO)) <> 0 Then
                Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
            Else
                sts = Den_No_Set_Proc(20, Last_JGYOBU, GET_DEN_NO)
                If sts Then
                    Syuko_Update_Proc = SYS_ERR
                    Exit Function
                End If
                Call UniCode_Conv(Y_SYUREC.DEN_NO, GET_DEN_NO)
        
            End If
    
                                                                        '出庫数量
            Call UniCode_Conv(Y_SYUREC.SURYO, Format(SUMI_JITU_QTY + MI_JITU_QTY, "0000000"))
            Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")                 '在庫収支
            Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")                '資産管理用在庫収支ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")                 '補助在庫収支ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.TANKA, "")                       '単価
    
            Call UniCode_Conv(Y_SYUREC.ODER_NO, "")                     'オーダー番号
            Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")                     'アイテム番号
            Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")                   '注文管理番号略号

            
            Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")                 '個装形態ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.TANABAN1, "")                    '棚番１
            Call UniCode_Conv(Y_SYUREC.TANABAN2, "")                    '棚番２
            Call UniCode_Conv(Y_SYUREC.TANABAN3, "")                    '棚番３
                                                                        '得意先名称
            Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(MTSREC.MUKE_NAME, vbUnicode))
    
            Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_T)         '注文区分名称
    
            Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")                     '原産国1
            Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")                     '原産国2
            Call UniCode_Conv(Y_SYUREC.BIKOU2, "")                      '備考2
  
    
            Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")                     '販売区分
            Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, "")                   '直送指示区分
            Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")                  'ﾕﾆｯﾄ修正管理番号
            Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")               '在庫引当順序
            Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")              '合梱管理番号
            Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")                  '受注残数量
            Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")                  '供給区分
            Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")                '商品化納品在庫収支ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")              '商品化納品資産管理収支ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")               '商品化納品補助収支ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.BIKOU1, "")                      '備考1
            Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")                   '帳端区分
            Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")                  '受付品目番号
                                                                        '品名
            Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
            Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")              '品目番号変更区分
            Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")             'ﾓｼﾞｭｰﾙ交換区分
            Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")                 '残在庫まとめ在庫収支ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")            '残在庫まとめ資産管理収支ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")             '残在庫まとめ補助収支ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")                   '指定納期
            Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")            'ｻｰﾋﾞｽ会社管理番号
            Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")                  '機種品目ｺｰﾄﾞ
            Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")             '環境企画部品区分
            Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")                '欠品解消区分
                                                                        '品番内部
            Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                                                                
                                                                        'ホスト棚番
            Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
    
            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")                   '印刷日付
            Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")                     '完了日付
            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")                  '検品日付
            Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")                     '特売り区分
    
            Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "00000000")          '実績数量
                                                                        '取込み日時
            Call UniCode_Conv(Y_SYUREC.INS_NOW, Format(Now, "YYYY/MM/DD HH:MM:SS"))
            
            
            Call UniCode_Conv(Y_SYUREC.FILLER, "")
        
'            SYUKA_QTY = SUMI_JITU_QTY + MI_JITU_QTY
'            SUMI_JITU_QTY = 0
'            MI_JITU_QTY = 0
        
        Case Else
                                    '出荷予定読込み
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)                  '事業部
'            Call UniCode_Conv(K0_Y_SYU.KEY_CYU_KBN, CYU_KBN)            '注文区分2004.04.08
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_NO)                'IDNo

            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        If CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) Then
                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "他端末でデータが変更されています。更新処理を中止します。", vbOKOnly, "確認入力"
                            End If
                            Exit Function
                        End If
                        Exit Do
                    Case BtErrKeyNotFound
                        If MESG_FLG = 1 Then
                            Beep
                            MsgBox "出荷予定が存在しません。更新処理を中止します。", vbOKOnly, "確認入力"
                        End If
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                                Syuko_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case BtErrDEAD_LOCK
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", MESG_FLG)
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            Loop
    End Select
'============================================================
    If Len(Trim(NYUKA_DT)) = 0 Then
    '*------------------------------------------------------'入荷日指定無し 在庫データ読込み（スキャナ処理）
    '
    
    
    
        '作業ﾛｸﾞ出力    '2008.08.06
        
        wkYOIN = YOIN   '2011.08.12
        
        If LOG_NON = 1 Then
        Else
            If SUMI_JITU_QTY = 0 And MI_JITU_QTY = 0 Then
                
                
                If Left(YOIN, 1) = ACT_DENPYO_ID Or _
                    Left(YOIN, 1) = ACT_DENPYO_ID2 Or _
                    Left(YOIN, 1) = ACT_SYUKA_HYO Then      'ACT_DENPYO_ID2追加　2015.02.21
                    wkYOIN = ACT_SYUKA_KEI & CYU_KBN
                End If
                
                
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
'                If P_SAGYO_LOG_OUTPUT_PROC(TANTO_CODE, _
'                                            ID, _
'                                            JGYOBU, _
'                                            NAIGAI, _
'                                            MENU_NO, _
'                                            wkYOIN, _
'                                            HIN_GAI, _
'                                            SYUKA_QTY, _
'                                            0, _
'                                            FROM_LOCATION, _
'                                            "", _
'                                            ID_NO, _
'                                            MUKE_CODE) Then
'                    Exit Function
'                End If
                Total_SUMI_JITU_QTY = 0
                Total_MI_JITU_QTY = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
            
            
            Else
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
                                            "") Then
                    Exit Function
                End If
            End If
        End If
    
    
    
    
    
    '---------------------------------------'商品化済み処理
        If SUMI_JITU_QTY <> 0 Then
            Zan_Qty = SUMI_JITU_QTY
            Do
                Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   '倉庫№
                Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '列
                Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '連
                Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '段
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
                                                '棚＋品＋商品／未商品ブレーク
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
                                    Call File_Error(sts, com + BtSNoWait, "在庫データ", 1)
                                    Syuko_Update_Proc = SYS_CANCEL
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
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, com + BtSNoWait, "在庫データ")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop

                If Upd_com = BtOpUpdate Then
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
                                    Syuko_Update_Proc = SYS_CANCEL
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
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, Upd_com, "在庫データ")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop
            '============================================================
                '*------------------------------------------------------'在庫移動歴出力
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                            Space(8), _
                                            JGYOBU, _
                                            NAIGAI, _
                                            HIN_GAI, _
                                            StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                            YOIN, _
                                            WK_Qty, _
                                            0, _
                                            ID, _
                                            TANTO_CODE, _
                                            RETRY, _
                                            CYU_KBN, _
                                            MEMO, _
                                            Ins_DateTime, _
                                            StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                            StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                            StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1)

                If sts Then
                    Syuko_Update_Proc = sts
                    Exit Function
                End If
                
                Zan_Qty = Zan_Qty - WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                Total_SUMI_JITU_QTY = Total_SUMI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                
                
                
                If Zan_Qty <= 0 Then
                    Exit Do                     '引き落とし終了
                End If
            Loop
        End If
                    
'************************************************************
    '
    '---------------------------------------'未商品処理
        If MI_JITU_QTY <> 0 Then
            Zan_Qty = MI_JITU_QTY
            Do
                Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   '倉庫№
                Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '列
                Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '連
                Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '段
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
                                                '棚＋品＋商品／未商品ブレーク
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
                                    Call File_Error(sts, com + BtSNoWait, "在庫データ", 1)
                                    Syuko_Update_Proc = SYS_CANCEL
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
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, com + BtSNoWait, "在庫データ")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop

                If Upd_com = BtOpUpdate Then
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
                                    Syuko_Update_Proc = SYS_CANCEL
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
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, Upd_com, "在庫データ")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop
            '============================================================
                '*------------------------------------------------------'在庫移動歴出力
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                            Space(8), _
                                            JGYOBU, _
                                            NAIGAI, _
                                            HIN_GAI, _
                                            StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                            YOIN, _
                                            0, _
                                            WK_Qty, _
                                            ID, _
                                            TANTO_CODE, _
                                            RETRY, _
                                            CYU_KBN, _
                                            MEMO, _
                                            Ins_DateTime, _
                                            StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                            StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                            StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1)
                If sts Then
                    Syuko_Update_Proc = sts
                    Exit Function
                End If
                
                Zan_Qty = Zan_Qty - WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                Total_MI_JITU_QTY = Total_MI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04

                If Zan_Qty <= 0 Then
                    Exit Do                     '引き落とし終了
                End If
            Loop
                    
        End If
'************************************************************
    '
    '---------------------------------------'商品化済み～未商品を順次引き当てる
        If SYUKA_QTY <> 0 Then
    
            Zan_Qty = SYUKA_QTY
            Do
                Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   '倉庫№
                Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '列
                Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '連
                Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '段
                Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '事業部
                Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '内外
                Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '品番（外部）
                Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")                        '商品／未商品
                Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")                        '入荷日
                
                
                SUMI_JITU_QTY = 0
                MI_JITU_QTY = 0
                
                RETRY_CNT = 0

                Do
                    sts = BTRV(BtOpGetGreater + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                                                '棚＋品＋商品／未商品ブレーク
                            If FROM_LOCATION <> (StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)) Or _
                                JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                                NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                                Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                                
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
                                    Call File_Error(sts, com + BtSNoWait, "在庫データ", 1)
                                    Syuko_Update_Proc = SYS_CANCEL
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
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, com + BtSNoWait, "在庫データ")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop

                If Upd_com = BtOpUpdate Then
                    Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - Zan_Qty, "00000000"))
                    Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '排他フラグ
                    Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '使用中子機ID
                    Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '使用中プログラム
                End If

                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                    SUMI_JITU_QTY = SUMI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                    Total_SUMI_JITU_QTY = Total_SUMI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                Else
                    MI_JITU_QTY = MI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                    Total_MI_JITU_QTY = Total_MI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
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
                                    Syuko_Update_Proc = SYS_CANCEL
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
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, Upd_com, "在庫データ")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop
            '============================================================
                '*------------------------------------------------------'在庫移動歴出力
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                            Space(8), _
                                            JGYOBU, _
                                            NAIGAI, _
                                            HIN_GAI, _
                                            StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                            YOIN, _
                                            SUMI_JITU_QTY, _
                                            MI_JITU_QTY, _
                                            ID, _
                                            TANTO_CODE, _
                                            RETRY, _
                                            CYU_KBN, _
                                            MEMO, _
                                            Ins_DateTime, _
                                            StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                            StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                            StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1)

                If sts Then
                    Syuko_Update_Proc = sts
                    Exit Function
                End If
                
                Zan_Qty = Zan_Qty - WK_Qty

                If Zan_Qty <= 0 Then
                    Exit Do                     '引き落とし終了
                End If
            Loop
        End If



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
        If LOG_NON = 1 Then
        Else
            If SYUKA_QTY <> 0 Then                              '2011.08.18
                
                If P_SAGYO_LOG_OUTPUT_PROC(TANTO_CODE, _
                                            ID, _
                                            JGYOBU, _
                                            NAIGAI, _
                                            MENU_NO, _
                                            wkYOIN, _
                                            HIN_GAI, _
                                            Total_SUMI_JITU_QTY, _
                                            Total_MI_JITU_QTY, _
                                            FROM_LOCATION, _
                                            "", _
                                            ID_NO, _
                                            MUKE_CODE) Then
                    Exit Function
                End If
            End If                                              '2011.08.18
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04



    Else
    '*------------------------------------------------------'入荷日指定有り 在庫データ読込み（画面処理）
    '
    '---------------------------------------'商品化済み処理
        If SUMI_JITU_QTY <> 0 Then
        
            Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   '倉庫№
            Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '列
            Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '連
            Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '段
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
                            If MESG_FLG = 1 Then
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
                        If MESG_FLG = 1 Then
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
                                Syuko_Update_Proc = SYS_CANCEL
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
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case BtErrDEAD_LOCK
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            
            Loop
                                        '在庫数
            If Upd_com = BtOpUpdate Then
                Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - SUMI_JITU_QTY, "00000000"))
                Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)            '排他フラグ
                Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                  '使用中子機ID
                Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                  '使用中プログラム
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
                                Syuko_Update_Proc = SYS_CANCEL
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
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case BtErrDEAD_LOCK
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    Case Else
                        Call File_Error(sts, Upd_com, "在庫データ")
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                        
                End Select
            Loop
'============================================================
    '*------------------------------------------------------'在庫移動歴出力
            sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                        Space(8), _
                                        JGYOBU, _
                                        NAIGAI, _
                                        HIN_GAI, _
                                        NYUKA_DT, _
                                        YOIN, _
                                        SUMI_JITU_QTY, _
                                        0, _
                                        ID, _
                                        TANTO_CODE, _
                                        RETRY, _
                                        CYU_KBN, _
                                        MEMO, _
                                        Ins_DateTime, _
                                        StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                        StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                        StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO)
            If sts Then
                Syuko_Update_Proc = sts
                Exit Function
            End If
        End If
'************************************************************
    '
    '---------------------------------------'商品化済み処理
    
        If MI_JITU_QTY <> 0 Then
        
            Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   '倉庫№
            Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '列
            Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '連
            Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '段
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
                        If SUMI_JITU_QTY > CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                            If MESG_FLG = 1 Then
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
                        If MESG_FLG = 1 Then
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
                                Syuko_Update_Proc = SYS_CANCEL
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
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case BtErrDEAD_LOCK
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            
            Loop
                                        '在庫数
            If Upd_com = BtOpUpdate Then
                Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - MI_JITU_QTY, "00000000"))
                Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)            '排他フラグ
                Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                  '使用中子機ID
                Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                  '使用中プログラム
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
                                Syuko_Update_Proc = SYS_CANCEL
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
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case BtErrDEAD_LOCK
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    Case Else
                        Call File_Error(sts, Upd_com, "在庫データ")
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                        
                End Select
            Loop
'============================================================
    '*------------------------------------------------------'在庫移動歴出力
            sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                        Space(8), _
                                        JGYOBU, _
                                        NAIGAI, _
                                        HIN_GAI, _
                                        NYUKA_DT, _
                                        YOIN, _
                                        0, _
                                        MI_JITU_QTY, _
                                        ID, _
                                        TANTO_CODE, _
                                        RETRY, _
                                        CYU_KBN, _
                                        MEMO, _
                                        Ins_DateTime, _
                                        StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                        StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                        StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO)
            If sts Then
                Syuko_Update_Proc = sts
                Exit Function
            End If
    
        End If
    End If
'============================================================
    If CYU_KBN = " " Then
    Else
        
        Wk_SYUKA_QTY = Wk_SYUKA_QTY + Wk_SUMI_JITU_QTY + Wk_MI_JITU_QTY
        If CYU_KBN <> CYU_KBN_KIN Then
        
        '*------------------------------------------------------'出荷予定更新
            If (CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))) = Wk_SYUKA_QTY Then
                Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_FIN)
                                                                '完了日付
                Call UniCode_Conv(Y_SYUREC.KAN_YMD, Format(Date, "yyyymmdd"))
                                                                
                '2011.03.30 更新日付の書き込み追加
                Call UniCode_Conv(Y_SYUREC.KAN_HMS, Format(Now, "HHMMSS"))
                                                                
                                                                '確定数量
                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(Y_SYUREC.SURYO, vbUnicode))
            Else
                                                                '確定数量
                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) + Wk_SYUKA_QTY, "0000000"))
            End If
            
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")              '使用端末ID（＝空白）
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")              '使用中プログラム（＝空白）
            com = BtOpUpdate
        Else
        
            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_FIN)
                                                                '完了日付
            Call UniCode_Conv(Y_SYUREC.KAN_YMD, Format(Date, "yyyymmdd"))
        
            
            '2011.03.30 更新日付の書き込み追加
            Call UniCode_Conv(Y_SYUREC.KAN_HMS, Format(Now, "HHMMSS"))
            
            Call UniCode_Conv(Y_SYUREC.JITU_SURYO, Format(Wk_SYUKA_QTY, "0000000"))
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")              '使用端末ID（＝空白）
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")              '使用中プログラム（＝空白）
            
            com = BtOpInsert
        End If


        '*------------------------------------------------------'出荷予定出力
        RETRY_CNT = 0
        Do
            sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'リトライ回数チェック
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '回数オーバー
                            Call File_Error(sts, com, "出荷予定", 0)
                            Syuko_Update_Proc = SYS_CANCEL
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
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case BtErrDuplicates
                    If com = BtOpUpdate Then            '更新時は異常
                        Call File_Error(sts, com, "出荷予定", MESG_FLG)
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                    End If
                    If Len(Trim(ID_NO)) <> 0 Then
                        Call File_Error(sts, com, "出荷予定", MESG_FLG)
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                    Else
                                                        'ID№再取込みしLOOP
                        sts = Den_No_Set_Proc(21, JGYOBU, GET_ID_NO)
                        If sts Then
                            Syuko_Update_Proc = sts
                            Exit Function
                        End If
                                                        'ID№
                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, GET_ID_NO)
                        Call UniCode_Conv(Y_SYUREC.ID_NO, GET_ID_NO)
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    End If
                Case BtErrDEAD_LOCK
                    Syuko_Update_Proc = SYS_CANCEL
                    Exit Function
                Case Else
                    Call File_Error(sts, com, "出荷予定", MESG_FLG)
                    Syuko_Update_Proc = SYS_ERR
                    Exit Function

            End Select
        Loop
    End If
'============================================================
                                        '最終出庫日
    Call UniCode_Conv(ITEMREC.LAST_SYU_DT, Format(Date, "yyyymmdd"))
    
    Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, Format(SUMI_JITU_QTY + MI_JITU_QTY, "00000000"))
    
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
                        Syuko_Update_Proc = SYS_CANCEL
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
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case BtErrDEAD_LOCK
                Syuko_Update_Proc = SYS_CANCEL
                Exit Function
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目マスタ")
                Syuko_Update_Proc = SYS_ERR
                Exit Function
                        
        End Select
    Loop
'============================================================
    
    Syuko_Update_Proc = False
    
End Function

Public Function wY_SYU_Open(Mode As Integer) As Integer
'****************************************************
'*      「出荷／出庫処理」    出荷予定ＯＰＥＮ処理
'*
'*  出荷予定ファイルを別ポインタでＯＰＥＮする
'*  (呼び元で起動時に１度だけ呼び出す)

'*  戻り値: false       :正常
'*          true        :異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    wY_SYU_Open = True
                                '在庫データ　フルパス取込み
    sts = GetIni("FILE", Y_SYU_ID, "SYS", c)
    
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wY_SYU_POS, wY_SYUREC, Len(wY_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
'-------------- ＯＰＥＮ処理での使用中は、立ち上げ時に１回だけのはずなので、常に画面入力とし、
'               ｷｬﾝｾﾙは、処理の起動ｷｬﾝｾﾙとする。
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    wY_SYU_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "出荷予定")
                Exit Function
        End Select
    Loop

    wY_SYU_Open = False

End Function

Public Function wY_SYU_CLOSE() As Integer

'****************************************************
'*      「出荷／出庫処理」    出荷予定ＣＬＯＳＥ処理
'*
'*  出荷予定ファイルを別ポインタでＣＬＯＳＥする
'*  (呼び元で終了時に１度だけ呼び出す)
'*  戻り値: false       :正常
'*          true        :異常
'****************************************************
Dim sts As Integer
    
    wY_SYU_CLOSE = True
    
    sts = BTRV(BtOpClose, wY_SYU_POS, wY_SYUREC, Len(wY_SYUREC), K3_wY_SYU, Len(K3_wY_SYU), 3)
    
    Select Case sts
        Case BtNoErr, BtErrNoOpen
        Case Else
            Call File_Error(sts, BtOpClose, "出荷予定")
            Exit Function
    End Select

    wY_SYU_CLOSE = False

End Function

