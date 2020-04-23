Attribute VB_Name = "SYUKO_SEK_UPDATE"
Option Explicit


Public Function Syuko_SEK_Update_Proc(JGYOBU As String, _
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
                                    Optional BIN_NO As String = "  ", _
                                    Optional LOG_NON As Integer = 0, _
                                    Optional Ins_DateTime As String, _
                                    Optional mode As Integer = 0) As Integer
'****************************************************
'*      「出荷／出庫処理」在庫データ更新
'*      大阪ＰＣ用　2007.03.17
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
'*          便№                        2007.05.16
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
    
Dim Wk_SUMI_JITU_QTY    As Long
Dim Wk_MI_JITU_QTY      As Long
Dim Wk_SYUKA_QTY        As Long

'Dim Ins_DateTime    As String * 14
    
Dim wkYOIN      As String * 2
    
    
    
'2010.01.06
Dim svSYUKA_YMD As String
Dim svDEN_NO    As String
Dim svTOK_KBN   As String
Dim svID_NO     As String
'2010.01.06
    
    
'''''   2011.04.11
Dim Total_SUMI_JITU_QTY     As Long
Dim Total_MI_JITU_QTY       As Long
'''''   2011.04.11
    
    
    
    
    Syuko_SEK_Update_Proc = True
                                                                      
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                                        
                                        
    Wk_SUMI_JITU_QTY = SUMI_JITU_QTY
    Wk_MI_JITU_QTY = MI_JITU_QTY
    Wk_SYUKA_QTY = SYUKA_QTY
                                        
'    Ins_DateTime = Format(Now, "YYYYMMDDHHMMSS")
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
                        Syuko_SEK_Update_Proc = SYS_CANCEL
                        Exit Function
                    
                    End If
                
                End If
                
                If MESG_FLG = 0 Then
                    DoEvents
                Else
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Syuko_SEK_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Syuko_SEK_Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
'============================================================
'*------------------------------------------------------'入荷日指定無し 在庫データ読込み（スキャナ処理）
'
'---------------------------------------'商品化済み～未商品を順次引き当てる

    If LOG_NON = 1 Then
    Else
        If SUMI_JITU_QTY = 0 And MI_JITU_QTY = 0 Then
            
            
            If Left(YOIN, 1) = ACT_DENPYO_ID Or _
                Left(YOIN, 1) = ACT_SYUKA_HYO Or _
                Left(YOIN, 1) = ACT_SYUKA_HYO_OSAKA Then
                wkYOIN = ACT_SYUKA_KEI & CYU_KBN
            End If
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




    Zan_Qty = SYUKA_QTY
    Do
        Call UniCode_Conv(K0_ZAIKO.Soko_No, Mid(FROM_LOCATION, 1, 2))   '倉庫№
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
                    If FROM_LOCATION <> (StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                        StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                        StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                        StrConv(ZAIKOREC.Dan, vbUnicode)) Or _
                        JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        
'                        If MESG_FLG = 1 Then
'                            Beep
'                            MsgBox "在庫数が不足しています。更新処理を中止します。", vbOKOnly, "確認入力"
'                        End If
'                        Exit Function
                    
                        mode = 1
                        GoTo SYUKA_UPDATE
                    
                    
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

'                    If MESG_FLG = 1 Then
'                        Beep
'                        MsgBox "在庫数が不足しています。更新処理を中止します。", vbOKOnly, "確認入力"
'                    End If
                    
'                    Exit Function
                    
                    mode = 1
                    GoTo SYUKA_UPDATE
                    

                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                'リトライ回数チェック
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '回数オーバー
                            Call File_Error(sts, com + BtSNoWait, "在庫データ", 1)
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If

                    End If

                    If MESG_FLG = 0 Then
                        DoEvents
                    Else
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "在庫データ")
                    Syuko_SEK_Update_Proc = SYS_ERR
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11
            Total_SUMI_JITU_QTY = Total_SUMI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11
        Else
            MI_JITU_QTY = MI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11
            Total_MI_JITU_QTY = Total_MI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11
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
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                
                    If MESG_FLG = 0 Then
                        DoEvents
                    Else
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, Upd_com, "在庫データ")
                    Syuko_SEK_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        Loop
    '============================================================
        '*------------------------------------------------------'在庫移動歴出力
        
        
        
        
        '出荷予定＆向け先から情報をクリア取得 2007.06.02
            
            
        '2010.01.06
        svSYUKA_YMD = StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode)
        svDEN_NO = StrConv(Y_SYUREC.DEN_NO, vbUnicode)
        svTOK_KBN = StrConv(Y_SYUREC.TOK_KBN, vbUnicode)
        svID_NO = StrConv(Y_SYUREC.ID_NO, vbUnicode)
        '2010.01.06
            
            
        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, "")
        Call UniCode_Conv(Y_SYUREC.DEN_NO, "")
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
        Call UniCode_Conv(IDOREC.TOKU_MARK, "")
        Call UniCode_Conv(Y_SYUREC.ID_NO, "")
            
    
        Call UniCode_Conv(MTSREC.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
        Call UniCode_Conv(MTSREC.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
                    
        Call UniCode_Conv(MTSREC.MUKE_NAME, StrConv(Y_SYUREC.MUKE_NAME, vbUnicode))
        Call UniCode_Conv(MTSREC.SS_NAME, "")
            
        Call UniCode_Conv(MTSREC.MUKE_DNAME, StrConv(Y_SYUREC.MUKE_NAME, vbUnicode))
            
            
        
        
        
        
        
        
        
        
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
                                    StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), _
                                    MENU_NO, MUKE_CODE, , ID_NO, _
                                    BIN_NO, _
                                    DEN_NO, SYUKA_YMD, 1)

        If sts Then
            Syuko_SEK_Update_Proc = sts
            Exit Function
        End If
        
        Zan_Qty = Zan_Qty - WK_Qty

        If Zan_Qty <= 0 Then
            Exit Do                     '引き落とし終了
        End If
    Loop

SYUKA_UPDATE:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11
    If LOG_NON = 1 Then
    Else
        If SYUKA_QTY <> 0 Then
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
        End If
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11


'============================================================
'   出荷予定(ﾎｽﾄｲﾒｰｼﾞ)--＞出荷予定の更新
'============================================================
    
    
    '今回出荷数のKEEP
    Wk_SYUKA_QTY = Wk_SYUKA_QTY + Wk_SUMI_JITU_QTY + Wk_MI_JITU_QTY
    
    Call UniCode_Conv(K4_Y_SYU_H.ID_NO, ID_NO)
    com = BtOpGetGreaterEqual
    
    
    Do
        DoEvents
        
        Do
            sts = BTRV(com + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
            Select Case sts
                Case BtNoErr
                    
                                        
                    If StrConv(Y_SYU_HREC.ID_NO, vbUnicode) <> ID_NO Then
                        sts = BtErrEOF
                    End If
                    
                    Exit Do
                Case BtErrEOF
                    If MESG_FLG = 1 Then
                        
                        
                        If Wk_SYUKA_QTY <> 0 Then
                            Beep
                            MsgBox "出荷予定が存在しません。更新処理を中止します。", vbOKOnly, "確認入力"
                        End If
                    End If
                    
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'リトライ回数チェック
                    If RETRY_SU <> 0 Then
                        
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                            '回数オーバー
                            Call File_Error(sts, com + BtSNoWait, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)", 0)
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        
                        End If
                    
                    End If
                    
                    If MESG_FLG = 0 Then
                        DoEvents
                    Else
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYU_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    Syuko_SEK_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        
        Loop
        
        '処理終了
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        
        If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) = "1" Then
            'ｷｬﾝｾﾙ分は未処理
        
            '2010.01.06
            Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, svSYUKA_YMD)
            Call UniCode_Conv(Y_SYUREC.DEN_NO, svDEN_NO)
            Call UniCode_Conv(Y_SYUREC.TOK_KBN, svTOK_KBN)
            Call UniCode_Conv(Y_SYUREC.ID_NO, svID_NO)
            '2010.01.06
        
        
        Else
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, StrConv(Y_SYU_HREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
        
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrEOF
                        If MESG_FLG = 1 Then
                            
                            
                            If Wk_SYUKA_QTY <> 0 Then
                                Beep
                                MsgBox "出荷予定が存在しません。更新処理を中止します。", vbOKOnly, "確認入力"
                            End If
                        End If
                        
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                            
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                                '回数オーバー
                                Call File_Error(sts, com + BtSNoWait, "出荷予定", 0)
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            
                            End If
                        
                        End If
                        
                        If MESG_FLG = 0 Then
                            DoEvents
                        Else
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定")
                        Syuko_SEK_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            Loop
            
            
            WK_Qty = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
            
            If Wk_SYUKA_QTY >= WK_Qty Then
                                  
                                  
                                  
                Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_FIN)
                                                                '完了日付
                Call UniCode_Conv(Y_SYUREC.KAN_YMD, Format(Date, "yyyymmdd"))
                                                                '確定数量
                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(Y_SYUREC.SURYO, vbUnicode))
                        
            
            
            Else
            
                
                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) + Wk_SYUKA_QTY, "0000000"))
            
            End If
        
        
                    
            Call UniCode_Conv(Y_SYU_HREC.J_SURYO, StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
        
        
            Wk_SYUKA_QTY = Wk_SYUKA_QTY - WK_Qty
        
        
        
        
        
        
        
        
        
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")              '使用端末ID（＝空白）
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")              '使用中プログラム（＝空白）
    
            '*------------------------------------------------------'出荷予定出力
            RETRY_CNT = 0
            Do
                
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '回数オーバー
                                Call File_Error(sts, BtOpUpdate, "出荷予定", 0)
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
    
                            End If
    
                        End If
    
                        If MESG_FLG = 0 Then
                            DoEvents
                        Else
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定", MESG_FLG)
                        Syuko_SEK_Update_Proc = SYS_ERR
                        Exit Function
    
                End Select
            Loop
        
        
        
        
        
            '*------------------------------------------------------'注文データ出力
            Call UniCode_Conv(K2_Y_SYU_TEI.KEN_NO, StrConv(Y_SYU_HREC.SEK_KEN_NO, vbUnicode))
            Call UniCode_Conv(K2_Y_SYU_TEI.HIN_NO, StrConv(Y_SYU_HREC.SEK_HIN_NO, vbUnicode))
        
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                                            
                    Case BtErrKeyNotFound
                        If MESG_FLG = 1 Then
                            
                            
                            If Wk_SYUKA_QTY <> 0 Then
                                Beep
                                MsgBox "注文データが存在しません。更新処理を中止します。", vbOKOnly, "確認入力"
                            End If
                        End If
                        
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                            
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                                '回数オーバー
                                Call File_Error(sts, com + BtSNoWait, "注文ﾃﾞｰﾀ")
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            
                            End If
                        
                        End If
                        
                        If MESG_FLG = 0 Then
                            DoEvents
                        Else
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<Y_SYU_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "注文ﾃﾞｰﾀ")
                        Syuko_SEK_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            
            Loop
        
        
            Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_TANTO, StrConv(App.EXEName, vbUpperCase))
            Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_DATETIME, Ins_DateTime)
        
            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_DATETIME, Ins_DateTime)
        
        
        
        
        
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '回数オーバー
                                Call File_Error(sts, BtOpUpdate, "注文ﾃﾞｰﾀ")
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
    
                            End If
    
                        End If
    
                        If MESG_FLG = 0 Then
                            DoEvents
                        Else
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<Y_SYU_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, com, "出荷予定", MESG_FLG)
                        Syuko_SEK_Update_Proc = SYS_ERR
                        Exit Function
    
                End Select
            Loop
        
        
        
        
        
        
        
        End If
        
        Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_TANTO, StrConv(App.EXEName, vbUpperCase))
        Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_DATETIME, Ins_DateTime)
    
        Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
        Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Ins_DateTime)

        '*------------------------------------------------------'出荷予定(ﾎｽﾄｲﾒｰｼﾞ)出力
        RETRY_CNT = 0
        Do
            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'リトライ回数チェック
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '回数オーバー
                            Call File_Error(sts, BtOpUpdate, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)", 0)
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If

                    If MESG_FLG = 0 Then
                        DoEvents
                    Else
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYU_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com, "出荷予定", MESG_FLG)
                    Syuko_SEK_Update_Proc = SYS_ERR
                    Exit Function

            End Select
        Loop



        If Wk_SYUKA_QTY <= 0 Then
            Exit Do
        End If


        com = BtOpGetNext
    Loop

'============================================================
    If mode = 1 Then
        sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        If sts Then
            Call File_Error(sts, BtOpUnlock, "品目ﾏｽﾀ", MESG_FLG)
            Syuko_SEK_Update_Proc = SYS_ERR
            Exit Function
        End If
        Syuko_SEK_Update_Proc = False
        Exit Function
    End If
                                        
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
                        Syuko_SEK_Update_Proc = SYS_CANCEL
                        Exit Function
                    
                    End If
                    
                End If
                
                If MESG_FLG = 0 Then
                    DoEvents
                Else
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Syuko_SEK_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目マスタ")
                Syuko_SEK_Update_Proc = SYS_ERR
                Exit Function
                        
        End Select
    Loop
'============================================================
    
    Syuko_SEK_Update_Proc = False
    
End Function


