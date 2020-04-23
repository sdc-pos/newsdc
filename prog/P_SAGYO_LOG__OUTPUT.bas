Attribute VB_Name = "P_SAGYO_LOG_OUTPUT"
Option Explicit
Public Function P_SAGYO_LOG_OUTPUT_PROC(TANTO_CODE As String, _
                                    WEL_ID As String, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    MENU_NO As String, _
                                    YOIN As String, _
                                    Optional HIN_GAI As String = "                    ", _
                                    Optional SUMI_QTY As Long = 0, _
                                    Optional MI_QTY As Long = 0, _
                                    Optional FROM_LOCATION As String = "        ", _
                                    Optional TO_LOCATION As String = "        ", _
                                    Optional ID_NO As String = "        ", _
                                    Optional MTS As String = "        ", _
                                    Optional SS As String = "        ", _
                                    Optional RETRY As Integer = 10, _
                                    Optional SHIJI_No As String = "        ", _
                                    Optional HIN_CHECK_LABEL_CNT As String = "   ", _
                                    Optional HIN_CHECK_GENPIN_CNT As String = "   ", _
                                    Optional wkMTS As String = "        ", _
                                    Optional JAN_CODE As String = "                    ", _
                                    Optional MEMO As String = "                                        ", _
                                    Optional HIN_CHECK_GAISOU_CNT As String = "   ", _
                                    Optional HINBAN_DAMMY As String = "                    ") As Integer
'****************************************************
'*      作業ログ出力処理更新
'*
'*  作業ログの出力を行う。
'*  (引数の設定ミスはこちらではチェックしない)
'*  引数：  担当者(省略不可)
'*          ID(省略不可)
'*          事業部（省略不可）
'*          国内外（省略不可）
'*          ﾒﾆｭｰ番号（省略不可）
'*          要因(省略不可)
'*          外部品番（省略可 TOPﾒﾆｭｰ時）
'*          商品化済み実績数（=0を可とする、履歴のみ出力）
'*          未商品実績数（=0を可とする、履歴のみ出力）
'*          FROM棚（XXXXXXXX(倉庫№+列+連+段)省略可
'*          TO棚（XXXXXXXX(倉庫№+列+連+段)省略可）
'*          伝票ID(省略可)
'*          MTS(省略可)
'*          SS(省略可)
'*          リトライ(省略可 １桁目:1=画面メッセージ有 0:無，２桁目:リトライ回数(0～9 0:無限))
'*          指図票№            2010.09.03
'*          品番ﾁｪｯｸﾗﾍﾞﾙ件数    2010.09.03
'*          品番ﾁｪｯｸ現品票件数  2010.09.03
'*          向け先
'*          JANｺｰﾄﾞ             2011.08.18
'*          メモ                2014.07.01
'*          商品化ﾁｪｯｸ外装ｶｳﾝﾄ  2015.11.07
'*
'*  戻り値: false       :正常
'*          true        :継続可能な異常
'*          SYS_ERR     :継続できない異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'*
'****************************************************
Dim sts                 As Integer
Dim Sumi_Zaiko_Qty      As Long
Dim Mi_Zaiko_Qty        As Long

Dim RETRY_CNT           As Integer
Dim MESG_FLG            As Integer
Dim RETRY_SU            As Integer
    
Dim ans                 As Integer
                                            
    P_SAGYO_LOG_OUTPUT_PROC = True
                                            
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                                            
                                        '作業ログ出力
    Call UniCode_Conv(P_SAGYO_LOG_REC.JITU_DT, Format(Now, "YYYYMMDD"))     '実績日付
    Call UniCode_Conv(P_SAGYO_LOG_REC.JITU_TM, Format(Now, "HHMMSS"))       '実績時刻
    
    Call UniCode_Conv(P_SAGYO_LOG_REC.TANTO_CODE, TANTO_CODE)               '担当者ｺｰﾄﾞ
    Call UniCode_Conv(P_SAGYO_LOG_REC.WEL_ID, WEL_ID)                       '端末ID
    Call UniCode_Conv(P_SAGYO_LOG_REC.JGYOBU, JGYOBU)                       '事業部
    Call UniCode_Conv(P_SAGYO_LOG_REC.NAIGAI, NAIGAI)                       '国内外
    Call UniCode_Conv(P_SAGYO_LOG_REC.MENU_NO, MENU_NO)                     'TOPﾒﾆｭｰ
    Call UniCode_Conv(P_SAGYO_LOG_REC.RIRK_ID, YOIN)                        '作業要因
    
    Call UniCode_Conv(P_SAGYO_LOG_REC.ID_NO, ID_NO)                         '伝票ID
    Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_GAI, HIN_GAI)                     '品番
    If Trim(HINBAN_DAMMY) = "." Then                                        '2017.10.30
        Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_GAI, ".")
    End If
                                                                            '商品化済分実績数量
    
    If SUMI_QTY >= 0 Then
        Call UniCode_Conv(P_SAGYO_LOG_REC.SUMI_JITU_QTY, Format(SUMI_QTY, "00000000"))
    Else
        Call UniCode_Conv(P_SAGYO_LOG_REC.SUMI_JITU_QTY, Format(SUMI_QTY, "0000000"))
    End If
                                                                            
                                                                            '未商品分実績数量
    If MI_QTY >= 0 Then
        Call UniCode_Conv(P_SAGYO_LOG_REC.MI_JITU_QTY, Format(MI_QTY, "00000000"))
    Else
        Call UniCode_Conv(P_SAGYO_LOG_REC.MI_JITU_QTY, Format(MI_QTY, "0000000"))
    End If
    Call UniCode_Conv(P_SAGYO_LOG_REC.MUKE_CODE, MTS)                       'MTS
    Call UniCode_Conv(P_SAGYO_LOG_REC.SS_CODE, SS)                          'SS
    
    Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_SOKO, Mid(FROM_LOCATION, 1, 2))  'FROM 棚番
    Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_RETU, Mid(FROM_LOCATION, 3, 2))  'FROM 棚番
    Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_REN, Mid(FROM_LOCATION, 5, 2))   'FROM 棚番
    Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_DAN, Mid(FROM_LOCATION, 7, 2))   'FROM 棚番
        
    Call UniCode_Conv(P_SAGYO_LOG_REC.TO_SOKO, Mid(TO_LOCATION, 1, 2))      'TO 棚番
    Call UniCode_Conv(P_SAGYO_LOG_REC.TO_RETU, Mid(TO_LOCATION, 3, 2))      'TO 棚番
    Call UniCode_Conv(P_SAGYO_LOG_REC.TO_REN, Mid(TO_LOCATION, 5, 2))       'TO 棚番
    Call UniCode_Conv(P_SAGYO_LOG_REC.TO_DAN, Mid(TO_LOCATION, 7, 2))       'TO 棚番
                                                                            '出力元ﾌﾟﾛｸﾞﾗﾑ
    Call UniCode_Conv(P_SAGYO_LOG_REC.PRG_ID, StrConv(App.EXEName, vbUpperCase))
        
    Call UniCode_Conv(P_SAGYO_LOG_REC.WORK_TM, "")
        
        
    Call UniCode_Conv(P_SAGYO_LOG_REC.SHIJI_No, SHIJI_No)                   '指示№ 2010.09.03
                                                                            '品番ﾁｪｯｸﾗﾍﾞﾙ件数 2010.09.03
    Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_CHECK_LABEL_CNT, HIN_CHECK_LABEL_CNT)
                                                                            '品番ﾁｪｯｸ現品票件数 2010.09.03
    Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_CHECK_GENPIN_CNT, HIN_CHECK_GENPIN_CNT)
        
        
        
    Call UniCode_Conv(P_SAGYO_LOG_REC.FILLER, "")
        
    '2011.01.19
    If Trim(wkMTS) <> "" Then
        Call UniCode_Conv(P_SAGYO_LOG_REC.MUKE_CODE, wkMTS)                     'MTS
    End If
        
        
        
    '2011.08.18
    Call UniCode_Conv(P_SAGYO_LOG_REC.JAN_CODE, JAN_CODE)                   'JANｺｰﾄﾞ    2011.08.18
        
        
    '2014.07.01
    Call UniCode_Conv(P_SAGYO_LOG_REC.MEMO, MEMO)                           'メモ    2014.07.01
        
    '2015.11.07
    Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_CHECK_GAISOU_CNT, HIN_CHECK_GAISOU_CNT)
        
        
    RETRY_CNT = 0
    Do
        
        sts = BTRV(BtOpInsert, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), K0_P_SAGYO_LOG, Len(K0_P_SAGYO_LOG), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                If RETRY_SU <> 0 Then
                    
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                        Call File_Error(sts, BtOpInsert, "作業ログ", 0)
                        P_SAGYO_LOG_OUTPUT_PROC = SYS_CANCEL
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
                    ans = MsgBox("他端末でデータ使用中です。<P_SAGYO_LOG.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        P_SAGYO_LOG_OUTPUT_PROC = SYS_CANCEL
                        Exit Function
                    End If
                End If
            
            
            Case BtErrDEAD_LOCK 'デッドロック   2010.11.10
                P_SAGYO_LOG_OUTPUT_PROC = SYS_CANCEL
                Exit Function
            
            
            
            Case Else
                Call File_Error(sts, BtOpInsert, "作業ﾛｸﾞ")
                P_SAGYO_LOG_OUTPUT_PROC = SYS_ERR
                Exit Function
        End Select
    Loop
                            '正常終了
    P_SAGYO_LOG_OUTPUT_PROC = False

End Function
