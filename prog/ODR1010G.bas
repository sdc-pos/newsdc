Attribute VB_Name = "ODR1010G"
Option Explicit
'********************************************************************
'*
'*              ＯＤＲ１０１０用　共通変数
'*
'********************************************************************

Public ODR10102_Return As Integer         '確認画面終了状態


Public GW_PURA(0 To 30)          As String       '在訂（＋）コード
Public GW_MAINA(0 To 30)         As String       '在訂（－）コード


Public DIS_ORDR_NO      As String       '親部品　注文№
Public DIS_BUNNO        As String       '分納回数
Public DIS_OYA_ITEM     As String       '親部品コード
Public DIS_ORDR_QTY     As String       '注文数量
Public DIS_NOUKI        As String       '親部品　注文納期
Public DIS_OK_DT        As String       '組立可能日
Public DIS_KAITO        As String       '親部品　回答納期
Public DIS_USE_YM       As String       '使用月
Public DIS_FIN_DT       As String       '完了日付
Public DIS_KEY          As String       'データＫｅｙ情報

Public DIS2_QTY         As String       '注文数量
Public DIS2_KAITO       As String       '親部品　回答納期

Public Key_SIMUKE       As String       '仕向け先
Public Key_JIGYOBU      As String       '事業部
Public Key_NAIGAI       As String       '国内外
Public Key_USE_YM       As String       '使用月（YYYYMM)
Public Key_INS_NO       As String       '登録順
Public Key_HinGai       As String       '親品番
Public Key_ORDER_NO     As String       '親品番　注文№
Public Key_BUN_NO       As String       '分納回数

Public Key_Ko_HinGai    As String       '子品番         2010/05/07追加
Public Key_Ko_JIGYOBU      As String    '子品番 事業部
Public Key_Ko_NAIGAI       As String    '子品番 国内外
Sub Main()
'2017.01.16 追加
    
    
Dim lngReturnValue      As Long
Dim strMyTitle          As String
Dim lngPrevHwnd         As Long
Dim lngTopHwnd          As Long
Dim lngThreadID1        As Long
Dim lngThreadID2        As Long
    
    
    
    
    Last_JGYOBU = Trim(Command)






    ' 2重起動の場合は、手前に持ってきて自分自身は終了する
    strMyTitle = App.Title
    App.Title = "$" & App.Title
    lngPrevHwnd = FindWindow("ThunderRT6Main", strMyTitle)
    If lngPrevHwnd <> 0 Then
    lngTopHwnd = GetLastActivePopup(lngPrevHwnd)
    If IsIconic(lngTopHwnd) = WIN32API_TRUE Then
    lngReturnValue = ShowWindow(lngTopHwnd, SW_NORMAL)
    End If
    lngThreadID1 = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
    lngThreadID2 = GetCurrentThreadId()
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 1)
    lngReturnValue = SetForegroundWindow(lngTopHwnd)
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 0)
    Exit Sub
    End If
    App.Title = strMyTitle










    ODR10101.Show
End Sub

Function OUT_TP1(HIN_GAI As String) As Integer
'
'           構成展開　→　所要量Ｆ出力
'

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_SeqKey    As String
Dim W_Ko_HinCD  As String

Dim W_QTY       As Double
Dim W_STR       As String
Dim W_Date      As String

Dim Fin_Qty     As Double           '2008.04.30

Dim W_A_Nouki   As String

    OUT_TP1 = True
        
    '
    '指定された親品番の構成子部品の全てに関して、展開情報を設定する！
    '
    W_SeqKey = ""
    
    If Trim(HIN_GAI) = "AD-HEPSC010" Then
        W_SeqKey = ""
    End If
    
    If Trim(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode)) = "" Then
        W_A_Nouki = "99999999"
    Else
        W_A_Nouki = StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode)
    End If
    
    '   最初に「親レコード」を取得。
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, GW_SIMUKE)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, HIN_GAI)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "0")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    com = BtOpGetGreaterEqual
    sts = BTRV(com, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                'Beep
                'MsgBox "指定された工程がありません。"
        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
            yn = MsgBox("他で使用中です！<構成Ｆ>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
            If yn = vbNo Then
                Exit Function
            End If
        Case Else
            Call File_Error(sts, com, "P_COMPO")
            Exit Function
    End Select
    If sts <> BtNoErr Then
        MsgBox "構成情報　未登録！ <" & HIN_GAI & ">", vbExclamation
        Exit Function
    End If
    
    If Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) = "AD-DLHS03A05" Then
        sts = BtNoErr
    End If
    
    '   ここから「子部品レコード」を読みながら展開Ｆを出力する。
    com = BtOpGetNext
    Do
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<構成Ｆ>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "P_COMPO")
                Exit Function
        End Select
        If sts <> BtNoErr Then
            Exit Do
        End If
        If Trim(StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(HIN_GAI) Then Exit Do
        'If Trim(StrConv(P_COMPO_O_REC.DATA_KBN, vbUnicode)) <> "0" Then Exit Do
        
        W_SeqKey = StrConv(P_COMPO_K_REC.SEQNO, vbUnicode)
        
        If CInt(W_SeqKey) <> 0 Then                 '構成部品レコード？
            
            Call ODR_TEMP1_CLR
    
            Call UniCode_Conv(ODR_TP1_R.KAITO_DT, W_A_Nouki)
            Call UniCode_Conv(ODR_TP1_R.CYUMON_DT, Trim(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode)))
            Call UniCode_Conv(ODR_TP1_R.USE_YM, Trim(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode)))
            
            Call UniCode_Conv(ODR_TP1_R.SHIMUKE, GW_SIMUKE)
            Call UniCode_Conv(ODR_TP1_R.JGYOBU, GW_JIGYOBU)
            Call UniCode_Conv(ODR_TP1_R.NAIGAI, GW_NAIGAI)
            Call UniCode_Conv(ODR_TP1_R.INS_NO, Trim(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode)))
            Call UniCode_Conv(ODR_TP1_R.ORDER_NO, Trim(StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode)))
            Call UniCode_Conv(ODR_TP1_R.BUN_NO, Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode)))
            Call UniCode_Conv(ODR_TP1_R.HIN_GAI, Trim(HIN_GAI))
            
            

'2008.04.04            Call UniCode_Conv(ODR_TP1_R.KO_NAIGAI, Trim(StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)))
'2008.04.04            Call UniCode_Conv(ODR_TP1_R.KO_JGYOBU, Trim(StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)))
            
            Call UniCode_Conv(ODR_TP1_R.KO_JGYOBU, SHIZAI)      '2008.04.04
            Call UniCode_Conv(ODR_TP1_R.KO_NAIGAI, NAIGAI_NAI)  '2008.04.04
            
            Call UniCode_Conv(ODR_TP1_R.KO_HIN_GAI, Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)))
            
            Call UniCode_Conv(ODR_TP1_R.KO_SYUBETSU, Trim(StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode)))
            If IsNull(StrConv(ODR_TP1_R.KO_SYUBETSU, vbUnicode)) Then
                Call UniCode_Conv(ODR_TP1_R.KO_SYUBETSU, "")
            End If
            If Left(StrConv(ODR_TP1_R.KO_SYUBETSU, vbUnicode), 1) < " " Then
                Call UniCode_Conv(ODR_TP1_R.KO_SYUBETSU, "")
            End If
            Call UniCode_Conv(ODR_TP1_R.KO_QTY, Trim(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)))
            
            
            
            W_QTY = CDbl(Trim(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)))
            W_QTY = W_QTY * CDbl(Trim(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)))
            W_STR = CStr(W_QTY)
            Call UniCode_Conv(ODR_TP1_R.ALL_QTY, W_STR)     '展開数
                        
            If W_QTY < 0 Then
                W_QTY = W_QTY * 1
            End If
            
            W_STR = CStr(Abs(W_QTY))
    
            Fin_Qty = 0
            
            If CDbl(Trim(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode))) <= 0 Then
                
                '08.11.27コメントに！
                'Call UniCode_Conv(ODR_TP1_R.USE_QTY, W_STR)                     '使用数
                
            
            Else
            
                'If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) = "" Then
                '    Call UniCode_Conv(ODR_TP1_R.NED_QTY, W_STR)                 '必要数
                '    Call UniCode_Conv(ODR_TP1_R.REQ_QTY, W_STR)                 '所要数
                'Else
                '                '　完了日　≦　繰越日：所要数（在庫引当対象）とする！
                '    If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) <= GW_SHIMEBI Then
                '        Call UniCode_Conv(ODR_TP1_R.USE_QTY, W_STR)             '使用数
                '        Call UniCode_Conv(ODR_TP1_R.REQ_QTY, W_STR)             '所要数
                '    Else
                '                '　完了日　＞　繰越日：何もなし！
                '
                '
                '    End If
                'End If
                '08.11.27上記を下記に変更
                If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) = "" Then
                    Call UniCode_Conv(ODR_TP1_R.NED_QTY, W_STR)                 '必要数
                    Call UniCode_Conv(ODR_TP1_R.REQ_QTY, W_STR)                 '所要数
                    
                    '08.12.12追加
                    Call UniCode_Conv(ODR_TP1_R.KAN_KB, "1")                    '未完＝１
                    Call UniCode_Conv(ODR_TP1_R.OK_DT, "")
                Else
                            '完成
                    Call UniCode_Conv(ODR_TP1_R.USE_QTY, W_STR)                 '使用数
                    
                    '08.12.12追加
                    Call UniCode_Conv(ODR_TP1_R.KAN_KB, "0")                    '完成＝０
                    Call UniCode_Conv(ODR_TP1_R.OK_DT, Trim(StrConv(ODR_ORDER_REC.KUMI_OK_DT, vbUnicode)))
                    '   ↑完成時、元々の組立可能日をセット！
                End If
            
            End If


            Call UniCode_Conv(ODR_TP1_R.UPDT_DT, Format(Date, "yyyymmdd"))
            Call UniCode_Conv(ODR_TP1_R.UPDT_TM, Format(Time, "hhmmss"))
            
            
            '2008/09/19 品目Ｍ登録チェック
            Call UniCode_Conv(K0_ITEM.JGYOBU, RTrim(StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode)))
            Call UniCode_Conv(K0_ITEM.NAIGAI, RTrim(StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode)))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)))
                
            Do
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        
                Select Case sts
                    Case BtNoErr
                        
                        Exit Do
                        
                    Case BtErrKeyNotFound, BtErrEOF
                        W_STR = StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode)
                        W_STR = W_STR & "-" & StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode)
                        W_STR = W_STR & "-" & Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
                        
                        'yn = MsgBox("品目未登録！<" & W_Str & ">" & Chr(13) & Chr(10) & _
                        '            "　続行しますか？", vbYesNo + vbDefaultButton1 + vbExclamation, "確認入力")
                        yn = vbYes
                        If yn = vbNo Then Exit Function
                        Exit Do
                        
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        yn = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If yn <> vbYes Then
                            Exit Function
                        End If
                        
                        
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
            Loop
    
    
    
                                    '2008/09/19 品目Ｍ無し：登録しない！
            If Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)) = "TEST" Then
                sts = BtNoErr
            End If
            
            
            If sts = BtNoErr Then
'                Call UniCode_Conv(ODR_TP1_R.FILLER, "ITEM未登録")   '2010/06/15 要望で追加！
'            End If
            '2010/06/15     再度、展開（出力）しないように修正！
            '               その分、子部品展開画面10103に構成Ｍ内容を表示！
            '
                Do
                    sts = BTRV(BtOpInsert, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, BtOpInsert, "ODR_TEMP1")
                            Exit Function
                    End Select
                Loop
            End If
            
            Key_SIMUKE = GW_SIMUKE
            
'2008.04.04            Key_JIGYOBU = Trim(StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04.            Key_NAIGAI = Trim(StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
            Key_JIGYOBU = SHIZAI        '2008.04.04
            Key_NAIGAI = NAIGAI_NAI     '2008.04.04
            
            GW_HINGAI_KO = Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
            
'2008.04.04            GW_JIGYOBU_KO = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
'2008.04.04            GW_NAIGAI_KO = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
            GW_JIGYOBU_KO = SHIZAI      '2008.04.04
            GW_NAIGAI_KO = NAIGAI_NAI   '2008.04.04
            
        End If
        
        com = BtOpGetNext
    Loop
        
    OUT_TP1 = False


End Function


Function SET_ALL() As Integer
'
'                                           TP1を基に在庫情報をセットする親SUB
'
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_S         As Integer
            
    SET_ALL = True
           
    GW_JIGYOBU_KO = ""
    GW_NAIGAI_KO = ""
    GW_HINGAI_KO = ""
    
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K2_ODR_TEMP1, Len(K2_ODR_TEMP1), 2)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<中間Ｆ１>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Function
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP1")
                    Exit Function
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        W_S = 0
        If GW_JIGYOBU_KO <> StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode) Then W_S = 1
        If GW_NAIGAI_KO <> StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode) Then W_S = 1
        If GW_HINGAI_KO <> StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode) Then W_S = 1
        
        
        GW_JIGYOBU_KO = StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode)
        GW_NAIGAI_KO = StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode)
        GW_HINGAI_KO = StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)
        
        
        If W_S = 1 Then
            
                '       在庫情報を検索・出力
            If SET_I_ZAIKO() Then
                MsgBox "在庫情報設定エラー！", vbExclamation
                Exit Function
            End If
            
                '       発注情報（入庫予定）情報　検索・出力
            If SET_ODR_ZAN() Then
                MsgBox "発注庫情報設定エラー！", vbExclamation
                Exit Function
            End If
        
                '2008/09/10 在訂情報の在庫情報を集計
            If SET_ZAITEI() Then
                MsgBox "在訂情報の在庫情報設定エラー！", vbExclamation
                Exit Function
            End If
                    
            '2008/05/31 半製品の在庫情報を集計
            If SET_H_SEIHIN() Then
                MsgBox "半製品の在庫情報設定エラー！", vbExclamation
                Exit Function
            End If
        
        End If
        
        
        com = BtOpGetNext
    Loop
        
                '       仕入実績情報　検索・出力
    If SET_UKEIRE() Then
        MsgBox "仕入実績情報設定エラー！", vbExclamation
        Exit Function
    End If
            
     
    SET_ALL = False

End Function
Function SET_I_ZAIKO() As Integer

        '       在庫情報を検索・出力（テーブル）
        
        '       在庫数＝０でも出力（08/09/11）
        
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String
Dim W_Edit      As String
Dim W_Zaiko     As Double
Dim W_QTY       As Double
    
    SET_I_ZAIKO = True

    Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)        '事業部
    Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)         '国内外
    Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, GW_HINGAI_KO)        '子品番
    Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "a")                      'io区分
    Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, GW_TOUGETU)              '使用月
    Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")                   '対象日付   YYYYMMDD
    Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")                    '注文№
    
    Do
        sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                'MsgBox "指定された工程がありません。"
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                Sleep (500)
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                Exit Function
        End Select
    Loop
    
    If sts = BtNoErr Then           '登録済み
        SET_I_ZAIKO = False
        Exit Function
    End If
    
    W_Zaiko = 0
    
            '-------------------------------------------------- '現在庫獲得
            
    Call UniCode_Conv(K0_ITEM.JGYOBU, GW_JIGYOBU_KO)     '事業部
    Call UniCode_Conv(K0_ITEM.NAIGAI, GW_NAIGAI_KO)      '国内外
    Call UniCode_Conv(K0_ITEM.HIN_GAI, GW_HINGAI_KO)     '子品番
    
    Do
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                Sleep (500)
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ITEM")
                'Exit Function
        End Select
    Loop
    
    If sts = BtNoErr Then
                
        If Trim(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) <> "" Then
            If IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
                W_Zaiko = CDbl(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode))
            End If
        End If
        
    End If
                        

    Call ODR_TEMP2_CLR
    
        '子　事業部
    Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, GW_JIGYOBU_KO)
        '子　国内外
    Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, GW_NAIGAI_KO)
        '子品番
    Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, GW_HINGAI_KO)
        'io区分
    Call UniCode_Conv(ODR_TP2_R.IO_KB, "a")
        '使用月
    Call UniCode_Conv(ODR_TP2_R.USE_YM, GW_TOUGETU)
        '納期
    Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, "")
        '注文№
    Call UniCode_Conv(ODR_TP2_R.ORDER_NO, "")
        

    W_STR = CStr(W_Zaiko)
    If Trim(W_STR) = "" Then W_STR = "0"
        
    Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)     '在庫数
    Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)    '元々の在庫数
        
    Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
    Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
    Do
        sts = BTRV(BtOpInsert, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                Sleep (500)
            Case Else
                Call File_Error(sts, BtOpInsert, "ODR_TEMP2")
                Exit Function
        End Select
    Loop
    
    SET_I_ZAIKO = False
    
End Function

Function SET_ODR_ZAN() As Integer

        '       発注情報（入庫予定）情報　検索＆出力    io区分＝ｆ


Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_Sw        As Integer

Dim W_STR       As String
Dim W_Edit      As String
Dim W_Zaiko     As Double
Dim W_QTY       As Double

Dim W_Date      As String       '当月の１日
Dim W_Today     As String       '本日
Dim wkYYMMDD     As String

Dim W_ZENGETU   As String


Dim W_Kan_DT    As String

    SET_ODR_ZAN = True
    
    
    W_Today = Format(Date, "yyyymmdd")
    W_Date = Left(W_Today, 6) & "01"
    W_STR = Left(W_Date, 4) & "/" & Mid(W_Date, 5, 2) & "/" & Right(W_Date, 2)
    
    
    wkYYMMDD = Left(GW_TOUGETU, 4) & "/" & Mid(GW_TOUGETU, 5, 2) & "/01"
    
    W_ZENGETU = Left(Format(DateAdd("d", -1, wkYYMMDD), "yyyymmdd"), 6) & "01"
    
    
    W_Zaiko = 0
    Call UniCode_Conv(K1_P_SHORDER.JGYOBU, GW_JIGYOBU_KO)
    Call UniCode_Conv(K1_P_SHORDER.NAIGAI, GW_NAIGAI_KO)
    Call UniCode_Conv(K1_P_SHORDER.HIN_GAI, GW_HINGAI_KO)
    Call UniCode_Conv(K1_P_SHORDER.ORDER_DT, "")
    Call UniCode_Conv(K1_P_SHORDER.ORDER_NO, "")
    com = BtOpGetGreaterEqual
    Do
        yn = 0
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K1_P_SHORDER, Len(K1_P_SHORDER), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
            Case Else
                Call File_Error(sts, com, "P_SHORDER")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(P_SHORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then Exit Do
        If Trim(StrConv(P_SHORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then Exit Do
        If Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI_KO) Then Exit Do
           
        '   2008/09 ＱＡの№１３による救済！
        If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) = "" Then
            Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, StrConv(P_SHORDER_REC.KAN_DT, vbUnicode))
        End If


If Trim(GW_HINGAI_KO) = "C215" Then
                W_Sw = True
End If
               
        W_Sw = True
               
        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = "1" Then W_Sw = False   'キャンセル？
            
            
            '   使用月：未設定　→　対象外！
        If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) = "" Then W_Sw = False
          
          
            '   使用月 ＜ 基準月　→　対象外！
        If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) < GW_TOUGETU Then W_Sw = False
        
        
            '   使用月　＞　２０ケ月　→　対象外！  2008/12/02
        If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) <> "" Then
            W_STR = Left(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 4) & "/" & _
                        Right(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 2)
            If W_STR > GW_MAX_YYMM Then
                W_Sw = False
            End If
       End If
       
       
       '                2008.12.17 完了F判定　追加
       If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = "1" Then W_Sw = False
       
            '   完了日      '2008/09/10 未完了のみ対象！？
        'If Trim(StrConv(P_SHORDER_REC.KAN_DT, vbUnicode)) <> "" Then
        '    W_Sw = False
        'End If
                       
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        'If W_Sw = True Then
        '    '2008/09/13とにかく、対象数は注文数に統一！（Ｑ＆Ａの№10より）
        '    W_Zaiko = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
        '    '08.12.04   仕入残の計算に変更！                                '08.12.04
        '    '               但し、＜０の場合は「０（ゼロ）」とする。        '08.12.04
        '    W_Zaiko = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - _
        '                CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))
        '    If W_Zaiko < 0 Then
        '        W_Zaiko = 0
        '    End If
        '    If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) <> "1" Then   '未完了
        '        '                            '発注数　＞　受入数
        '        'If CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) > CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) Then
        '        '
        '        '    W_Zaiko = W_Zaiko + CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
        '        'Else
        '        '                            '発注数　≦　受入数
        '        '    W_Zaiko = W_Zaiko + CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))
        '        'End If
        '    Else
        '                                    '完了：受入数
        '        'W_Zaiko = W_Zaiko + CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))
        '
        '        '2008/11/22「完了 ＆ 使用月＝当月 ＆ 受入日≧繰越日」は受入済み→月初在庫在庫！
        '        If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) = GW_TOUGETU Then
        '            'If Trim(StrConv(P_SHORDER_REC.KAN_DT, vbUnicode)) >= GW_SHIMEBI Then
        '            '2008.11.29                                     不等号が逆！            (*_*;
        '            If Trim(StrConv(P_SHORDER_REC.KAN_DT, vbUnicode)) < GW_SHIMEBI Then
        '                W_Zaiko = 0
        '            End If
        '        End If
        '        '2008.12.02 仕入完了は日付判定無しで、引当可能在庫としない！
        '        W_Zaiko = 0
        '
        '    End If
        'End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '       2008.12.04上記ブロックを下記に変更！
        '
        '               仕入残＝発注数から基準月までの受入数を減算。
        '
        W_Zaiko = 0
        If W_Sw = True Then
            W_Zaiko = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
            
            Call UniCode_Conv(K0_P_SHUKEIRE.ORDER_NO, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
            Call UniCode_Conv(K0_P_SHUKEIRE.SEQNO, "")
            com = BtOpGetGreaterEqual
            Do
                Do
                    sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, com, "P_SHUKEIRE")
                            Exit Function
                    End Select
                Loop
                If sts <> BtNoErr Then Exit Do
                
                '   受入データの計上月を判定に加味が正解？
                
                If Trim(StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode)) <> _
                        Trim(StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)) Then Exit Do
                    
                
                If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <= GW_TOUGETU Then
                    W_Zaiko = W_Zaiko - CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                End If
                
                
                If W_Zaiko <= 0 Then Exit Do
                
                com = BtOpGetNext
            Loop
            
        End If
                    '   ここまで！
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        If W_Zaiko < 0 Then
            W_Zaiko = 0
        End If
        
        If W_Zaiko <> 0 Then
        
            Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)    '子　事業部
            Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)     '子　国内外
            Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, GW_HINGAI_KO)    '子品番
            
            If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) = "" Then
                Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "g")                  'io区分
            Else
                Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "f")                  'io区分
            End If
            
            W_STR = StrConv(P_SHORDER_REC.USE_YM, vbUnicode)
            Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, W_STR)               '使用月
                
            W_STR = Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode))
            Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, W_STR)            '注文日
                
            W_STR = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)
            Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, W_STR)             '注文№
                
            sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                        
                    Call ODR_TEMP2_CLR
                    com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                    Exit Function
            End Select
                
                '子　事業部
            Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, GW_JIGYOBU_KO)
                '子　国内外
            Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, GW_NAIGAI_KO)
                '子品番
            Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, GW_HINGAI_KO)
            
            
                                                '2008/09 回答納期の有無で区分が異なる！
                'io区分         2008.05.02
            If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) = "" Then
                Call UniCode_Conv(ODR_TP2_R.IO_KB, "g")                  'io区分
            Else
                Call UniCode_Conv(ODR_TP2_R.IO_KB, "f")                  'io区分
            End If
                            
                
            W_STR = StrConv(P_SHORDER_REC.USE_YM, vbUnicode)
            Call UniCode_Conv(ODR_TP2_R.USE_YM, W_STR)                  '使用月
                    
If Trim(GW_HINGAI_KO) = "C029" Then
    Debug.Print
End If
                    
            W_STR = StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)
                '納期
            Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, W_STR)
                '注文№
            Call UniCode_Conv(ODR_TP2_R.ORDER_NO, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
                
                
            W_QTY = W_Zaiko + CDbl(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode))
                                
                                
            W_STR = CStr(W_QTY)
                
            Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)
            Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)
                
            Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
            Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
            
            If com <> BtOpUpdate Then
                Do
                    sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, com, "ODR_TEMP2")
                            Exit Function
                    End Select
                Loop
            Else
                '   これは、発生する？　(･･;)
                W_QTY = CDbl(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode))
                
            End If
            
            
        End If
    
    
        W_Zaiko = 0
        com = BtOpGetNext
    Loop
    
    
    SET_ODR_ZAN = False
    
End Function

Function SET_UKEIRE() As Integer

        '       受入実績情報　検索＆出力        io区分＝ｃ

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_Sw        As Integer

Dim W_STR       As String
Dim W_Zaiko     As Double

Dim W_Moto      As Double

    SET_UKEIRE = True
    
    W_Zaiko = 0
    
    Call UniCode_Conv(K1_P_SHUKEIRE.KEIJYO_YM, GW_TOUGETU)
    Call UniCode_Conv(K1_P_SHUKEIRE.ORDER_CODE, "")
    Call UniCode_Conv(K1_P_SHUKEIRE.UKEIRE_DT, "")
    
    com = BtOpGetGreaterEqual
    Do
        yn = 0
        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K1_P_SHUKEIRE, Len(K1_P_SHUKEIRE), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
            Case Else
                Call File_Error(sts, com, "P_SHORDER")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode)) = "05975" Then
            sts = 0
        End If
        
                
                        '   計上年月　≠　当月　→　終了
        If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <> GW_TOUGETU Then Exit Do
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        'W_STR = Trim(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode))
        '
        ''2008.12.02
        ''                   受入年月　＝ 　繰越年月？
        'If Left(W_STR, 6) <> Left(GW_SHIMEBI, 6) Then
        '&'   W_STR = ""
        'End If
        '
        '                    '   受入日　<=　繰越日？
        'If W_STR <= GW_SHIMEBI Then
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '                                   2008/12/06  上記ブロックの受入日付判定不要！とした。
        
                        '   発注情報確認
            Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
            
            sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "P_SHORDER")
                    Exit Function
            End Select
            
            If sts = BtNoErr Then
                W_Sw = True
                If Trim(StrConv(P_SHORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then
                    W_Sw = False
                End If
                
                If Trim(StrConv(P_SHORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then
                    W_Sw = False
                End If
                                'キャンセル？
                If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = "1" Then
                    W_Sw = False
                End If
                
                If Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) = "C215" Then
                    
                    W_STR = StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode)
                    sts = 0
                End If
                            '事業部、国内外が一致　→　展開データ中の品目の有無確認
                If W_Sw = True Then
                    Call UniCode_Conv(K2_ODR_TEMP1.KO_JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K2_ODR_TEMP1.KO_NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K2_ODR_TEMP1.KO_HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K2_ODR_TEMP1.SHIMUKE, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.JGYOBU, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.NAIGAI, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.HIN_GAI, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.ORDER_NO, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.INS_NO, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.BUN_NO, "")
        
                    sts = BTRV(BtOpGetGreaterEqual, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), _
                                                K2_ODR_TEMP1, Len(K2_ODR_TEMP1), 2)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetGreaterEqual, "ODR_TEMP1")
                            'Exit Function
                    End Select
                    If sts <> BtNoErr Then W_Sw = False
                    
                    If Trim(StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then
                        W_Sw = False
                    End If
                    
                    If Trim(StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then
                        W_Sw = False
                    End If
                    
                    If Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)) <> _
                                    Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) Then
                        W_Sw = False
                    End If
                    
        
                    '>>>>>>>>>>>>>>>>>>>>   受入実績（受入数）の累計
                    If W_Sw Then
                        W_Zaiko = CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                        
                        If Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) = "C215" Then
                            W_Zaiko = W_Zaiko * 1
                        End If
                            
                        Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)    '子　事業部
                        Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)     '子　国内外
                        Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))    '子品番
                        Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "c")                  'io区分
                                        
                        Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, GW_TOUGETU)               '使用月
                            
                        Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")            '注文日
                            
                        Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")             '注文№
                            
                        sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, _
                                            Len(K0_ODR_TEMP2), 0)
                        Select Case sts
                            Case BtNoErr
                                com = BtOpUpdate
                            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                                    
                                Call ODR_TEMP2_CLR
                                com = BtOpInsert
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                                Exit Function
                        End Select
                            
                            '子　事業部
                        Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, GW_JIGYOBU_KO)
                            '子　国内外
                        Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, GW_NAIGAI_KO)
                            '子品番
                        Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                            'io区分         2008.05.02
                        Call UniCode_Conv(ODR_TP2_R.IO_KB, "c")
                            '使用月
                        Call UniCode_Conv(ODR_TP2_R.USE_YM, GW_TOUGETU)
                            '納期
                        Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, "")
                            '注文№
                        Call UniCode_Conv(ODR_TP2_R.ORDER_NO, "")
                           
                        W_Moto = CDbl(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode))
                        W_STR = CStr(W_Zaiko + W_Moto)
                            
                        Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)
                        Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)
                            
                        Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
                        Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
                            
                        Do
                            sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                                    Sleep (500)
                                Case Else
                                    Call File_Error(sts, com, "ODR_TEMP2")
                                    Exit Function
                            End Select
                        Loop
            
                    End If
                    
                End If
                
                
            End If
            
        'End If
        
        
        com = BtOpGetNext
    Loop
    

    SET_UKEIRE = False

End Function

Function SET_ZAITEI() As Integer

        '       在訂（±）情報　検索＆出力        io区分＝ｄ
        
        '       移動歴より検索
        
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_Sw        As Integer

Dim W_STR       As String
Dim W_Zaiko     As Double
Dim W_QTY       As Double

Dim W_Date      As String       '当月の１日
Dim W_Today     As String       '本日
Dim wkYYMMDD     As String

Dim W_ZENGETU   As String


Dim W_Kan_DT    As String

Dim X_i         As Integer

    SET_ZAITEI = True
    
    W_Zaiko = 0
    
    
    W_Today = Format(Date, "yyyymmdd")                  '本日（PC-DATE)
    
    Call UniCode_Conv(K1_IDO.JGYOBU, GW_JIGYOBU_KO)
    Call UniCode_Conv(K1_IDO.NAIGAI, GW_NAIGAI_KO)
    Call UniCode_Conv(K1_IDO.HIN_GAI, GW_HINGAI_KO)
    Call UniCode_Conv(K1_IDO.JITU_DT, GW_SHIMEBI)       '対象：繰越日以降！
    Call UniCode_Conv(K1_IDO.JITU_TM, "")
    com = BtOpGetGreaterEqual
    
    Do
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
            Case Else
                Call File_Error(sts, com, "IDO")
                Exit Function
        End Select
        
        
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(IDOREC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then Exit Do
        If Trim(StrConv(IDOREC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then Exit Do
        If Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI_KO) Then Exit Do
        
        
        If Trim(StrConv(IDOREC.JITU_DT, vbUnicode)) > W_Today Then Exit Do
        
        
    'SUMI_JITU_QTY(0 To 7)               As Byte     '実績数量(商品化済み)
    'MI_JITU_QTY(0 To 7)                 As Byte     '実績数量(未商品)
        
        'W_QTY = 0
        'Select Case StrConv(IDOREC.RIRK_ID, vbUnicode)
        '    Case GW_PURA
        '        If IsNumeric(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) Then
        '            W_QTY = CDbl(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
        '        End If
        '        If IsNumeric(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) Then
        '            W_QTY = W_QTY + CDbl(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
        '        End If
        '
        '    Case GW_MAINA
        '        If IsNumeric(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) Then
        '            W_QTY = CDbl(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
        '        End If
        '        If IsNumeric(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) Then
        '            W_QTY = W_QTY + CDbl(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
        '        End If
        '
        '        W_QTY = W_QTY * -1
        '
        '    Case Else
        '        W_QTY = 0
        '
        'End Select
        '               2009/03/04  GW_PURA,GW_MAINAをテーブルにした！
        W_QTY = 0
        For X_i = 0 To UBound(GW_PURA)
            If GW_PURA(X_i) <> "" Then
                If StrConv(IDOREC.RIRK_ID, vbUnicode) = GW_PURA(X_i) Then
                    If IsNumeric(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) Then
                        W_QTY = CDbl(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
                    End If
                    If IsNumeric(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) Then
                        W_QTY = W_QTY + CDbl(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
                    End If
                End If
            End If
        Next X_i
        
        For X_i = 0 To UBound(GW_MAINA)
            If GW_MAINA(X_i) <> "" Then
                If StrConv(IDOREC.RIRK_ID, vbUnicode) = GW_MAINA(X_i) Then
                    If IsNumeric(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) Then
                        W_QTY = W_QTY - CDbl(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
                    End If
                    If IsNumeric(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) Then
                        W_QTY = W_QTY - CDbl(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
                    End If
                End If
            End If
        Next X_i
        
        
        W_Zaiko = W_Zaiko + W_QTY
        
        com = BtOpGetNext
        
    Loop
    
    If W_Zaiko <> 0 Then
    
        
            Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)    '子　事業部
            Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)     '子　国内外
            Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, GW_HINGAI_KO)    '子品番
            Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "d")                  'io区分
                            
            Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, GW_TOUGETU)               '使用月
                
            Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")            '注文日
                
            Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")             '注文№
                
            sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                        
                    Call ODR_TEMP2_CLR
                    com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                    Exit Function
            End Select
                
                '子　事業部
            Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, GW_JIGYOBU_KO)
                '子　国内外
            Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, GW_NAIGAI_KO)
                '子品番
            Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, GW_HINGAI_KO)
                'io区分         2008.05.02
            Call UniCode_Conv(ODR_TP2_R.IO_KB, "d")
                '使用月
            Call UniCode_Conv(ODR_TP2_R.USE_YM, GW_TOUGETU)
                '納期
            Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, "")
                '注文№
            Call UniCode_Conv(ODR_TP2_R.ORDER_NO, "")
                
            W_STR = CStr(W_Zaiko)
                
            Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)
            Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)
                
            Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
            Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
            If com <> BtOpUpdate Then
                Do
                    sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, com, "ODR_TEMP2")
                            Exit Function
                    End Select
                Loop
            End If
            
    End If
    SET_ZAITEI = False

End Function


Function SET_H_SEIHIN() As Integer
                    '半製品情報の集約（キーで集計する） 'io区分＝ｅ

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_Sw        As Integer

Dim W_STR       As String
Dim W_Zaiko     As Double
Dim W_QTY       As Double


    SET_H_SEIHIN = True
    
    
    Call ODR_TEMP2_CLR
    
    W_Zaiko = 0
        
    Call UniCode_Conv(K1_ODR_HANSEIHIN.KO_JGYOBU, GW_JIGYOBU_KO)
    Call UniCode_Conv(K1_ODR_HANSEIHIN.KO_NAIGAI, GW_NAIGAI_KO)
    Call UniCode_Conv(K1_ODR_HANSEIHIN.KO_HIN_GAI, GW_HINGAI_KO)
    
    If Trim(GW_HINGAI_KO) = "AD-HESB66AZ" Then
        W_Zaiko = 0
    End If
    
    com = BtOpGetGreaterEqual
    
    Do
                                            
        sts = BTRV(com, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_K_REC, Len(ODR_HANSEIHIN_K_REC), _
                            K1_ODR_HANSEIHIN, Len(K1_ODR_HANSEIHIN), 1)
                            
                            
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
            Case Else
                Call File_Error(sts, com, "ODR_HANSEIHIN")
                Exit Do
        End Select
                
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_JGYOBU, vbUnicode)) > Trim(GW_JIGYOBU_KO) Then Exit Do
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_NAIGAI, vbUnicode)) > Trim(GW_NAIGAI_KO) Then Exit Do
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_HIN_GAI, vbUnicode)) > Trim(GW_HINGAI_KO) Then Exit Do
        
        W_Sw = True
        
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then
            W_Sw = False
        End If
        
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then
            W_Sw = False
        End If
        
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI_KO) Then
            W_Sw = False
        End If
        
        
                    '2008/11/11 日付を判定しない！
''''        If Trim(StrConv(ODR_HANSEIHIN_K_REC.USE_YM, vbUnicode)) <> Trim(GW_TOUGETU) Then
''''            W_Sw = False
''''        End If
        
        
        If StrConv(ODR_HANSEIHIN_K_REC.SEQNO, vbUnicode) = "000" Then
            W_Sw = False            '親レコード
        Else
            If StrConv(ODR_HANSEIHIN_K_REC.ZAITEI_F, vbUnicode) = "1" Then
                W_Sw = False        '実在庫登録済み
            End If
        End If
        
                
'        If CDbl(StrConv(ODR_HANSEIHIN_K_REC.USE_QTY, vbUnicode)) <= 0 Then
'            W_Sw = False            '半製品の戻し
'        End If
        
        W_QTY = 0
                                            '発注数     9(5)v9(2)
        If W_Sw = True Then
            If Trim(StrConv(ODR_HANSEIHIN_K_REC.USE_QTY, vbUnicode)) <> "" Then
                If IsNumeric(Trim(StrConv(ODR_HANSEIHIN_K_REC.USE_QTY, vbUnicode))) Then
                    W_QTY = CDbl(StrConv(ODR_HANSEIHIN_K_REC.USE_QTY, vbUnicode))
                End If
            End If
        End If
        W_Zaiko = W_Zaiko + W_QTY
            
        com = BtOpGetNext
    Loop
               
    If W_Zaiko <> 0 Then
                      
        Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)   '子　事業部
        Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)     '子　国内外
        Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, GW_HINGAI_KO)    '子品番
        Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "e")                          'io区分 2008.05.02
        Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, GW_TOUGETU)              '使用月
            
        Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")           '注文日
        Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")            '注文№
            
            
            
        sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    'MsgBox "指定された工程がありません。"
                    
                Call ODR_TEMP2_CLR
                com = BtOpInsert
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                Exit Function
        End Select
            
            '子　事業部
        Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, GW_JIGYOBU_KO)   '子　事業部
            '子　国内外
        Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, GW_NAIGAI_KO)     '子　国内外
            '子品番
        Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, GW_HINGAI_KO)    '子品番
            'io区分
        Call UniCode_Conv(ODR_TP2_R.IO_KB, "e")
            '使用月
        Call UniCode_Conv(ODR_TP2_R.USE_YM, GW_TOUGETU)
            
            '2008/05/31 納期   無視！に変更。
        Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, "")
            
            '2008/05/31 注文№は無視！に変更。
        Call UniCode_Conv(ODR_TP2_R.ORDER_NO, "")
        
        W_Zaiko = W_Zaiko + CDbl(Trim(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode)))
        W_STR = CStr(W_Zaiko)
    
        Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)
        Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)
            
        Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
        Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
            
        Do
            sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP2")
                    Exit Function
            End Select
        Loop
            
    End If
        
    SET_H_SEIHIN = False


End Function




Function ZAN_CALC() As Integer

        '       在庫、所要量＆発注情報が表示されたので、日付順に差し引き残数を設定。

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String

Dim W_QTY       As Double
Dim W_Date      As String
Dim W_NOW       As String

    ZAN_CALC = True
    
    W_NOW = GW_TOUGETU & "01"
    com = BtOpGetFirst
    Do
        sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K3_ODR_TEMP1, Len(K3_ODR_TEMP1), 3)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                'Beep
                'MsgBox "指定された工程がありません。"
            Case Else
                Call File_Error(sts, com, "ODR_TEMP1")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        
                        '↓Trim 2008/07/02追加
        GW_JIGYOBU_KO = Trim(StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
        GW_NAIGAI_KO = Trim(StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
        GW_HINGAI_KO = Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
        
        If StrConv(ODR_TP1_R.KAN_KB, vbUnicode) = "0" Then
            W_STR = StrConv(ODR_TP1_R.HIN_GAI, vbUnicode)
        End If
        
        W_STR = StrConv(ODR_TP1_R.HIN_GAI, vbUnicode) & StrConv(ODR_TP1_R.ORDER_NO, vbUnicode)
        
        W_QTY = CDbl(StrConv(ODR_TP1_R.REQ_QTY, vbUnicode))     '対象数：所要数
        
        '       08.12.12 下記に変更
        W_QTY = CDbl(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode))     '対象数：展開数
        
        
        
        W_Date = StrConv(ODR_TP1_R.CYUMON_DT, vbUnicode)
        
        'If W_Date >= W_NOW Then                         '「必要日≧当月」は、引当処理する。
        '2008/09/10         ↑この判定は無意味！
        
            If W_QTY = 0 Then
                W_Date = ""                             '所要数＝０　→　完成品の子部品
            Else
                If GW_HINGAI_KO = "B016" Then
                    W_QTY = W_QTY * 1
                End If
                
                
                                                    '2010/03/04 マイナスは引当しない！ ：バグ！　(*_*;
                If W_QTY > 0 Then
                    If Zaiko_Hikiate(W_Date, W_QTY) Then
                        MsgBox "在庫引当処理エラー！", vbExclamation
                        Exit Function
                    End If
                End If
                
            End If
            
            
            '   08.12.12変更：未完の時に引当結果の日付を設定する。
            If StrConv(ODR_TP1_R.KAN_KB, vbUnicode) = "1" Then
                Call UniCode_Conv(ODR_TP1_R.OK_DT, W_Date)
            Else
                W_STR = StrConv(ODR_TP1_R.KAN_KB, vbUnicode)
            End If
            
            
            W_STR = CStr(W_QTY)
            
            Call UniCode_Conv(ODR_TP1_R.FUSOKU_QTY, W_STR)
            
            
            Call UniCode_Conv(ODR_TP1_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
            Call UniCode_Conv(ODR_TP1_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
            
            Do
                sts = BTRV(BtOpUpdate, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K3_ODR_TEMP1, Len(K3_ODR_TEMP1), 3)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        Sleep (500)
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "ODR_TEMP1")
                        Exit Function
                End Select
            Loop
            
        'End If
        
        com = BtOpGetNext
    Loop
    
    
    
    ZAN_CALC = False

End Function
Function SET_O_MAINA(HIN_GAI As String) As Integer
'
'                                 '       親注文のマイナスデータを在庫見做しで加算する。
'           構成展開　→　在庫数加算！      io区分＝ｂ
'

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_SeqKey    As String
Dim W_Ko_HinCD  As String

Dim Z_QTY       As Double
Dim W_QTY       As Double
Dim W_STR       As String
Dim W_Date      As String


Dim W_YOKU_YM   As String


    SET_O_MAINA = True
    
                                                                    
    'Z_QTY = CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)) * -1       '親の注文数
    Z_QTY = Abs(CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)))       '親の注文数
    '
    '指定された親品番の構成子部品の全てに関して、展開情報を設定する！
    '
    W_SeqKey = ""
    
    '   最初に「親レコード」を取得。
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, GW_SIMUKE)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, HIN_GAI)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "0")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    com = BtOpGetGreaterEqual
    sts = BTRV(com, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                'Beep
                'MsgBox "指定された工程がありません。"
        Case Else
            Call File_Error(sts, com, "P_COMPO")
            Exit Function
    End Select
    If sts <> BtNoErr Then
        MsgBox "構成情報　未登録！ <" & HIN_GAI & ">", vbExclamation
        Exit Function
    End If
    
    '   ここから「子部品レコード」を読みながら展開Ｆを出力する。
    com = BtOpGetNext
    Do
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                'Beep
                'MsgBox "指定された工程がありません。"
            Case Else
                Call File_Error(sts, com, "P_COMPO")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(HIN_GAI) Then Exit Do
        'If Trim(StrConv(P_COMPO_O_REC.DATA_KBN, vbUnicode)) <> "0" Then Exit Do
        
        W_SeqKey = StrConv(P_COMPO_K_REC.SEQNO, vbUnicode)
        
        If CInt(W_SeqKey) <> 0 Then                 '構成部品レコード？
            Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))    '子　事業部
            Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))    '子　国内外
            
            
            '   2008/10/06  2008.04.04の展開処理の修正と同一にした！
            Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, SHIZAI)    '子　事業部
            Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, NAIGAI_NAI)    '子　国内外
            
            Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))  '子　品番
            
            Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "b")             'io区分
            
                                               '使用月
            Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, StrConv(ODR_ORDER_REC.USE_YM, vbUnicode))
            
            Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")           '注文日
            Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")            '注文№
            
            sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    'MsgBox "指定された工程がありません。"
                    
                    Call ODR_TEMP2_CLR
                    '子　事業部
                    Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                    '子　国内外
                    Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    
                    
                    '   2008/10/06  2008.04.04の展開処理の修正と同一にした！
                    '子　事業部
                    Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, SHIZAI)
                    '子　国内外
                    Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, NAIGAI_NAI)

                    '子品番
                    Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    'io区分
                    Call UniCode_Conv(ODR_TP2_R.IO_KB, "b")
                    '使用月
                    Call UniCode_Conv(ODR_TP2_R.USE_YM, StrConv(ODR_ORDER_REC.USE_YM, vbUnicode))
                    
                    com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                    Exit Function
            End Select
                                    
            If Trim(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = "" Or _
                Not IsNumeric(Trim(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))) Then
                    'このような異常データの場合は「１」とみなす。
                W_QTY = Z_QTY * 1
            Else
                W_QTY = Z_QTY * CDbl(Trim(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)))
            End If
            W_QTY = W_QTY + CDbl(Trim(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode)))
            
            
            W_STR = CStr(W_QTY)
            
            If Trim(W_STR) = "" Then W_STR = "0"
            Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)     '在庫数
            Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)    '元々の在庫数
            
            Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
            Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
            Do
                sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        Sleep (500)
                    Case Else
                        Call File_Error(sts, com, "ODR_TEMP2")
                        Exit Function
                End Select
            Loop
            
        End If
        
        com = BtOpGetNext
    Loop
    
    
    SET_O_MAINA = False
    
End Function

Function Zaiko_Hikiate(OK_DT As String, W_QTY As Double) As Integer

        '       在庫、所要量＆発注情報が表示されたので、日付順に差し引き残数を設定。

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String

Dim W_Key       As String
Dim W_Date      As String

Dim W_K_io      As String
Dim W_K_DT      As String
Dim W_K_No      As String

Dim W_IN        As Double


If Trim(GW_HINGAI_KO) = "B016" Then
    Debug.Print
End If


    Zaiko_Hikiate = True
    
    W_Date = OK_DT
    
    
    OK_DT = ""
    W_K_io = ""
    W_K_DT = ""
    W_K_No = ""
    com = BtOpGetGreater
    Do
    
        Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)    '子　事業部
        Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)     '子　国内外
        Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, GW_HINGAI_KO)    '子品番
        Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, W_K_io)               'io区分
        Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, W_K_DT)        '注文日
        Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, W_K_No)            '注文№
        
        'com = BtOpGetGreater
        sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                'MsgBox "指定された工程がありません。"
            Case Else
                Call File_Error(sts, com, "ODR_TEMP2")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        
        
        If Trim(StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then Exit Do
        If Trim(StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then Exit Do
        If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI_KO) Then Exit Do
        
        
        
        
'''2008.04.10        If StrConv(ODR_TP2_R.CYUMON_DT, vbUnicode) > W_DATE Then Exit Do
        
        
If Trim(GW_HINGAI_KO) = "B123" Then
    If StrConv(ODR_TP1_R.USE_YM, vbUnicode) = "200812" Then
    Debug.Print
    End If
End If
        
        
        W_K_io = StrConv(ODR_TP2_R.IO_KB, vbUnicode)
        W_K_DT = StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode)
        W_K_No = StrConv(ODR_TP2_R.ORDER_NO, vbUnicode)
        
        W_IN = CDbl(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode))
        
        
        '           回答納期＝空白　→　対象外！？
        If StrConv(ODR_TP2_R.IO_KB, vbUnicode) = "g" Then
                
            W_IN = 0
        
        End If
        
        
        '2008.12.02     仕入残の内、使用月の異なる仕入残は引当対象外とする！？
        'If StrConv(ODR_TP2_R.IO_KB, vbUnicode) = "f" Then
        '    If StrConv(ODR_TP2_R.USE_YM, vbUnicode) <> "" Then
        '        If StrConv(ODR_TP2_R.USE_YM, vbUnicode) <> StrConv(ODR_TP1_R.USE_YM, vbUnicode) Then
        '            W_IN = 0
        '        End If
        '    End If
        'End If
        
        '2009.07.13
        '           上記の「仕入残異なる･･･」はミス！
        '               「異なる」ではなく、「使用月＜仕入残の使用月」じゃないと、前月残が使用されない！？
        If StrConv(ODR_TP2_R.IO_KB, vbUnicode) = "f" Then
            If StrConv(ODR_TP2_R.USE_YM, vbUnicode) <> "" Then
                If StrConv(ODR_TP2_R.USE_YM, vbUnicode) > StrConv(ODR_TP1_R.USE_YM, vbUnicode) Then
                    W_IN = 0
                End If
            End If
        End If
        
        
        
        If W_IN > 0 Then
            
            If W_QTY <= W_IN Then
                W_IN = W_IN - W_QTY
                W_QTY = 0
                OK_DT = Trim(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode))
                
                If Trim(OK_DT) = "" Then
                    OK_DT = Format(Date, "yyyymmdd")        '在庫データで可能！
                End If
                
            Else
                W_QTY = W_QTY - W_IN
                W_IN = 0
                
            End If
            
            
            W_STR = CStr(W_IN)
            
            Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)
            
            Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
            Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
    
            Do
                sts = BTRV(BtOpUpdate, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        Sleep (500)
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "ODR_TEMP2")
                        Exit Do
                End Select
            Loop
        
        End If
        
        
        If W_QTY <= 0 Then              '08/11/14 「＜」を「≦」に変更！　(*_*;
            Exit Do
        End If
        
        
        com = BtOpGetNext '+ BtSNoWait
    Loop
    
    
    Zaiko_Hikiate = False

End Function

Function OK_DT_SRCH(OK_DT As String) As Integer
'
'           組立可能日　検索
'

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    OK_DT_SRCH = True
    OK_DT = ""
        
If "TEST-1" = Trim(Key_HinGai) Then
    Debug.Print
End If
        
        '
    '指定された親品番の構成子部品の全てに関して、在庫・発注情報を確認する！？　(･･;)
    '
    Call UniCode_Conv(K1_ODR_TEMP1.SHIMUKE, Key_SIMUKE)    '仕向け先
    Call UniCode_Conv(K1_ODR_TEMP1.JGYOBU, Key_JIGYOBU)         '事業部
    Call UniCode_Conv(K1_ODR_TEMP1.NAIGAI, Key_NAIGAI)          '国内外
    Call UniCode_Conv(K1_ODR_TEMP1.HIN_GAI, Key_HinGai)         '親品番
    Call UniCode_Conv(K1_ODR_TEMP1.ORDER_NO, Key_ORDER_NO)      '親品番　注文№
    Call UniCode_Conv(K1_ODR_TEMP1.INS_NO, Key_INS_NO)          '登録順
    Call UniCode_Conv(K1_ODR_TEMP1.BUN_NO, Key_BUN_NO)          '分納回数
    Call UniCode_Conv(K1_ODR_TEMP1.OK_DT, "")                   '取り揃う日
        
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K1_ODR_TEMP1, Len(K1_ODR_TEMP1), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                
                
            Case Else
                Call File_Error(sts, com, "ODR_TEMP1")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        If Trim(StrConv(ODR_TP1_R.SHIMUKE, vbUnicode)) <> Trim(Key_SIMUKE) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.JGYOBU, vbUnicode)) <> Trim(Key_JIGYOBU) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.NAIGAI, vbUnicode)) <> Trim(Key_NAIGAI) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.HIN_GAI, vbUnicode)) <> Trim(Key_HinGai) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.ORDER_NO, vbUnicode)) <> Trim(Key_ORDER_NO) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.INS_NO, vbUnicode)) <> Trim(Key_INS_NO) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.BUN_NO, vbUnicode)) <> Trim(Key_BUN_NO) Then Exit Do
        
        
        'OK_DT = Trim(StrConv(ODR_TP1_R.OK_DT, vbUnicode))
        
        'If Trim(OK_DT) = "" Then Exit Do            '在庫不足のデータ有り！！
        
        '2008/12.16
        '               下記に変更
        If Trim(StrConv(ODR_TP1_R.OK_DT, vbUnicode)) = "" Then
            OK_DT = ""
            Exit Do
        End If
        
        If OK_DT < Trim(StrConv(ODR_TP1_R.OK_DT, vbUnicode)) Then
            OK_DT = Trim(StrConv(ODR_TP1_R.OK_DT, vbUnicode))
        End If
        
        com = BtOpGetNext
    Loop
    
    OK_DT_SRCH = False


End Function


Function OUT_KENTO() As Integer
'
'           各使用月別、子部品ごとに、月初在庫数～必要数など各種項目を設定
'
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer
Dim W_STR       As String
Dim W_Dbl       As Double
Dim W_QTY       As Double
Dim W_ODR       As Double

Dim X_i         As Integer
Dim X_j         As Integer

Dim W_From      As String
Dim W_To        As String
Dim W_Key1      As String
Dim W_Key2      As String
Dim W_Key3      As String
Dim W_Key4      As String

Dim W_YYMM      As String


Dim LAST_ORDER_DT   As String       '2016.12.14
Dim LAST_ORDER_QTY  As String       '2016.12.14
Dim i               As Integer      '2016.12.14



    OUT_KENTO = True
    
        
    W_From = Left(GW_TOUGETU, 4) & "/" & Right(GW_TOUGETU, 2) & "/01" '基準の年月（yyyymm）
    W_To = ""
    
    
            '発注検討Ｆ Close →　占有Open → Close → KILL → 占有Open
    
    If ODR_KENTO_Open(BtOpenExec) Then
        MsgBox "処理を中断します。", vbExclamation
        Exit Function
    End If
    
    sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_KENTO")
        End If
    End If
    
    If ODR_KENTO_KILL Then
        Exit Function
    End If
    
    If ODR_KENTO_Open(BtOpenExec) Then
        MsgBox "処理を中断します。", vbExclamation
        Exit Function
    End If
    
    Call ODR_KENTO_CLR
    W_STR = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:nn:ss")
    Call UniCode_Conv(ODR_KNT_R.ITEM_NM, W_STR)
    
    sts = BTRV(BtOpInsert, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    Select Case sts
        Case BtNoErr
                    
        Case Else
            Call File_Error(sts, BtOpInsert, "ODR_KENTO")
            Exit Function
    End Select
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                '   月初在庫の出力
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_ZAIKO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        Call ODR_KENTO_CLR
        
        If Trim(StrConv(ODR_ZK_R.KO_HIN_GAI, vbUnicode)) = "B015" Then
            sts = BtNoErr
        End If
            
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ODR_ZK_R.KO_JGYOBU, vbUnicode))    '事業部
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ODR_ZK_R.KO_NAIGAI, vbUnicode))      '国内外
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ODR_ZK_R.KO_HIN_GAI, vbUnicode))     '子品番
                
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    '*********************************************************************
                            'TEST的に編集！ (^_^;)
                    '品名
                Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録")
                    '発注ロット
                W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).LOT), "0") & "1"
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, W_STR)
                    '仕入先
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                    
                    '仕入単価
                W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).TANKA), "0") & "1"
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, W_STR)
                
                sts = BtNoErr
                
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ITEM")
                'Exit Function
        End Select
    
        If sts = BtNoErr Then
                
            Call UniCode_Conv(ODR_KNT_R.ITEM_NM, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)))
                
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>> 最新注文日を使用する    2016.12.14
            
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode)) Then
                Call UniCode_Conv(ODR_KNT_R.LOT_QTY, StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))
            Else
                Call UniCode_Conv(ODR_KNT_R.LOT_QTY, "00000000.00")
            End If
                
                
                
If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "C200" Then
    Debug.Print
End If
                
                
            LAST_ORDER_DT = ""
            LAST_ORDER_QTY = ""


            For i = 0 To 2
                If StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode) > LAST_ORDER_DT Then
                    LAST_ORDER_DT = StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode)
                    LAST_ORDER_QTY = StrConv(ITEMREC.G_SHIIRE_TBL(i).LOT, vbUnicode)
                End If
            Next i

            If IsNumeric(LAST_ORDER_QTY) Then
                Call UniCode_Conv(ODR_KNT_R.LOT_QTY, LAST_ORDER_QTY)
            Else
                Call UniCode_Conv(ODR_KNT_R.LOT_QTY, "00000000.00")
            End If

'>>>>>>>>>>>>>>>>>>>>>>>>>> 最新注文日を使用する    2016.12.14
                
                
                
            Call UniCode_Conv(ODR_KNT_R.SECT, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                Call UniCode_Conv(ODR_KNT_R.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
            Else
                Call UniCode_Conv(ODR_KNT_R.TANKA, "00000000.00")
            End If
            
            '一括発注区分の設定（09.05.22)
            Call UniCode_Conv(ODR_KNT_R.IKKATU_MK, StrConv(ITEMREC.AVE_SYUKA, vbUnicode))
                        
        End If
        
        
        
        Call UniCode_Conv(ODR_KNT_R.KO_JGYOBU, StrConv(ODR_ZK_R.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(ODR_KNT_R.KO_NAIGAI, StrConv(ODR_ZK_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(ODR_KNT_R.KO_HIN_GAI, StrConv(ODR_ZK_R.KO_HIN_GAI, vbUnicode))
        

        For X_i = 0 To UBound(ODR_ZK_R.ALL_ZAI)
            W_To = Left(Format(DateAdd("m", X_i, W_From), "yyyymmdd"), 6)
            Call UniCode_Conv(ODR_KNT_R.USE_YM, W_To)
            
            W_STR = CStr(CDbl(Trim(StrConv(ODR_ZK_R.ALL_ZAI(X_i).Z_QTY, vbUnicode))))
            Call UniCode_Conv(ODR_KNT_R.ZAI_QTY, W_STR)
        
            sts = BTRV(BtOpInsert, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    
                Case Else
                    Call File_Error(sts, BtOpInsert, "ODR_KENTO")
                    Exit Do
            End Select
               
        Next X_i
        
        com = BtOpGetNext
    Loop
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                '   展開済みデータ（TP1）より展開数・所要数・使用数　出力
    
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K2_ODR_TEMP1, Len(K2_ODR_TEMP1), 2)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<中間所要量Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP1")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        Call UniCode_Conv(K0_ODR_KENTO.USE_YM, StrConv(ODR_TP1_R.USE_YM, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
        
        
        If Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)) = "B015" Then
            sts = BtNoErr
        End If
        
        
        Do
            sts = BTRV(BtOpGetEqual, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF
                    Call ODR_KENTO_CLR
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))    '事業部
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))      '国内外
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))     '子品番
                            
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                                '*********************************************************************
                                        'TEST的に編集！ (^_^;)
                                '品名
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録")
                                '発注ロット
                            W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).LOT), "0") & "1"
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, W_STR)
                                '仕入先
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                                
                                '仕入単価
                            W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).TANKA), "0") & "1"
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, W_STR)
                            
                            sts = BtNoErr
                            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ITEM")
                            'Exit Function
                    End Select
                
                    If sts = BtNoErr Then
                            
                        Call UniCode_Conv(ODR_KNT_R.ITEM_NM, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)))
                            
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode)) Then
                            Call UniCode_Conv(ODR_KNT_R.LOT_QTY, StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))
                        Else
                            Call UniCode_Conv(ODR_KNT_R.LOT_QTY, "00000000.00")
                        End If
                            
                        Call UniCode_Conv(ODR_KNT_R.SECT, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                            
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                            Call UniCode_Conv(ODR_KNT_R.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
                        Else
                            Call UniCode_Conv(ODR_KNT_R.TANKA, "00000000.00")
                        End If
                        
                    End If
                                
                    
                    Call UniCode_Conv(ODR_KNT_R.USE_YM, StrConv(ODR_TP1_R.USE_YM, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
                    com = BtOpInsert
                    sts = BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
                                            '展開数
        W_QTY = CDbl(Trim(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode)))
        W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.ALL_QTY, vbUnicode)))
        W_STR = CStr(W_Dbl)
        Call UniCode_Conv(ODR_KNT_R.ALL_QTY, W_STR)
        
                                            '使用数
        W_QTY = CDbl(Trim(StrConv(ODR_TP1_R.USE_QTY, vbUnicode)))
        
        
    
        
        W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.USE_QTY, vbUnicode)))
'debug2019



If Trim(StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode)) = "D550" Then
    Debug.Print StrConv(ODR_KNT_R.USE_YM, vbUnicode) & " " & W_Dbl & " " & CDbl(Trim(StrConv(ODR_KNT_R.USE_QTY, vbUnicode)))
End If
        
        W_STR = CStr(W_Dbl)
        Call UniCode_Conv(ODR_KNT_R.USE_QTY, W_STR)
    
                                            '必要数
        W_QTY = CDbl(Trim(StrConv(ODR_TP1_R.NED_QTY, vbUnicode)))
        W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.NED_QTY, vbUnicode)))
        W_STR = CStr(W_Dbl)
        Call UniCode_Conv(ODR_KNT_R.NED_QTY, W_STR)
    
    
        '08.11.27   注文数＜０対応
                                    '展開数
        If CDbl(Trim(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode))) < 0 Then
                                                '展開数･･･絶対値（＋）にする！
            W_QTY = Abs(CDbl(Trim(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode))))
            W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.MINASHI1, vbUnicode)))
            W_STR = CStr(W_Dbl)
            Call UniCode_Conv(ODR_KNT_R.MINASHI1, W_STR)
            
            
            
        End If
    
    
    
    
        Do
            sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
    
        
        com = BtOpGetNext
    Loop
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '   展開済みデータ（TP2）より半製品、在訂、仕入残など　出力
    
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K1_ODR_TEMP2, Len(K1_ODR_TEMP2), 1)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<TEMP2>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP2")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        
        If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) = "B015" Then
            sts = BtNoErr
        End If
        
        
        Call UniCode_Conv(K0_ODR_KENTO.USE_YM, StrConv(ODR_TP2_R.USE_YM, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))
    
        Do
            sts = BTRV(BtOpGetEqual, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF
                    Call ODR_KENTO_CLR
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))    '事業部
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))      '国内外
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))     '子品番
                            
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                                '*********************************************************************
                                        'TEST的に編集！ (^_^;)
                                '品名
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録")
                                '発注ロット
                            W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).LOT), "0") & "1"
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, W_STR)
                                '仕入先
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                                
                                '仕入単価
                            W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).TANKA), "0") & "1"
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, W_STR)
                            
                            sts = BtNoErr
                            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ITEM")
                            'Exit Function
                    End Select
                
                    If sts = BtNoErr Then
                            
                        Call UniCode_Conv(ODR_KNT_R.ITEM_NM, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)))
                            
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode)) Then
                            Call UniCode_Conv(ODR_KNT_R.LOT_QTY, StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))
                        Else
                            Call UniCode_Conv(ODR_KNT_R.LOT_QTY, "00000000.00")
                        End If
                            
                        Call UniCode_Conv(ODR_KNT_R.SECT, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                            
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                            Call UniCode_Conv(ODR_KNT_R.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
                        Else
                            Call UniCode_Conv(ODR_KNT_R.TANKA, "00000000.00")
                        End If
                        
                    End If
                    
                    Call UniCode_Conv(ODR_KNT_R.USE_YM, StrConv(ODR_TP2_R.USE_YM, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))
                    com = BtOpInsert
                    sts = BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        W_QTY = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
        
        If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) = "AD-KZ061X" Then
            W_QTY = W_QTY * 1
        End If
        
    'ZAI_QTY(0 To 10)            As Byte         '月初在庫数     9(8)v9(2)
    'MAI_QTY(0 To 10)            As Byte         '不足数         9(8)v9(2)
    'ODR_QTY(0 To 10)            As Byte         '注文数         9(8)v9(2)
    'SHI_QTY(0 To 10)            As Byte         '仕入残数       9(8)v9(2)
    'HANSEIHIN_QTY(0 To 10)      As Byte         '半製品数       9(8)v9(2)
    'ZAITEI_QTY(0 To 10)         As Byte         '在訂±数       9(8)v9(2)
    'KAITO(0 To 7)               As Byte         '回答納期
    'ZAN_CNT(0 To 2)             As Byte         '仕入残　件数
    
        Select Case StrConv(ODR_TP2_R.IO_KB, vbUnicode)
            
            Case "a"            '在庫
                
                'W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.ZAI_QTY, vbUnicode)))
                'W_Str = CStr(W_Dbl)
                'Call UniCode_Conv(ODR_KNT_R.ZAI_QTY, W_Str)
                
                W_YYMM = StrConv(ODR_TP2_R.USE_YM, vbUnicode)
                
                If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) = "K142" Then
                    W_STR = ""
                End If
                
                If W_YYMM = GW_TOUGETU Then
                    W_QTY = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
                    W_STR = CStr(W_QTY)
                    Call UniCode_Conv(ODR_KNT_R.ITEM_Z_QTY, W_STR)
                End If
            
            Case "b"            '親注文＜０
                
                '2008/10/06
                'W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.ZAI_QTY, vbUnicode)))
                'W_Str = CStr(W_Dbl)
                'Call UniCode_Conv(ODR_KNT_R.ZAI_QTY, W_Str)
                
            
            Case "c"            '仕入済み
            
                W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.MINASHI2, vbUnicode)))
                W_STR = CStr(W_Dbl)
                'Call UniCode_Conv(ODR_KNT_R.ZAI_QTY, W_Str)
                
                                '2008.12.02
                Call UniCode_Conv(ODR_KNT_R.MINASHI2, W_STR)
            
            Case "d"            '在訂
                
                W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.ZAITEI_QTY, vbUnicode)))
                W_STR = CStr(W_Dbl)
                Call UniCode_Conv(ODR_KNT_R.ZAITEI_QTY, W_STR)
            
            Case "e"            '半製品
            
                W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.HANSEIHIN_QTY, vbUnicode)))
                W_STR = CStr(W_Dbl)
                Call UniCode_Conv(ODR_KNT_R.HANSEIHIN_QTY, W_STR)
            
            Case "f"            '発注残
            
                W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY1, vbUnicode)))
                W_STR = CStr(W_Dbl)
                Call UniCode_Conv(ODR_KNT_R.SHI_QTY1, W_STR)
                
                
                'If Trim(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode)) <> "" Then
                If Trim(StrConv(ODR_KNT_R.KAITO, vbUnicode)) = "" Then
                    Call UniCode_Conv(ODR_KNT_R.KAITO, StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode))
                End If
                    
                If Trim(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode)) < StrConv(ODR_KNT_R.KAITO, vbUnicode) Then
                    Call UniCode_Conv(ODR_KNT_R.KAITO, StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode))
                End If
                    
                W_Dbl = CDbl(Trim(StrConv(ODR_KNT_R.ZAN_CNT, vbUnicode))) + 1
                W_STR = CStr(W_Dbl)
                Call UniCode_Conv(ODR_KNT_R.ZAN_CNT, W_STR)
                'End If
                
            Case "g"            '発注残（回答納期無し）
            
                W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY2, vbUnicode)))
                W_STR = CStr(W_Dbl)
                Call UniCode_Conv(ODR_KNT_R.SHI_QTY2, W_STR)
                
            Case Else
         
         
         
        End Select
            
        Do
            sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        
        
        com = BtOpGetNext
    Loop
        
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '不足数の計算
    
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<ODR_KENTO>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        W_Key1 = StrConv(ODR_KNT_R.USE_YM, vbUnicode)
        W_Key2 = StrConv(ODR_KNT_R.KO_JGYOBU, vbUnicode)
        W_Key3 = StrConv(ODR_KNT_R.KO_NAIGAI, vbUnicode)
        W_Key4 = StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode)
        
        If Trim(W_Key4) = "B533" Then
            sts = BtNoErr
        End If
        
        
        W_Dbl = CDbl(Trim(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode)))           '元々の不足数
        
        W_Dbl = W_Dbl - CDbl(Trim(StrConv(ODR_KNT_R.NED_QTY, vbUnicode)))   '－必要数
        
        W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.ZAI_QTY, vbUnicode)))   '＋在庫数
        W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.ODR_QTY, vbUnicode)))   '＋注文数
        W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY1, vbUnicode)))   '＋仕入残数
        
        
        '2008/12/10 使用数を減算！
        W_Dbl = W_Dbl - CDbl(Trim(StrConv(ODR_KNT_R.USE_QTY, vbUnicode)))   '－使用数
        
        
        '       半製品数は、月初在庫に加算されている！
        '       もう一度加算すると、ダブッて加算してしまう！
        'W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.HANSEIHIN_QTY, vbUnicode)))   '＋半製品数
        
        
        '       在訂±数は、月初在庫に加算されている！
        '       もう一度加算すると、ダブッて加算してしまう！
        'W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.ZAITEI_QTY, vbUnicode)))   '＋在訂±数
        
        
        W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY2, vbUnicode)))   '＋仕入残数（回答納期無し）
        
        W_STR = CStr(W_Dbl)
        Call UniCode_Conv(ODR_KNT_R.MAI_QTY, W_STR)         '不足数
        
        
        
        '注文数の計算
        If IsNumeric(Trim(StrConv(ODR_KNT_R.LOT_QTY, vbUnicode))) Then
            W_QTY = CDbl(Trim(StrConv(ODR_KNT_R.LOT_QTY, vbUnicode)))
            If W_QTY = 0 Then W_QTY = 1
        Else
            W_QTY = 1
        End If
        
        If W_Dbl < 0 Then
            W_Dbl = W_Dbl * -1
            W_ODR = 0
            Do
                If W_ODR >= W_Dbl Then Exit Do
                W_ODR = W_ODR + W_QTY
            Loop
            W_STR = CStr(W_ODR)
            Call UniCode_Conv(ODR_KNT_R.ODR_QTY, W_STR)
        End If
        
        
        W_Dbl = CDbl(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode))
        
        Do
            sts = BTRV(BtOpUpdate, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpUpdate, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        
    com = 0
    If com <> 0 Then
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  次月以降の不足数を計算
        Call UniCode_Conv(K1_ODR_KENTO.KO_JGYOBU, W_Key2)
        Call UniCode_Conv(K1_ODR_KENTO.KO_NAIGAI, W_Key3)
        Call UniCode_Conv(K1_ODR_KENTO.KO_HIN_GAI, W_Key4)
        Call UniCode_Conv(K1_ODR_KENTO.USE_YM, W_Key1)
        com = BtOpGetGreater
        Do
            Do
                sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K1_ODR_KENTO, Len(K1_ODR_KENTO), 1)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                        
                        
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        yn = MsgBox("他で使用中です！<ODR_KENTO>" & Chr(13) & Chr(10) & _
                                    "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                        If yn = vbNo Then Exit Do
                    Case Else
                        Call File_Error(sts, com, "ODR_KENTO")
                        Exit Do
                End Select
            Loop
            If sts <> BtNoErr Then Exit Do
            If StrConv(ODR_KNT_R.KO_JGYOBU, vbUnicode) <> W_Key2 Then Exit Do
            If StrConv(ODR_KNT_R.KO_NAIGAI, vbUnicode) <> W_Key3 Then Exit Do
            If StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode) <> W_Key4 Then Exit Do
            
            If StrConv(ODR_KNT_R.USE_YM, vbUnicode) > W_Key1 Then
                W_QTY = CDbl(Trim(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode))) + W_Dbl
                W_STR = CStr(W_QTY)
                Call UniCode_Conv(ODR_KNT_R.MAI_QTY, W_STR)
                
                
                
                
                Do
                    sts = BTRV(BtOpUpdate, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
                    Select Case sts
                        Case BtNoErr
                            
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "ODR_KENTO")
                            Exit Do
                    End Select
                Loop
            End If
            
            com = BtOpGetNext
        Loop
    End If
        
        Call UniCode_Conv(K0_ODR_KENTO.USE_YM, W_Key1)
        Call UniCode_Conv(K0_ODR_KENTO.KO_JGYOBU, W_Key2)
        Call UniCode_Conv(K0_ODR_KENTO.KO_NAIGAI, W_Key3)
        Call UniCode_Conv(K0_ODR_KENTO.KO_HIN_GAI, W_Key4)
        
        com = BtOpGetGreater
    Loop

    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '一括発注の場合、注文数＝０とする。     '09.05.22
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<ODR_KENTO>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        
        '一括       2009.05.22
        'If StrConv(ODR_KNT_R.IKKATU_MK, vbUnicode) = "1" Then
        '    Call UniCode_Conv(ODR_KNT_R.ODR_QTY, "0")
        'End If
        
        Do
            sts = BTRV(BtOpUpdate, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpUpdate, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        

        com = BtOpGetNext
    Loop

    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    

    OUT_KENTO = False
    
    sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_KENTO")
        End If
    End If
    
End Function
Function GESSYO_SET() As Integer
'
'           子部品ごとに、使用月単位の月初在庫設定
'
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer
Dim W_STR       As String
Dim W_Dbl       As Double
Dim W_QTY       As Double

Dim W_Key1      As String
Dim W_Key2      As String

Dim X_i         As Integer
Dim X_j         As Integer

Dim W_From      As String
Dim W_To        As String


    GESSYO_SET = True
                                
    W_Key1 = ""
    W_Key2 = ""
    
    Call ODR_ZAIKO_CLR
    
    W_From = Left(GW_TOUGETU, 4) & "/" & Right(GW_TOUGETU, 2) & "/01" '基準の年月（yyyymm）
    W_To = ""
    
                            '月初在庫Ｆ Close →　占有Open → Close → KILL → 占有Open
                   
    sts = BTRV(BtOpClose, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_ZAIKO")
        End If
    End If
    
    
    If ODR_ZAIKO_Open(BtOpenExec) Then
        MsgBox "処理を中断します。", vbExclamation
        Exit Function
    End If
    sts = BTRV(BtOpClose, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_ZAIKO")
        End If
    End If
    
    If ODR_ZAIKO_KILL Then
        Exit Function
    End If
    
    If ODR_ZAIKO_Open(BtOpenExec) Then
        MsgBox "処理を中断します。", vbExclamation
        Exit Function
    End If
    
    '2008/10/17
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ODR_ZAIKO_CLR
    Call UniCode_Conv(ODR_ZK_R.FILLER, W_From)
    
    sts = BTRV(BtOpInsert, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
    If sts <> BtNoErr Then
        MsgBox "月初在庫　基準月追加失敗！", vbExclamation
        sts = BTRV(BtOpClose, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "ODR_ZAIKO")
            End If
        End If
        If ODR_ZAIKO_Open(BtOpenNomal) Then
            MsgBox "処理を中断します。", vbExclamation
        End If

        Exit Function
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    
    com = BtOpGetFirst
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '   使用月＝空白の場合、基準年月を設定する。
    Do
        Do
            sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K1_ODR_TEMP2, Len(K1_ODR_TEMP2), 1)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<中間所要量Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP2")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(ODR_TP2_R.USE_YM, vbUnicode)) = "" Then
            Call UniCode_Conv(ODR_TP2_R.USE_YM, GW_TOUGETU)
            Do
                sts = BTRV(BtOpUpdate, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K1_ODR_TEMP2, Len(K1_ODR_TEMP2), 1)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        Sleep (500)
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "ODR_TEMP2")
                        Exit Do
                End Select
            Loop
            If sts <> BtNoErr Then Exit Do
        End If
        
        com = BtOpGetNext
    Loop
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
          
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '   先頭から読み、使用月単位の月初在庫を設定する。
                    '   使用月が等しいデータは、在庫数を加算する。
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K1_ODR_TEMP2, Len(K1_ODR_TEMP2), 1)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<中間所要量Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP2")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) = "B015" Then
            sts = BtNoErr
        End If
        If Trim(StrConv(ODR_TP2_R.IO_KB, vbUnicode)) = "b" Then
            sts = BtNoErr
        End If
        
        
        '2008/11/14
        Call UniCode_Conv(K0_ODR_ZK.KO_JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ODR_ZK.KO_NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ODR_ZK.KO_HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))
        Do
            sts = BTRV(BtOpGetEqual, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Call ODR_ZAIKO_CLR
                    Call UniCode_Conv(ODR_ZK_R.KO_JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ODR_ZK_R.KO_NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ODR_ZK_R.KO_HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))
                    com = BtOpInsert
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<ODR_ZAIKO>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Function
                Case Else
                    Call File_Error(sts, com, "ODR_ZAIKO")
                    Exit Function
            End Select
        Loop
        
        W_STR = StrConv(ODR_TP2_R.USE_YM, vbUnicode)
        W_To = Left(W_STR, 4) & "/" & Right(W_STR, 2) & "/01"
        X_i = DateDiff("m", W_From, W_To)
             
        '2008/12/01 下記判定を追加！
        If X_i >= UBound(ODR_ZK_R.ALL_ZAI) Then
            'MsgBox "[" & W_From & "] ～ [" & W_To & "]  使用月　期間設定異常！？", vbExclamation
        
        Else
                    
                    '   仕入残は、翌月の月初在庫に反映！！                  '2008/09/27
            If StrConv(ODR_TP2_R.IO_KB, vbUnicode) >= "f" Then
                X_i = X_i + 1
            End If
                    
                    
            W_Dbl = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
            W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_ZK_R.ALL_ZAI(X_i).Z_QTY, vbUnicode)))
                    
            W_STR = CStr(W_Dbl)
            Call Numeric_Check(EDIT_ONLY, UBound(ODR_ZK_R.ALL_ZAI(0).Z_QTY) + 1, 2, NEGA_ENA, ZSUP_DIS, _
                                COMA_DIS, CStr(W_Dbl), W_STR)
                    
            For X_j = X_i To UBound(ODR_ZK_R.ALL_ZAI)
                W_Dbl = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
                W_QTY = W_Dbl + CDbl(Trim(StrConv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, vbUnicode)))
                W_STR = CStr(W_QTY)
                Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, W_STR)            '在庫数     9(5)v9(2)
            Next X_j
        
        End If
        
        Do
            sts = BTRV(com, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_ZAIKO")
                    Exit Do
            End Select
        Loop
        
        '>>>>>>>>>>>>>  2008/11/14 下記をコメントに！
        '
        '
        '                                                    '回答納期の無い仕入は除外!08/09/19
        '                                                    '含む！！！     2008/09/27
        ''If StrConv(ODR_TP2_R.IO_KB, vbUnicode) <> "g" Then
        '
        '    W_Key1 = StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode)
        '    W_Key1 = W_Key1 & StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode)
        '    W_Key1 = W_Key1 & StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)
        '    W_Key1 = W_Key1 & StrConv(ODR_TP2_R.USE_YM, vbUnicode)
        '
        '    If W_Key1 = W_Key2 Then
        '
        '        W_Str = StrConv(ODR_TP2_R.USE_YM, vbUnicode)
        '        W_To = Left(W_Str, 4) & "/" & Right(W_Str, 2) & "/01"
        '        X_i = DateDiff("m", W_From, W_To)
        '
        '
        '        '   仕入残は、翌月の月初在庫に反映！！                  '2008/09/27
        '        If StrConv(ODR_TP2_R.IO_KB, vbUnicode) >= "f" Then
        '            X_i = X_i + 1
        '        End If
        '
        '
        '
        '        W_Dbl = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
        '        W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_ZK_R.ALL_ZAI(X_i).Z_QTY, vbUnicode)))
        '
        '        W_Str = CStr(W_Dbl)
        '        Call Numeric_Check(EDIT_ONLY, UBound(ODR_ZK_R.ALL_ZAI(0).Z_QTY) + 1, 2, NEGA_ENA, ZSUP_DIS, _
        '                    COMA_DIS, CStr(W_Dbl), W_Str)
        '
        '        For X_j = X_i To UBound(ODR_ZK_R.ALL_ZAI)
        '            W_Dbl = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
        '            W_QTY = W_Dbl + CDbl(Trim(StrConv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, vbUnicode)))
        '            W_Str = CStr(W_QTY)
        '            Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, W_Str)            '在庫数     9(5)v9(2)
        '        Next X_j
        '
        '    Else
        '        If W_Key2 <> "" Then
        '            Do
        '                sts = BTRV(BtOpInsert, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
        '                Select Case sts
        '                    Case BtNoErr
        '                        Exit Do
        '                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
        '                        Sleep (500)
        '                    Case Else
        '                        Call File_Error(sts, BtOpInsert, "ODR_ZAIKO")
        '                        Exit Do
        '                End Select
        '            Loop
        '            If sts <> BtNoErr Then
        '                MsgBox "月初在庫　追加失敗！", vbExclamation
        '                'Exit Do
        '            End If
        '
        '        End If
        '
        '        Call ODR_ZAIKO_CLR
        '        '子　事業部
        '        Call UniCode_Conv(ODR_ZK_R.KO_JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))
        '        '子　国内外
        '        Call UniCode_Conv(ODR_ZK_R.KO_NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))
        '        '子品番
        '        Call UniCode_Conv(ODR_ZK_R.KO_HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))
        '
        '        W_Dbl = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
        '
        '
        '        W_Str = StrConv(ODR_TP2_R.USE_YM, vbUnicode)
        '        W_To = Left(W_Str, 4) & "/" & Right(W_Str, 2) & "/01"
        '
        '        X_i = DateDiff("m", W_From, W_To)
        '
        '        W_Str = CStr(W_Dbl)
        '        Call Numeric_Check(EDIT_ONLY, UBound(ODR_ZK_R.ALL_ZAI(0).Z_QTY) + 1, 2, NEGA_ENA, ZSUP_DIS, _
        '                    COMA_DIS, CStr(W_Dbl), W_Str)
        '        W_Str = CStr(W_Dbl)
        '        For X_j = X_i To UBound(ODR_ZK_R.ALL_ZAI)
        '            Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, W_Str)            '在庫数     9(5)v9(2)
        '        Next X_j
        '
        '    End If
        '
        ''End If
        
        
        W_Key2 = W_Key1
        com = BtOpGetNext
        
    Loop
    
    If W_Key2 <> "" Then


        Do
            sts = BTRV(BtOpInsert, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpInsert, "ODR_ZAIKO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then
            MsgBox "月初在庫　追加失敗！", vbExclamation
            'Exit Do
        End If
            
    End If
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '   先頭から読み、使用月単位の必要数を計算し、
                    '   使用月の翌月在庫から減算して月初在庫を計算する。
    com = BtOpGetFirst
    
    Do
        Do
            sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K2_ODR_TEMP1, Len(K2_ODR_TEMP1), 2)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<中間所要量Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP1")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        Call UniCode_Conv(K0_ODR_ZK.KO_JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ODR_ZK.KO_NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ODR_ZK.KO_HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
        W_STR = StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)
        Do
            sts = BTRV(BtOpGetEqual, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF
                    Call ODR_ZAIKO_CLR
                    Call UniCode_Conv(ODR_ZK_R.KO_JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ODR_ZK_R.KO_NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ODR_ZK_R.KO_HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
                    com = BtOpInsert
                    sts = BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_ZAIKO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
                     
                     
        W_Dbl = CDbl(Trim(StrConv(ODR_TP1_R.NED_QTY, vbUnicode)))   '所要数
        If W_Dbl = 0 Then
            W_STR = StrConv(ODR_TP1_R.HIN_GAI, vbUnicode)
            W_Dbl = 0
        End If
        
        W_STR = StrConv(ODR_TP1_R.USE_YM, vbUnicode)
        W_To = Left(W_STR, 4) & "/" & Right(W_STR, 2) & "/01"
        X_i = DateDiff("m", W_From, W_To) + 1                   '翌月の月初在庫から減算
          
             
        '2008/12/01 下記判定を追加！
        If X_i >= UBound(ODR_ZK_R.ALL_ZAI) Then
            'MsgBox "[" & W_From & "] ～ [" & W_To & "]  使用月　期間設定異常！？", vbExclamation
        
        Else
                        
              
            For X_j = X_i To UBound(ODR_ZK_R.ALL_ZAI)
                W_Dbl = CDbl(Trim(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode)))   '展開数
                W_QTY = CDbl(StrConv(ODR_ZK_R.ALL_ZAI(X_j).Y_QTY, vbUnicode))
                W_QTY = W_QTY + W_Dbl
                
                W_STR = CStr(W_QTY)
                Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_j).Y_QTY, W_STR)       'Σ展開数
            
                
                W_Dbl = CDbl(Trim(StrConv(ODR_TP1_R.REQ_QTY, vbUnicode)))   '必要数
                
                W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_TP1_R.USE_QTY, vbUnicode)))   '使用数 2008/12/10
                
                W_QTY = CDbl(StrConv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, vbUnicode))
                W_QTY = W_QTY - W_Dbl                                       '在庫数　―　必要数
                
                W_STR = CStr(W_QTY)
                Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, W_STR)       '在庫数
            
            Next X_j
                           
        End If
        
        
        sts = BTRV(com, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
        Select Case sts
            Case BtNoErr
                
            Case Else
                Call File_Error(sts, com, "ODR_ZAIKO")
                Exit Do
        End Select
        
        com = BtOpGetNext
    Loop
    
    'End If
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                            '月初在庫Ｆ Close → 共用Open
    sts = BTRV(BtOpClose, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_ZAIKO")
        End If
    End If
    If ODR_ZAIKO_Open(BtOpenNomal) Then
        MsgBox "処理を中断します。", vbExclamation
        Exit Function
    End If
    
    GESSYO_SET = False
    
End Function

