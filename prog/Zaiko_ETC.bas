Attribute VB_Name = "Zaiko_ETC"

Option Explicit

'--------------------------------------------------------   在庫数集計時、除外する倉庫  2012.03.20
Public Zaiko_Syukei_Jyogai_Soko_No      As Variant
Public Zaiko_Syukei_Jyogai_Soko_No2     As Variant          '2014.11.07







'Public KASO_NYUKA_Soko      As String * 2   '仮想 入荷倉庫
'Public KASO_NYUKABA_Soko    As String * 2   '仮想 入荷場倉庫
'Public KASO_SYOHN_Soko      As String * 2   '仮想 商品化室
'Public KASO_NAI_Soko        As String * 2   '仮想 内職
'Public KASO_IDO_Soko        As String * 2   '仮想 移動
'Public KASO_FURIKAE_Soko    As String * 2   '仮想 国内外振替
'Public KASO_SYUKA_Soko      As String * 2   '仮想 出荷場（未使用）
'Public GLB_SUMI_YUKO_ZAIKO_QTY  As Long     '有効在庫(商品化済み)
'Public GLB_MI_YUKO_ZAIKO_QTY    As Long     '有効在庫(未商品)

Public Function Zaiko_Syukei_Proc(Sumi_Zaiko_Qty As Long, _
                                    Mi_Zaiko_Qty As Long, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    Optional Location As String = "        ", _
                                    Optional GOODS_OFF_SOKO_NO, _
                                    Optional GOODS_OFF_SOKO_NO_F As Integer = False, _
                                    Optional Jyogai_Soko_On = False, _
                                    Optional Jyogai_Soko_On2 = False, _
                                    Optional Bt_sts As Integer = BtNoErr, _
                                    Optional mesg_mode As Integer = 1) As Integer
'****************************************************
'*      在庫数集計
'*
'*  品番または品番＋棚番毎の在庫数を集計する。
'*
'*  引数 :  在庫数（商品化済み）
'*          在庫数（未商品）
'*          事業部（省略不可）
'*          国内外（省略不可）
'*          品番外部(省略不可)
'*          棚番(省略可 省略=空白)
'*          商品化済み除外倉庫(省略可)                          2011.12.16
'*          商品化済み除外倉庫有無(省略可　省略=除外倉庫なし）  2011.12.16
'*          除外倉庫有無                                        2012.03.20
'*          除外倉庫有無２                                      2014.11.07
'*          Btrieve 戻り値                                      2015.03.13
'*          エラーＢＯＸ表示有無 0:表示しない                   2015.03.13
'*
'*  戻り値: false    正常
'*          SYS_ERR  継続できない異常
'****************************************************
Dim sts     As Integer
Dim com     As Integer
Dim Soko_No As String * 2
Dim Retu    As String * 2
Dim Ren     As String * 2
Dim Dan     As String * 2

Dim GOODS_OFF_T() _
            As String * 2
Dim i       As Long


Dim Found_Flg   As Boolean                  '2012.03.20


    Zaiko_Syukei_Proc = SYS_ERR


    If Not GOODS_OFF_SOKO_NO_F Then
        ReDim GOODS_OFF_T(0 To 0)
        GOODS_OFF_T(0) = "**"
    Else
        ReDim GOODS_OFF_T(0 To UBound(GOODS_OFF_SOKO_NO))
        For i = 0 To UBound(GOODS_OFF_SOKO_NO)
            GOODS_OFF_T(i) = GOODS_OFF_SOKO_NO(i)
        Next i
    End If

    Sumi_Zaiko_Qty = 0
    Mi_Zaiko_Qty = 0

    com = BtOpGetGreater

    If Len(Trim(Location)) = 0 Then
                                '倉庫番号空白は棚番省略とみなす
        Call UniCode_Conv(K1_ZAIKO.JGYOBU, JGYOBU)
        Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K1_ZAIKO.HIN_GAI, HIN_GAI)
        Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
        Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
        Call UniCode_Conv(K1_ZAIKO.Retu, "")
        Call UniCode_Conv(K1_ZAIKO.Ren, "")
        Call UniCode_Conv(K1_ZAIKO.Dan, "")

        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
            Select Case sts
                Case BtNoErr
                    If JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Bt_sts = sts                                '2015.03.13
                    Call File_Error(sts, com, "在庫データ", mesg_mode)
                    Exit Function
            End Select

'--------------------------------------------------------   在庫数集計時、除外する倉庫  2012.03.20
            Found_Flg = False

            If Not Jyogai_Soko_On Then
            Else
                For i = 0 To UBound(Zaiko_Syukei_Jyogai_Soko_No)
                    If Zaiko_Syukei_Jyogai_Soko_No(i) = StrConv(ZAIKOREC.Soko_No, vbUnicode) Then
                        Found_Flg = True
                        Exit For
                    End If
                Next i
            End If
            
            
            '除外する倉庫２ 2014.11.07
            If Not Jyogai_Soko_On2 Then
            Else
                For i = 0 To UBound(Zaiko_Syukei_Jyogai_Soko_No2)
                    If Zaiko_Syukei_Jyogai_Soko_No2(i) = StrConv(ZAIKOREC.Soko_No, vbUnicode) Then
                        Found_Flg = True
                        Exit For
                    End If
                Next i
            End If
            
            
            
            If Not Found_Flg Then
            
'--------------------------------------------------------   在庫数集計時、除外する倉庫  2012.03.20
            
                Select Case StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
                    Case "0"
                        '---------------    商品化済み除外倉庫対応  2011.12.16
                        For i = 0 To UBound(GOODS_OFF_T)
                            If Trim(GOODS_OFF_T(i)) = Trim(StrConv(ZAIKOREC.Soko_No, vbUnicode)) Then
                                Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                Exit For
                            End If
                        Next i
                        If i > UBound(GOODS_OFF_T) Then
                            Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                        End If
                        '---------------    商品化済み除外倉庫対応  2011.12.16
                    Case "1"
                        Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                End Select

            End If

            com = BtOpGetNext

'            DoEvents
            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                DoEvents                                                    '2016.01.26
            End If                                                          '2016.01.26
        Loop

    Else

        Soko_No = Mid(Location, 1, 2)
        Retu = Mid(Location, 3, 2)
        Ren = Mid(Location, 5, 2)
        Dan = Mid(Location, 7, 2)

        Call UniCode_Conv(K0_ZAIKO.Soko_No, Soko_No)
        Call UniCode_Conv(K0_ZAIKO.Retu, Retu)
        Call UniCode_Conv(K0_ZAIKO.Ren, Ren)
        Call UniCode_Conv(K0_ZAIKO.Dan, Dan)
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(Retu)) = 0 Then
                        Retu = StrConv(ZAIKOREC.Retu, vbUnicode)
                    End If
                    If Len(Trim(Ren)) = 0 Then
                        Ren = StrConv(ZAIKOREC.Ren, vbUnicode)
                    End If
                    If Len(Trim(Dan)) = 0 Then
                        Dan = StrConv(ZAIKOREC.Dan, vbUnicode)
                    End If

                    If Soko_No <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Or _
                        Retu <> StrConv(ZAIKOREC.Retu, vbUnicode) Or _
                        Ren <> StrConv(ZAIKOREC.Ren, vbUnicode) Or _
                        Dan <> StrConv(ZAIKOREC.Dan, vbUnicode) Or _
                        JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Bt_sts = sts                                '2015.03.13
                    Call File_Error(sts, com, "在庫データ", mesg_mode)
                    Exit Function
            End Select

            Select Case StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
                Case "0"
                    Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                Case "1"
                    Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            End Select

            com = BtOpGetNext

'            DoEvents
            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                DoEvents                                                    '2016.01.26
            End If                                                          '2016.01.26
        Loop
    End If

    Zaiko_Syukei_Proc = False

End Function


'Public Function Kaso_Soko_No_Set() As Integer
'****************************************************
'*      仮想倉庫№の取込み
'*
'*  引数 :  なし
'*  戻り値: false       正常
'*          SYS_ERR     継続できない異常
'****************************************************
'Dim c As String
'
'    Kaso_Soko_No_Set = SYS_ERR
'                                    '仮想   入荷倉庫
'    If GetIni("SYSTEM", "KASO_NYUKA", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [KASO_NYUKA] READ ERROR")
'        Exit Function
'    End If
'    KASO_NYUKA_Soko = Trim(c)
'                                    '仮想   入荷場
'    If GetIni("SYSTEM", "KASO_NYUKABA", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [KASO_NYUKABA] READ ERROR")
'        Exit Function
'    End If
'    KASO_NYUKABA_Soko = Trim(c)
'                                    '仮想   商品化室
'    If GetIni("SYSTEM", "KASO_SYOHN", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [KASO_SYOHN] READ ERROR")
'        Exit Function
'    End If
'    KASO_SYOHN_Soko = Trim(c)
'                                    '仮想   内職
'    If GetIni("SYSTEM", "KASO_NAI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [KASO_NAI] READ ERROR")
'        Exit Function
'    End If
'    KASO_NAI_Soko = Trim(c)
'                                    '仮想   移動
'    If GetIni("SYSTEM", "KASO_IDO", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [KASO_IDO] READ ERROR")
'        Exit Function
'    End If
'    KASO_IDO_Soko = Trim(c)
'                                    '仮想   国内外振替
'    If GetIni("SYSTEM", "KASO_FURIKAE", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [KASO_FURIKAE] READ ERROR")
'        Exit Function
'    End If
'    KASO_FURIKAE_Soko = Trim(c)
'                                    '仮想   出荷場（未使用）
'    If GetIni("SYSTEM", "KASO_SYUKA", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [KASO_SYUKA] READ ERROR")
'        Exit Function
'    End If
'    KASO_SYUKA_Soko = Trim(c)
'
'    Kaso_Soko_No_Set = False
'End Function

Public Function Zaiko_Lock_Proc(Location As String, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    ID As String, _
                                    Optional Sumi_Zaiko_Qty As Long = 0, _
                                    Optional Mi_Zaiko_Qty As Long = 0, _
                                    Optional RETRY As Integer = 10) As Integer
'****************************************************
'*      在庫データの使用予約
'*
'*    引数：棚番（XXXXXXXX(倉庫№+列+連+段)省略不可）
'*          事業部（省略不可）
'*          国内外（省略不可）
'*          品番外部(省略不可)
'*          ID(省略不可)
'*          リトライ(省略可 １桁目:1=画面メッセージ有 0:無，２桁目:リトライ回数(0～9 0:無限))
'*
'*  引数 :  なし
'*  戻り値: false       :正常
'*          true        :継続できる異常
'*          SYS_ERR     :継続できない異常
'*          SYS_CANCEL  :ｷｬﾝｾﾙ
'****************************************************
Dim sts         As Integer
Dim com         As Integer

Dim RETRY_CNT   As Integer
Dim MESG_FLG    As Integer
Dim RETRY_SU    As Integer
    
Dim ans         As Integer

    Zaiko_Lock_Proc = True
    
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
    
    Call UniCode_Conv(K5_ZAIKO.Soko_No, Mid(Location, 1, 2))    '倉庫№
    Call UniCode_Conv(K5_ZAIKO.Retu, Mid(Location, 3, 2))       '列
    Call UniCode_Conv(K5_ZAIKO.Ren, Mid(Location, 5, 2))        '連
    Call UniCode_Conv(K5_ZAIKO.Dan, Mid(Location, 7, 2))        '段
    Call UniCode_Conv(K5_ZAIKO.JGYOBU, JGYOBU)                  '事業部
    Call UniCode_Conv(K5_ZAIKO.NAIGAI, NAIGAI)                  '内外
    Call UniCode_Conv(K5_ZAIKO.HIN_GAI, HIN_GAI)                '品番（外部）（両方）
    Call UniCode_Conv(K5_ZAIKO.NYUKA_DT, "")                    '入荷日（空白固定）

    Sumi_Zaiko_Qty = 0
    Mi_Zaiko_Qty = 0

    com = BtOpGetGreaterEqual

    Do
        RETRY_CNT = 0
        Do
            sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
            Select Case sts
                Case BtNoErr
                    If Mid(Location, 1, 2) <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Or _
                        Mid(Location, 3, 2) <> StrConv(ZAIKOREC.Retu, vbUnicode) Or _
                        Mid(Location, 5, 2) <> StrConv(ZAIKOREC.Ren, vbUnicode) Or _
                        Mid(Location, 7, 2) <> StrConv(ZAIKOREC.Dan, vbUnicode) Or _
                        JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        RTrim(HIN_GAI) <> RTrim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        
                        sts = BTRV(BtOpUnlock, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
                        If sts Then
                                                    
                            Call File_Error(sts, BtOpUnlock, "在庫データ")
                            Zaiko_Lock_Proc = SYS_ERR
                            Exit Function
                        
                        End If
                                        'ＫＥＹブレーク
                        Zaiko_Lock_Proc = False
                        Exit Function
                    End If
                
                    If StrConv(ZAIKOREC.LOCK_F, vbUnicode) = LOCK_ON Then
                                        
                                        
                                        '自タスクで占有中
'2016.06.17                        If StrConv(ZAIKOREC.WEL_ID, vbUnicode) = ID And
                        If Trim(StrConv(ZAIKOREC.WEL_ID, vbUnicode)) = Trim(ID) And _
                            StrConv(Trim(StrConv(ZAIKOREC.PRG_ID, vbUnicode)), vbUpperCase) = StrConv(App.EXEName, vbUpperCase) Then
                            Exit Do
                        Else
                                        
                                        
                            sts = BTRV(BtOpUnlock, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
                            If sts Then
                                                        
                                Call File_Error(sts, BtOpUnlock, "在庫データ")
                                Zaiko_Lock_Proc = SYS_ERR
                                Exit Function
                            
                            End If
                                        
                                        'リトライ回数チェック
'                            If RETRY_SU <> 0 Then
'
'                                RETRY_CNT = RETRY_CNT + 1
'                                If RETRY_CNT > RETRY_SU Then
'
'
'
'
'                                            '回数オーバー
'                                    Call File_Error(sts, com + BtSNoWait, "在庫データ", 0)
'                                    Zaiko_Lock_Proc = SYS_CANCEL
'                                    Exit Function
'
'                                End If
'
'                            End If
                
                            If MESG_FLG = 0 Then
                                
                                        
                                
                                
                                
                                Zaiko_Lock_Proc = SYS_CANCEL
                                Exit Function

'                                DoEvents
                            Else
                                
                                
                                
                                Beep
'2015.05.12                                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                ans = MsgBox("『" & Trim(StrConv(ZAIKOREC.WEL_ID, vbUnicode)) & "』で作業中です。処理を完了させて下さい。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")  '2015.05.12
                                If ans = vbCancel Then
                                    Zaiko_Lock_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
'---------------------------' 2001.08.07
'                            com = BtOpGetEqual
                            com = BtOpGetGreaterEqual
'---------------------------' 2001.08.07
                        
                        End If
                    Else
                        Exit Do
                    End If
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                    If RETRY_SU <> 0 Then
                    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                            Call File_Error(sts, com + BtSNoWait, "在庫データ", 0)
                            Zaiko_Lock_Proc = SYS_CANCEL
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
                            Zaiko_Lock_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                
                Case BtErrEOF
                    Zaiko_Lock_Proc = False
                    Exit Function
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "在庫データ")
                    Zaiko_Lock_Proc = SYS_ERR
                    Exit Function
            End Select
        Loop
    
        Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_ON)     '排他ﾌﾗｸﾞ（ON）
        Call UniCode_Conv(ZAIKOREC.WEL_ID, ID)          '使用子機ID
                                                        '使用ﾌﾟﾛｸﾞﾗﾑ
        Call UniCode_Conv(ZAIKOREC.PRG_ID, StrConv(App.EXEName, vbUpperCase))
        
        RETRY_CNT = 0
        Do
            sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                    If RETRY_SU <> 0 Then
                    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                            Call File_Error(sts, BtOpUpdate, "在庫ﾃﾞｰﾀ", 0)
                            Zaiko_Lock_Proc = SYS_CANCEL
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
                            Zaiko_Lock_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "在庫データ")
                    Zaiko_Lock_Proc = SYS_ERR
                    Exit Function
                
            End Select
        Loop
        
        
        If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
            Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
        Else
            Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
        End If
        
        com = BtOpGetNext
    
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
    Loop
End Function
Public Function Zaiko_UNLock_Proc(Location As String, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    ID As String, _
                                    Optional RETRY As Integer = 10) As Integer
'****************************************************
'*      在庫データの使用予約解除
'*
'*    引数：棚番（XXXXXXXX(倉庫№+列+連+段)省略可）
'*          事業部（省略可）
'*          国内外（省略可）
'*          品番外部(省略可)
'*          ID(省略可)
'*          リトライ(省略可 １桁目:1=画面メッセージ有 0:無，２桁目:リトライ回数(0～9 0:無限))
'*  ※棚番・事業部・国内外・品番外部又はIDの何れかが必須！！
'*  引数 :  なし
'*  戻り値: false       :正常
'*          true        :継続可能な異常
'*          SYS_ERR     :継続できない異常
'*          SYS_CANCEL  :ｷｬﾝｾﾙ
'****************************************************
Dim sts         As Integer
Dim com         As Integer

Dim RETRY_CNT   As Integer
Dim MESG_FLG    As Integer
Dim RETRY_SU    As Integer
    
Dim ans         As Integer

    Zaiko_UNLock_Proc = True
    
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
    
    
    If Len(Trim(Location)) = 0 Then
'---------------------------------------------------------------'プログラムIDでﾛｯｸ解除
        Call UniCode_Conv(K3_ZAIKO.WEL_ID, ID)                  '使用子機ID
                                                                '使用プログラムID
        Call UniCode_Conv(K3_ZAIKO.PRG_ID, StrConv(App.EXEName, vbUpperCase))

        com = BtOpGetGreaterEqual

        Do
            RETRY_CNT = 0
            Do
                sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
                Select Case sts
                    Case BtNoErr
                        If ID <> StrConv(ZAIKOREC.WEL_ID, vbUnicode) Or _
                            StrConv(App.EXEName, vbUpperCase) <> Trim(StrConv(ZAIKOREC.PRG_ID, vbUnicode)) Then
                            
                            sts = BTRV(BtOpUnlock, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "在庫データ")
                                Zaiko_UNLock_Proc = SYS_ERR
                                Exit Function
                            End If
                                        'ＫＥＹブレーク
                            Zaiko_UNLock_Proc = False
                            Exit Function
                        End If
                        
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, com + BtSNoWait, "在庫データ", 0)
                                Zaiko_UNLock_Proc = SYS_CANCEL
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
                                Zaiko_UNLock_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If

'                        com = BtOpGetEqual
                
                    Case BtErrEOF
                        Zaiko_UNLock_Proc = False
                        Exit Function
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "在庫データ")
                        Zaiko_UNLock_Proc = SYS_ERR
                        Exit Function
                End Select
            Loop
    
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)    '排他ﾌﾗｸﾞ（OFF）
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")          '使用子機ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")          '使用ﾌﾟﾛｸﾞﾗﾑ
        
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), BtNCC)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, BtOpUpdate, "在庫ﾃﾞｰﾀ", 0)
                                Zaiko_UNLock_Proc = SYS_CANCEL
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
                                Zaiko_UNLock_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "在庫データ")
                        Zaiko_UNLock_Proc = SYS_ERR
                        Exit Function
                
                End Select
            Loop
        
            com = BtOpGetNext
    
'            DoEvents
            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                DoEvents                                                    '2016.01.26
            End If                                                          '2016.01.26
        Loop
    
    Else
'---------------------------------------------------------------'棚番＋品番でﾛｯｸ解除
        Call UniCode_Conv(K5_ZAIKO.Soko_No, Mid(Location, 1, 2))    '倉庫№
        Call UniCode_Conv(K5_ZAIKO.Retu, Mid(Location, 3, 2))       '列
        Call UniCode_Conv(K5_ZAIKO.Ren, Mid(Location, 5, 2))        '連
        Call UniCode_Conv(K5_ZAIKO.Dan, Mid(Location, 7, 2))        '段
        Call UniCode_Conv(K5_ZAIKO.JGYOBU, JGYOBU)                  '事業部
        Call UniCode_Conv(K5_ZAIKO.NAIGAI, NAIGAI)                  '内外
        Call UniCode_Conv(K5_ZAIKO.HIN_GAI, HIN_GAI)                '品番（外部）
        Call UniCode_Conv(K5_ZAIKO.NYUKA_DT, "")                    '入荷日（空白固定）

        com = BtOpGetGreater

        Do
            RETRY_CNT = 0
            Do
                sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
                Select Case sts
                    Case BtNoErr
                        If Mid(Location, 1, 2) <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Or _
                            Mid(Location, 3, 2) <> StrConv(ZAIKOREC.Retu, vbUnicode) Or _
                            Mid(Location, 5, 2) <> StrConv(ZAIKOREC.Ren, vbUnicode) Or _
                            Mid(Location, 7, 2) <> StrConv(ZAIKOREC.Dan, vbUnicode) Or _
                            JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                            NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                            RTrim(HIN_GAI) <> RTrim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                                        'ＫＥＹブレーク
                            sts = BTRV(BtOpUnlock, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
                            If sts Then
                                                    
                                Call File_Error(sts, BtOpUnlock, "在庫データ")
                                Zaiko_UNLock_Proc = SYS_ERR
                                Exit Function
                        
                            End If
                            
                            Zaiko_UNLock_Proc = False
                            Exit Function
                        End If
                        
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                        
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, com + BtSNoWait, "在庫データ", 0)
                                Zaiko_UNLock_Proc = SYS_CANCEL
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
                                Zaiko_UNLock_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                
                    Case BtErrEOF
                        Zaiko_UNLock_Proc = False
                        Exit Function
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "在庫データ")
                        Zaiko_UNLock_Proc = SYS_ERR
                        Exit Function
                End Select
            Loop
    
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)    '排他ﾌﾗｸﾞ（ON）
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")          '使用子機ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")          '使用ﾌﾟﾛｸﾞﾗﾑ
        
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                Call File_Error(sts, BtOpUpdate, "在庫ﾃﾞｰﾀ", 0)
                                Zaiko_UNLock_Proc = SYS_CANCEL
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
                                Zaiko_UNLock_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, com, "在庫データ")
                        Zaiko_UNLock_Proc = SYS_ERR
                        Exit Function
                
                End Select
            Loop
        
            com = BtOpGetNext
'            DoEvents
            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                DoEvents                                                    '2016.01.26
            End If                                                          '2016.01.26
    
        Loop
    End If
End Function


Public Function SOKO_Zaiko_Syukei_Proc(Sumi_Zaiko_Qty As Long, _
                                    Mi_Zaiko_Qty As Long, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    Optional Location As String = "        ") As Integer
'****************************************************
'*      在庫数集計
'*
'*  品番または品番＋棚番毎の在庫数を集計する。
'*
'*  引数 :  在庫数（商品化済み）
'*          在庫数（未商品）
'*          事業部（省略不可）
'*          国内外（省略不可）
'*          品番外部(省略不可)
'*          棚番(省略可 省略=空白)
'*          商品化済み除外倉庫(省略可)                          2011.12.16
'*          商品化済み除外倉庫有無(省略可　省略=除外倉庫なし）  2011.12.16
'*          除外倉庫有無                                        2012.03.20
'*
'*  戻り値: false    正常
'*          SYS_ERR  継続できない異常
'*
'*          2014.07.01 奈良用に新設
'****************************************************
Dim sts     As Integer
Dim com     As Integer
Dim Soko_No As String * 2
Dim Retu    As String * 2
Dim Ren     As String * 2
Dim Dan     As String * 2

Dim i       As Long




    SOKO_Zaiko_Syukei_Proc = SYS_ERR



    Sumi_Zaiko_Qty = 0
    Mi_Zaiko_Qty = 0

    com = BtOpGetGreaterEqual


    Soko_No = Mid(Location, 1, 2)
    Retu = Mid(Location, 3, 2)
    Ren = Mid(Location, 5, 2)
    Dan = Mid(Location, 7, 2)

    Call UniCode_Conv(K4_ZAIKO.JGYOBU, JGYOBU)
    Call UniCode_Conv(K4_ZAIKO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K4_ZAIKO.HIN_GAI, HIN_GAI)
    
    Call UniCode_Conv(K4_ZAIKO.Soko_No, Soko_No)
    Call UniCode_Conv(K4_ZAIKO.Retu, Retu)
    Call UniCode_Conv(K4_ZAIKO.Ren, Ren)
    Call UniCode_Conv(K4_ZAIKO.Dan, Dan)
        
        
    Do
            
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
        Select Case sts
        Case BtNoErr
            If Len(Trim(Retu)) = 0 Then
                Retu = StrConv(ZAIKOREC.Retu, vbUnicode)
            End If
            If Len(Trim(Ren)) = 0 Then
                Ren = StrConv(ZAIKOREC.Ren, vbUnicode)
            End If
            If Len(Trim(Dan)) = 0 Then
                Dan = StrConv(ZAIKOREC.Dan, vbUnicode)
            End If

            If Soko_No <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Or _
                Retu <> StrConv(ZAIKOREC.Retu, vbUnicode) Or _
                Ren <> StrConv(ZAIKOREC.Ren, vbUnicode) Or _
                Dan <> StrConv(ZAIKOREC.Dan, vbUnicode) Or _
                JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                Exit Do
            End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫データ")
                Exit Function
        End Select

        Select Case StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
            Case "0"
                Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            Case "1"
                Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
        End Select

        com = BtOpGetNext

'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
    Loop

    SOKO_Zaiko_Syukei_Proc = False

End Function



Public Function NEW_SOKO_Zaiko_Syukei_Proc(Sumi_Zaiko_Qty As Long, _
                                    Mi_Zaiko_Qty As Long, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    Optional Location As String = "        ") As Integer
'****************************************************
'*      在庫数集計
'*
'*  品番または品番＋棚番毎の在庫数を集計する。
'*
'*  引数 :  在庫数（商品化済み）
'*          在庫数（未商品）
'*          事業部（省略不可）
'*          国内外（省略不可）
'*          品番外部(省略不可)
'*          棚番(省略可 省略=空白)
'*          商品化済み除外倉庫(省略可)                          2011.12.16
'*          商品化済み除外倉庫有無(省略可　省略=除外倉庫なし）  2011.12.16
'*          除外倉庫有無                                        2012.03.20
'*
'*  戻り値: false    正常
'*          SYS_ERR  継続できない異常
'*
'*          2014.07.01 奈良用に新設
'*          2018.09.18 奈良用に新設(修正)
'****************************************************
Dim sts     As Integer
Dim com     As Integer
Dim Soko_No As String * 2
Dim Retu    As String * 2
Dim Ren     As String * 2
Dim Dan     As String * 2


Dim S_Retu  As String * 2
Dim S_Ren   As String * 2
Dim S_Dan   As String * 2


Dim i       As Long




    NEW_SOKO_Zaiko_Syukei_Proc = SYS_ERR



    Sumi_Zaiko_Qty = 0
    Mi_Zaiko_Qty = 0

    com = BtOpGetGreaterEqual


    Soko_No = Mid(Location, 1, 2)
    S_Retu = Mid(Location, 3, 2)
    S_Ren = Mid(Location, 5, 2)
    S_Dan = Mid(Location, 7, 2)

    Call UniCode_Conv(K4_ZAIKO.JGYOBU, JGYOBU)
    Call UniCode_Conv(K4_ZAIKO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K4_ZAIKO.HIN_GAI, HIN_GAI)
    
    Call UniCode_Conv(K4_ZAIKO.Soko_No, Soko_No)
    Call UniCode_Conv(K4_ZAIKO.Retu, Retu)
    Call UniCode_Conv(K4_ZAIKO.Ren, Ren)
    Call UniCode_Conv(K4_ZAIKO.Dan, Dan)
        
        
    Do
            
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
        Select Case sts
        Case BtNoErr
            If Len(Trim(S_Retu)) = 0 Then
                Retu = StrConv(ZAIKOREC.Retu, vbUnicode)
            End If
            If Len(Trim(S_Ren)) = 0 Then
                Ren = StrConv(ZAIKOREC.Ren, vbUnicode)
            End If
            If Len(Trim(S_Dan)) = 0 Then
                Dan = StrConv(ZAIKOREC.Dan, vbUnicode)
            End If

            If Soko_No <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Or _
                Retu <> StrConv(ZAIKOREC.Retu, vbUnicode) Or _
                Ren <> StrConv(ZAIKOREC.Ren, vbUnicode) Or _
                Dan <> StrConv(ZAIKOREC.Dan, vbUnicode) Or _
                JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                Exit Do
            End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫データ")
                Exit Function
        End Select

        Select Case StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
            Case "0"
                Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            Case "1"
                Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
        End Select

        com = BtOpGetNext

'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
    Loop

    NEW_SOKO_Zaiko_Syukei_Proc = False

End Function


