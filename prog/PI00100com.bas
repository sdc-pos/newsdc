Attribute VB_Name = "PI00100com"
Option Explicit

Private Type Item_Key_tag
    JGYOBU  As String * 1
    NAIGAI  As String * 1
End Type

Public K_Item_Tbl() As Item_Key_tag   '個装資材品目情報
Public G_Item_Tbl() As Item_Key_tag   '外装資材品目情報



Private Type D_Item_Tbl_Tag
    SYUBETSU    As String * 2               '種別
    JGYOBU      As String * 1               '事業部
    NAIGAI      As String * 1               '国内外
    HIN_GAI     As String * 20              '品番
    QTY         As Double                   '員数
    SHIJI_QTY   As Double                   '数量（指示数）
    BIKOU       As String * 40              '備考（入力値）
    ID_NO       As String * 12              'ID_No(出荷予定ID_No)
End Type



Public D_Item_Tbl()     As D_Item_Tbl_Tag   '同梱／構成品目情報


Public Taget_Key        As String * 8       '更新対象の指図票№ 2008.02.13

Public Taget_SHIMUKE_CODE_KEY _
                        As String * 2       '印刷対象　仕向け先 2008.02.02

Public Taget_Hin_key    As String * 20      '印刷対象　品番     2008.02.02
Public Taget_JGYOBU_key As String * 1       '印刷対象　事業部   2008.02.02
Public Taget_NAIGAI_key As String * 1       '印刷対象　国内外   2008.02.02


Public Doukon_Tbl_No(0 To 19) _
                        As String * 1

Public Doukon_Start     As Integer          '画面開始行№

Public POS_UMU          As Boolean

Public PRI_S_TANTO      As Boolean      '収支／担当者印刷 OFF:印刷なし ON:印刷あり
Public PRI_MAIN_BCR     As Boolean      'ﾒｲﾝﾊﾞｰｺｰﾄﾞ OFF:印刷なし ON:印刷あり
Public PRI_BIKOU_BCR    As Boolean      '備考欄　OFF：入力値　ON:出荷BCR

'2011.08.04
'Public PRI_DOUKON       As Boolean     '商品化検査　同梱 OFF:印刷なし ON:印刷あり
Public PRI_DOUKON       As Integer      '商品化検査　同梱 0:同梱・刻印印刷なし 1:同梱印字 2:刻印印字
'2011.08.04

Public PRI_NYUKO_IN     As Boolean      '入庫完了印　同梱 OFF:印刷なし ON:印刷あり

Public PRI_INPUT_IN     As Boolean      '入力完了印　同梱 OFF:印刷なし ON:印刷あり



Public PRI_SAGYO_DAY    As Boolean      '作業日／数量／担当 OFF:印刷なし ON:印刷あり 2007.05.22
Public PRI_HINBAN_BIKOU As Boolean      '下部　品番／№／数量 OFF:印刷なし ON:印刷あり 2007.05.22


Public JISEKI_TITLE     As Variant      '自責の名称タイトル
Public TASEKI_TITLE     As Variant      '他責の名称タイトル




Public HIN_INV          As Boolean      '未登録品番可否


Public LabelPrint_F     As String       '2008.05.30




Public JISSEKI_DSP      As String * 1   '2008.08.19



Public chk_TORI_GENSANKOKU  As String * 20 '原産国有無ﾁｪｯｸ用   2013.01.08


Public KAIKON_PRI       As Boolean      '開梱・リード巻線・粘着防止・の表示 2013.01.16


Public GENSANKOKU_MSG_F As Boolean      '原産国ﾒｯｾｰｼﾞ表示有無   2013.02.19



Public KAISYA_DEF_VALUE     As String   'ﾃﾞﾌｫﾙﾄ会社ｺｰﾄﾞ     2013.03.28
Public JIGYOBU_DEF_VALUE    As String   'ﾃﾞﾌｫﾙﾄ事業部ｺｰﾄﾞ   2013.03.28

Public NYUKA_KANSYOZAI  As Boolean      '入荷時緩衝材の表示 2013.11.05
    
    
    
Public PRINT_STOP_F     As Boolean      '印刷中止　2015.03.26
    
Public LABEL_PLUS        As Integer      'ラベル発行枚数 2015.04.02
    
Public GAI_BUHIN_CHK    As Boolean      '海外供給区分チェック 2015.07.23
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   前回修正年月日
'Public Const Last_Update_Day$ = "(PI00010 2015.12.14 17:30)"



Public PI000104_Error_F     As Integer      '2019.03.14
Public PI000104_HIN_GAI     As String * 20  '2019.03.14
Public PI000104_OLD_HIN_GAI As String * 20  '2019.03.14

Public PI000104_CANCEL_F    As Integer      '2019.03.14

'---------------------------------------------- *商品化指図ﾃﾞｰﾀ（親）別ポインタ
'ポジショニング
Public wP_SSHIJI_O_POS  As POSBLK
'データ・バッファ
Public wP_SSHIJI_O_REC  As P_SSHIJI_O_REC_Tag
'キー・データ
Public K0_wP_SSHIJI_O   As KEY0_P_SSHIJI_O
Public K1_wP_SSHIJI_O   As KEY1_P_SSHIJI_O

' ------------------------------------------------------------------------
'       指定した精度の数値に四捨五入します。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に四捨五入された数値。
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function

Public Function File_Open_Proc() As Integer
'----------------------------------------------------------------------------
'               ファイル　ＯＰＥＮ処理
'           2015.03.13
'           2015.04.24 Sub --> Function
'----------------------------------------------------------------------------
                                
Dim sts     As Integer
                                
    File_Open_Proc = True
                                
    DoEvents
                                
Call LOG_OUT(LOG_F, "File 再オープン処理 　開始")           '2015.03.26
                                
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
        Exit Function
    End If
                                
                                
                                
                                '発番マスタＯＰＥＮ
    If HATUBAN_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If

                                '商品ﾗﾍﾞﾙ用品目マスタＯＰＥＮ
    If L_ITEM_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If

                                'クラスマスタＯＰＥＮ
    If P_Class_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '商品化指図（子）ﾃﾞｰﾀＯＰＥＮ
    If P_SSHIJI_K_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '商品化指図（親）ﾃﾞｰﾀＯＰＥＮ
    If P_SSHIJI_O_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '出荷予定ﾃﾞｰﾀＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If


    '2010.07.20 ▽
                                '原産国マスタＯＰＥＮ
    If GENSAN_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
    '2010.07.20 △
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If

                                '商品化指図（親）ﾜｰｸＯＰＥＮ
    If wP_SSHIJI_O_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If

                                '入出庫単価設定マスタＯＰＥＮ   2008.09.20
    If SE_LOC_TANKA_M_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If


    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
                                'PNマスタＯＰＥＮ
    If PN_M_Open(0) Then
'        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload PI000101
        Exit Function
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Call LOG_OUT(LOG_F, "File 再オープン処理 　正常終了")           '2015.03.26

    File_Open_Proc = False

End Function

Public Function wP_SSHIJI_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              商品化指図(親)ワーク  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    wP_SSHIJI_O_Open = True
                                            '商品化指図(親)ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_O]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "商品化指図(親)ﾜｰｸ")
                Exit Function
        End Select
    Loop

    wP_SSHIJI_O_Open = False

End Function

Sub Main()
    
    
    
Dim lngReturnValue      As Long
Dim strMyTitle          As String
Dim lngPrevHwnd         As Long
Dim lngTopHwnd          As Long
Dim lngThreadID1        As Long
Dim lngThreadID2        As Long
    
    
    
    




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










    PI000101.Show
End Sub

