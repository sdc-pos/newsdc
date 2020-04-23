Attribute VB_Name = "HS_SIJ"
Option Explicit
'********************************************************************
'*
'*              ホスト受信データ ファイル定義
'*
'*          CREATE 2004.03.04
'********************************************************************
'ファイルＩＤ
Public Const HS_IN_SIJ_ID$ = "HS_IN_SIJ"
Public Const HS_OUT_SIJ_ID$ = "HS_OUT_SIJ"
'ファイル№
Public HS_SIJ_No As Integer
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義(入庫)
Type HS_IN_SIJREC_Tag
    
    
    
    TEXT_NO(0 To 8) As Byte         'ﾃｷｽﾄ№
    JGYOBU(0 To 0) As Byte          '事業部区分
    CYOK_KBN(0 To 0) As Byte        '直送区分
    DEN_DT(0 To 7) As Byte          '伝票日付
    IO_KBN(0 To 0) As Byte          '入出庫区分
    PM_KBN(0 To 0) As Byte          '赤黒区分
    DEN_SYU(0 To 0) As Byte         '伝票種別
    DEN_NO(0 To 5) As Byte          '伝票№
    CYU_KBN(0 To 0) As Byte         '注文区分
    HIN_GAI(0 To 12) As Byte        '品番（外部）
    HIN_NAI(0 To 12) As Byte        '品番（内部）
    HIN_NAME(0 To 24) As Byte       '品名
    YOTEI_QTY(0 To 5) As Byte       '数量
    YOSAN_FROM(0 To 4) As Byte      '予算単位（元）
    YOSAN_TO(0 To 4) As Byte        '予算単位（先）
    HOST_SOKO(0 To 1) As Byte       '倉庫区分（ﾎｽﾄ）
    HOST_TANA(0 To 7) As Byte       '棚番（ﾎｽﾄ）
    SYUK_CODE(0 To 4) As Byte       '支給先／出荷先
    SYUK_NAME(0 To 19) As Byte      '支給先／出荷先名
    REC_END(0 To 0) As Byte         'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
    CR_LF(0 To 1) As Byte           'CR.LF
    
    
    
End Type




'データ・バッファ
Public HS_IN_SIJREC As HS_IN_SIJREC_Tag
'-------------------------------------------'
'レコード定義(出荷)
Type HS_OUT_SIJREC_Tag
'    JGYOBA(0 To 7)              As Byte     '事業場
'    DATA_KBN(0 To 0)            As Byte     'データ区分
'    TORI_KBN(0 To 1)            As Byte     '取引区分
'    ID_NO(0 To 7)               As Byte     'ID-NO
'    HIN_NO(0 To 19)             As Byte     '品目番号
'    DEN_NO(0 To 9)              As Byte     '伝票番号
'    SURYO(0 To 6)               As Byte     '出庫数量
'    MUKE_CODE(0 To 7)           As Byte     '得意先コード
'    SYUKO_SYUSI(0 To 1)         As Byte     '出庫収支
'    SYUKA_YMD(0 To 7)           As Byte     '出荷日付
'    ODER_NO(0 To 11)            As Byte     'オーダー番号
'    ITEM_NO(0 To 4)             As Byte     'アイテム番号
'    MUKE_NAME(0 To 23)          As Byte     '得意先名称
'    CHU_KBN(0 To 0)             As Byte     '注文区分
'    CHU_KBN_NAME(0 To 9)        As Byte     '注文区分名称
'    EXPORT_KBN(0 To 0)          As Byte     '輸出出荷検査区分
'    LABEL_ISSUE_KBN(0 To 0)     As Byte     '個装ラベル発行区分
'    LABEL_ISSUE_UNIT(0 To 4)    As Byte     '個装ラベル発行単位数
'    LABEL_TANKA_KBN(0 To 0)     As Byte     '個装ラベル単価表示区分
'    TANKA(0 To 9)               As Byte     '単価
'    KINGAKU(0 To 9)             As Byte     '金額
'    BIKOU2(0 To 19)             As Byte     '備考２
'    REBATE_KBN(0 To 0)          As Byte     'リベート区分
'    CHOHA_KBN(0 To 0)           As Byte     '帳端区分
'    ATAISA_KBN(0 To 0)          As Byte     '値差区分
'    REP_KISHU(0 To 19)          As Byte     '代表機種
'    NS_KANRI_NO(0 To 8)         As Byte     'ＮＳ管理番号
'    MTS_HIN_CODE(0 To 10)       As Byte     'ＭＴＳ部品コード
'    BIKOU1(0 To 39)             As Byte     '備考１
'    CHOKU_KBN(0 To 0)           As Byte     '直送区分
'    REBATE_RATE(0 To 4)         As Byte     'リベート率
'    HIN_NAME(0 To 19)           As Byte     '品名
'    JGYOBU_GAI(0 To 7)          As Byte     '対外事業場
'    SS_CODE(0 To 7)             As Byte     '直送先コード
'    KISHU_HIN_NO(0 To 2)        As Byte     '機種品目コード
'    HIN_NAI(0 To 19)            As Byte     '品番（内部）
'    CRLF(0 To 1)                As Byte     'CRLF

    JGYOBA(0 To 7)              As Byte     '事業場
    DATA_KBN(0 To 0)            As Byte     'データ区分
    TORI_KBN(0 To 1)            As Byte     '取引区分
    ID_NO(0 To 7)               As Byte     'ID-NO
    HIN_NO(0 To 19)             As Byte     '品目番号
    DEN_NO(0 To 9)              As Byte     '伝票番号
    SURYO(0 To 6)               As Byte     '出庫数量
    MUKE_CODE(0 To 7)           As Byte     '出庫先
    SYUKO_SYUSI(0 To 1)         As Byte     '出庫収支
    SYUKO_YMD(0 To 7)           As Byte     '出庫日付
    TANKA(0 To 9)               As Byte     '単価
    ODER_NO(0 To 11)            As Byte     'オーダー番号
    ITEM_NO(0 To 4)             As Byte     'アイテム番号
    ODER_R_NO(0 To 4)           As Byte     'オーダー略号
    KOSO_KEITAI(0 To 9)         As Byte     '個装形態
    SYUKA_YMD(0 To 7)           As Byte     '出荷日
    TANABAN1(0 To 9)            As Byte     '棚番１
    TANABAN2(0 To 9)            As Byte     '棚番２
    TANABAN3(0 To 9)            As Byte     '棚番３
    MUKE_NAME(0 To 23)          As Byte     '出庫先名称
    CHU_KBN(0 To 0)             As Byte     '注文区分
    CHU_KBN_NAME(0 To 9)        As Byte     '注文区分名称
    ORIGIN1(0 To 9)             As Byte     '原産国１
    ORIGIN2(0 To 9)             As Byte     '原産国２
    BIKOU2(0 To 39)             As Byte     '備考２
    HAN_KBN(0 To 0)             As Byte     '販売区分
    CHOKU_KBN(0 To 0)           As Byte     '直送区分
    UNIT_ID_NO(0 To 7)          As Byte     'ﾕﾆｯﾄ修理ID-NO
    ZAIKO_HIKIATE(0 To 2)       As Byte     '在庫引当順序
    GOKON_KANRI_NO(0 To 8)      As Byte     '合梱管理番号
    JUCHU_ZAN(0 To 6)           As Byte     '受注残数量
    KYOKYU_KBN(0 To 0)          As Byte     '供給区分
    SHOHIN_SYUSI(0 To 1)        As Byte     '商品化納入先収支
    BIKOU1(0 To 39)             As Byte     '備考１
    CHOHA_KBN(0 To 0)           As Byte     '帳端区分
    JYU_HIN_NO(0 To 19)         As Byte     '受注品目番号
    HIN_NAME(0 To 19)           As Byte     '品名
    HIN_CHANGE_KBN(0 To 0)      As Byte     '品番変更区分
    MODULE_EXCHANGE(0 To 0)     As Byte     'モジュール交換区分
    ZAIKO_SYUSI(0 To 1)         As Byte     '残在庫まとめ在庫収支コード
    NOUKI_YMD(0 To 7)           As Byte     '指定納期
    SERVICE_KANRI_NO(0 To 8)    As Byte     'サービス会社管理番号
    KISHU_CODE(0 To 2)          As Byte     '機種品目コード
    ENVIRONMENT_KBN(0 To 0)     As Byte     '環境規格部品区分
    SS_CODE(0 To 7)             As Byte     '直送先コード
    FILLER(0 To 4)              As Byte
    CRLF(0 To 1)                As Byte     'CRLF









End Type

'データ・バッファ
Public HS_OUT_SIJREC As HS_OUT_SIJREC_Tag
Public Function HS_SIJ_Open(Mode As Integer, Data_Type As Integer) As Integer
'********************************************************************
'*
'*      ホスト受信データ  ＯＰＥＮ
'*
'*      引数　:OPENモード（0:参照　1:更新）
'*             ﾃﾞｰﾀﾀｲﾌﾟ   (1:入庫　2:出荷)
'*
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2004.03.05
'********************************************************************

Dim ans         As Integer
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    On Error GoTo HS_SIJ_Op_Err     'ｴﾗｰﾄﾗｯﾌﾟON

    HS_SIJ_Open = True
                                    
    Select Case Data_Type
        Case 1          '入庫
            If GetIni("FILE", HS_IN_SIJ_ID, "SYS", c) Then
                Call LOG_OUT(LOG_F, "SYS.INI [HS_IN_SIJ]読み込みエラー")
                Exit Function
            End If
        Case 2          '出庫
            If GetIni("FILE", HS_OUT_SIJ_ID, "SYS", c) Then
                Call LOG_OUT(LOG_F, "SYS.INI [HS_OUT_SIJ]読み込みエラー")
                Exit Function
            End If
    End Select
                                    
    FullPath = RTrim(c)
    
    HS_SIJ_No = FreeFile

    If Mode = 0 Then
        Open FullPath For Input As #HS_SIJ_No
    Else
        Open FullPath For Binary As #HS_SIJ_No
    End If
    
    HS_SIJ_Open = False

    Exit Function

HS_SIJ_Op_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case ErrDiskNotReady
            If Mode = 1 Then
                Beep
                ans = MsgBox("ドライブを確認して下さい", vbYesNo + vbExclamation + vbDefaultButton1, "確認入力")
                If ans = vbYes Then
                    Resume
                End If
            End If
        Case ErrDeviceUnavailable
            If Mode = 1 Then
                Beep
                ans = MsgBox("ドライブまたはパスが見つかりません" & FullPath, vbExclamation)
            End If
        Case ErrNotFound
            If Mode = 1 Then
                Beep
                ans = MsgBox("ファイルが見つかりません" & FullPath, vbExclamation)
            End If
        Case Else
            If Mode = 1 Then
                Beep
                ans = MsgBox("エラー [HS_SIJ Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
End Function
