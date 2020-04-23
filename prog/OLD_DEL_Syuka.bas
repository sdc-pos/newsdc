Attribute VB_Name = "OLD_DEL_SYU"
Option Explicit
'********************************************************************
'*
'*              （旧）削除済み出荷予定データ  ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const OLD_DEL_SYU_ID$ = "OLD_DEL_SYU"

'ページサイズ
Public Const OLD_DEL_SYU_PG_SIZ% = 2048

'ポジション・ブロック
Public OLD_DEL_SYU_POS    As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type OLD_DEL_SYUREC_Tag
    WEL_ID(0 To 2)              As Byte     '使用子機ID
    PRG_ID(0 To 7)              As Byte     '使用中プログラム
    KAN_KBN(0 To 0)             As Byte     '完了区分
    DT_SYU(0 To 0)              As Byte     'データ種別
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    KEY_ID_NO(0 To 7)           As Byte     'ID-NO
    NAIGAI(0 To 0)              As Byte     '国内外
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
    JGYOBA(0 To 7)              As Byte     '事業場
    DATA_KBN(0 To 0)            As Byte     'データ区分
    TORI_KBN(0 To 1)            As Byte     '取引区分
    ID_NO(0 To 7)               As Byte     'ID-NO
    HIN_NO(0 To 19)             As Byte     '品目番号
    DEN_NO(0 To 9)              As Byte     '伝票番号
    SURYO(0 To 6)               As Byte     '出庫数量
    MUKE_CODE(0 To 7)           As Byte     '得意先コード
    SYUKO_SYUSI(0 To 1)         As Byte     '出庫収支
    SYUKA_YMD(0 To 7)           As Byte     '出荷日付
    ODER_NO(0 To 11)            As Byte     'オーダー番号
    ITEM_NO(0 To 4)             As Byte     'アイテム番号
    MUKE_NAME(0 To 23)          As Byte     '得意先名称
    CYU_KBN(0 To 0)             As Byte     '注文区分
    CYU_KBN_NAME(0 To 9)        As Byte     '注文区分名称
    EXPORT_KBN(0 To 0)          As Byte     '輸出出荷検査区分
    LABEL_ISSUE_KBN(0 To 0)     As Byte     '個装ラベル発行区分
    LABEL_ISSUE_UNIT(0 To 4)    As Byte     '個装ラベル発行単位数
    LABEL_TANKA_KBN(0 To 0)     As Byte     '個装ラベル単価表示区分
    TANKA(0 To 9)               As Byte     '単価
    KINGAKU(0 To 9)             As Byte     '金額
    BIKOU2(0 To 19)             As Byte     '備考２
    REBATE_KBN(0 To 0)          As Byte     'リベート区分
    CHOHA_KBN(0 To 0)           As Byte     '帳端区分
    ATAISA_KBN(0 To 0)          As Byte     '値差区分
    REP_KISHU(0 To 19)          As Byte     '代表機種
    NS_KANRI_NO(0 To 8)         As Byte     'ＮＳ管理番号
    MTS_HIN_CODE(0 To 10)       As Byte     'ＭＴＳ部品コード
    BIKOU1(0 To 39)             As Byte     '備考１
    CHOKU_KBN(0 To 0)           As Byte     '直送区分
    REBATE_RATE(0 To 4)         As Byte     'リベート率
    HIN_NAME(0 To 19)           As Byte     '品名
    JGYOBA_GAI(0 To 7)          As Byte     '対外事業場
    KISHU_CODE(0 To 2)          As Byte     '機種コード
    SS_CODE(0 To 7)             As Byte     '直送先コード
    HIN_NAI(0 To 12)            As Byte     '品番（内部）
    HTANABAN(0 To 7)            As Byte     'ホスト棚番
    PRINT_YMD(0 To 7)           As Byte     '出庫表印刷日付
    KAN_YMD(0 To 7)             As Byte     '完了日付
    KENPIN_YMD(0 To 7)          As Byte     '検品日付
    TOK_KBN(0 To 0)             As Byte     '特売り区分
    JITU_SURYO(0 To 6)          As Byte     '出庫実績数量
    INS_NOW(0 To 13)            As Byte     '取込み日時
    FILLER(0 To 74)             As Byte     'FILLER
End Type

'データ・バッファ
Public OLD_DEL_SYUREC           As OLD_DEL_SYUREC_Tag

'キー定義
Type KEY0_OLD_DEL_SYU            'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
End Type



'キー・データ
Public K0_OLD_DEL_SYU               As KEY0_OLD_DEL_SYU

Function OLD_DEL_SYU_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*             （旧）削除済み出荷予定データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_DEL_SYU_Open = True
                                            '削除済み出荷予定データフルパス取込み
    sts = GetIni("FILE", OLD_DEL_SYU_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI [OLD_DEL_SYU]読み込みエラー ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_DEL_SYU_POS, OLD_DEL_SYUREC, Len(OLD_DEL_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            
            Case BtErrFileNotFound
            
                OLD_DEL_SYU_Open = sts
            
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "（旧）削除済み出荷予定データ")
                Exit Function
        End Select
    Loop
    
    OLD_DEL_SYU_Open = False

End Function
