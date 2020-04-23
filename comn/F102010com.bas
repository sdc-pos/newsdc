Attribute VB_Name = "F102010com"
Option Explicit

Type wkSyukaRec_tag
    JGYOBA(0 To 7)              As Byte             '事業場
    DATA_KBN(0 To 0)            As Byte             'データ区分
    TORI_KBN(0 To 1)            As Byte             '取引区分
    ID_NO(0 To 11)              As Byte             'ID-NO
    KAIKEI_JGYOBA(0 To 7)       As Byte             '会計用事業場ｺｰﾄﾞ
    SHISAN_JGYOBA(0 To 7)       As Byte             '資産管理事業場ｺｰﾄﾞ
    HIN_NO(0 To 19)             As Byte             '品目番号
    DEN_NO(0 To 9)              As Byte             '伝票番号
    SURYO(0 To 6)               As Byte             '出庫数量
    MUKE_CODE(0 To 7)           As Byte             '出庫先
    SYUKO_SYUSI(0 To 7)         As Byte             '出庫収支
    SHISAN_SYUSI(0 To 7)        As Byte             '資産管理用在庫収支ｺｰﾄﾞ
    HOJYO_SYUSI(0 To 7)         As Byte             '補助在庫収支ｺｰﾄﾞ
    SYUKO_YMD(0 To 7)           As Byte             '出庫日付
    TANKA(0 To 9)               As Byte             '単価
    ODER_NO(0 To 11)            As Byte             'オーダー番号
    ITEM_NO(0 To 4)             As Byte             'アイテム番号
    ODER_NO_R(0 To 4)           As Byte             'オーダー略号
    KOSO_KEITAI(0 To 13)        As Byte             '個装形態       10-->14 2011.10.31
    SYUKA_YMD(0 To 7)           As Byte             '出荷日
    TANABAN1(0 To 9)            As Byte             '棚番１
    TANABAN2(0 To 9)            As Byte             '棚番２
    TANABAN3(0 To 9)            As Byte             '棚番３
    MUKE_NAME(0 To 23)          As Byte             '出庫先名称
    CYU_KBN(0 To 0)             As Byte             '注文区分
    CYU_KBN_NAME(0 To 39)       As Byte             '注文区分名称
    ORIGIN1(0 To 9)             As Byte             '原産国１
    ORIGIN2(0 To 9)             As Byte             '原産国２
    BIKOU2(0 To 39)             As Byte             '備考２
    HAN_KBN(0 To 0)             As Byte             '販売区分
    CHOKU_KBN(0 To 0)           As Byte             '直送区分
    UNIT_ID_NO(0 To 11)         As Byte             'ﾕﾆｯﾄ修理ID-NO
    ZAIKO_HIKIATE(0 To 2)       As Byte             '在庫引当順序
    GOKON_KANRI_NO(0 To 7)      As Byte             '合梱管理番号
    JYUCHU_ZAN(0 To 6)          As Byte             '受注残数量
    KYOKYU_KBN(0 To 0)          As Byte             '供給区分
    SHOHIN_SYUSI(0 To 7)        As Byte             '商品化納入先収支
    S_SHISAN_SYUSI(0 To 7)      As Byte             '商品化納品資産管理収支ｺｰﾄﾞ
    S_HOJYO_SYUSI(0 To 7)       As Byte             '商品化納品補助収支ｺｰﾄﾞ
    BIKOU1(0 To 39)             As Byte             '備考１
    CHOHA_KBN(0 To 0)           As Byte             '帳端区分
    JYU_HIN_NO(0 To 39)         As Byte             '受注品目番号
    HIN_NAME(0 To 39)           As Byte             '品名
    HIN_CHANGE_KBN(0 To 0)      As Byte             '品番変更区分
    MODULE_EXCHANGE(0 To 0)     As Byte             'モジュール交換区分
    ZAIKO_SYUSI(0 To 7)         As Byte             '残在庫まとめ在庫収支コード
    ZAN_SHISAN_SYUSI(0 To 7)    As Byte             '残在庫まとめ資産管理収支ｺｰﾄﾞ
    ZAN_HOJYO_SYUSI(0 To 7)     As Byte             '残在庫まとめ補助収支ｺｰﾄﾞ
    NOUKI_YMD(0 To 7)           As Byte             '指定納期
    SERVICE_KANRI_NO(0 To 8)    As Byte             'サービス会社管理番号
    KISHU_CODE(0 To 2)          As Byte             '機種品目コード
    ENVIRONMENT_KBN(0 To 0)     As Byte             '環境規格部品区分
    SS_CODE(0 To 7)             As Byte             '直送先コード
    KEPIN_KAIJYO(0 To 0)        As Byte             '欠品解消区分
'    FILLER(0 To 3)              As Byte
    CRLF(0 To 1)                As Byte             'CRLF
End Type

Public RYOHEN      As String * 2       '良品返品の要因 2009.07.10


Public Const WEL_MAEGARI_TANA_S_OSAKA$ = "H2"       '「WEL 資材前借入庫」の要因 2016.05.30

