Attribute VB_Name = "F1100801bas"
Option Explicit

Type INREC_Tag
    JGYOBA(0 To 7)                    As Byte     '事業場コード
    SISAN_JGYOBA(0 To 7)    As Byte     '資産管理事業場コード
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    KISYU_HIN(0 To 2)       As Byte     '代表機種品目コード
    HINMOKU_CD(0 To 2)      As Byte     '品目コード
    SOKO_CD(0 To 1)         As Byte     '倉庫コード
    KOSO_CD(0 To 9)         As Byte     '個装形態コード
    BUHIN_SIZ(0 To 0)       As Byte     '部品サイズ区分
    KONPO_SAISU(0 To 13)    As Byte     '部品梱包才数
    LABEL_HAKKO(0 To 0)     As Byte     '適用機種ラベル発行区分
    KOBAI_TANTO(0 To 4)     As Byte     '購買担当者コード
    UNIT_BUHIN(0 To 0)      As Byte     'ユニット部品区分
    NAI_BUHIN(0 To 0)       As Byte     '国内供給部品区分
    GAI_BUHIN(0 To 0)       As Byte     '海外供給部品区分
    HIN_BETU_NM(0 To 19)    As Byte     '品目別名
    HIN_NAME(0 To 19)       As Byte     '品目名
    U_TANKA2(0 To 9)        As Byte     '売上単価２
    U_TANKA3(0 To 9)        As Byte     '売上単価３
    U_TANKA4(0 To 9)        As Byte     '売上単価４
    LOC_NO1(0 To 9)         As Byte     'ﾛｹｰｼｮﾝ番号１
    LOC_NO2(0 To 9)         As Byte     'ﾛｹｰｼｮﾝ番号２
    LOC_NO3(0 To 9)         As Byte     'ﾛｹｰｼｮﾝ番号３
    HIN_NAI(0 To 19)        As Byte     '工場品目番号（内部品番）
    GENSANKOKU(0 To 9)      As Byte     '現物表示原産国名
    HYO_TANKA(0 To 9)       As Byte     '標準単価
    FILLER(0 To 3)          As Byte     'FILLER
'    JGYOBA(0 To 7)          As Byte     '事業場コード
'    SISAN_JGYOBA(0 To 7)    As Byte     '資産管理事業場コード
'    HIN_GAI(0 To 19)        As Byte     '品番（外部）
'    KISYU_HIN(0 To 2)       As Byte     '代表機種品目コード
'    HINMOKU_CD(0 To 2)      As Byte     '品目コード
'    SOKO_CD(0 To 1)         As Byte     '倉庫コード
'    KOSO_CD(0 To 9)         As Byte     '個装形態コード
'    BUHIN_SIZ(0 To 0)       As Byte     '部品サイズ区分
'    KONPO_SAISU(0 To 13)    As Byte     '部品梱包才数
'    LABEL_HAKKO(0 To 0)     As Byte     '適用機種ラベル発行区分
'    KOBAI_TANTO(0 To 4)     As Byte     '購買担当者コード
'    UNIT_BUHIN(0 To 0)      As Byte     'ユニット部品区分
'    NAI_BUHIN(0 To 0)       As Byte     '国内供給部品区分
'    GAI_BUHIN(0 To 0)       As Byte     '海外供給部品区分
'    HIN_BETU_NM(0 To 19)    As Byte     '品目別名
'    HIN_NAME(0 To 19)       As Byte     '品目名
'    U_TANKA2(0 To 9)        As Byte     '売上単価２
'    U_TANKA3(0 To 9)        As Byte     '売上単価３
'    U_TANKA4(0 To 9)        As Byte     '売上単価４
'    LOC_NO1(0 To 9)         As Byte     'ﾛｹｰｼｮﾝ番号１
'    LOC_NO2(0 To 9)         As Byte     'ﾛｹｰｼｮﾝ番号２
'    LOC_NO3(0 To 9)         As Byte     'ﾛｹｰｼｮﾝ番号３
'    HIN_NAI(0 To 19)        As Byte     '工場品目番号（内部品番）
'    GENSANKOKU(0 To 9)      As Byte     '現物表示原産国名
'    HYO_TANKA(0 To 9)       As Byte     '標準単価
'    FILLER(0 To 3)          As Byte     'FILLER
End Type

'データ・バッファ
Public INREC    As INREC_Tag

