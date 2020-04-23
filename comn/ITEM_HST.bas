Attribute VB_Name = "ITEM_HST"
Option Explicit
'********************************************************************
'*
'*              品目単価変更履歴  ファイル定義
'*
'*          CREATE 2008.07.19
'********************************************************************
'ファイルＩＤ
Public Const ITEM_HST_ID$ = "ITEM_HST"

'ページサイズ
Public Const ITEM_HST_PG_SIZ% = 4096

'ポジション・ブロック
Public ITEM_HST_POS         As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************


Private Type SHIIRE_TBL_Tag         '仕入情報定義用のﾃｰﾌﾞﾙ
    CODE(0 To 4)            As Byte     'ｺｰﾄﾞ
    TANKA(0 To 10)          As Byte     '単価 9(8)V99
    TANKA_DT(0 To 7)        As Byte     '単価設定日
    LOT(0 To 7)             As Byte     'ﾛｯﾄ数
    LEAD_TIME(0 To 2)       As Byte     'ﾘｰﾄﾞﾀｲﾑ
    LAST_ORDER_DT(0 To 7)   As Byte     '前回注文日
    LAST_ORDER_QTY(0 To 10)  As Byte    '前回注文数
End Type

Private Type BEF_KOUTEI_tag
    BEF_KOUTEI(0 To 5)          As Byte     '前工程 2008.09.19
End Type


Private Type MAIN_KOUTEI_tag
    MAIN_KOUTEI(0 To 5)         As Byte     '作業工程 2008.09.19
End Type

Private Type AFT_KOUTEI_tag
    AFT_KOUTEI(0 To 5)          As Byte     '後工程 2008.09.19
End Type


'レコード定義
Type ITEM_HSTREC_Tag
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    '2005.11.15 桁数変更 13---> 20
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
    '2005.11.15 桁数変更 25---> 40
    HIN_NAME(0 To 39)           As Byte     '品名
    ST_SET_DT(0 To 7)           As Byte     '標準倉庫設定日付
    ST_SOKO(0 To 1)             As Byte     '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)             As Byte     '             列
    ST_REN(0 To 1)              As Byte     '             連
    ST_DAN(0 To 1)              As Byte     '             段
    BEF_SOKO(0 To 1)            As Byte     '前回入庫倉庫 倉庫
    BEF_RETU(0 To 1)            As Byte     '             列
    BEF_REN(0 To 1)             As Byte     '             連
    BEF_DAN(0 To 1)             As Byte     '             段
    LAST_NYU_DT(0 To 7)         As Byte     '最終入庫日付
    LAST_SYU_DT(0 To 7)         As Byte     '最終出庫日付
    '2005.11.15 桁数変更 13---> 20
    HIN_NAI(0 To 19)            As Byte     '品番（内部）
    BIKOU_SOKO(0 To 1)          As Byte     '備考 ホスト倉庫
    BIKOU_TANA(0 To 7)          As Byte     '備考 ホスト棚番
    '未使用のため削除 2005.11.15 SIZAI_CD(0 To 4)    As Byte     '資材コード
    HOJYU_P(0 To 7)             As Byte     '補充点（危険在庫）
    AVE_SYUKA(0 To 7)           As Byte     '月平均出荷数
    SAMPLE_QTY(0 To 0)          As Byte     'サンプル数
    LAST_INP_DT(0 To 7)         As Byte     '最終入荷日付
'*------------------------------------------ 2001.02.15 追加 ▽
    '未使用のため削除 2005.11.15 LOCK_F(0 To 0)      As Byte     '排他フラグ
    '未使用のため削除 2005.11.15 WEL_ID(0 To 2)      As Byte     '使用子機ID
    '未使用のため削除 2005.11.15 PRG_ID(0 To 7)      As Byte     '使用中プログラム
'*------------------------------------------ 2001.02.15 追加 △
    LAST_CHK_DT(0 To 7)         As Byte     '最終照合日付2001.06.12
    LAST_CHK_QTY(0 To 7)        As Byte     '最終照合時在庫数2001.06.12
    '未使用のため削除 2005.11.15 MOTO_JIGYOBU(0 To 0) As Byte    '元事事業部     '未使用2004.02
    BIKOU(0 To 14)              As Byte     '印刷備考
    IRI_QTY(0 To 7)             As Byte     '印刷入り数

    '2005.11.15 桁数変更 13---> 20
    JAN_CODE(0 To 19)           As Byte     'Janコード      2004.02
    '2005.11.15 桁数変更 13---> 20
    HIN_CHANGE(0 To 19)         As Byte     '品番読み替え   2004.02
    GOODS_KBN(0 To 0)           As Byte     '商品化有無     2004.02
    PACKING_NO(0 To 3)          As Byte     '個装箱№       2004.02
    RANK(0 To 2)                As Byte     '現在ランク     2004.06
    NEW_RANK(0 To 2)            As Byte     '現在ランク     2004.06
    GLICS1_TANA(0 To 9)         As Byte     'グリックス棚番１   2005.05
    GLICS2_TANA(0 To 9)         As Byte     'グリックス棚番２   2005.05
    GLICS3_TANA(0 To 9)         As Byte     'グリックス棚番３   2005.05
'*------------------------------------------ 2005.11.15 追加(業務管理項目) ▽
    G_SHIIRE_KBN(0 To 1)        As Byte     '業務管理　 仕入区分
    G_HANBAI_KBN(0 To 1)        As Byte     '           販売区分
    G_SYUSHI(0 To 2)            As Byte     '           収支単位
    G_KUMITATE(0 To 0)          As Byte     '           組立製品
    G_ST_URITAN(0 To 10)        As Byte     '           標準粗利売価単価　9(8)V99
    G_ST_URITAN_DT(0 To 7)      As Byte     '           標準粗利売価設定日
    G_ST_SHITAN(0 To 10)        As Byte     '           標準粗利原価単価  9(8)V99
    G_ST_SHITAN_DT(0 To 7)      As Byte     '           標準粗利原価設定日
                                            '           仕入先情報
    G_SHIIRE_TBL(0 To 2)        As SHIIRE_TBL_Tag
    G_ZEN_ZAIKO_KIN(0 To 10)    As Byte     '           前月在庫金額
    G_SHIZAI_KBN(0 To 0)        As Byte     '           資材区分
    G_LABEL_NON(0 To 0)         As Byte     '           ﾗﾍﾞﾙ貼り付け計上なし
'*------------------------------------------ 2005.11.15 追加(業務管理項目) △

'*------------------------------------------ 2005.11.15 追加(商品ﾗﾍﾞﾙ項目) ▽
    L_HIN_NAME_E(0 To 29)       As Byte     '商品ﾗﾍﾞﾙ   品名
    L_BIKOU(0 To 19)            As Byte     '           備考
    L_KAISHA_CODE(0 To 1)       As Byte     '           会社コード
    L_KISHU1(0 To 24)           As Byte     '           機種(1)
    xL_KISHU2(0 To 39)          As Byte     '           機種(2) 未使用 2006.01.24
    L_KISHU3(0 To 149)          As Byte     '           機種(3)(→適用機種備考)
    L_PAPER(0 To 0)             As Byte     '           紙
    L_PLASTIC(0 To 0)           As Byte     '           プラスチック
    L_URIKIN1(0 To 9)           As Byte     '           価格(1)
    L_URIKIN2(0 To 9)           As Byte     '           価格(2)
    L_URIKIN3(0 To 9)           As Byte     '           価格(3)
    L_LABEL(0 To 0)             As Byte     '           適用機種ﾗﾍﾞﾙ
    L_MAISU(0 To 0)             As Byte     '           枚数ﾗﾍﾞﾙ
    L_KISHU_BIKOU(0 To 449)     As Byte     '           適用機種備考(→機種（３）)
    L_SAGYO_SHIJI(0 To 449)     As Byte     '           作業指示
    L_BIKOU3(0 To 4)            As Byte     '           備考３
    L_JGYOBU_CODE(0 To 1)       As Byte     '           事業部コード
    L_IRI_QTY(0 To 7)           As Byte     '           入り数
    L_TANA1(0 To 19)            As Byte     '           棚番(1)
    L_TANA2(0 To 19)            As Byte     '           棚番(2)
'*------------------------------------------ 2005.11.15 追加(商品ﾗﾍﾞﾙ項目) △
    S_TANTO(0 To 1)             As Byte     '収単／担当者コード
    ZAIKO_F(0 To 0)             As Byte     '在庫管理対象有無 1:対象 0:対象外

    L_KISHU2(0 To 51)           As Byte     '           機種(2)

    G_ZEN_ZAIKO_QTY(0 To 7)     As Byte     '           前月在庫数量
    G_LAST_SYUKA_QTY(0 To 7)    As Byte     '           最終出荷数

    G_S2_ZAI_QTY(0 To 7)        As Byte     'GLICS在庫(S2) 袋井用
    G_P2_ZAI_QTY(0 To 7)        As Byte     'GLICS在庫(P2) 袋井用


    K_KEITAI(0 To 9)            As Byte     '個装形態


    UNIT_BUHIN(0 To 0)          As Byte     'ﾕﾆｯﾄ部品区分       2006.07.28
    NAI_BUHIN(0 To 0)           As Byte     '国内供給部品区分   2006.07.28
    GAI_BUHIN(0 To 0)           As Byte     '海外供給部品区分   2006.07.28
    HYO_TANKA(0 To 9)           As Byte     '標準単価   2006.07.28

    LAST_CODE(0 To 4)           As Byte     '最終仕入先コード   2007.05.29
    LAST_TANKA(0 To 10)         As Byte     '最終仕入単価       2007.05.29

    MAKER_CODE(0 To 7)          As Byte     'ﾒｰｶｰｺｰﾄﾞ           2007.06.06
    MAKER_NAME(0 To 39)         As Byte     'ﾒｰｶｰ名称           2007.06.06


    L_MARK(0 To 0)              As Byte     '再梱包ﾏｰｸ          2007.11.08

    xSAI_SU(0 To 3)              As Byte     '才数               2008.02.14

    D_KEISHIKI(0 To 19)         As Byte     '形式               2008.02.14
    D_MATERIAL(0 To 19)         As Byte     '材質               2008.02.14
    D_THICKNESS(0 To 9)         As Byte     'ﾀﾞﾝﾎﾞｰﾙ　厚さ      2008.02.14


    D_SIZE_W(0 To 7)            As Byte     'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（W）   2008.02.14
    D_SIZE_D(0 To 7)            As Byte     'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（D）   2008.02.14
    D_SIZE_H(0 To 7)            As Byte     'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（H）   2008.02.14

    D_PRINT(0 To 3)             As Byte     '印刷する／しない    2008.02.14

    S_KOUSU(0 To 7)             As Byte     '商品化　工数       2008.02.14

    S_KOUSU_GENKA(0 To 10)      As Byte     '商品化　工数原価   2008.02.14
    S_KOUSU_BAIKA(0 To 10)      As Byte     '商品化　工数売価   2008.02.14
    S_KOUSU_SET_DATE(0 To 7)    As Byte     '商品化　単価設定日 2008.02.14

    S_SHIZAI_GENKA(0 To 10)     As Byte     '商品化　資材原価   2008.02.14
    S_SHIZAI_BAIKA(0 To 10)     As Byte     '商品化　資材売価   2008.02.14
    S_SHIZAI_SET_DATE(0 To 7)   As Byte     '商品化　単価設定日 2008.02.14

    SE_USOU_F(0 To 1)           As Byte     '輸送箱　出力ﾌﾗｸﾞ   2008.02.14

    USE_TAPE_KIND(0 To 19)      As Byte     '使用テープ種類     2008.02.14
    USE_TAPE_LNG(0 To 7)        As Byte     '使用テープ長       2008.02.14

    H_TANA_MAKE(0 To 0)         As Byte     '棚番マーク         2008.04.02

    SE_TANKA_MEMO(0 To 39)      As Byte     '請求単価　メモ     2008.04.15

    xGENSANKOKU(0 To 9)         As Byte     '原産国             2008.06.11-->2009.07.16 未使用

    S_GAISO_TANKA(0 To 10)      As Byte     '外装単価 9(8)V99   2008.06.12
    S_PPSC_KAKO_KOSU(0 To 7)    As Byte     'PPSC加工単価9(8)   2008.06.12
    S_BU_KAKO_KOSU(0 To 7)      As Byte     'BU加工単価9(8)     2008.06.12

    SEI_LOT(0 To 7)             As Byte     '生産ロット         2008.07.07
    SEI_RATE(0 To 6)            As Byte     '分レート           2008.07.07
    SEI_SYU_KON(0 To 5)         As Byte     '集合梱包           2008.07.07

    SEI_TANKA_TANTO(0 To 4)     As Byte     '単価設定担当者     2008.07.09

    SHIMUKE_CODE(0 To 1)        As Byte     '仕向け先           2008.07.09

    SEI_KBN(0 To 0)             As Byte     '請求区分           2008.07.16

    SEI_LABEL_QTY(0 To 1)       As Byte     'ラベル貼り枚数     2008.07.19

    SEI_SZI_CNT(0 To 1)         As Byte     '資材件数     　    2008.08.20追加
    SEI_DKN_CNT(0 To 1)         As Byte     '同梱件数           2008.08.20追加

                                            '前工程             2008.09.19  2011.12.12 未使用とする
    BEF_KOUTEI(0 To 9)          As BEF_KOUTEI_tag
                                            '作業工程           2008.09.19
    MAIN_KOUTEI(0 To 9)         As MAIN_KOUTEI_tag
                                            '後工程             2008.09.19  2011.12.12 未使用とする
    AFT_KOUTEI(0 To 9)          As AFT_KOUTEI_tag

    SE_IO_TANKA_No(0 To 1)      As Byte     '棚区分             200.09.19

    STAT(0 To 0)                As Byte     '状態区分           2009.01.21

    INSP_MESSAGE(0 To 39)       As Byte     '出荷検品ﾒｯｾｰｼﾞ     2009.04.17

    S_SEIKYU_F(0 To 0)          As Byte     '商品化請求ﾌﾗｸﾞ     2009.04.28

    
    
'---------
    
    BEF_S_KOUSU_BAIKA(0 To 10)  As Byte     '商品化　工数売価   2009.06.02
    BEF_S_SHIZAI_BAIKA(0 To 10) As Byte     '商品化　資材売価   2009.06.02
    BEF_S_GAISO_TANKA(0 To 10)  As Byte     '外装単価 9(8)V99   2009.06.02
    BEF_S_PPSC_KAKO_KOSU(0 To 7) As Byte    'PPSC加工単価9(8)   2009.06.02
    BEF_S_BU_KAKO_KOSU(0 To 7)  As Byte     'BU加工単価9(8)     2009.06.02
    
    M_BIKOU(0 To 119)           As Byte     '見積書備考         2009.06.02
    SHIYOU_NO(0 To 9)           As Byte     '仕様書№           2009.06.02
    MITSUMORI_KBN(0 To 0)       As Byte     '見積り区分         2009.06.02
    TANKA_KIRIKAE_DT(0 To 7)    As Byte     '単価切替日付       2009.06.02
    KIRIKAE_KBN(0 To 0)         As Byte     '切替区分           2009.06.02
    
    
'---------
    
    GENSANKOKU(0 To 19)         As Byte     '原産国             '2009.07.16
    
    
    
    PLUS_KOUSU(0 To 5)          As Byte     'プラス分工数       2009.09.17  2011.12.12 未使用とする
    
    
    
    KUTI_SU(0 To 3)             As Byte     '口数               2010.01.18
    KONPOU_F(0 To 0)            As Byte     '梱包区分           2010.01.18
    
    SAI_SU(0 To 4)              As Byte     '才数               2010.01.18
    
    
    
    TORI_GENSANKOKU(0 To 19)    As Byte     '取込み時原産国     2010.07.20
    TORI_GEN_GENSANKOKU(0 To 19) _
                                As Byte     '取込み時原産国表示 2010.07.20
    TORI_SHIIRE_WORK_CENTER(0 To 7) _
                                As Byte     '仕入ﾜｰｸセンター    2010.07.20
        
    
    
    KANKYO_KBN(0 To 2)          As Byte     '環境種類区分       2010.07.27
    KANKYO_KBN_ST(0 To 7)       As Byte     '環境種類区分適用開始 2010.07.27
    KANKYO_KBN_SURYO(0 To 9)    As Byte     '環境種類区分数量   2010.07.27
    
    BEF_L_LABEL(0 To 0)         As Byte     '''''
    
    BEF_1_L_PAPER(0 To 0)       As Byte     '           紙
    BEF_1_L_PLASTIC(0 To 0)     As Byte     '           プラスチック
    BEF_2_L_PAPER(0 To 0)       As Byte     '           紙
    BEF_2_L_PLASTIC(0 To 0)     As Byte     '           プラスチック
    BEF_3_L_PAPER(0 To 0)       As Byte     '           紙
    BEF_3_L_PLASTIC(0 To 0)     As Byte     '           プラスチック
    BEF_4_L_PAPER(0 To 0)       As Byte     '           紙
    BEF_4_L_PLASTIC(0 To 0)     As Byte     '           プラスチック
    BEF_LAST_L_PAPER(0 To 0)    As Byte     '           紙
    BEF_LAST_L_PLASTIC(0 To 0)  As Byte     '           プラスチック
    
    
    BIKOU20(0 To 19)            As Byte     '印刷備考
    
    
    PRT_GENSANKOKU(0 To 0)      As Byte     '原産国印字有無     2010.11.10
    GAISO_IRI_QTY(0 To 7)       As Byte     '外装箱入り数 (9(8)) 2010.11.10
    
    
    GOODS_OUT_F(0 To 0)         As Byte     '「商品化計画」除外ﾌﾗｸﾞ "1":除外    2011.06.30
    
    
    PLN_KOUSU(0 To 10)          As Byte     '「商品化ｼｽﾃﾑ」用標準工数           2011.10.02
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   見積書改造(品名ｶﾃｺﾞﾘｰ対応)  2011.12.12
    G_SPTAN(0 To 10)            As Byte     ' 「請求ｼｽﾃﾑ」特別単価 9(8).99
    
    CATE_ST_KOUTEI(0 To 5)      As Byte     ' 「請求ｼｽﾃﾑ」前後工程（秒）    標準    9(3).99
    CATE_ST_FUKA(0 To 5)        As Byte     ' 「請求ｼｽﾃﾑ」付加工数（秒）    標準    9(3).99
    CATE_ST_JITU1(0 To 5)       As Byte     ' 「請求ｼｽﾃﾑ」 実作業工程（秒） 標準    9(3).99
    CATE_ST_YOYU_RITU(0 To 5)   As Byte     ' 「請求ｼｽﾃﾑ」 余裕率（率）     標準    9(3).99
    CATE_ST_JITU2(0 To 5)       As Byte     ' 「請求ｼｽﾃﾑ」 実作業工程（秒） 標準    9(3).99
    CATE_ST_TOTAL(0 To 5)       As Byte     ' 「請求ｼｽﾃﾑ」 作業時間計（秒） 標準    9(3).99
    CATE_ST_FUN(0 To 5)         As Byte     ' 「請求ｼｽﾃﾑ」 分/個（分/個）   標準    9(3).99
    CATE_ST_FUN_RATE(0 To 6)    As Byte     ' 「請求ｼｽﾃﾑ」 分ﾚｰﾄ（円/分）   標準    9(4).99
    CATE_ST_KOURYO(0 To 12)     As Byte     ' 「請求ｼｽﾃﾑ」 工料＠（円/個）  標準    9(10).99
    
    
    
    
    CATE_AD_KOUTEI(0 To 5)      As Byte     ' 「請求ｼｽﾃﾑ」前後工程（秒）    調整    9(3).99
    CATE_AD_FUKA(0 To 5)        As Byte     ' 「請求ｼｽﾃﾑ」 付加工数（秒）   調整    9(3).99
    CATE_AD_JITU1(0 To 5)       As Byte     ' 「請求ｼｽﾃﾑ」 実作業工程（秒） 調整    9(3).99
    CATE_AD_YOYU_RITU(0 To 5)   As Byte     ' 「請求ｼｽﾃﾑ」 余裕率（率）     調整    9(3).99
    CATE_AD_JITU2(0 To 5)       As Byte     ' 「請求ｼｽﾃﾑ」 実作業工程（秒） 調整    9(3).99
    CATE_AD_TOTAL(0 To 5)       As Byte     ' 「請求ｼｽﾃﾑ」 作業時間計（秒） 調整    9(3).99
    CATE_AD_FUN(0 To 5)         As Byte     ' 「請求ｼｽﾃﾑ」  分/個（分/個）  調整    9(3).99
    CATE_AD_FUN_RATE(0 To 6)    As Byte     ' 「請求ｼｽﾃﾑ」  分ﾚｰﾄ（円/分）  調整    9(4).99
    CATE_AD_KOURYO(0 To 12)     As Byte     ' 「請求ｼｽﾃﾑ」  工料＠（円/個） 調整    9(10).99
    
    CATEGORY_CODE(0 To 7)       As Byte
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   見積書改造(品名ｶﾃｺﾞﾘｰ対応)  2011.12.12
    CS_TANTO_CD(0 To 7)         As Byte     'CS担当ｺｰﾄﾞ 2011.12.12
        
    FILLER(0 To 90)            As Byte      'FILLER   2011.12.12  項目追加によりサイズ変更

    INS_TANTO(0 To 4)           As Byte     '追加　担当者　     2009.01.21
    Ins_DateTime(0 To 13)       As Byte     '追加　日時         2009.01.21

    UPD_TANTO(0 To 4)           As Byte     '更新　担当者　     2005.11.15
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時         2005.11.15

End Type
'データ・バッファ
Public ITEM_HSTREC As ITEM_HSTREC_Tag

'キー定義

Type KEY0_ITEM_HST                  'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type

Type KEY1_ITEM_HST                  'ＫＥＹ１
    TANKA_KIRIKAE_DT(0 To 7)    As Byte     '単価切替日付       2009.06.02
End Type




'キー・データ
Public K0_ITEM_HST      As KEY0_ITEM_HST

Type ITEM_HST_FSpeck
    fs      As BtFileSpeck                 ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                 ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck

    ks3     As BtKeySpeck
End Type

Private ITEM_HST_Speck  As ITEM_HST_FSpeck
Private Function ITEM_HST_Create() As Integer
'********************************************************************
'*
'*              品目単価変更履歴  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ITEM_HST_Create = True
                                            '品目単価変更履歴   フルパス取込み
    sts = GetIni("FILE", ITEM_HST_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_HST]読み込みエラー ")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    ITEM_HST_Speck.fs.recoleng = Len(ITEM_HSTREC)   ' レコード長
    ITEM_HST_Speck.fs.PageSize = ITEM_HST_PG_SIZ    ' ページサイズ
    ITEM_HST_Speck.fs.idexnumb = 2                  ' インデックス数
    ITEM_HST_Speck.fs.fileflag = 0                  ' ファイルフラグ
    ITEM_HST_Speck.fs.reserve = &H0                 ' 予約済み
'-----------------------------------------------
                                                ' キー０
    ITEM_HST_Speck.ks0.keypos = 1                   ' キーポジション
    ITEM_HST_Speck.ks0.keyleng = 1                  ' キー長
                                                    ' キーフラグ
    ITEM_HST_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    ITEM_HST_Speck.ks0.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_HST_Speck.ks0.reserve = &H0                ' 予約済み
                                                
    ITEM_HST_Speck.ks1.keypos = 2                   ' キーポジション
    ITEM_HST_Speck.ks1.keyleng = 1                  ' キー長
                                                    ' キーフラグ
    ITEM_HST_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    ITEM_HST_Speck.ks1.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_HST_Speck.ks1.reserve = &H0                ' 予約済み
                                                
    ITEM_HST_Speck.ks2.keypos = 3                   ' キーポジション
    ITEM_HST_Speck.ks2.keyleng = 20                 ' キー長
                                                    ' キーフラグ
    ITEM_HST_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_HST_Speck.ks2.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_HST_Speck.ks2.reserve = &H0                ' 予約済み
'-----------------------------------------------

'-----------------------------------------------
                                                ' キー１

    ITEM_HST_Speck.ks3.keypos = 2627                   ' キーポジション
    ITEM_HST_Speck.ks3.keyleng = 8                 ' キー長
                                                    ' キーフラグ
    ITEM_HST_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_HST_Speck.ks3.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_HST_Speck.ks3.reserve = &H0                ' 予約済み
'-----------------------------------------------



    sts = BTRV(BtOpCreate, ITEM_HST_POS, ITEM_HST_Speck, Len(ITEM_HST_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "品目単価変更履歴")
        Exit Function
    End If

    ITEM_HST_Create = False

End Function

Public Function ITEM_HST_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              品目単価変更履歴  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    ITEM_HST_Open = True
    
    sts = GetIni("FILE", ITEM_HST_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_HST]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, ITEM_HST_POS, ITEM_HSTREC, Len(ITEM_HSTREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_HST_Create()        '品目単価変更履歴 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_HST_POS, ITEM_HSTREC, Len(ITEM_HSTREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "品目単価変更履歴")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "品目単価変更履歴")
                Exit Function
        End Select
    Loop

    ITEM_HST_Open = False

End Function


