Attribute VB_Name = "ITEM_B"
Option Explicit
'********************************************************************
'*
'*              品目マスタ  ファイル定義
'*
'*          CREATE 2004.02.19
'********************************************************************
'ファイルＩＤ
Public Const ITEM_B_ID$ = "ITEM_B"

'ページサイズ
Public Const ITEM_B_PG_SIZ% = 4096

'ポジション・ブロック
Public ITEM_B_POS         As POSBLK
'=
'====================================================================
'=          レコード初期化プロシージャ     ( Rclr_ITEM_BREC )
'====================================================================
'=
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************


Private Type SHIIRE_TBL_Tag         '仕入情報定義用のﾃｰﾌﾞﾙ
    CODE(0 To 4)                As Byte     'ｺｰﾄﾞ
    TANKA(0 To 10)              As Byte     '単価 9(8)V99
    TANKA_DT(0 To 7)            As Byte     '単価設定日
    LOT(0 To 7)                 As Byte     'ﾛｯﾄ数
    LEAD_TIME(0 To 2)           As Byte     'ﾘｰﾄﾞﾀｲﾑ
    LAST_ORDER_DT(0 To 7)       As Byte     '前回注文日
    LAST_ORDER_QTY(0 To 10)     As Byte     '前回注文数
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
Type ITEM_BREC_Tag
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

                                            '前工程             2008.09.19
    BEF_KOUTEI(0 To 9)          As BEF_KOUTEI_tag
                                            '作業工程           2008.09.19
    MAIN_KOUTEI(0 To 9)         As MAIN_KOUTEI_tag
                                            '後工程             2008.09.19
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
    
    
    
    PLUS_KOUSU(0 To 5)          As Byte     'プラス分工数       2009.09.17
    
    
    
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
    
    
    FILLER(0 To 282)            As Byte     'FILLER             2010.07.27    サイズ変更

    INS_TANTO(0 To 4)           As Byte     '追加　担当者　     2009.01.21
    Ins_DateTime(0 To 13)       As Byte     '追加　日時         2009.01.21

    UPD_TANTO(0 To 4)           As Byte     '更新　担当者　     2005.11.15
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時         2005.11.15

End Type
'データ・バッファ
Public ITEM_BREC As ITEM_BREC_Tag

'キー定義

Type KEY0_ITEM_B            'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.11.15 桁数変更 13---> 20
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type

Type KEY1_ITEM_B            'ＫＥＹ１
    LAST_SYU_DT(0 To 7) As Byte     '最終出庫日付
End Type

Type KEY2_ITEM_B            'ＫＥＹ２
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.11.15 桁数変更 13---> 20
    HIN_NAI(0 To 19)    As Byte     '品番（内部）
End Type

Type KEY3_ITEM_B            'ＫＥＹ３
    JGYOBU(0 To 0)      As Byte     '事業部区分
    ST_SET_DT(0 To 7)   As Byte     '標準倉庫設定日付
End Type


Type KEY4_ITEM_B            'ＫＥＹ４ 2004.02
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.11.15 桁数変更 13---> 20
    JAN_CODE(0 To 19)   As Byte     'Janコード
End Type

Type KEY5_ITEM_B            'ＫＥＹ５ 2004.02
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.11.15 桁数変更 13---> 20
    HIN_CHANGE(0 To 19) As Byte     '品番読み替え
End Type

Type KEY6_ITEM_B            'ＫＥＹ６ 2004.02
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    ST_SOKO(0 To 1)     As Byte     '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)     As Byte     '             列
    ST_REN(0 To 1)      As Byte     '             連
    ST_DAN(0 To 1)      As Byte     '             段
    '2005.11.15 桁数変更 13---> 20
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type



'キー・データ
Public K0_ITEM_B As KEY0_ITEM_B
Public K1_ITEM_B As KEY1_ITEM_B
Public K2_ITEM_B As KEY2_ITEM_B
Public K3_ITEM_B As KEY3_ITEM_B
Public K4_ITEM_B As KEY4_ITEM_B
Public K5_ITEM_B As KEY5_ITEM_B
Public K6_ITEM_B As KEY6_ITEM_B

Type ITEM_B_FSpeck
    fs      As BtFileSpeck                 ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                 ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck
    ks8     As BtKeySpeck
    ks9     As BtKeySpeck
    ks10    As BtKeySpeck
    ks11    As BtKeySpeck
    ks12    As BtKeySpeck
    ks13    As BtKeySpeck
    ks14    As BtKeySpeck
    ks15    As BtKeySpeck
    ks16    As BtKeySpeck
    ks17    As BtKeySpeck
    ks18    As BtKeySpeck
    ks19    As BtKeySpeck
    ks20    As BtKeySpeck
    ks21    As BtKeySpeck
End Type

Private ITEM_B_Speck  As ITEM_B_FSpeck

Private Function ITEM_B_Create() As Integer
'********************************************************************
'*
'*              品目マスタ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ITEM_B_Create = True
                                            '品目マスタフルパス取込み
    sts = GetIni("FILE", ITEM_B_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_B]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    ITEM_B_Speck.fs.recoleng = Len(ITEM_BREC)       ' レコード長
    ITEM_B_Speck.fs.PageSize = ITEM_B_PG_SIZ        ' ページサイズ
    ITEM_B_Speck.fs.idexnumb = 7                  ' インデックス数
    ITEM_B_Speck.fs.fileflag = 0                  ' ファイルフラグ
    ITEM_B_Speck.fs.reserve = &H0                 ' 予約済み
'-----------------------------------------------
                                                ' キー０
    ITEM_B_Speck.ks0.keypos = 1                   ' キーポジション
    ITEM_B_Speck.ks0.keyleng = 1                  ' キー長
    ITEM_B_Speck.ks0.keyflag = BtKfExt + BtKfSeg  ' キーフラグ
    ITEM_B_Speck.ks0.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_B_Speck.ks0.reserve = &H0                ' 予約済み

    ITEM_B_Speck.ks1.keypos = 2                   ' キーポジション
    ITEM_B_Speck.ks1.keyleng = 1                  ' キー長
    ITEM_B_Speck.ks1.keyflag = BtKfExt + BtKfSeg  ' キーフラグ
    ITEM_B_Speck.ks1.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_B_Speck.ks1.reserve = &H0                ' 予約済み

    ITEM_B_Speck.ks2.keypos = 3                   ' キーポジション
    ITEM_B_Speck.ks2.keyleng = 20                 ' キー長
    ITEM_B_Speck.ks2.keyflag = BtKfExt            ' キーフラグ
    ITEM_B_Speck.ks2.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_B_Speck.ks2.reserve = &H0                ' 予約済み
'-----------------------------------------------
                                                ' キー１
    ITEM_B_Speck.ks3.keypos = 95                  ' キーポジション
    ITEM_B_Speck.ks3.keyleng = 8                  ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_B_Speck.ks3.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_B_Speck.ks3.reserve = &H0                ' 予約済み
'-----------------------------------------------
                                                ' キー２
    ITEM_B_Speck.ks4.keypos = 1                   ' キーポジション
    ITEM_B_Speck.ks4.keyleng = 1                  ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks4.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_B_Speck.ks4.reserve = &H0                ' 予約済み

    ITEM_B_Speck.ks5.keypos = 2                   ' キーポジション
    ITEM_B_Speck.ks5.keyleng = 1                  ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks5.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_B_Speck.ks5.reserve = &H0                ' 予約済み

    ITEM_B_Speck.ks6.keypos = 103                 ' キーポジション
    ITEM_B_Speck.ks6.keyleng = 20                 ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_B_Speck.ks6.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_B_Speck.ks6.reserve = &H0                ' 予約済み
'-----------------------------------------------
                                                ' キー３
    ITEM_B_Speck.ks7.keypos = 1                   ' キーポジション
    ITEM_B_Speck.ks7.keyleng = 1                  ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks7.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_B_Speck.ks7.reserve = &H0                ' 予約済み

    ITEM_B_Speck.ks8.keypos = 63                  ' キーポジション
    ITEM_B_Speck.ks8.keyleng = 8                  ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_B_Speck.ks8.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_B_Speck.ks8.reserve = &H0                ' 予約済み
'-----------------------------------------------
                                                ' キー４
    ITEM_B_Speck.ks9.keypos = 1                   ' キーポジション
    ITEM_B_Speck.ks9.keyleng = 1                  ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks9.keytype = Chr(BtKtString)    ' キータイプ
    ITEM_B_Speck.ks9.reserve = &H0                ' 予約済み

    ITEM_B_Speck.ks10.keypos = 2                  ' キーポジション
    ITEM_B_Speck.ks10.keyleng = 1                 ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks10.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks10.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks10.reserve = &H0               ' 予約済み

    ITEM_B_Speck.ks11.keypos = 197                ' キーポジション
    ITEM_B_Speck.ks11.keyleng = 20                ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks11.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_B_Speck.ks11.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks11.reserve = &H0               ' 予約済み
'-----------------------------------------------
                                                ' キー５
    ITEM_B_Speck.ks12.keypos = 1                  ' キーポジション
    ITEM_B_Speck.ks12.keyleng = 1                 ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks12.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks12.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks12.reserve = &H0               ' 予約済み

    ITEM_B_Speck.ks13.keypos = 2                  ' キーポジション
    ITEM_B_Speck.ks13.keyleng = 1                 ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks13.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ITEM_B_Speck.ks13.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks13.reserve = &H0               ' 予約済み

    ITEM_B_Speck.ks14.keypos = 217                ' キーポジション
    ITEM_B_Speck.ks14.keyleng = 20                ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks14.keyflag = BtKfExt + BtKfDup + BtKfChg
    ITEM_B_Speck.ks14.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks14.reserve = &H0               ' 予約済み
'-----------------------------------------------
                                                ' キー６
    ITEM_B_Speck.ks15.keypos = 1                  ' キーポジション
    ITEM_B_Speck.ks15.keyleng = 1                 ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks15.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks15.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks15.reserve = &H0               ' 予約済み

    ITEM_B_Speck.ks16.keypos = 2                  ' キーポジション
    ITEM_B_Speck.ks16.keyleng = 1                 ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks16.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks16.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks16.reserve = &H0               ' 予約済み

    ITEM_B_Speck.ks17.keypos = 71                  ' キーポジション
    ITEM_B_Speck.ks17.keyleng = 2                 ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks17.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks17.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks17.reserve = &H0               ' 予約済み

    ITEM_B_Speck.ks18.keypos = 73                 ' キーポジション
    ITEM_B_Speck.ks18.keyleng = 2                 ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks18.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks18.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks18.reserve = &H0               ' 予約済み

    ITEM_B_Speck.ks19.keypos = 75                 ' キーポジション
    ITEM_B_Speck.ks19.keyleng = 2                 ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks19.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks19.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks19.reserve = &H0               ' 予約済み

    ITEM_B_Speck.ks20.keypos = 77                 ' キーポジション
    ITEM_B_Speck.ks20.keyleng = 2                 ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks20.keyflag = BtKfExt + BtKfSeg + BtKfChg
    ITEM_B_Speck.ks20.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks20.reserve = &H0               ' 予約済み

    ITEM_B_Speck.ks21.keypos = 3                  ' キーポジション
    ITEM_B_Speck.ks21.keyleng = 20                ' キー長
                                                ' キーフラグ
    ITEM_B_Speck.ks21.keyflag = BtKfExt + BtKfChg
    ITEM_B_Speck.ks21.keytype = Chr(BtKtString)   ' キータイプ
    ITEM_B_Speck.ks21.reserve = &H0               ' 予約済み
'-----------------------------------------------

    sts = BTRV(BtOpCreate, ITEM_B_POS, ITEM_B_Speck, Len(ITEM_B_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "品目マスタ")
        Exit Function
    End If

    ITEM_B_Create = False

End Function

Public Function ITEM_B_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              品目マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_B_Open = True
                                            '品目マスタフルパス取込み
    sts = GetIni("FILE", ITEM_B_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_B]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_B_Create()        '品目マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "品目マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "品目マスタ")
                Exit Function
        End Select
    Loop

    ITEM_B_Open = False

End Function

Public Sub Rclr_ITEM_BREC()
'********************************************************************
'*
'*              品目マスタ  レコード初期化
'*
'********************************************************************
Dim i       As Long


    Call UniCode_Conv(ITEM_BREC.JGYOBU, "")           '事業部区分
    Call UniCode_Conv(ITEM_BREC.NAIGAI, "")           '国内外
    Call UniCode_Conv(ITEM_BREC.HIN_GAI, "")          '品番（外部）
    Call UniCode_Conv(ITEM_BREC.HIN_NAME, "")         '品名
    Call UniCode_Conv(ITEM_BREC.ST_SET_DT, "")        '標準倉庫設定日付
    Call UniCode_Conv(ITEM_BREC.ST_SOKO, "")          '標準入庫倉庫 倉庫
    Call UniCode_Conv(ITEM_BREC.ST_RETU, "")          '             列
    Call UniCode_Conv(ITEM_BREC.ST_REN, "")           '             連
    Call UniCode_Conv(ITEM_BREC.ST_DAN, "")           '             段
    Call UniCode_Conv(ITEM_BREC.BEF_SOKO, "")         '前回入庫倉庫 倉庫
    
    Call UniCode_Conv(ITEM_BREC.BEF_RETU, "")         '             列
    Call UniCode_Conv(ITEM_BREC.BEF_REN, "")          '             連
    Call UniCode_Conv(ITEM_BREC.BEF_DAN, "")          '             段
    Call UniCode_Conv(ITEM_BREC.LAST_NYU_DT, "")      '最終入庫日付
    Call UniCode_Conv(ITEM_BREC.LAST_SYU_DT, "")      '最終出庫日付
    Call UniCode_Conv(ITEM_BREC.HIN_NAI, "")          '品番（内部）
    Call UniCode_Conv(ITEM_BREC.BIKOU_SOKO, "")       '備考 ホスト倉庫
    Call UniCode_Conv(ITEM_BREC.BIKOU_TANA, "")       '備考 ホスト棚番
    Call UniCode_Conv(ITEM_BREC.LAST_INP_DT, "")      '最終入荷日付
    Call UniCode_Conv(ITEM_BREC.LAST_CHK_DT, "")      '最終照合日付       2001.06.12
    
    Call UniCode_Conv(ITEM_BREC.BIKOU, "")            '印刷備考
    Call UniCode_Conv(ITEM_BREC.JAN_CODE, "")         'Janコード      2004.02
    Call UniCode_Conv(ITEM_BREC.HIN_CHANGE, "")       '品番読み替え   2004.02
    Call UniCode_Conv(ITEM_BREC.GOODS_KBN, GOODS_ON)  '商品化有無     2004.02
    Call UniCode_Conv(ITEM_BREC.PACKING_NO, "")       '個装箱№       2004.02
    Call UniCode_Conv(ITEM_BREC.RANK, "")             '現在ランク     2004.06
    Call UniCode_Conv(ITEM_BREC.NEW_RANK, "")         '現在ランク     2004.06
    Call UniCode_Conv(ITEM_BREC.GLICS1_TANA, "")      'グリックス棚番１   2005.05
    Call UniCode_Conv(ITEM_BREC.GLICS2_TANA, "")      'グリックス棚番２   2005.05
    Call UniCode_Conv(ITEM_BREC.GLICS3_TANA, "")      'グリックス棚番３   2005.05
    
    Call UniCode_Conv(ITEM_BREC.G_SHIIRE_KBN, "")     '業務管理　 仕入区分
    Call UniCode_Conv(ITEM_BREC.G_HANBAI_KBN, "")     '           販売区分
    Call UniCode_Conv(ITEM_BREC.G_SYUSHI, "")         '           収支単位
    Call UniCode_Conv(ITEM_BREC.G_KUMITATE, "")       '           組立製品
    Call UniCode_Conv(ITEM_BREC.G_ST_URITAN_DT, "")   '           標準粗利売価設定日
    Call UniCode_Conv(ITEM_BREC.G_ST_SHITAN_DT, "")   '           標準粗利原価設定日
                                                    '           仕入先情報
    For i = 0 To UBound(ITEM_BREC.G_SHIIRE_TBL)
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).CODE, "")             'ｺｰﾄﾞ
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '単価設定日
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ﾘｰﾄﾞﾀｲﾑ
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    '前回注文日
    Next i
    
    Call UniCode_Conv(ITEM_BREC.G_SHIZAI_KBN, "")     '           資材区分
    Call UniCode_Conv(ITEM_BREC.G_LABEL_NON, "")      '           ﾗﾍﾞﾙ貼り付け計上なし
    
    Call UniCode_Conv(ITEM_BREC.L_HIN_NAME_E, "")     '商品ﾗﾍﾞﾙ   品名
    Call UniCode_Conv(ITEM_BREC.L_BIKOU, "")          '           備考
    Call UniCode_Conv(ITEM_BREC.L_KAISHA_CODE, "")    '           会社コード
    Call UniCode_Conv(ITEM_BREC.L_KISHU1, "")         '           機種(1)
    Call UniCode_Conv(ITEM_BREC.xL_KISHU2, "")        '           機種(2) 未使用 2006.01.24
    Call UniCode_Conv(ITEM_BREC.L_KISHU3, "")         '           機種(3)(→適用機種備考)
    Call UniCode_Conv(ITEM_BREC.L_PAPER, "0")         '           紙
    Call UniCode_Conv(ITEM_BREC.L_PLASTIC, "0")       '           プラスチック
    Call UniCode_Conv(ITEM_BREC.L_LABEL, "0")         '           適用機種ﾗﾍﾞﾙ
    Call UniCode_Conv(ITEM_BREC.L_MAISU, "0")         '           枚数ﾗﾍﾞﾙ
    Call UniCode_Conv(ITEM_BREC.L_KISHU_BIKOU, "")    '           適用機種備考(→機種（３）)
    Call UniCode_Conv(ITEM_BREC.L_SAGYO_SHIJI, "")    '           作業指示
    Call UniCode_Conv(ITEM_BREC.L_BIKOU3, "")         '           備考３
    Call UniCode_Conv(ITEM_BREC.L_JGYOBU_CODE, "")    '           事業部コード
    Call UniCode_Conv(ITEM_BREC.L_TANA1, "")          '           棚番(1)
    Call UniCode_Conv(ITEM_BREC.L_TANA2, "")          '           棚番(2)
    
    Call UniCode_Conv(ITEM_BREC.S_TANTO, "")          '収単／担当者コード
    Call UniCode_Conv(ITEM_BREC.ZAIKO_F, "")          '在庫管理対象有無 1:対象 0:対象外
    Call UniCode_Conv(ITEM_BREC.L_KISHU2, "")         '           機種(2)
    Call UniCode_Conv(ITEM_BREC.K_KEITAI, "")         '個装形態
    Call UniCode_Conv(ITEM_BREC.UNIT_BUHIN, "")       'ﾕﾆｯﾄ部品区分       2006.07.28
    Call UniCode_Conv(ITEM_BREC.NAI_BUHIN, "")        '国内供給部品区分   2006.07.28
    Call UniCode_Conv(ITEM_BREC.GAI_BUHIN, "")        '海外供給部品区分   2006.07.28
    Call UniCode_Conv(ITEM_BREC.LAST_CODE, "")        '最終仕入先コード   2007.05.29
    Call UniCode_Conv(ITEM_BREC.MAKER_CODE, "")       'ﾒｰｶｰｺｰﾄﾞ           2007.06.06
    Call UniCode_Conv(ITEM_BREC.MAKER_NAME, "")       'ﾒｰｶｰ名称           2007.06.06
    
    Call UniCode_Conv(ITEM_BREC.L_MARK, "")           '再梱包ﾏｰｸ          2007.11.08
    Call UniCode_Conv(ITEM_BREC.SAI_SU, "")           '才数               2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_KEISHIKI, "")       '形式               2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_MATERIAL, "")       '材質               2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_THICKNESS, "")      'ﾀﾞﾝﾎﾞｰﾙ　厚さ      2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_SIZE_W, "")         'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（W）   2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_SIZE_D, "")         'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（D）   2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_SIZE_H, "")         'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（H）   2008.02.14
    Call UniCode_Conv(ITEM_BREC.D_PRINT, "")          '印刷する／しない    2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_KOUSU_SET_DATE, "") '商品化　単価設定日 2008.02.14
    
    Call UniCode_Conv(ITEM_BREC.S_SHIZAI_SET_DATE, "") '商品化　単価設定日 2008.02.14
    Call UniCode_Conv(ITEM_BREC.SE_USOU_F, "")        '輸送箱　出力ﾌﾗｸﾞ   2008.02.14
    Call UniCode_Conv(ITEM_BREC.USE_TAPE_KIND, "")    '使用テープ種類     2008.02.14
    Call UniCode_Conv(ITEM_BREC.USE_TAPE_LNG, "")     '使用テープ長       2008.02.14
    Call UniCode_Conv(ITEM_BREC.H_TANA_MAKE, "")      '棚番マーク         2008.04.02
    Call UniCode_Conv(ITEM_BREC.SE_TANKA_MEMO, "")    '請求単価　メモ     2008.04.15
    Call UniCode_Conv(ITEM_BREC.GENSANKOKU, "")       '原産国             2008.06.11
    Call UniCode_Conv(ITEM_BREC.SEI_LOT, "")          '生産ロット         2008.07.07
    Call UniCode_Conv(ITEM_BREC.SEI_SYU_KON, "")      '集合梱包           2008.07.07
    Call UniCode_Conv(ITEM_BREC.SEI_TANKA_TANTO, "")  '単価設定担当者     2008.07.09
    Call UniCode_Conv(ITEM_BREC.SHIMUKE_CODE, "")     '仕向け先           2008.07.09
    Call UniCode_Conv(ITEM_BREC.SEI_KBN, "")          '請求区分           2008.07.16
                                            '前工程             2008.09.19
    For i = 0 To UBound(ITEM_BREC.BEF_KOUTEI)
        Call UniCode_Conv(ITEM_BREC.BEF_KOUTEI(i).BEF_KOUTEI, "")     '前工程 2008.09.19
    Next i
                                            '作業工程           2008.09.19
    For i = 0 To UBound(ITEM_BREC.MAIN_KOUTEI)
        Call UniCode_Conv(ITEM_BREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")   '作業工程 2008.09.19
    Next i
                                            '後工程             2008.09.19
    For i = 0 To UBound(ITEM_BREC.AFT_KOUTEI)
        Call UniCode_Conv(ITEM_BREC.AFT_KOUTEI(i).AFT_KOUTEI, "")     '後工程 2008.09.19
    Next i

    Call UniCode_Conv(ITEM_BREC.SE_IO_TANKA_No, "")   '棚区分             200.09.19
    Call UniCode_Conv(ITEM_BREC.STAT, "")             '状態区分           2009.01.21
    Call UniCode_Conv(ITEM_BREC.INSP_MESSAGE, "")     '出荷検品ﾒｯｾｰｼﾞ     2009.04.17
    Call UniCode_Conv(ITEM_BREC.S_SEIKYU_F, "")       '商品化請求ﾌﾗｸﾞ     2009.04.28
    Call UniCode_Conv(ITEM_BREC.FILLER, "")           'FILLER             2009.04.28サイズ変更
    
    Call UniCode_Conv(ITEM_BREC.INS_TANTO, "")        '追加　担当者　     2009.01.21
    Call UniCode_Conv(ITEM_BREC.Ins_DateTime, "")     '追加　日時         2009.01.21
    Call UniCode_Conv(ITEM_BREC.UPD_TANTO, "")        '更新　担当者　     2005.11.15
    Call UniCode_Conv(ITEM_BREC.UPD_DATETIME, "")     '更新　日時         2005.11.15

'-------------------------------------------------------------------------------------------
'               ０クリア項目

                                                    '補充点（危険在庫）
    Call UniCode_Conv(ITEM_BREC.HOJYU_P, String(UBound(ITEM_BREC.HOJYU_P) + 1, "0"))
                                                    '月平均出荷数
    Call UniCode_Conv(ITEM_BREC.AVE_SYUKA, String(UBound(ITEM_BREC.AVE_SYUKA) + 1, "0"))
                                                    'サンプル数
    Call UniCode_Conv(ITEM_BREC.SAMPLE_QTY, String(UBound(ITEM_BREC.SAMPLE_QTY) + 1, "0"))
                                                    '最終照合時在庫数   2001.06.12
    Call UniCode_Conv(ITEM_BREC.LAST_CHK_QTY, String(UBound(ITEM_BREC.LAST_CHK_QTY) + 1, "0"))
                                                    '印刷入り数
    Call UniCode_Conv(ITEM_BREC.IRI_QTY, String(UBound(ITEM_BREC.IRI_QTY) + 1, "0"))
                                                    '           標準粗利売価単価　9(8)V99
    Call UniCode_Conv(ITEM_BREC.G_ST_URITAN, String(UBound(ITEM_BREC.G_ST_URITAN) + 1, "0"))
                                                    '           標準粗利原価単価  9(8)V99
    Call UniCode_Conv(ITEM_BREC.G_ST_SHITAN, String(UBound(ITEM_BREC.G_ST_SHITAN) + 1, "0"))

    For i = 0 To UBound(ITEM_BREC.G_SHIIRE_TBL)
                                                                        '単価 9(8)V99
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).TANKA, _
                   String(UBound(ITEM_BREC.G_SHIIRE_TBL(i).TANKA) + 1, "0"))
                                                                        'ﾛｯﾄ数
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LOT, _
                   String(UBound(ITEM_BREC.G_SHIIRE_TBL(i).LOT) + 1, "0"))
                                                                        '前回注文数
        Call UniCode_Conv(ITEM_BREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, _
                   String(UBound(ITEM_BREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY) + 1, "0"))
    Next i
                                                    '           前月在庫金額
    Call UniCode_Conv(ITEM_BREC.G_ZEN_ZAIKO_KIN, String(UBound(ITEM_BREC.G_ZEN_ZAIKO_KIN) + 1, "0"))
                                                    '           価格(1)
    Call UniCode_Conv(ITEM_BREC.L_URIKIN1, String(UBound(ITEM_BREC.L_URIKIN1) + 1, "0"))
                                                    '           価格(2)
    Call UniCode_Conv(ITEM_BREC.L_URIKIN2, String(UBound(ITEM_BREC.L_URIKIN2) + 1, "0"))
                                                    '           価格(3)
    Call UniCode_Conv(ITEM_BREC.L_URIKIN3, String(UBound(ITEM_BREC.L_URIKIN3) + 1, "0"))
                                                    '           入り数
    Call UniCode_Conv(ITEM_BREC.L_IRI_QTY, String(UBound(ITEM_BREC.L_IRI_QTY) + 1, "0"))
                                                    '           前月在庫数量
    Call UniCode_Conv(ITEM_BREC.G_ZEN_ZAIKO_QTY, String(UBound(ITEM_BREC.G_ZEN_ZAIKO_QTY) + 1, "0"))
                                                    '           最終出荷数
    Call UniCode_Conv(ITEM_BREC.G_LAST_SYUKA_QTY, String(UBound(ITEM_BREC.G_LAST_SYUKA_QTY) + 1, "0"))
                                                    'GLICS在庫(S2) 袋井用
    Call UniCode_Conv(ITEM_BREC.G_S2_ZAI_QTY, String(UBound(ITEM_BREC.G_S2_ZAI_QTY) + 1, "0"))
                                                    'GLICS在庫(P2) 袋井用
    Call UniCode_Conv(ITEM_BREC.G_P2_ZAI_QTY, String(UBound(ITEM_BREC.G_P2_ZAI_QTY) + 1, "0"))
                                                    '標準単価   2006.07.28
    Call UniCode_Conv(ITEM_BREC.HYO_TANKA, String(UBound(ITEM_BREC.HYO_TANKA) + 1, "0"))
                                                    '最終仕入単価       2007.05.29
    Call UniCode_Conv(ITEM_BREC.LAST_TANKA, String(UBound(ITEM_BREC.LAST_TANKA) + 1, "0"))
                                                    '商品化　工数       2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_KOUSU, String(UBound(ITEM_BREC.S_KOUSU) + 1, "0"))
                                                    '商品化　工数原価   2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_KOUSU_GENKA, String(UBound(ITEM_BREC.S_KOUSU_GENKA) + 1, "0"))
                                                    '商品化　工数売価   2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_KOUSU_BAIKA, String(UBound(ITEM_BREC.S_KOUSU_BAIKA) + 1, "0"))
                                                    '商品化　資材原価   2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_SHIZAI_GENKA, String(UBound(ITEM_BREC.S_SHIZAI_GENKA) + 1, "0"))
                                                    '商品化　資材売価   2008.02.14
    Call UniCode_Conv(ITEM_BREC.S_SHIZAI_BAIKA, String(UBound(ITEM_BREC.S_SHIZAI_BAIKA) + 1, "0"))

                                                    '外装単価 9(8)V99   2008.06.12
    Call UniCode_Conv(ITEM_BREC.S_GAISO_TANKA, String(UBound(ITEM_BREC.S_GAISO_TANKA) + 1, "0"))
                                                    'PPSC加工単価9(8)   2008.06.12
    Call UniCode_Conv(ITEM_BREC.S_PPSC_KAKO_KOSU, String(UBound(ITEM_BREC.S_PPSC_KAKO_KOSU) + 1, "0"))
                                                    'BU加工単価9(8)     2008.06.12
    Call UniCode_Conv(ITEM_BREC.S_BU_KAKO_KOSU, String(UBound(ITEM_BREC.S_BU_KAKO_KOSU) + 1, "0"))
    
                                                    '分レート           2008.07.07
    Call UniCode_Conv(ITEM_BREC.SEI_RATE, String(UBound(ITEM_BREC.SEI_RATE) + 1, "0"))
    
                                                    'ラベル貼り枚数     2008.07.19
    Call UniCode_Conv(ITEM_BREC.SEI_LABEL_QTY, String(UBound(ITEM_BREC.SEI_LABEL_QTY) + 1, "0"))

                                                    '資材件数     　    2008.08.20追加
    Call UniCode_Conv(ITEM_BREC.SEI_SZI_CNT, String(UBound(ITEM_BREC.SEI_SZI_CNT) + 1, "0"))
                                                    '同梱件数           2008.08.20追加
    Call UniCode_Conv(ITEM_BREC.SEI_DKN_CNT, String(UBound(ITEM_BREC.SEI_DKN_CNT) + 1, "0"))


'-------------------------------------------------------------------------------------------
'               2009.06.02
                                                    '商品化　工数売価
    Call UniCode_Conv(ITEM_BREC.BEF_S_KOUSU_BAIKA, String(UBound(ITEM_BREC.BEF_S_KOUSU_BAIKA) + 1, "0"))
                                                    '商品化　資材売価
    Call UniCode_Conv(ITEM_BREC.BEF_S_SHIZAI_BAIKA, String(UBound(ITEM_BREC.BEF_S_SHIZAI_BAIKA) + 1, "0"))
                                                    '外装単価
    Call UniCode_Conv(ITEM_BREC.BEF_S_GAISO_TANKA, String(UBound(ITEM_BREC.BEF_S_GAISO_TANKA) + 1, "0"))
                                                    'PPSC加工単価
    Call UniCode_Conv(ITEM_BREC.BEF_S_PPSC_KAKO_KOSU, String(UBound(ITEM_BREC.BEF_S_PPSC_KAKO_KOSU) + 1, "0"))
                                                    'BU加工単価
    Call UniCode_Conv(ITEM_BREC.BEF_S_BU_KAKO_KOSU, String(UBound(ITEM_BREC.BEF_S_BU_KAKO_KOSU) + 1, "0"))
    
    Call UniCode_Conv(ITEM_BREC.M_BIKOU, "")              '見積書備考
    
    Call UniCode_Conv(ITEM_BREC.SHIYOU_NO, "")            '仕様書№
    
    Call UniCode_Conv(ITEM_BREC.MITSUMORI_KBN, "")        '見積り区分
    
    Call UniCode_Conv(ITEM_BREC.TANKA_KIRIKAE_DT, "")    '単価切替日付
    
    Call UniCode_Conv(ITEM_BREC.KIRIKAE_KBN, "")          '切替区分
    
'               2009.06.02
'-------------------------------------------------------------------------------------------

End Sub
