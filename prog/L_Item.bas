Attribute VB_Name = "L_ITEM"
Option Explicit
'********************************************************************
'*
'*              品目マスタ  ファイル定義
'*
'*          CREATE 2004.02.19
'********************************************************************
'ファイルＩＤ
Public Const L_ITEM_ID$ = "L_ITEM"

'ページサイズ
Public Const L_ITEM_PG_SIZ% = 4096

'ポジション・ブロック
Public L_ITEM_POS         As POSBLK
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


'レコード定義
Type L_ITEMREC_Tag
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.11.15 桁数変更 13---> 20
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    '2005.11.15 桁数変更 25---> 40
    HIN_NAME(0 To 39)   As Byte     '品名
    ST_SET_DT(0 To 7)   As Byte     '標準倉庫設定日付
    ST_SOKO(0 To 1)     As Byte     '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)     As Byte     '             列
    ST_REN(0 To 1)      As Byte     '             連
    ST_DAN(0 To 1)      As Byte     '             段
    BEF_SOKO(0 To 1)    As Byte     '前回入庫倉庫 倉庫
    BEF_RETU(0 To 1)    As Byte     '             列
    BEF_REN(0 To 1)     As Byte     '             連
    BEF_DAN(0 To 1)     As Byte     '             段
    LAST_NYU_DT(0 To 7) As Byte     '最終入庫日付
    LAST_SYU_DT(0 To 7) As Byte     '最終出庫日付
    '2005.11.15 桁数変更 13---> 20
    HIN_NAI(0 To 19)    As Byte     '品番（内部）
    BIKOU_SOKO(0 To 1)  As Byte     '備考 ホスト倉庫
    BIKOU_TANA(0 To 7)  As Byte     '備考 ホスト棚番
    '未使用のため削除 2005.11.15 SIZAI_CD(0 To 4)    As Byte     '資材コード
    HOJYU_P(0 To 7)     As Byte     '補充点（危険在庫）
    AVE_SYUKA(0 To 7)   As Byte     '月平均出荷数
    SAMPLE_QTY(0 To 0)  As Byte     'サンプル数
    LAST_INP_DT(0 To 7) As Byte     '最終入荷日付
'*------------------------------------------ 2001.02.15 追加 ▽
    '未使用のため削除 2005.11.15 LOCK_F(0 To 0)      As Byte     '排他フラグ
    '未使用のため削除 2005.11.15 WEL_ID(0 To 2)      As Byte     '使用子機ID
    '未使用のため削除 2005.11.15 PRG_ID(0 To 7)      As Byte     '使用中プログラム
'*------------------------------------------ 2001.02.15 追加 △
    LAST_CHK_DT(0 To 7) As Byte     '最終照合日付2001.06.12
    LAST_CHK_QTY(0 To 7) As Byte    '最終照合時在庫数2001.06.12
    '未使用のため削除 2005.11.15 MOTO_JIGYOBU(0 To 0) As Byte    '元事事業部     '未使用2004.02
    BIKOU(0 To 14)      As Byte     '印刷備考
    IRI_QTY(0 To 7)     As Byte     '印刷入り数
    
    '2005.11.15 桁数変更 13---> 20
    JAN_CODE(0 To 19)   As Byte     'Janコード      2004.02
    '2005.11.15 桁数変更 13---> 20
    HIN_CHANGE(0 To 19) As Byte     '品番読み替え   2004.02
    GOODS_KBN(0 To 0)   As Byte     '商品化有無     2004.02
    PACKING_NO(0 To 3)  As Byte     '個装箱№       2004.02
    RANK(0 To 2)        As Byte     '現在ランク     2004.06
    NEW_RANK(0 To 2)    As Byte     '現在ランク     2004.06
    GLICS1_TANA(0 To 9) As Byte     'グリックス棚番１   2005.05
    GLICS2_TANA(0 To 9) As Byte     'グリックス棚番２   2005.05
    GLICS3_TANA(0 To 9) As Byte     'グリックス棚番３   2005.05
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
    xL_KISHU2(0 To 39)           As Byte     '           機種(2)
    L_KISHU3(0 To 149)          As Byte     '           機種(3)
    L_PAPER(0 To 0)             As Byte     '           紙
    L_PLASTIC(0 To 0)           As Byte     '           プラスチック
    L_URIKIN1(0 To 9)           As Byte     '           価格(1)
    L_URIKIN2(0 To 9)           As Byte     '           価格(2)
    L_URIKIN3(0 To 9)           As Byte     '           価格(3)
    L_LABEL(0 To 0)             As Byte     '           適用機種ﾗﾍﾞﾙ
    L_MAISU(0 To 0)             As Byte     '           枚数ﾗﾍﾞﾙ
    L_KISHU_BIKOU(0 To 449)     As Byte     '           適用機種備考
    L_SAGYO_SHIJI(0 To 449)     As Byte     '           作業指示
    L_BIKOU3(0 To 4)            As Byte     '           備考３
    L_JGYOBU_CODE(0 To 1)       As Byte     '           事業部コード
    L_IRI_QTY(0 To 7)           As Byte     '           入り数
    L_TANA1(0 To 19)            As Byte     '           棚番(1)
    L_TANA2(0 To 19)            As Byte     '           棚番(2)
'*------------------------------------------ 2005.11.15 追加(商品ﾗﾍﾞﾙ項目) △
    S_TANTO(0 To 1)             As Byte     '収単／担当者コード
    ZAIKO_F(0 To 0)             As Byte     '在庫管理対象有無 0:対象 1:対象外
    
    
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
    
    
    SAI_SU(0 To 3)              As Byte     '才数               2008.02.14
    
    D_KEISHIKI(0 To 19)         As Byte     '形式               2008.02.14
    D_MATERIAL(0 To 19)         As Byte     '材質               2008.02.14
    D_THICKNESS(0 To 9)         As Byte     'ﾀﾞﾝﾎﾞｰﾙ　厚さ      2008.02.14
    
    
    D_SIZE_W(0 To 7)            As Byte     'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（W）   2008.02.14
    D_SIZE_D(0 To 7)            As Byte     'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（D）   2008.02.14
    D_SIZE_H(0 To 7)            As Byte     'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（H）   2008.02.14
        
    D_PRINT(0 To 3)            As Byte      '印刷する／しない   2008.02.14
            
        
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
    
    
    GENSANKOKU(0 To 9)          As Byte     '原産国             2008.06.11
    
    S_GAISO_TANKA(0 To 10)      As Byte     '外装単価 9(8)V99   2008.06.12
    S_PPSC_KAKO_KOSU(0 To 7)    As Byte     'PPSC加工単価9(8)   2008.06.12
    S_BU_KAKO_KOSU(0 To 7)      As Byte     'BU加工単価9(8)   2008.06.12
    
    FILLER(0 To 865)           As Byte     'FILLER
    
    
    
    

    UPD_TANTO(0 To 4)           As Byte     '更新　担当者　 2005.11.15
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時     2005.11.15

End Type
'データ・バッファ
Public L_ITEMREC As L_ITEMREC_Tag

'キー定義

Type KEY0_L_ITEM            'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.11.15 桁数変更 13---> 20
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type

Type KEY1_L_ITEM            'ＫＥＹ１
    LAST_SYU_DT(0 To 7) As Byte     '最終出庫日付
End Type

Type KEY2_L_ITEM            'ＫＥＹ２
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.11.15 桁数変更 13---> 20
    HIN_NAI(0 To 19)    As Byte     '品番（内部）
End Type

Type KEY3_L_ITEM            'ＫＥＹ３
    JGYOBU(0 To 0)      As Byte     '事業部区分
    ST_SET_DT(0 To 7)   As Byte     '標準倉庫設定日付
End Type


Type KEY4_L_ITEM            'ＫＥＹ４ 2004.02
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.11.15 桁数変更 13---> 20
    JAN_CODE(0 To 19)   As Byte     'Janコード
End Type

Type KEY5_L_ITEM            'ＫＥＹ５ 2004.02
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.11.15 桁数変更 13---> 20
    HIN_CHANGE(0 To 19) As Byte     '品番読み替え
End Type

Type KEY6_L_ITEM            'ＫＥＹ６ 2004.02
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
Public K0_L_ITEM As KEY0_L_ITEM
Public K1_L_ITEM As KEY1_L_ITEM
Public K2_L_ITEM As KEY2_L_ITEM
Public K3_L_ITEM As KEY3_L_ITEM
Public K4_L_ITEM As KEY4_L_ITEM
Public K5_L_ITEM As KEY5_L_ITEM
Public K6_L_ITEM As KEY6_L_ITEM

Type L_ITEM_FSpeck
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

Private L_ITEM_Speck  As L_ITEM_FSpeck
Private Function L_ITEM_Create() As Integer
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

    L_ITEM_Create = True
                                            '品目マスタフルパス取込み
    sts = GetIni("FILE", L_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [L_ITEM]読み込みエラー ")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    L_ITEM_Speck.fs.recoleng = Len(L_ITEMREC)   ' レコード長
    L_ITEM_Speck.fs.PageSize = ITEM_PG_SIZ      ' ページサイズ
    L_ITEM_Speck.fs.idexnumb = 7                  ' インデックス数
    L_ITEM_Speck.fs.fileflag = 0                  ' ファイルフラグ
    L_ITEM_Speck.fs.reserve = &H0                 ' 予約済み
'-----------------------------------------------
                                                ' キー０
    L_ITEM_Speck.ks0.keypos = 1                   ' キーポジション
    L_ITEM_Speck.ks0.keyleng = 1                  ' キー長
    L_ITEM_Speck.ks0.keyflag = BtKfExt + BtKfSeg  ' キーフラグ
    L_ITEM_Speck.ks0.keytype = Chr(BtKtString)    ' キータイプ
    L_ITEM_Speck.ks0.reserve = &H0                ' 予約済み
                                                
    L_ITEM_Speck.ks1.keypos = 2                   ' キーポジション
    L_ITEM_Speck.ks1.keyleng = 1                  ' キー長
    L_ITEM_Speck.ks1.keyflag = BtKfExt + BtKfSeg  ' キーフラグ
    L_ITEM_Speck.ks1.keytype = Chr(BtKtString)    ' キータイプ
    L_ITEM_Speck.ks1.reserve = &H0                ' 予約済み
                                                
    L_ITEM_Speck.ks2.keypos = 3                   ' キーポジション
    L_ITEM_Speck.ks2.keyleng = 20                 ' キー長
    L_ITEM_Speck.ks2.keyflag = BtKfExt            ' キーフラグ
    L_ITEM_Speck.ks2.keytype = Chr(BtKtString)    ' キータイプ
    L_ITEM_Speck.ks2.reserve = &H0                ' 予約済み
'-----------------------------------------------
                                                ' キー１
    L_ITEM_Speck.ks3.keypos = 95                  ' キーポジション
    L_ITEM_Speck.ks3.keyleng = 8                  ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    L_ITEM_Speck.ks3.keytype = Chr(BtKtString)    ' キータイプ
    L_ITEM_Speck.ks3.reserve = &H0                ' 予約済み
'-----------------------------------------------
                                                ' キー２
    L_ITEM_Speck.ks4.keypos = 1                   ' キーポジション
    L_ITEM_Speck.ks4.keyleng = 1                  ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks4.keytype = Chr(BtKtString)    ' キータイプ
    L_ITEM_Speck.ks4.reserve = &H0                ' 予約済み
                                                    
    L_ITEM_Speck.ks5.keypos = 2                   ' キーポジション
    L_ITEM_Speck.ks5.keyleng = 1                  ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks5.keytype = Chr(BtKtString)    ' キータイプ
    L_ITEM_Speck.ks5.reserve = &H0                ' 予約済み
                                                
    L_ITEM_Speck.ks6.keypos = 103                 ' キーポジション
    L_ITEM_Speck.ks6.keyleng = 20                 ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg
    L_ITEM_Speck.ks6.keytype = Chr(BtKtString)    ' キータイプ
    L_ITEM_Speck.ks6.reserve = &H0                ' 予約済み
'-----------------------------------------------
                                                ' キー３
    L_ITEM_Speck.ks7.keypos = 1                   ' キーポジション
    L_ITEM_Speck.ks7.keyleng = 1                  ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks7.keytype = Chr(BtKtString)    ' キータイプ
    L_ITEM_Speck.ks7.reserve = &H0                ' 予約済み
                                                
    L_ITEM_Speck.ks8.keypos = 63                  ' キーポジション
    L_ITEM_Speck.ks8.keyleng = 8                  ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg
    L_ITEM_Speck.ks8.keytype = Chr(BtKtString)    ' キータイプ
    L_ITEM_Speck.ks8.reserve = &H0                ' 予約済み
'-----------------------------------------------
                                                ' キー４
    L_ITEM_Speck.ks9.keypos = 1                   ' キーポジション
    L_ITEM_Speck.ks9.keyleng = 1                  ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks9.keytype = Chr(BtKtString)    ' キータイプ
    L_ITEM_Speck.ks9.reserve = &H0                ' 予約済み
                                                
    L_ITEM_Speck.ks10.keypos = 2                  ' キーポジション
    L_ITEM_Speck.ks10.keyleng = 1                 ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks10.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks10.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks10.reserve = &H0               ' 予約済み
                                                
    L_ITEM_Speck.ks11.keypos = 197                ' キーポジション
    L_ITEM_Speck.ks11.keyleng = 20                ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks11.keyflag = BtKfExt + BtKfDup + BtKfChg
    L_ITEM_Speck.ks11.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks11.reserve = &H0               ' 予約済み
'-----------------------------------------------
                                                ' キー５
    L_ITEM_Speck.ks12.keypos = 1                  ' キーポジション
    L_ITEM_Speck.ks12.keyleng = 1                 ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks12.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks12.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks12.reserve = &H0               ' 予約済み
                                                
    L_ITEM_Speck.ks13.keypos = 2                  ' キーポジション
    L_ITEM_Speck.ks13.keyleng = 1                 ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks13.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    L_ITEM_Speck.ks13.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks13.reserve = &H0               ' 予約済み
                                                
    L_ITEM_Speck.ks14.keypos = 217                ' キーポジション
    L_ITEM_Speck.ks14.keyleng = 20                ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks14.keyflag = BtKfExt + BtKfDup + BtKfChg
    L_ITEM_Speck.ks14.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks14.reserve = &H0               ' 予約済み
'-----------------------------------------------
                                                ' キー６
    L_ITEM_Speck.ks15.keypos = 1                  ' キーポジション
    L_ITEM_Speck.ks15.keyleng = 1                 ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks15.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks15.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks15.reserve = &H0               ' 予約済み

    L_ITEM_Speck.ks16.keypos = 2                  ' キーポジション
    L_ITEM_Speck.ks16.keyleng = 1                 ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks16.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks16.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks16.reserve = &H0               ' 予約済み

    L_ITEM_Speck.ks17.keypos = 71                  ' キーポジション
    L_ITEM_Speck.ks17.keyleng = 2                 ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks17.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks17.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks17.reserve = &H0               ' 予約済み

    L_ITEM_Speck.ks18.keypos = 73                 ' キーポジション
    L_ITEM_Speck.ks18.keyleng = 2                 ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks18.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks18.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks18.reserve = &H0               ' 予約済み

    L_ITEM_Speck.ks19.keypos = 75                 ' キーポジション
    L_ITEM_Speck.ks19.keyleng = 2                 ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks19.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks19.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks19.reserve = &H0               ' 予約済み

    L_ITEM_Speck.ks20.keypos = 77                 ' キーポジション
    L_ITEM_Speck.ks20.keyleng = 2                 ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks20.keyflag = BtKfExt + BtKfSeg + BtKfChg
    L_ITEM_Speck.ks20.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks20.reserve = &H0               ' 予約済み

    L_ITEM_Speck.ks21.keypos = 3                  ' キーポジション
    L_ITEM_Speck.ks21.keyleng = 20                ' キー長
                                                ' キーフラグ
    L_ITEM_Speck.ks21.keyflag = BtKfExt + BtKfChg
    L_ITEM_Speck.ks21.keytype = Chr(BtKtString)   ' キータイプ
    L_ITEM_Speck.ks21.reserve = &H0               ' 予約済み
'-----------------------------------------------

    sts = BTRV(BtOpCreate, L_ITEM_POS, L_ITEM_Speck, Len(L_ITEM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "L品目マスタ")
        Exit Function
    End If

    L_ITEM_Create = False

End Function

Public Function L_ITEM_Open(Mode As Integer) As Integer
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
    
    L_ITEM_Open = True
                                            '品目マスタフルパス取込み
    sts = GetIni("FILE", L_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [L_ITEM]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = L_ITEM_Create()        '品目マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "L_品目マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "L_品目マスタ")
                Exit Function
        End Select
    Loop

    L_ITEM_Open = False

End Function


