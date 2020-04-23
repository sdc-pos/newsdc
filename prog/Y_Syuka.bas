Attribute VB_Name = "Y_SYU"
Option Explicit
'********************************************************************
'*
'*              出荷予定データ  ファイル定義
'*              新ﾚｲｱｳﾄ対応     2006.05.24
'*
'********************************************************************
'ファイルＩＤ
Public Const Y_SYU_ID$ = "Y_SYU"

'ページサイズ
Public Const Y_SYU_PG_SIZ% = 4096

'ポジション・ブロック
Public Y_SYU_POS    As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type Y_SYUREC_Tag
    WEL_ID(0 To 2)              As Byte     '使用子機ID
    PRG_ID(0 To 7)              As Byte     '使用中プログラム
    KAN_KBN(0 To 0)             As Byte     '完了区分
    DT_SYU(0 To 0)              As Byte     'データ種別
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    KEY_ID_NO(0 To 11)          As Byte     'ID-NO
    NAIGAI(0 To 0)              As Byte     '国内外
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
    '-----------------  ﾎｽﾄ出荷ﾃﾞｰﾀｲﾒｰｼﾞ　▽
    JGYOBA(0 To 7)              As Byte     '事業場ｺｰﾄﾞ
    DATA_KBN(0 To 0)            As Byte     'データ区分
    TORI_KBN(0 To 1)            As Byte     '取引区分
    ID_NO(0 To 11)              As Byte     'ID-NO
    KAIKEI_JGYOBA(0 To 7)       As Byte     '会計用事業場ｺｰﾄﾞ
    SHISAN_JGYOBA(0 To 7)       As Byte     '資産管理用事業場ｺｰﾄﾞ
    HIN_NO(0 To 19)             As Byte     '品目番号
    DEN_NO(0 To 9)              As Byte     '伝票番号
    SURYO(0 To 6)               As Byte     '出荷数量
    MUKE_CODE(0 To 7)           As Byte     '得意先コード
    SYUKO_SYUSI(0 To 7)         As Byte     '在庫収支
    SHISAN_SYUSI(0 To 7)        As Byte     '資産管理用在庫収支ｺｰﾄﾞ
    HOJYO_SYUSI(0 To 7)         As Byte     '補助在庫収支ｺｰﾄﾞ
    SYUKO_YMD(0 To 7)           As Byte     '出庫日
    TANKA(0 To 9)               As Byte     '実際単価
    ODER_NO(0 To 11)            As Byte     'オーダー番号
    ITEM_NO(0 To 4)             As Byte     'アイテム番号
    ODER_NO_R(0 To 4)           As Byte     '注文管理番号略号
    KOSO_KEITAI(0 To 9)         As Byte     '個装形態ｺｰﾄﾞ
    SYUKA_YMD(0 To 7)           As Byte     '出荷予定日
    TANABAN1(0 To 9)            As Byte     'ﾛｹｰｼｮﾝ1
    TANABAN2(0 To 9)            As Byte     'ﾛｹｰｼｮﾝ2
    TANABAN3(0 To 9)            As Byte     'ﾛｹｰｼｮﾝ3
    MUKE_NAME(0 To 23)          As Byte     '得意先名称
    CYU_KBN(0 To 0)             As Byte     '注文区分
    CYU_KBN_NAME(0 To 39)       As Byte     '注文区分名称
    ORIGIN1(0 To 9)             As Byte     '原産国1
    ORIGIN2(0 To 9)             As Byte     '原産国2
    BIKOU2(0 To 39)             As Byte     '備考2
    HAN_KBN(0 To 0)             As Byte     '販売区分
    CHOKU_KBN(0 To 0)           As Byte     '直送指示区分
    UNIT_ID_NO(0 To 11)         As Byte     'ﾕﾆｯﾄ修正管理番号
    ZAIKO_HIKIATE(0 To 2)       As Byte     '在庫引当順序
    GOKON_KANRI_NO(0 To 7)      As Byte     '合梱管理番号
    JYUCHU_ZAN(0 To 6)          As Byte     '受注残数量
    KYOKYU_KBN(0 To 0)          As Byte     '供給区分
    SHOHIN_SYUSI(0 To 7)        As Byte     '商品化納品在庫収支ｺｰﾄﾞ
    S_SHISAN_SYUSI(0 To 7)      As Byte     '商品化納品資産管理収支ｺｰﾄﾞ
    S_HOJYO_SYUSI(0 To 7)       As Byte     '商品化納品補助収支ｺｰﾄﾞ
    BIKOU1(0 To 39)             As Byte     '備考1
    CHOHA_KBN(0 To 0)           As Byte     '帳端区分
    JYU_HIN_NO(0 To 39)         As Byte     '受付品目番号
    HIN_NAME(0 To 39)           As Byte     '品名
    HIN_CHANGE_KBN(0 To 0)      As Byte     '品目番号変更区分
    MODULE_EXCHANGE(0 To 0)     As Byte     'ﾓｼﾞｭｰﾙ交換区分
    ZAIKO_SYUSI(0 To 7)         As Byte     '残在庫まとめ在庫収支ｺｰﾄﾞ
    ZAN_SHISAN_SYUSI(0 To 7)    As Byte     '残在庫まとめ資産管理収支ｺｰﾄﾞ
    ZAN_HOJYO_SYUSI(0 To 7)     As Byte     '残在庫まとめ補助収支ｺｰﾄﾞ
    NOUKI_YMD(0 To 7)           As Byte     '指定納期
    SERVICE_KANRI_NO(0 To 8)    As Byte     'ｻｰﾋﾞｽ会社管理番号
    KISHU_CODE(0 To 2)          As Byte     '機種品目ｺｰﾄﾞ
    ENVIRONMENT_KBN(0 To 0)     As Byte     '環境企画部品区分
    SS_CODE(0 To 7)             As Byte     '直送相手先ｺｰﾄﾞ
    KEPIN_KAIJYO(0 To 0)        As Byte     '欠品解消区分
    '-----------------  ﾎｽﾄ出荷ﾃﾞｰﾀｲﾒｰｼﾞ　△
    HIN_NAI(0 To 19)            As Byte     '品番（内部）
    HTANABAN(0 To 7)            As Byte     'ホスト棚番
    PRINT_YMD(0 To 7)           As Byte     '出庫表印刷日付
    KAN_YMD(0 To 7)             As Byte     '完了日付
    KENPIN_YMD(0 To 7)          As Byte     '検品日付
    TOK_KBN(0 To 0)             As Byte     '特売り区分
    JITU_SURYO(0 To 6)          As Byte     '出庫実績数量
    INS_NOW(0 To 13)            As Byte     '取込み日時
    KENPIN_TANTO_CODE(0 To 4)   As Byte     '検品担当者ｺｰﾄﾞ 2006.08.07
    KENPIN_HMS(0 To 5)          As Byte     '検品時刻       2006.08.07
    LK_MUKE_CODE(0 To 7)        As Byte     '上位ﾘﾝｸ用向け先2006.08.07
    LK_SEQ_NO(0 To 11)          As Byte     '上位ﾘﾝｸ用連番  2006.08.07
    G_KENPIN_F(0 To 0)          As Byte     '画面検品ﾌﾗｸﾞ   2006.08.07
    
    
    KENPIN_SURYO(0 To 6)        As Byte     '検品時数量     2006.08.07
    
    
    
    
    H_IO_KBN(0 To 0)            As Byte     '入出庫区分(引数、処理区分) 2007.12.11
    H_SOKO_CODE(0 To 0)         As Byte     '倉庫ｺｰﾄﾞ(在庫収支ﾃｰﾌﾞﾙの倉庫ｺｰﾄﾞ) 2007.12.11
    
    
    UPD_NOW(0 To 13)            As Byte     '更新日時   2008.11.28
    
    KAN_HMS(0 To 5)             As Byte     '完了時刻   2011.03.30
    
    FILLER(0 To 5)              As Byte      'FILLER     2011.03.30
End Type

'データ・バッファ
Public Y_SYUREC                 As Y_SYUREC_Tag

'キー定義
Type KEY0_Y_SYU            'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_ID_NO(0 To 11)          As Byte     'ID-NO
End Type

Type KEY1_Y_SYU            'ＫＥＹ１
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KAN_KBN(0 To 0)             As Byte     '完了区分
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    KEY_ID_NO(0 To 11)          As Byte     'ID-NO
    NAIGAI(0 To 0)              As Byte     '国内外
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
End Type

Type KEY2_Y_SYU            'ＫＥＹ２
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
End Type

Type KEY3_Y_SYU            'ＫＥＹ３
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
    NAIGAI(0 To 0)              As Byte     '国内外
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
    KEY_ID_NO(0 To 11)          As Byte     'ID-NO
End Type

Type KEY4_Y_SYU            'ＫＥＹ４
    WEL_ID(0 To 2)              As Byte     '使用子機ID
    PRG_ID(0 To 7)              As Byte     '使用中プログラム
End Type

Type KEY5_Y_SYU            'ＫＥＹ５
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
    HTANABAN(0 To 7)            As Byte     'ホスト棚番         '2004.06.08
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号           '2004.06.29
End Type

Type KEY6_Y_SYU            'ＫＥＹ６
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    HTANABAN(0 To 7)            As Byte     'ホスト棚番
    NAIGAI(0 To 0)              As Byte     '国内外
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
End Type

Type KEY7_Y_SYU            'ＫＥＹ７
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
End Type

Type KEY8_Y_SYU            'ＫＥＹ８        2011.08.02
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
    KAN_YMD(0 To 7)             As Byte     '完了日付
    KAN_HMS(0 To 13)            As Byte     '完了日時
End Type

Type KEY9_Y_SYU            'ＫＥＹ９        2011.08.02
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
End Type

Type KEY10_Y_SYU            'ＫＥＹ10       2012.06.15
    KAN_KBN(0 To 0)             As Byte     '完了区分
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
    PRINT_YMD(0 To 7)           As Byte     '出庫表印刷日付
    DEN_NO(0 To 9)              As Byte     '伝票番号
End Type



'キー・データ
Public K0_Y_SYU                 As KEY0_Y_SYU
Public K1_Y_SYU                 As KEY1_Y_SYU
Public K2_Y_SYU                 As KEY2_Y_SYU
Public K3_Y_SYU                 As KEY3_Y_SYU
Public K4_Y_SYU                 As KEY4_Y_SYU
Public K5_Y_SYU                 As KEY5_Y_SYU
Public K6_Y_SYU                 As KEY6_Y_SYU
Public K7_Y_SYU                 As KEY7_Y_SYU

Public K8_Y_SYU                 As KEY8_Y_SYU       '2011.08.02
Public K9_Y_SYU                 As KEY9_Y_SYU       '2011.08.02

Public K10_Y_SYU                As KEY10_Y_SYU      '2012.06.15


Type Y_SYU_FSpeck
    fs      As BtFileSpeck                  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks7     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks8     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks9     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks10    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks11    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks12    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks13    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks14    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks15    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks16    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks17    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks18    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks19    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks20    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks21    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks22    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks23    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks24    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks25    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks26    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks27    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks28    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks29    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks30    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks31    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks32    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks33    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks34    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks35    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks36    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks37    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks38    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体

    ks39    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2011.08.02
    ks40    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2011.08.02
    ks41    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2011.08.02
    ks42    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2011.08.02
    ks43    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2011.08.02
    ks44    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2011.08.02

    ks45    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.06.15
    ks46    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.06.15
    ks47    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.06.15
    ks48    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.06.15
    ks49    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.06.15

End Type

Private Y_SYU_Speck As Y_SYU_FSpeck

Private Function Y_SYU_Create() As Integer

'********************************************************************
'*
'*              出荷予定データ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    Y_SYU_Create = True
                                            '出荷予定データフルパス取込み
    sts = GetIni("FILE", Y_SYU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_SYU]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    Y_SYU_Speck.fs.recoleng = Len(Y_SYUREC)         ' レコード長
    Y_SYU_Speck.fs.PageSize = Y_SYU_PG_SIZ          ' ページサイズ
    Y_SYU_Speck.fs.idexnumb = 11                    ' インデックス数    8-->10 2011.08.02   10-->11 2012.06.15
    Y_SYU_Speck.fs.fileflag = 0                     ' ファイルフラグ
    Y_SYU_Speck.fs.reserve = &H0                    ' 予約済み
'---------------------------------------------------' キー０
    Y_SYU_Speck.ks0.keypos = 14                     ' キーポジション
    Y_SYU_Speck.ks0.keyleng = 1                     ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_Speck.ks0.reserve = &H0                   ' 予約済み
    
    Y_SYU_Speck.ks1.keypos = 16                     ' キーポジション
    Y_SYU_Speck.ks1.keyleng = 12                    ' キー長
    Y_SYU_Speck.ks1.keyflag = BtKfExt + BtKfChg     ' キーフラグ
    Y_SYU_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_Speck.ks1.reserve = &H0                   ' 予約済み

'---------------------------------------------------' キー１
    Y_SYU_Speck.ks2.keypos = 14                     ' キーポジション
    Y_SYU_Speck.ks2.keyleng = 1                     ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_Speck.ks2.reserve = &H0                   ' 予約済み
    
    Y_SYU_Speck.ks3.keypos = 12                     ' キーポジション
    Y_SYU_Speck.ks3.keyleng = 1                     ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_Speck.ks3.reserve = &H0                   ' 予約済み
    
    Y_SYU_Speck.ks4.keypos = 49                     ' キーポジション
    Y_SYU_Speck.ks4.keyleng = 8                     ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks4.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_Speck.ks4.reserve = &H0                   ' 予約済み
    
    Y_SYU_Speck.ks5.keypos = 57                     ' キーポジション
    Y_SYU_Speck.ks5.keyleng = 8                     ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks5.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_Speck.ks5.reserve = &H0                   ' 予約済み
    
    Y_SYU_Speck.ks6.keypos = 15                     ' キーポジション
    Y_SYU_Speck.ks6.keyleng = 1                     ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks6.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_Speck.ks6.reserve = &H0                   ' 予約済み
    
    Y_SYU_Speck.ks7.keypos = 16                     ' キーポジション
    Y_SYU_Speck.ks7.keyleng = 12                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks7.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_Speck.ks7.reserve = &H0                   ' 予約済み
    
    Y_SYU_Speck.ks8.keypos = 28                     ' キーポジション
    Y_SYU_Speck.ks8.keyleng = 1                     ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks8.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_Speck.ks8.reserve = &H0                   ' 予約済み
    
    Y_SYU_Speck.ks9.keypos = 29                    ' キーポジション
    Y_SYU_Speck.ks9.keyleng = 20                   ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks9.keyflag = BtKfExt + BtKfChg
    Y_SYU_Speck.ks9.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks9.reserve = &H0                  ' 予約済み
'---------------------------------------------------' キー２
    Y_SYU_Speck.ks10.keypos = 14                    ' キーポジション
    Y_SYU_Speck.ks10.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks10.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks10.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks11.keypos = 15                    ' キーポジション
    Y_SYU_Speck.ks11.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks11.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks11.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks12.keypos = 49                    ' キーポジション
    Y_SYU_Speck.ks12.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks12.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks12.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks13.keypos = 57                    ' キーポジション
    Y_SYU_Speck.ks13.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks13.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks13.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks13.reserve = &H0                  ' 予約済み
'---------------------------------------------------' キー３
    Y_SYU_Speck.ks14.keypos = 14                    ' キーポジション
    Y_SYU_Speck.ks14.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks14.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks14.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks15.keypos = 15                    ' キーポジション
    Y_SYU_Speck.ks15.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks15.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks15.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks16.keypos = 49                    ' キーポジション
    Y_SYU_Speck.ks16.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks16.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks16.reserve = &H0                  ' 予約済み
                                                    
    Y_SYU_Speck.ks17.keypos = 57                    ' キーポジション
    Y_SYU_Speck.ks17.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks17.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks17.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks17.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks18.keypos = 28                    ' キーポジション
    Y_SYU_Speck.ks18.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks18.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks18.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks19.keypos = 29                    ' キーポジション
    Y_SYU_Speck.ks19.keyleng = 20                   ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks19.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks19.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks19.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks20.keypos = 16                    ' キーポジション
    Y_SYU_Speck.ks20.keyleng = 12                   ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks20.keyflag = BtKfExt + BtKfChg
    Y_SYU_Speck.ks20.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks20.reserve = &H0                  ' 予約済み
'---------------------------------------------------' キー４
    Y_SYU_Speck.ks21.keypos = 1                     ' キーポジション
    Y_SYU_Speck.ks21.keyleng = 3                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks21.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks21.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks21.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks22.keypos = 4                     ' キーポジション
    Y_SYU_Speck.ks22.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks22.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks22.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks22.reserve = &H0                  ' 予約済み
'---------------------------------------------------' キー５
    Y_SYU_Speck.ks23.keypos = 14                    ' キーポジション
    Y_SYU_Speck.ks23.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks23.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks23.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks23.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks24.keypos = 15                    ' キーポジション
    Y_SYU_Speck.ks24.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks24.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks24.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks24.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks25.keypos = 49                    ' キーポジション
    Y_SYU_Speck.ks25.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks25.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks25.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks25.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks26.keypos = 57                    ' キーポジション
    Y_SYU_Speck.ks26.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks26.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks26.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks26.reserve = &H0                  ' 予約済み
    
    
    Y_SYU_Speck.ks27.keypos = 648                   ' キーポジション
    Y_SYU_Speck.ks27.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks27.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks27.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks27.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks28.keypos = 65                    ' キーポジション
    Y_SYU_Speck.ks28.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks28.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks28.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks28.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks29.keypos = 29                    ' キーポジション
    Y_SYU_Speck.ks29.keyleng = 20                   ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks29.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks29.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks29.reserve = &H0                  ' 予約済み
'---------------------------------------------------' キー６
    Y_SYU_Speck.ks30.keypos = 14                    ' キーポジション
    Y_SYU_Speck.ks30.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks30.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks30.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks30.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks31.keypos = 15                    ' キーポジション
    Y_SYU_Speck.ks31.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks31.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks31.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks31.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks32.keypos = 648                   ' キーポジション
    Y_SYU_Speck.ks32.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks32.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks32.keytype = Chr(BtKtString)      ' キータイプ3
    Y_SYU_Speck.ks32.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks33.keypos = 28                    ' キーポジション
    Y_SYU_Speck.ks33.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks33.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_SYU_Speck.ks33.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks33.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks34.keypos = 29                    ' キーポジション
    Y_SYU_Speck.ks34.keyleng = 20                   ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks34.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks34.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks34.reserve = &H0                  ' 予約済み
'---------------------------------------------------' キー７
    Y_SYU_Speck.ks35.keypos = 14                    ' キーポジション
    Y_SYU_Speck.ks35.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks35.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks35.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks35.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks36.keypos = 28                    ' キーポジション
    Y_SYU_Speck.ks36.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks36.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks36.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks36.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks37.keypos = 29                    ' キーポジション
    Y_SYU_Speck.ks37.keyleng = 20                   ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks37.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks37.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks37.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks38.keypos = 65                    ' キーポジション
    Y_SYU_Speck.ks38.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks38.keyflag = BtKfExt + BtKfDup + BtKfChg
    Y_SYU_Speck.ks38.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks38.reserve = &H0                  ' 予約済み
    
    
'---------------------------------------------------' キー８    2011.08.02
    Y_SYU_Speck.ks39.keypos = 14                    ' キーポジション
    Y_SYU_Speck.ks39.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks39.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks39.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks39.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks40.keypos = 28                    ' キーポジション
    Y_SYU_Speck.ks40.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks40.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks40.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks40.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks41.keypos = 29                    ' キーポジション
    Y_SYU_Speck.ks41.keyleng = 20                   ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks41.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks41.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks41.reserve = &H0                  ' 予約済み
    
    Y_SYU_Speck.ks42.keypos = 664                   ' キーポジション
    Y_SYU_Speck.ks42.keyleng = 8                     ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks42.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks42.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks42.reserve = &H0                  ' 予約済み
    
    

    Y_SYU_Speck.ks43.keypos = 757                    ' キーポジション
    Y_SYU_Speck.ks43.keyleng = 6                     ' キー長
                                                            ' キーフラグ
    Y_SYU_Speck.ks43.keyflag = BtKfExt + BtKfDup + BtKfChg
    Y_SYU_Speck.ks43.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks43.reserve = &H0                  ' 予約済み
    
    
'---------------------------------------------------' キー９    2011.08.02
    Y_SYU_Speck.ks44.keypos = 65                    ' キーポジション
    Y_SYU_Speck.ks44.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks44.keyflag = BtKfExt + BtKfDup + BtKfChg
    Y_SYU_Speck.ks44.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks44.reserve = &H0                  ' 予約済み
    
    
'---------------------------------------------------' キー10    2012.06.15
    
    
    Y_SYU_Speck.ks45.keypos = 12                    ' キーポジション
    Y_SYU_Speck.ks45.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks45.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks45.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks45.reserve = &H0                  ' 予約済み
    
    
    
    Y_SYU_Speck.ks46.keypos = 14                    ' キーポジション
    Y_SYU_Speck.ks46.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks46.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks46.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks46.reserve = &H0                  ' 予約済み
    
    
    Y_SYU_Speck.ks47.keypos = 65                    ' キーポジション
    Y_SYU_Speck.ks47.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks47.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks47.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks47.reserve = &H0                  ' 予約済み
    
    
    Y_SYU_Speck.ks48.keypos = 656                   ' キーポジション
    Y_SYU_Speck.ks48.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks48.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    Y_SYU_Speck.ks48.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks48.reserve = &H0                  ' 予約済み
    
    
    Y_SYU_Speck.ks49.keypos = 132                   ' キーポジション
    Y_SYU_Speck.ks49.keyleng = 10                   ' キー長
                                                    ' キーフラグ
    Y_SYU_Speck.ks49.keyflag = BtKfExt + BtKfDup + BtKfChg
    Y_SYU_Speck.ks49.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_Speck.ks49.reserve = &H0                  ' 予約済み
    
    
    
    sts = BTRV(BtOpCreate, Y_SYU_POS, Y_SYU_Speck, Len(Y_SYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "出荷予定データ")
        Exit Function
    End If

    Y_SYU_Create = False

End Function

Function Y_SYU_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              出荷予定データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    Y_SYU_Open = True
                                            '出荷予定データフルパス取込み
    sts = GetIni("FILE", Y_SYU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_SYU]読み込みエラー ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Y_SYU_Create()        '出荷予定データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "出荷予定データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "出荷予定データ")
                Exit Function
        End Select
    Loop
    Y_SYU_Open = False
End Function
