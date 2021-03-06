Attribute VB_Name = "Y_GLICS"
Option Explicit
'********************************************************************
'*                                                                  *
'*              入荷予定データ  ファイル定義                        *
'*                                                                  *
'********************************************************************
'ファイルＩＤ
Public Const Y_GLICS_ID$ = "Y_GLICS"

'ページサイズ
Public Const Y_GLICS_PG_SIZ% = 2048

'ポジション・ブロック
Public Y_GLICS_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type Y_GLICSREC_Tag
    KAN_KBN(0 To 0)             As Byte     '完了区分
    DT_SYU(0 To 0)              As Byte     'データ種別
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    TEXT_NO(0 To 8)             As Byte     'テキスト��
    
    '-----------------  ﾎｽﾄ入庫ﾃﾞｰﾀｲﾒｰｼﾞ　▽
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
    SYUKO_YMD(0 To 7)           As Byte     '出庫日付
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
    CYU_KBN_NAME(0 To 9)        As Byte     '注文区分名称
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
    JYU_HIN_NO(0 To 19)         As Byte     '受付品目番号
    HIN_NAME(0 To 19)           As Byte     '品名
    HIN_CHANGE_KBN(0 To 0)      As Byte     '品目番号変更区分
    MODULE_EXCHANGE(0 To 0)     As Byte     'ﾓｼﾞｭｰﾙ交換区分
    ZAIKO_SYUSI(0 To 7)         As Byte     '残在庫まとめ在庫収支ｺｰﾄﾞ
    ZAN_SHISAN_SYUSI(0 To 7)    As Byte     '残在庫まとめ資産管理収支ｺｰﾄﾞ
    ZAN_HOJYO_SYUSI(0 To 7)     As Byte     '残在庫まとめ補助収支ｺｰﾄﾞ
    NOUKI_YMD(0 To 7)           As Byte     '指定納期
    SERVICE_KANRI_NO(0 To 8)    As Byte     'ｻｰﾋﾞｽ会社管理番号
    KI_HIN_NO(0 To 2)           As Byte     '機種品目ｺｰﾄﾞ
    ENVIRONMENT_KBN(0 To 0)     As Byte     '環境企画部品区分
    SS_CODE(0 To 7)             As Byte     '直送相手先ｺｰﾄﾞ
    KEPIN_KAIJYO(0 To 0)        As Byte     '欠品解消区分
    '-----------------  ﾎｽﾄ入庫ﾃﾞｰﾀｲﾒｰｼﾞ　△
    
    KAN_DT(0 To 7)              As Byte     '完了日付
    BEF_NYU_QTY(0 To 7)         As Byte     '先行入荷数
    YOSAN_FROM(0 To 4)          As Byte     '予算単位（元）
    YOSAN_TO(0 To 4)            As Byte     '予算単位（先）
    HTANABAN(0 To 7)            As Byte     '標準棚番
    HIN_NAI(0 To 12)            As Byte     '品番（内部）
    H_SOKO(0 To 7)              As Byte     'ﾎｽﾄ倉庫 2006.10.17
    
    NYU_LIST_OUT(0 To 0)        As Byte     '入庫予定出力ﾌﾗｸﾞ 2007.06.12
    
    
    
    '-----------------  旧GLICS項目追加 2007.06.15
    CYOK_KBN(0 To 0)            As Byte     '直送区分
    IO_KBN(0 To 0)              As Byte     '入出庫区分
    PM_KBN(0 To 0)              As Byte     '赤黒区分
    DEN_SYU(0 To 0)             As Byte     '伝票種別
    SYUK_CODE(0 To 4)           As Byte     '支給先／出荷先
    SYUK_NAME(0 To 19)          As Byte     '支給先／出荷先名
    
    
    INS_NOW(0 To 13)            As Byte     '挿入年月日時分秒
    '-----------------  旧GLICS項目追加 2007.06.15
    
    '----------------   2010.07.08 ▽
    GENSANKOKU(0 To 19)         As Byte     '原産国名
    GEN_GENSANKOKU(0 To 19)     As Byte     '現物表示原産国名
    SHIIRE_WORK_CENTER(0 To 7)  As Byte     '資材仕入先ﾜｰｸｾﾝﾀｰ
    KANKYO_KBN(0 To 2)          As Byte     '環境種類区分
    KANKYO_KBN_ST(0 To 7)       As Byte     '環境種類区分適用開始
    KANKYO_KBN_SURYO(0 To 9)    As Byte     '環境種類区分数量
    ID_NO2(0 To 11)             As Byte     'ID_NO
    AITESAKI_CODE(0 To 15)      As Byte     '相手先ｺｰﾄﾞ
    JYUCHU_YMD(0 To 7)          As Byte     '受注年月日
    SHITEI_NOUKI_YMD(0 To 7)    As Byte     '指定納期年月日
    LIST_OUT_END_F(0 To 0)      As Byte     '入庫ﾘｽﾄ出力F
    NYUKO_TANABAN(0 To 7)       As Byte     '入庫棚番
    MAEGARI_SURYO(0 To 7)       As Byte     '前借相殺数
    '----------------   2010.07.08 △
    
    '----------------   2011.03.23 ▽
    MOTO_PROG_ID(0 To 7)        As Byte     '発生元プログラム
    MOTO_TEXT_NO(0 To 8)        As Byte     '元テキスト��
    '----------------   2011.03.23 △
    
    
    
    FILLER(0 To 23)            As Byte      '40-->23    2011.03.23
End Type

'データ・バッファ
Public Y_GLICSREC                  As Y_GLICSREC_Tag

'キー定義
Type KEY0_Y_GLICS            'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    SYUKA_YMD(0 To 7)           As Byte     '出荷日
    TEXT_NO(0 To 8)             As Byte     'テキスト��
End Type

Type KEY1_Y_GLICS            'ＫＥＹ１
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KAN_KBN(0 To 0)             As Byte     '完了区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_NO(0 To 19)             As Byte     '品目番号
    SYUKA_YMD(0 To 7)           As Byte     '出荷日
    TEXT_NO(0 To 8)             As Byte     'テキスト��
End Type

Type KEY2_Y_GLICS            'ＫＥＹ２
    JGYOBU(0 To 0)              As Byte     '事業部区分
    SYUKA_YMD(0 To 7)           As Byte     '出荷日
    HIN_NO(0 To 19)             As Byte     '品目番号
    NAIGAI(0 To 0)              As Byte     '国内外
End Type

Type KEY3_Y_GLICS            'ＫＥＹ３
    SYUKA_YMD(0 To 7)           As Byte     '出荷日
End Type



'キー・データ
Public K0_Y_GLICS                 As KEY0_Y_GLICS
Public K1_Y_GLICS                 As KEY1_Y_GLICS
Public K2_Y_GLICS                 As KEY2_Y_GLICS
Public K3_Y_GLICS                 As KEY3_Y_GLICS

Private Type Y_GLICS_FSpeck
    fs      As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
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
End Type

Private Y_GLICS_Speck As Y_GLICS_FSpeck

Private Function Y_GLICS_Create() As Integer
'********************************************************************
'*                                                                  *
'*              入荷予定データ  ＣＲＥＡＴＥ                        *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    Y_GLICS_Create = True
                                            '入荷予定データフルパス取込み
    sts = GetIni("FILE", Y_GLICS_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_GLICS]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    Y_GLICS_Speck.fs.recoleng = Len(Y_GLICSREC)     ' レコード長
    Y_GLICS_Speck.fs.PageSize = Y_GLICS_PG_SIZ      ' ページサイズ
    Y_GLICS_Speck.fs.idexnumb = 4                 ' インデックス数
    Y_GLICS_Speck.fs.fileflag = 0                 ' ファイルフラグ
    Y_GLICS_Speck.fs.reserve = &H0                ' 予約済み
    '-------------------------------------------
                                                ' キー０
    Y_GLICS_Speck.ks0.keypos = 3                  ' キーポジション
    Y_GLICS_Speck.ks0.keyleng = 1                 ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    Y_GLICS_Speck.ks0.reserve = &H0               ' 予約済み
                                                ' キー０
    Y_GLICS_Speck.ks1.keypos = 172                ' キーポジション
    Y_GLICS_Speck.ks1.keyleng = 8                 ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks1.keytype = Chr(BtKtString)   ' キータイプ
    Y_GLICS_Speck.ks1.reserve = &H0               ' 予約済み
                                                ' キー０
    Y_GLICS_Speck.ks2.keypos = 5                  ' キーポジション
    Y_GLICS_Speck.ks2.keyleng = 9                 ' キー長
    Y_GLICS_Speck.ks2.keyflag = BtKfExt + BtKfChg ' キーフラグ
    Y_GLICS_Speck.ks2.keytype = Chr(BtKtString)   ' キータイプ
    Y_GLICS_Speck.ks2.reserve = &H0               ' 予約済み
    '-------------------------------------------
                                                
                                                ' キー１
    Y_GLICS_Speck.ks3.keypos = 3                  ' キーポジション
    Y_GLICS_Speck.ks3.keyleng = 1                 ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_GLICS_Speck.ks3.keytype = Chr(BtKtString)   ' キータイプ
    Y_GLICS_Speck.ks3.reserve = &H0               ' 予約済み
                                                ' キー１
    Y_GLICS_Speck.ks4.keypos = 1                  ' キーポジション
    Y_GLICS_Speck.ks4.keyleng = 1                 ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_GLICS_Speck.ks4.keytype = Chr(BtKtString)   ' キータイプ
    Y_GLICS_Speck.ks4.reserve = &H0               ' 予約済み
                                                ' キー１
    Y_GLICS_Speck.ks5.keypos = 4                 ' キーポジション
    Y_GLICS_Speck.ks5.keyleng = 1                 ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_GLICS_Speck.ks5.keytype = Chr(BtKtString)   ' キータイプ
    Y_GLICS_Speck.ks5.reserve = &H0               ' 予約済み
                                                ' キー１
    Y_GLICS_Speck.ks6.keypos = 53                 ' キーポジション
    Y_GLICS_Speck.ks6.keyleng = 20                ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_GLICS_Speck.ks6.keytype = Chr(BtKtString)   ' キータイプ
    Y_GLICS_Speck.ks6.reserve = &H0               ' 予約済み
                                                ' キー１
    Y_GLICS_Speck.ks7.keypos = 172                ' キーポジション
    Y_GLICS_Speck.ks7.keyleng = 8                 ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_GLICS_Speck.ks7.keytype = Chr(BtKtString)   ' キータイプ
    Y_GLICS_Speck.ks7.reserve = &H0               ' 予約済み
                                                ' キー１
    Y_GLICS_Speck.ks8.keypos = 5                ' キーポジション
    Y_GLICS_Speck.ks8.keyleng = 9                 ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks8.keyflag = BtKfExt + BtKfChg
    Y_GLICS_Speck.ks8.keytype = Chr(BtKtString)   ' キータイプ
    Y_GLICS_Speck.ks8.reserve = &H0               ' 予約済み
    '-------------------------------------------
                                                
                                                ' キー２
    Y_GLICS_Speck.ks9.keypos = 3                  ' キーポジション
    Y_GLICS_Speck.ks9.keyleng = 1                 ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks9.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks9.keytype = Chr(BtKtString)   ' キータイプ
    Y_GLICS_Speck.ks9.reserve = &H0               ' 予約済み
                                                ' キー２
    Y_GLICS_Speck.ks10.keypos = 172               ' キーポジション
    Y_GLICS_Speck.ks10.keyleng = 8                ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks10.keytype = Chr(BtKtString)  ' キータイプ
    Y_GLICS_Speck.ks10.reserve = &H0              ' 予約済み
                                                ' キー２
    Y_GLICS_Speck.ks11.keypos = 53                ' キーポジション
    Y_GLICS_Speck.ks11.keyleng = 20               ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks11.keytype = Chr(BtKtString)  ' キータイプ
    Y_GLICS_Speck.ks11.reserve = &H0              ' 予約済み
                                                ' キー２
    Y_GLICS_Speck.ks12.keypos = 4                 ' キーポジション
    Y_GLICS_Speck.ks12.keyleng = 1                ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfSeg
    Y_GLICS_Speck.ks12.keytype = Chr(BtKtString)  ' キータイプ
    Y_GLICS_Speck.ks12.reserve = &H0              ' 予約済み
                                                ' キー２
    Y_GLICS_Speck.ks13.keypos = 5                 ' キーポジション
    Y_GLICS_Speck.ks13.keyleng = 9                ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks13.keyflag = BtKfExt + BtKfChg
    Y_GLICS_Speck.ks13.keytype = Chr(BtKtString)  ' キータイプ
    Y_GLICS_Speck.ks13.reserve = &H0              ' 予約済み
    '-------------------------------------------
                                                
                                                ' キー３
    Y_GLICS_Speck.ks14.keypos = 172               ' キーポジション
    Y_GLICS_Speck.ks14.keyleng = 8                ' キー長
                                                ' キーフラグ
    Y_GLICS_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_GLICS_Speck.ks14.keytype = Chr(BtKtString)  ' キータイプ
    Y_GLICS_Speck.ks14.reserve = &H0              ' 予約済み
    '-------------------------------------------
    
    sts = BTRV(BtOpCreate, Y_GLICS_POS, Y_GLICS_Speck, Len(Y_GLICS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "入荷予定データ")
        Y_GLICS_Create = True
        Exit Function
    End If

    Y_GLICS_Create = False

End Function

Function Y_GLICS_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              入荷予定データ  ＯＰＥＮ                            *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    Y_GLICS_Open = True
                                            '入荷予定データフルパス取込み
    sts = GetIni("FILE", Y_GLICS_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_GLICS]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Y_GLICS_Create()        '入荷予定データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "入荷予定データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "入荷予定データ")
                Exit Function
        End Select
    Loop
    
    Y_GLICS_Open = False

End Function


