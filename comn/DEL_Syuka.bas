Attribute VB_Name = "DEL_SYU"
Option Explicit
'********************************************************************
'*
'*              削除済み出荷予定データ  ファイル定義
'*              新　ﾚｲｱｳﾄ対応 2006.05.24
'********************************************************************
'ファイルＩＤ
Public Const DEL_SYU_ID$ = "DEL_SYU"

'ページサイズ
Public Const DEL_SYU_PG_SIZ% = 4096

'ポジション・ブロック
Public DEL_SYU_POS    As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type DEL_SYUREC_Tag
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
    KENPIN_TANTO_CODE(0 To 4)   As Byte     '検品担当者ｺｰﾄﾞ 2006.07.20
    KENPIN_HMS(0 To 5)          As Byte     '検品時刻       2006.07.20
    
    LK_MUKE_CODE(0 To 7)        As Byte     '上位ﾘﾝｸ用向け先2006.07.20
        
    FILLER(0 To 47)             As Byte     'FILLER
End Type

'データ・バッファ
Public DEL_SYUREC As DEL_SYUREC_Tag

'キー定義
Type KEY0_DEL_SYU            'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
End Type

Type KEY1_DEL_SYU           'ＫＥＹ１
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
End Type

Type KEY2_DEL_SYU            'ＫＥＹ２
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    NAIGAI(0 To 0)              As Byte     '国内外
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
    KEY_ID_NO(0 To 11)           As Byte     'ID-NO
End Type


'キー・データ
Public K0_DEL_SYU               As KEY0_DEL_SYU
Public K1_DEL_SYU               As KEY1_DEL_SYU
Public K2_DEL_SYU               As KEY2_DEL_SYU

Type DEL_SYU_FSpeck
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
End Type

Private DEL_SYU_Speck As DEL_SYU_FSpeck

Private Function DEL_SYU_Create() As Integer
'********************************************************************
'*
'*              削除済み出荷予定データ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    DEL_SYU_Create = True
                                            '削除済み出荷予定データフルパス取込み
    sts = GetIni("FILE", DEL_SYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [DEL_SYU]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    DEL_SYU_Speck.fs.recoleng = Len(DEL_SYUREC)               ' レコード長
    DEL_SYU_Speck.fs.PageSize = DEL_SYU_PG_SIZ              ' ページサイズ
    DEL_SYU_Speck.fs.idexnumb = 3                           ' インデックス数
    DEL_SYU_Speck.fs.fileflag = 0                           ' ファイルフラグ
    DEL_SYU_Speck.fs.reserve = &H0                          ' 予約済み
'---------------------------------------------------        キー０
    
    DEL_SYU_Speck.ks0.keypos = 14                           ' キーポジション
    DEL_SYU_Speck.ks0.keyleng = 1                           ' キー長
    DEL_SYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks0.keytype = Chr(BtKtString)             ' キータイプ
    DEL_SYU_Speck.ks0.reserve = &H0                         ' 予約済み
    
    DEL_SYU_Speck.ks1.keypos = 15                           ' キーポジション
    DEL_SYU_Speck.ks1.keyleng = 1                           ' キー長
    DEL_SYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks1.keytype = Chr(BtKtString)             ' キータイプ
    DEL_SYU_Speck.ks1.reserve = &H0                         ' 予約済み
    
    DEL_SYU_Speck.ks2.keypos = 49                           ' キーポジション
    DEL_SYU_Speck.ks2.keyleng = 8                           ' キー長
    DEL_SYU_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks2.keytype = Chr(BtKtString)             ' キータイプ
    DEL_SYU_Speck.ks2.reserve = &H0                         ' 予約済み
    
    DEL_SYU_Speck.ks3.keypos = 57                           ' キーポジション
    DEL_SYU_Speck.ks3.keyleng = 8                           ' キー長
    DEL_SYU_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks3.keytype = Chr(BtKtString)             ' キータイプ
    DEL_SYU_Speck.ks3.reserve = &H0                         ' 予約済み
    
    DEL_SYU_Speck.ks4.keypos = 29                           ' キーポジション
    DEL_SYU_Speck.ks4.keyleng = 20                          ' キー長
    DEL_SYU_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks4.keytype = Chr(BtKtString)             ' キータイプ
    DEL_SYU_Speck.ks4.reserve = &H0                         ' 予約済み
    
    DEL_SYU_Speck.ks5.keypos = 65                           ' キーポジション
    DEL_SYU_Speck.ks5.keyleng = 8                           ' キー長
    DEL_SYU_Speck.ks5.keyflag = BtKfExt + BtKfDup           ' キーフラグ
    DEL_SYU_Speck.ks5.keytype = Chr(BtKtString)             ' キータイプ
    DEL_SYU_Speck.ks5.reserve = &H0                         ' 予約済み

'---------------------------------------------------        キー１
    
    DEL_SYU_Speck.ks6.keypos = 65                           ' キーポジション
    DEL_SYU_Speck.ks6.keyleng = 8                           ' キー長
    DEL_SYU_Speck.ks6.keyflag = BtKfExt + BtKfDup           ' キーフラグ
    DEL_SYU_Speck.ks6.keytype = Chr(BtKtString)             ' キータイプ
    DEL_SYU_Speck.ks6.reserve = &H0                         ' 予約済み
    
'---------------------------------------------------        キー２
    
    DEL_SYU_Speck.ks7.keypos = 14                           ' キーポジション
    DEL_SYU_Speck.ks7.keyleng = 1                           ' キー長
    DEL_SYU_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks7.keytype = Chr(BtKtString)             ' キータイプ
    DEL_SYU_Speck.ks7.reserve = &H0                         ' 予約済み
    
    DEL_SYU_Speck.ks8.keypos = 49                           ' キーポジション
    DEL_SYU_Speck.ks8.keyleng = 8                           ' キー長
    DEL_SYU_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks8.keytype = Chr(BtKtString)             ' キータイプ
    DEL_SYU_Speck.ks8.reserve = &H0                         ' 予約済み
    
    DEL_SYU_Speck.ks9.keypos = 57                           ' キーポジション
    DEL_SYU_Speck.ks9.keyleng = 8                           ' キー長
    DEL_SYU_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks9.keytype = Chr(BtKtString)             ' キータイプ
    DEL_SYU_Speck.ks9.reserve = &H0                         ' 予約済み
    
    DEL_SYU_Speck.ks10.keypos = 15                          ' キーポジション
    DEL_SYU_Speck.ks10.keyleng = 1                          ' キー長
    DEL_SYU_Speck.ks10.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks10.keytype = Chr(BtKtString)            ' キータイプ
    DEL_SYU_Speck.ks10.reserve = &H0                        ' 予約済み
    
    DEL_SYU_Speck.ks11.keypos = 28                          ' キーポジション
    DEL_SYU_Speck.ks11.keyleng = 1                          ' キー長
    DEL_SYU_Speck.ks11.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks11.keytype = Chr(BtKtString)            ' キータイプ
    DEL_SYU_Speck.ks11.reserve = &H0                        ' 予約済み
    
    DEL_SYU_Speck.ks12.keypos = 29                          ' キーポジション
    DEL_SYU_Speck.ks12.keyleng = 20                         ' キー長
    DEL_SYU_Speck.ks12.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    DEL_SYU_Speck.ks12.keytype = Chr(BtKtString)            ' キータイプ
    DEL_SYU_Speck.ks12.reserve = &H0                        ' 予約済み
    
    DEL_SYU_Speck.ks13.keypos = 16                          ' キーポジション
    DEL_SYU_Speck.ks13.keyleng = 12                          ' キー長
    DEL_SYU_Speck.ks13.keyflag = BtKfExt + BtKfDup          ' キーフラグ
    DEL_SYU_Speck.ks13.keytype = Chr(BtKtString)            ' キータイプ
    DEL_SYU_Speck.ks13.reserve = &H0                        ' 予約済み
    
    sts = BTRV(BtOpCreate, DEL_SYU_POS, DEL_SYU_Speck, Len(DEL_SYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "削除済み出荷予定データ")
        Exit Function
    End If

    DEL_SYU_Create = False

End Function
Function DEL_SYU_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*             削除済み出荷予定データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    DEL_SYU_Open = True
                                            '削除済み出荷予定データフルパス取込み
    sts = GetIni("FILE", DEL_SYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [DEL_SYU]読み込みエラー ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = DEL_SYU_Create()        '削除済み出荷予定データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "削除済み出荷予定データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "削除済み出荷予定データ")
                Exit Function
        End Select
    Loop
    
    DEL_SYU_Open = False

End Function
