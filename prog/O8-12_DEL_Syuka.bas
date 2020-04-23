Attribute VB_Name = "O_DEL_SYU"
Option Explicit
'********************************************************************
'*
'*              削除済み出荷予定データ  ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const O_DEL_SYU_ID$ = "O_DEL_SYU"

'ページサイズ
Public Const O_DEL_SYU_PG_SIZ% = 2048

'ポジション・ブロック
Public O_DEL_SYU_POS    As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type O_DEL_SYUREC_Tag
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
    HIN_NAI(0 To 19)            As Byte     '品番（内部）
    HTANABAN(0 To 7)            As Byte     'ホスト棚番
    PRINT_YMD(0 To 7)           As Byte     '出庫表印刷日付
    KAN_YMD(0 To 7)             As Byte     '完了日付
    KENPIN_YMD(0 To 7)          As Byte     '検品日付
    TOK_KBN(0 To 0)             As Byte     '特売り区分
    JITU_SURYO(0 To 6)          As Byte     '出庫実績数量
    INS_NOW(0 To 13)            As Byte     '取込み日時
    FILLER(0 To 67)             As Byte     'FILLER
End Type

'データ・バッファ
Public O_DEL_SYUREC As O_DEL_SYUREC_Tag

'キー定義
Type KEY0_O_DEL_SYU            'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
End Type

Type KEY1_O_DEL_SYU           'ＫＥＹ１
    KEY_SYUKA_YMD(0 To 7)       As Byte     '出荷日付
End Type

Type KEY2_O_DEL_SYU            'ＫＥＹ２
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KEY_MUKE_CODE(0 To 7)       As Byte     '得意先コード
    KEY_SS_CODE(0 To 7)         As Byte     '直送先コード
    KEY_CYU_KBN(0 To 0)         As Byte     '注文区分
    NAIGAI(0 To 0)              As Byte     '国内外
    KEY_HIN_NO(0 To 19)         As Byte     '品目番号
    KEY_ID_NO(0 To 7)           As Byte     'ID-NO
End Type


'キー・データ
Public K0_O_DEL_SYU               As KEY0_O_DEL_SYU
Public K1_O_DEL_SYU               As KEY1_O_DEL_SYU
Public K2_O_DEL_SYU               As KEY2_O_DEL_SYU

Type O_DEL_SYU_FSpeck
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

Private O_DEL_SYU_Speck As O_DEL_SYU_FSpeck

Private Function O_DEL_SYU_Create() As Integer
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

    O_DEL_SYU_Create = True
                                            '削除済み出荷予定データフルパス取込み
    sts = GetIni("FILE", O_DEL_SYU_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_DEL_SYU]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    O_DEL_SYU_Speck.fs.recoleng = Len(O_DEL_SYUREC)               ' レコード長
    O_DEL_SYU_Speck.fs.PageSize = O_DEL_SYU_PG_SIZ              ' ページサイズ
    O_DEL_SYU_Speck.fs.idexnumb = 3                           ' インデックス数
    O_DEL_SYU_Speck.fs.fileflag = 0                           ' ファイルフラグ
    O_DEL_SYU_Speck.fs.reserve = &H0                          ' 予約済み
'---------------------------------------------------        キー０
    
    O_DEL_SYU_Speck.ks0.keypos = 14                           ' キーポジション
    O_DEL_SYU_Speck.ks0.keyleng = 1                           ' キー長
    O_DEL_SYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks0.keytype = Chr(BtKtString)             ' キータイプ
    O_DEL_SYU_Speck.ks0.reserve = &H0                         ' 予約済み
    
    O_DEL_SYU_Speck.ks1.keypos = 15                           ' キーポジション
    O_DEL_SYU_Speck.ks1.keyleng = 1                           ' キー長
    O_DEL_SYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks1.keytype = Chr(BtKtString)             ' キータイプ
    O_DEL_SYU_Speck.ks1.reserve = &H0                         ' 予約済み
    
    O_DEL_SYU_Speck.ks2.keypos = 45                           ' キーポジション
    O_DEL_SYU_Speck.ks2.keyleng = 8                           ' キー長
    O_DEL_SYU_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks2.keytype = Chr(BtKtString)             ' キータイプ
    O_DEL_SYU_Speck.ks2.reserve = &H0                         ' 予約済み
    
    O_DEL_SYU_Speck.ks3.keypos = 53                           ' キーポジション
    O_DEL_SYU_Speck.ks3.keyleng = 8                           ' キー長
    O_DEL_SYU_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks3.keytype = Chr(BtKtString)             ' キータイプ
    O_DEL_SYU_Speck.ks3.reserve = &H0                         ' 予約済み
    
    O_DEL_SYU_Speck.ks4.keypos = 25                           ' キーポジション
    O_DEL_SYU_Speck.ks4.keyleng = 20                          ' キー長
    O_DEL_SYU_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks4.keytype = Chr(BtKtString)             ' キータイプ
    O_DEL_SYU_Speck.ks4.reserve = &H0                         ' 予約済み
    
    O_DEL_SYU_Speck.ks5.keypos = 61                           ' キーポジション
    O_DEL_SYU_Speck.ks5.keyleng = 8                           ' キー長
    O_DEL_SYU_Speck.ks5.keyflag = BtKfExt + BtKfDup           ' キーフラグ
    O_DEL_SYU_Speck.ks5.keytype = Chr(BtKtString)             ' キータイプ
    O_DEL_SYU_Speck.ks5.reserve = &H0                         ' 予約済み

'---------------------------------------------------        キー１
    
    O_DEL_SYU_Speck.ks6.keypos = 61                           ' キーポジション
    O_DEL_SYU_Speck.ks6.keyleng = 8                           ' キー長
    O_DEL_SYU_Speck.ks6.keyflag = BtKfExt + BtKfDup           ' キーフラグ
    O_DEL_SYU_Speck.ks6.keytype = Chr(BtKtString)             ' キータイプ
    O_DEL_SYU_Speck.ks6.reserve = &H0                         ' 予約済み
    
'---------------------------------------------------        キー２
    
    O_DEL_SYU_Speck.ks7.keypos = 14                           ' キーポジション
    O_DEL_SYU_Speck.ks7.keyleng = 1                           ' キー長
    O_DEL_SYU_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks7.keytype = Chr(BtKtString)             ' キータイプ
    O_DEL_SYU_Speck.ks7.reserve = &H0                         ' 予約済み
    
    O_DEL_SYU_Speck.ks8.keypos = 45                           ' キーポジション
    O_DEL_SYU_Speck.ks8.keyleng = 8                           ' キー長
    O_DEL_SYU_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks8.keytype = Chr(BtKtString)             ' キータイプ
    O_DEL_SYU_Speck.ks8.reserve = &H0                         ' 予約済み
    
    O_DEL_SYU_Speck.ks9.keypos = 53                           ' キーポジション
    O_DEL_SYU_Speck.ks9.keyleng = 8                           ' キー長
    O_DEL_SYU_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks9.keytype = Chr(BtKtString)             ' キータイプ
    O_DEL_SYU_Speck.ks9.reserve = &H0                         ' 予約済み
    
    O_DEL_SYU_Speck.ks10.keypos = 15                          ' キーポジション
    O_DEL_SYU_Speck.ks10.keyleng = 1                          ' キー長
    O_DEL_SYU_Speck.ks10.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks10.keytype = Chr(BtKtString)            ' キータイプ
    O_DEL_SYU_Speck.ks10.reserve = &H0                        ' 予約済み
    
    O_DEL_SYU_Speck.ks11.keypos = 24                          ' キーポジション
    O_DEL_SYU_Speck.ks11.keyleng = 1                          ' キー長
    O_DEL_SYU_Speck.ks11.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks11.keytype = Chr(BtKtString)            ' キータイプ
    O_DEL_SYU_Speck.ks11.reserve = &H0                        ' 予約済み
    
    O_DEL_SYU_Speck.ks12.keypos = 25                          ' キーポジション
    O_DEL_SYU_Speck.ks12.keyleng = 20                         ' キー長
    O_DEL_SYU_Speck.ks12.keyflag = BtKfExt + BtKfSeg + BtKfDup ' キーフラグ
    O_DEL_SYU_Speck.ks12.keytype = Chr(BtKtString)            ' キータイプ
    O_DEL_SYU_Speck.ks12.reserve = &H0                        ' 予約済み
    
    O_DEL_SYU_Speck.ks13.keypos = 16                          ' キーポジション
    O_DEL_SYU_Speck.ks13.keyleng = 8                          ' キー長
    O_DEL_SYU_Speck.ks13.keyflag = BtKfExt + BtKfDup          ' キーフラグ
    O_DEL_SYU_Speck.ks13.keytype = Chr(BtKtString)            ' キータイプ
    O_DEL_SYU_Speck.ks13.reserve = &H0                        ' 予約済み
    
    sts = BTRV(BtOpCreate, O_DEL_SYU_POS, O_DEL_SYU_Speck, Len(O_DEL_SYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "削除済み出荷予定データ")
        Exit Function
    End If

    O_DEL_SYU_Create = False

End Function
Function O_DEL_SYU_Open(Mode As Integer) As Integer
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
    
    O_DEL_SYU_Open = True
                                            '削除済み出荷予定データフルパス取込み
    sts = GetIni("FILE", O_DEL_SYU_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_DEL_SYU]読み込みエラー ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, O_DEL_SYU_POS, O_DEL_SYUREC, Len(O_DEL_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_DEL_SYU_Create()        '削除済み出荷予定データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_DEL_SYU_POS, O_DEL_SYUREC, Len(O_DEL_SYUREC), ByVal FullPath, Len(FullPath), Mode)
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
    
    O_DEL_SYU_Open = False

End Function
