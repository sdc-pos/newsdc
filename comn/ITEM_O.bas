Attribute VB_Name = "ITEM_O"
Option Explicit
'********************************************************************
'*
'*              大阪事　見積用品目マスタ  ファイル定義
'*
'*          CREATE 2016.05.24
'********************************************************************
'ファイルＩＤ
Public Const ITEM_O_ID$ = "ITEM_O"

'ページサイズ
Public Const ITEM_O_PG_SIZ% = 4096

'ポジション・ブロック
Public ITEM_O_POS               As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************



'レコード定義
Type ITEM_O_REC_Tag
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番(外部)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    SEQ_NO(0 To 2)              As Byte     'SEQ_NO
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    KO_JGYOBU(0 To 0)           As Byte     '子　事業部区分
    KO_NAIGAI(0 To 0)           As Byte     '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte     '子　品番(外部)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    
    
    NAKANISHI_TANI(0 To 3)      As Byte     '中西工料　単位
    NAKANISHI_KIN(0 To 10)      As Byte     '中西工料　金額

    SHOHIN_TANI(0 To 3)         As Byte     '商品化工料　単位
    SHOHIN_KIN(0 To 10)         As Byte     '商品化工料　金額

    PF_KAKOU_TANI(0 To 3)       As Byte     'PF加工　単位
    PF_KAKOU_KIN(0 To 10)       As Byte     'PF加工　金額

    PE_KAKOU_TANI(0 To 3)       As Byte     'PE加工　単位
    PE_KAKOU_KIN(0 To 10)       As Byte     'PE加工　金額

    PE_SHIZAI_TANI(0 To 3)      As Byte     'PF資材　単位
    PE_SHIZAI_KIN(0 To 10)      As Byte     'PF資材　金額

    HINBAN_LABEL_TANI(0 To 3)   As Byte     '品番表示ﾗﾍﾞﾙ　単位
    HINBAN_LABEL_KIN(0 To 10)   As Byte     '品番表示ﾗﾍﾞﾙ　金額

    KOUJI_SETSU_TANI(0 To 3)    As Byte     '設置工事説明書　単位
    KOUJI_SETSU_KIN(0 To 10)    As Byte     '設置工事説明書　金額

    KONPOU_TANI(0 To 3)         As Byte     '梱包材　単位
    KONPOU_KIN(0 To 10)         As Byte     '梱包材　金額

    FUKU_SHIZAI_TANI(0 To 3)    As Byte     '副資材　単位
    FUKU_SHIZAI_KIN(0 To 10)    As Byte     '副資材　金額

    KONPOU_ASSY_TANI(0 To 3)    As Byte     '梱包ASSY　単位
    KONPOU_ASSY_KIN(0 To 10)    As Byte     '梱包ASSY　金額

    KANRI_TANI(0 To 3)          As Byte     '管理費　単位
    KANRI_KIN(0 To 10)          As Byte     '管理費　金額
    
    GOUKEI_KIN(0 To 10)         As Byte     '合計金額
    
        
    
    INPUT_TANTO_CODE(0 To 4)    As Byte     '入力担当者ｺｰﾄﾞ
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    BUZAI_TANTO_NAME(0 To 19)   As Byte     '部材担当者名
    T_HIN_NAME(0 To 39)         As Byte     '提出品名
    TANI(0 To 3)                As Byte     '単位
    T_TANKA(0 To 10)            As Byte     '提出単価
    T_KINGAKU(0 To 10)          As Byte     '提出金額
    NAKANISHI_T_KIN(0 To 10)    As Byte     '中西工料　提出金額
    SHOHIN_T_KIN(0 To 10)       As Byte     '商品化工料　提出金額
    PF_KAKOU_T_KIN(0 To 10)     As Byte     'PF加工　提出金額
    PE_KAKOU_T_KIN(0 To 10)     As Byte     'PE加工　提出金額
    PE_SHIZAI_T_KIN(0 To 10)    As Byte     'PF資材　提出金額
    HINBAN_LABEL_T_KIN(0 To 10) As Byte     '品番表示ﾗﾍﾞﾙ　提出金額
    KOUJI_SETSU_T_KIN(0 To 10)  As Byte     '設置工事説明書　提出金額
    KONPOU_T_KIN(0 To 10)       As Byte     '梱包材　提出金額
    FUKU_SHIZAI_T_KIN(0 To 10)  As Byte     '副資材　提出金額
    KONPOU_ASSY_T_KIN(0 To 10)  As Byte     '梱包ASSY　提出金額
    KANRI_T_KIN(0 To 10)        As Byte     '管理費　提出金額
    GOUKEI_T_KIN(0 To 10)       As Byte     '提出合計金額

    NAKANISHI_F(0 To 0)         As Byte     '中西工料 見積書表示ﾌﾗｸﾞ
    SHOHIN_F(0 To 0)            As Byte     '商品化工料 見積書表示ﾌﾗｸﾞ
    PF_KAKOU_F(0 To 0)          As Byte     'PF加工 見積書表示ﾌﾗｸﾞ
    PE_KAKOU_F(0 To 0)          As Byte     'PE加工 見積書表示ﾌﾗｸﾞ
    PE_SHIZAI_F(0 To 0)         As Byte     'PF資材 見積書表示ﾌﾗｸﾞ
    HINBAN_LABEL_F(0 To 0)      As Byte     '品番表示ﾗﾍﾞﾙ 見積書表示ﾌﾗｸﾞ
    KOUJI_SETSU_F(0 To 0)       As Byte     '設置工事説明書 見積書表示ﾌﾗｸﾞ
    KONPOU_F(0 To 0)            As Byte     '梱包材 見積書表示ﾌﾗｸﾞ
    FUKU_SHIZAI_F(0 To 0)       As Byte     '副資材 見積書表示ﾌﾗｸﾞ
    KONPOU_ASSY_F(0 To 0)       As Byte     '梱包ASSY 見積書表示ﾌﾗｸﾞ
    KANRI_F(0 To 0)             As Byte     '管理費 見積書表示ﾌﾗｸﾞ

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    KO_QTY(0 To 5)              As Byte     '員数
    NAKANISHI_QTY(0 To 5)       As Byte     '中西工料　数量
    SHOHIN_QTY(0 To 5)          As Byte     '商品化工料　数量
    PF_KAKOU_QTY(0 To 5)        As Byte     'PF加工　数量
    PE_KAKOU_QTY(0 To 5)        As Byte     'PE加工　数量
    PE_SHIZAI_QTY(0 To 5)       As Byte     'PF資材　数量
    HINBAN_LABEL_QTY(0 To 5)    As Byte     '品番表示ﾗﾍﾞﾙ　数量
    KOUJI_SETSU_QTY(0 To 5)     As Byte     '設置工事説明書　数量
    KONPOU_QTY(0 To 5)          As Byte     '梱包材　数量
    FUKU_SHIZAI_QTY(0 To 5)     As Byte     '副資材　数量
    KONPOU_ASSY_QTY(0 To 5)     As Byte     '梱包ASSY　数量
    KANRI_QTY(0 To 5)           As Byte     '管理費　数量
    
    
    NAKANISHI_T_TAN(0 To 10)    As Byte     '中西工料　提出単価
    SHOHIN_T_TAN(0 To 10)       As Byte     '商品化工料　提出単価
    PF_KAKOU_T_TAN(0 To 10)     As Byte     'PF加工　提出単価
    PE_KAKOU_T_TAN(0 To 10)     As Byte     'PE加工　提出単価
    PE_SHIZAI_T_TAN(0 To 10)    As Byte     'PF資材　提出単価
    HINBAN_LABEL_T_TAN(0 To 10) As Byte     '品番表示ﾗﾍﾞﾙ　提出単価
    KOUJI_SETSU_T_TAN(0 To 10)  As Byte     '設置工事説明書　提出単価
    KONPOU_T_TAN(0 To 10)       As Byte     '梱包材　提出単価
    FUKU_SHIZAI_T_TAN(0 To 10)  As Byte     '副資材　提出単価
    KONPOU_ASSY_T_TAN(0 To 10)  As Byte     '梱包ASSY　提出単価
    KANRI_T_TAN(0 To 10)        As Byte     '管理費　提出単価
    
    KO_SYUBETSU(0 To 1)         As Byte         '子　種別
    
    FILLER(0 To 323)            As Byte
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    
    INS_TANTO(0 To 9)           As Byte     '追加　担当者
    Ins_DateTime(0 To 13)       As Byte     '追加　日時
    UPD_TANTO(0 To 9)           As Byte     '更新　担当者
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時

End Type
'データ・バッファ
Public ITEM_O_REC               As ITEM_O_REC_Tag

'キー定義

Type KEY0_ITEM_O                'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番(外部)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    SEQ_NO(0 To 2)              As Byte     'SEQ_NO
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28-->2017.11.07
'    KO_JGYOBU(0 To 0)           As Byte     '子　事業部区分
'    KO_NAIGAI(0 To 0)           As Byte     '子　国内外
'    KO_HIN_GAI(0 To 19)         As Byte     '子　品番(外部)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28-->2017.11.07

End Type


Type KEY1_ITEM_O                'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番(外部)

    KO_JGYOBU(0 To 0)           As Byte     '子　事業部区分
    KO_NAIGAI(0 To 0)           As Byte     '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte     '子　品番(外部)

    SEQ_NO(0 To 2)              As Byte     'SEQ_NO

End Type




'キー・データ
Public K0_ITEM_O                As KEY0_ITEM_O
Public K1_ITEM_O                As KEY1_ITEM_O

Type ITEM_O_FSpeck
    fs      As BtFileSpeck                  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2017.09.28
    
    ks4     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2017.09.28
    ks5     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2017.09.28
    ks6     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2017.09.28
    ks7     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2017.09.28
    ks8     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2017.09.28
    ks9     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2017.09.28
    ks10     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体    2017.09.28

End Type

Private ITEM_O_Speck            As ITEM_O_FSpeck

Private Function ITEM_O_Create() As Integer
'********************************************************************
'*
'*              大阪事　見積用品目マスタ  CREATE
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************

Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ITEM_O_Create = True
                                        '大阪事　見積用品目マスタ フルパス取込み
    sts = GetIni("FILE", ITEM_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_O]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    ITEM_O_Speck.fs.recoleng = Len(ITEM_O_REC)          ' レコード長
    ITEM_O_Speck.fs.PageSize = ITEM_O_PG_SIZ            ' ページサイズ
    ITEM_O_Speck.fs.idexnumb = 1                        ' インデックス数
    ITEM_O_Speck.fs.fileflag = 0                        ' ファイルフラグ
    ITEM_O_Speck.fs.reserve = &H0                       ' 予約済み
'-----------------------------------------------
                                                ' キー０
    ITEM_O_Speck.ks0.keypos = 1                         ' キーポジション
    ITEM_O_Speck.ks0.keyleng = 1                        ' キー長
    ITEM_O_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    ITEM_O_Speck.ks0.keytype = Chr(BtKtString)          ' キータイプ
    ITEM_O_Speck.ks0.reserve = &H0                      ' 予約済み

    ITEM_O_Speck.ks1.keypos = 2                         ' キーポジション
    ITEM_O_Speck.ks1.keyleng = 1                        ' キー長
    ITEM_O_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    ITEM_O_Speck.ks1.keytype = Chr(BtKtString)          ' キータイプ
    ITEM_O_Speck.ks1.reserve = &H0                      ' 予約済み

    ITEM_O_Speck.ks2.keypos = 3                         ' キーポジション
    ITEM_O_Speck.ks2.keyleng = 20                       ' キー長
    ITEM_O_Speck.ks2.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    ITEM_O_Speck.ks2.keytype = Chr(BtKtString)          ' キータイプ
    ITEM_O_Speck.ks2.reserve = &H0                      ' 予約済み


    ITEM_O_Speck.ks3.keypos = 23                        ' キーポジション
    ITEM_O_Speck.ks3.keyleng = 3                        ' キー長
    ITEM_O_Speck.ks3.keyflag = BtKfExt                  ' キーフラグ
    ITEM_O_Speck.ks3.keytype = Chr(BtKtString)          ' キータイプ
    ITEM_O_Speck.ks3.reserve = &H0                      ' 予約済み



'-----------------------------------------------
    sts = BTRV(BtOpCreate, ITEM_O_POS, ITEM_O_Speck, Len(ITEM_O_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "大阪事　見積用品目マスタ")
        Exit Function
    End If

    ITEM_O_Create = False

End Function

Public Function ITEM_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              大阪事　見積用品目マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_O_Open = True
                                            '大阪事　見積用品目マスタ フルパス取込み
    sts = GetIni("FILE", ITEM_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_O]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_O_Create()    '大阪事　見積用品目マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "大阪事　見積用品目マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "大阪事　見積用品目マスタ")
                Exit Function
        End Select
    Loop

    ITEM_O_Open = False

End Function

Public Sub Rclr_ITEM_O_REC()

'********************************************************************
'*
'*              大阪事　見積用品目マスタ  レコード初期化
'*
'********************************************************************

    Call UniCode_Conv(ITEM_O_REC.JGYOBU, "")            '事業部区分
    Call UniCode_Conv(ITEM_O_REC.NAIGAI, "")            '国内外
    Call UniCode_Conv(ITEM_O_REC.HIN_GAI, "")           '品番（外部）


    Call UniCode_Conv(ITEM_O_REC.KO_JGYOBU, "")         '事業部区分     2017.09.28
    Call UniCode_Conv(ITEM_O_REC.KO_NAIGAI, "")         '国内外         2017.09.28
    Call UniCode_Conv(ITEM_O_REC.KO_HIN_GAI, "")        '品番（外部）   2017.09.28


    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_TANI, "")    '中西工料　単位
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_KIN, "")     '中西工料　金額

    Call UniCode_Conv(ITEM_O_REC.SHOHIN_TANI, "")       '商品化工料　単位
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_KIN, "")        '商品化工料　金額

    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_TANI, "")     'PF加工　単位
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_KIN, "")      'PF加工　金額

    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_TANI, "")     'PE加工　単位
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_KIN, "")      'PE加工　金額

    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_TANI, "")    'PF資材　単位
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_KIN, "")     'PF資材　金額

    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_TANI, "") '品番表示ﾗﾍﾞﾙ　単位
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_KIN, "") '品番表示ﾗﾍﾞﾙ　金額

    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_TANI, "")  '設置工事説明書　単位
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_KIN, "")  '設置工事説明書　金額

    Call UniCode_Conv(ITEM_O_REC.KONPOU_TANI, "")       '梱包材　単位
    Call UniCode_Conv(ITEM_O_REC.KONPOU_KIN, "")        '梱包材　金額

    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_TANI, "")  '副資材　単位
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_KIN, "")   '副資材　金額

    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_TANI, "")  '梱包ASSY　単位
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_KIN, "")   '梱包ASSY　金額

    Call UniCode_Conv(ITEM_O_REC.KANRI_TANI, "")        '管理費　単位
    Call UniCode_Conv(ITEM_O_REC.KANRI_KIN, "")         '管理費　金額
    
    Call UniCode_Conv(ITEM_O_REC.GOUKEI_KIN, "")        '合計　金額
    Call UniCode_Conv(ITEM_O_REC.INPUT_TANTO_CODE, "")  '入力担当者ｺｰﾄﾞ
    
    
 '>>>>>>>>>>>>>>>>>>>>> 2017.09.27
    Call UniCode_Conv(ITEM_O_REC.BUZAI_TANTO_NAME, "")  '部材担当者名
    Call UniCode_Conv(ITEM_O_REC.T_HIN_NAME, "")        '提出品名
    Call UniCode_Conv(ITEM_O_REC.TANI, "")              '単位
    Call UniCode_Conv(ITEM_O_REC.T_TANKA, "")           '提出単位
    Call UniCode_Conv(ITEM_O_REC.T_KINGAKU, "")         '提出金額
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_T_KIN, "")   '中西工料　提出金額
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_T_KIN, "")      '商品化工料　提出金額
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_T_KIN, "")    'PF加工　提出金額
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_T_KIN, "")    'PE加工　提出金額
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_T_KIN, "")   'PF資材　提出金額
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_T_KIN, "") '品番表示ﾗﾍﾞﾙ　提出金額
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_T_KIN, "") '設置工事説明書　提出金額
    Call UniCode_Conv(ITEM_O_REC.KONPOU_T_KIN, "")      '梱包材　提出金額
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_T_KIN, "") '副資材　提出金額
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_T_KIN, "") '梱包ASSY　提出金額
    Call UniCode_Conv(ITEM_O_REC.KANRI_T_KIN, "")       '管理費　提出金額
    Call UniCode_Conv(ITEM_O_REC.GOUKEI_T_KIN, "")      '提出合計金額

    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_F, "")       '中西工料 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_F, "")          '商品化工料 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_F, "")        'PF加工 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_F, "")        'PE加工 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_F, "")       'PF資材 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_F, "")    '品番表示ﾗﾍﾞﾙ 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_F, "")     '設置工事説明書 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.KONPOU_F, "")          '梱包材 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_F, "")     '副資材 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_F, "")     '梱包ASSY 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.KANRI_F, "")           '管理費 見積書表示ﾌﾗｸﾞ
 '>>>>>>>>>>>>>>>>>>>>> 2017.09.27
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    Call UniCode_Conv(ITEM_O_REC.KO_QTY, "")            '員数
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_QTY, "")     '中西工料　数量
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_QTY, "")        '商品化工料　数量
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_QTY, "")      'PF加工　数量
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_QTY, "")      'PE加工　数量
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_QTY, "")     'PF資材　数量
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_QTY, "")  '品番表示ﾗﾍﾞﾙ　数量
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_QTY, "")   '設置工事説明書　数量
    Call UniCode_Conv(ITEM_O_REC.KONPOU_QTY, "")        '梱包材　数量
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_QTY, "")   '副資材　数量
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_QTY, "")   '梱包ASSY　数量
    Call UniCode_Conv(ITEM_O_REC.KANRI_QTY, "")         '管理費　数量
    
    
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_T_TAN, "")   '中西工料　提出単価
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_T_TAN, "")      '商品化工料　提出単価
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_T_TAN, "")    'PF加工　提出単価
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_T_TAN, "")    'PE加工　提出単価
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_T_TAN, "")   'PF資材　提出単価
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_T_TAN, "") '品番表示ﾗﾍﾞﾙ　提出単価
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_T_TAN, "") '設置工事説+明書　提出単価
    Call UniCode_Conv(ITEM_O_REC.KONPOU_T_TAN, "")      '梱包材　提出単価
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_T_TAN, "") '副資材　提出単価
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_T_TAN, "") '梱包ASSY　提出単価
    Call UniCode_Conv(ITEM_O_REC.KANRI_T_TAN, "")       '管理費　提出単価
    
    
    
    Call UniCode_Conv(ITEM_O_REC.FILLER, "")
    
    Call UniCode_Conv(ITEM_O_REC.INS_TANTO, "")         '追加　担当者
    Call UniCode_Conv(ITEM_O_REC.Ins_DateTime, "")      '追加　日時
    Call UniCode_Conv(ITEM_O_REC.UPD_TANTO, "")         '更新　担当者
    Call UniCode_Conv(ITEM_O_REC.UPD_DATETIME, "")      '更新　日時



End Sub
