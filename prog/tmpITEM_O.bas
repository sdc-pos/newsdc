Attribute VB_Name = "tmpITEM_O"
Option Explicit
'********************************************************************
'*
'*              大阪事　見積用品目マスタ  ファイル定義
'*
'*          CREATE 2016.05.24
'********************************************************************
'ファイルＩＤ
Public Const tmpITEM_O_ID$ = "tmpITEM_O"

'ページサイズ
Public Const tmpITEM_O_PG_SIZ% = 4096

'ポジション・ブロック
Public tmpITEM_O_POS               As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************



'データ・バッファ
Public tmpITEM_O_REC                As ITEM_O_REC_Tag

'キー定義
'キー・データ
Public K0_tmpITEM_O                 As KEY0_ITEM_O
Public K1_tmpITEM_O                 As KEY1_ITEM_O




Private tmpITEM_O_Speck            As ITEM_O_FSpeck

Private Function tmpITEM_O_Create() As Integer
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

    tmpITEM_O_Create = True
                                        '大阪事　見積用品目マスタ フルパス取込み
    sts = GetIni("FILE", tmpITEM_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpITEM_O]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    tmpITEM_O_Speck.fs.recoleng = Len(tmpITEM_O_REC)          ' レコード長
    tmpITEM_O_Speck.fs.PageSize = tmpITEM_O_PG_SIZ            ' ページサイズ
    tmpITEM_O_Speck.fs.idexnumb = 2                        ' インデックス数
    tmpITEM_O_Speck.fs.fileflag = 0                        ' ファイルフラグ
    tmpITEM_O_Speck.fs.reserve = &H0                       ' 予約済み
'-----------------------------------------------
                                                ' キー０
    tmpITEM_O_Speck.ks0.keypos = 1                         ' キーポジション
    tmpITEM_O_Speck.ks0.keyleng = 1                        ' キー長
    tmpITEM_O_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    tmpITEM_O_Speck.ks0.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks0.reserve = &H0                      ' 予約済み

    tmpITEM_O_Speck.ks1.keypos = 2                         ' キーポジション
    tmpITEM_O_Speck.ks1.keyleng = 1                        ' キー長
    tmpITEM_O_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    tmpITEM_O_Speck.ks1.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks1.reserve = &H0                      ' 予約済み

    tmpITEM_O_Speck.ks2.keypos = 3                         ' キーポジション
    tmpITEM_O_Speck.ks2.keyleng = 20                       ' キー長
    tmpITEM_O_Speck.ks2.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    tmpITEM_O_Speck.ks2.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks2.reserve = &H0                      ' 予約済み


    tmpITEM_O_Speck.ks3.keypos = 23                        ' キーポジション
    tmpITEM_O_Speck.ks3.keyleng = 3                        ' キー長
    tmpITEM_O_Speck.ks3.keyflag = BtKfExt                  ' キーフラグ
    tmpITEM_O_Speck.ks3.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks3.reserve = &H0                      ' 予約済み


'-----------------------------------------------
                                                ' キー１
    tmpITEM_O_Speck.ks4.keypos = 1                         ' キーポジション
    tmpITEM_O_Speck.ks4.keyleng = 1                        ' キー長
    tmpITEM_O_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' キーフラグ
    tmpITEM_O_Speck.ks4.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks4.reserve = &H0                      ' 予約済み

    tmpITEM_O_Speck.ks5.keypos = 2                         ' キーポジション
    tmpITEM_O_Speck.ks5.keyleng = 1                        ' キー長
    tmpITEM_O_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' キーフラグ
    tmpITEM_O_Speck.ks5.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks5.reserve = &H0                      ' 予約済み

    tmpITEM_O_Speck.ks6.keypos = 3                         ' キーポジション
    tmpITEM_O_Speck.ks6.keyleng = 20                       ' キー長
    tmpITEM_O_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' キーフラグ
    tmpITEM_O_Speck.ks6.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks6.reserve = &H0                      ' 予約済み

    tmpITEM_O_Speck.ks7.keypos = 26                         ' キーポジション
    tmpITEM_O_Speck.ks7.keyleng = 1                        ' キー長
    tmpITEM_O_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' キーフラグ
    tmpITEM_O_Speck.ks7.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks7.reserve = &H0                      ' 予約済み

    tmpITEM_O_Speck.ks8.keypos = 27                         ' キーポジション
    tmpITEM_O_Speck.ks8.keyleng = 1                        ' キー長
    tmpITEM_O_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' キーフラグ
    tmpITEM_O_Speck.ks8.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks8.reserve = &H0                      ' 予約済み

    tmpITEM_O_Speck.ks9.keypos = 28                         ' キーポジション
    tmpITEM_O_Speck.ks9.keyleng = 20                       ' キー長
    tmpITEM_O_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' キーフラグ
    tmpITEM_O_Speck.ks9.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks9.reserve = &H0                      ' 予約済み


    tmpITEM_O_Speck.ks10.keypos = 23                        ' キーポジション
    tmpITEM_O_Speck.ks10.keyleng = 3                        ' キー長
    tmpITEM_O_Speck.ks10.keyflag = BtKfExt + BtKfDup + BtKfChg              ' キーフラグ
    tmpITEM_O_Speck.ks10.keytype = Chr(BtKtString)          ' キータイプ
    tmpITEM_O_Speck.ks10.reserve = &H0                      ' 予約済み

'-----------------------------------------------
    sts = BTRV(BtOpCreate, tmpITEM_O_POS, tmpITEM_O_Speck, Len(tmpITEM_O_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "大阪事　見積用品目マスタ")
        Exit Function
    End If

    tmpITEM_O_Create = False

End Function

Public Function tmpITEM_O_Open(Mode As Integer) As Integer
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

    tmpITEM_O_Open = True
                                            '大阪事　見積用品目マスタ フルパス取込み
    sts = GetIni("FILE", tmpITEM_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpITEM_O]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpITEM_O_Create()    '大阪事　見積用品目マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), ByVal FullPath, Len(FullPath), Mode)
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

    tmpITEM_O_Open = False

End Function

Public Sub Rclr_tmpITEM_O_REC()

'********************************************************************
'*
'*              大阪事　見積用品目マスタ  レコード初期化
'*
'********************************************************************

    Call UniCode_Conv(tmpITEM_O_REC.JGYOBU, "")            '事業部区分
    Call UniCode_Conv(tmpITEM_O_REC.NAIGAI, "")            '国内外
    Call UniCode_Conv(tmpITEM_O_REC.HIN_GAI, "")           '品番（外部）


    Call UniCode_Conv(tmpITEM_O_REC.KO_JGYOBU, "")         '事業部区分     2017.09.28
    Call UniCode_Conv(tmpITEM_O_REC.KO_NAIGAI, "")         '国内外         2017.09.28
    Call UniCode_Conv(tmpITEM_O_REC.KO_HIN_GAI, "")        '品番（外部）   2017.09.28


    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_TANI, "")    '中西工料　単位
    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_KIN, "")     '中西工料　金額

    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_TANI, "")       '商品化工料　単位
    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_KIN, "")        '商品化工料　金額

    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_TANI, "")     'PF加工　単位
    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_KIN, "")      'PF加工　金額

    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_TANI, "")     'PE加工　単位
    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_KIN, "")      'PE加工　金額

    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_TANI, "")    'PF資材　単位
    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_KIN, "")     'PF資材　金額

    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_TANI, "") '品番表示ﾗﾍﾞﾙ　単位
    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_KIN, "") '品番表示ﾗﾍﾞﾙ　金額

    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_TANI, "")  '設置工事説明書　単位
    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_KIN, "")  '設置工事説明書　金額

    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_TANI, "")       '梱包材　単位
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_KIN, "")        '梱包材　金額

    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_TANI, "")  '副資材　単位
    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_KIN, "")   '副資材　金額

    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_TANI, "")  '梱包ASSY　単位
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_KIN, "")   '梱包ASSY　金額

    Call UniCode_Conv(tmpITEM_O_REC.KANRI_TANI, "")        '管理費　単位
    Call UniCode_Conv(tmpITEM_O_REC.KANRI_KIN, "")         '管理費　金額
    
    Call UniCode_Conv(tmpITEM_O_REC.GOUKEI_KIN, "")        '合計　金額
    Call UniCode_Conv(tmpITEM_O_REC.INPUT_TANTO_CODE, "")  '入力担当者ｺｰﾄﾞ
    
    
 '>>>>>>>>>>>>>>>>>>>>> 2017.09.27
    Call UniCode_Conv(tmpITEM_O_REC.BUZAI_TANTO_NAME, "")  '部材担当者名
    Call UniCode_Conv(tmpITEM_O_REC.T_HIN_NAME, "")        '提出品名
    Call UniCode_Conv(tmpITEM_O_REC.TANI, "")              '単位
    Call UniCode_Conv(tmpITEM_O_REC.T_TANKA, "")           '提出単位
    Call UniCode_Conv(tmpITEM_O_REC.T_KINGAKU, "")         '提出金額
    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_T_KIN, "")   '中西工料　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_T_KIN, "")      '商品化工料　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_T_KIN, "")    'PF加工　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_T_KIN, "")    'PE加工　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_T_KIN, "")   'PF資材　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_T_KIN, "") '品番表示ﾗﾍﾞﾙ　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_T_KIN, "") '設置工事説明書　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_T_KIN, "")      '梱包材　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_T_KIN, "") '副資材　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_T_KIN, "") '梱包ASSY　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.KANRI_T_KIN, "")       '管理費　提出金額
    Call UniCode_Conv(tmpITEM_O_REC.GOUKEI_T_KIN, "")      '提出合計金額

    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_F, "")       '中西工料 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_F, "")          '商品化工料 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_F, "")        'PF加工 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_F, "")        'PE加工 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_F, "")       'PF資材 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_F, "")    '品番表示ﾗﾍﾞﾙ 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_F, "")     '設置工事説明書 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_F, "")          '梱包材 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_F, "")     '副資材 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_F, "")     '梱包ASSY 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(tmpITEM_O_REC.KANRI_F, "")           '管理費 見積書表示ﾌﾗｸﾞ
 '>>>>>>>>>>>>>>>>>>>>> 2017.09.27
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    Call UniCode_Conv(tmpITEM_O_REC.KO_QTY, "")            '員数
    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_QTY, "")     '中西工料　数量
    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_QTY, "")        '商品化工料　数量
    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_QTY, "")      'PF加工　数量
    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_QTY, "")      'PE加工　数量
    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_QTY, "")      'PF資材　数量
    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_QTY, "")  '品番表示ﾗﾍﾞﾙ　数量
    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_QTY, "")   '設置工事説明書　数量
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_QTY, "")        '梱包材　数量
    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_QTY, "")   '副資材　数量
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_QTY, "")   '梱包ASSY　数量
    Call UniCode_Conv(tmpITEM_O_REC.KANRI_QTY, "")         '管理費　数量
    
    
    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_T_TAN, "")   '中西工料　提出単価
    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_T_TAN, "")      '商品化工料　提出単価
    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_T_TAN, "")    'PF加工　提出単価
    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_T_TAN, "")    'PE加工　提出単価
    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_T_TAN, "")   'PF資材　提出単価
    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_T_TAN, "") '品番表示ﾗﾍﾞﾙ　提出単価
    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_T_TAN, "") '設置工事説+明書　提出単価
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_T_TAN, "")      '梱包材　提出単価
    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_T_TAN, "") '副資材　提出単価
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_T_TAN, "") '梱包ASSY　提出単価
    Call UniCode_Conv(tmpITEM_O_REC.KANRI_T_TAN, "")       '管理費　提出単価
    
    
    
    Call UniCode_Conv(tmpITEM_O_REC.FILLER, "")
    
    Call UniCode_Conv(tmpITEM_O_REC.INS_TANTO, "")         '追加　担当者
    Call UniCode_Conv(tmpITEM_O_REC.Ins_DateTime, "")      '追加　日時
    Call UniCode_Conv(tmpITEM_O_REC.UPD_TANTO, "")         '更新　担当者
    Call UniCode_Conv(tmpITEM_O_REC.UPD_DATETIME, "")      '更新　日時



End Sub

