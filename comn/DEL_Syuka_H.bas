Attribute VB_Name = "DEL_SYU_H"
Option Explicit
'********************************************************************
'*
'*              削除済出荷予定（ﾎｽﾄ）データ  ファイル定義
'*              大阪ＰＣ専用    2006.12.02
'*
'********************************************************************
'ファイルＩＤ
Public Const DEL_SYU_H_ID$ = "DEL_SYU_H"

'ページサイズ
Public Const DEL_SYU_H_PG_SIZ% = 4096

'ポジション・ブロック
Public DEL_SYU_H_POS    As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type DEL_SYU_HREC_Tag
    ID_NO(0 To 11)              As Byte     'ID_NO(有効 :伝票№ 7 桁+追番 2桁)
    SYUKA_NO(0 To 2)            As Byte     '№
    SYUKA_YMD(0 To 7)           As Byte     '出荷予定日
    OKURISAKI(0 To 39)          As Byte     '送り先名
    URIDEN(0 To 0)              As Byte     '売伝
    DEN_NO(0 To 6)              As Byte     '伝票№
    SEQ_NO(0 To 0)              As Byte     '追番
    HIN_NO(0 To 19)             As Byte     '品番
    SURYO(0 To 6)               As Byte     '出荷数量
    ODER_NO(0 To 9)             As Byte     '注文番号
    MUKE_CODE(0 To 7)           As Byte     '得意先コード
    MUKE_NAME(0 To 39)          As Byte     '得意先名称
    BIKOU(0 To 99)              As Byte     '備考
    UNSOU_KAISHA(0 To 39)       As Byte     '運送会社名
    
    INS_NOW(0 To 13)            As Byte     '取込み日時
    PRINT_NOW(0 To 13)          As Byte     '出荷ﾗﾍﾞﾙ印刷日時

    DATA_CNT(0 To 4)            As Byte     'ﾃﾞｰﾀ発生順
    
    OKURI_NO(0 To 19)           As Byte     '送り状№
    KENPIN_NOW(0 To 13)         As Byte     '検品日時
    KENPIN_TANTO_CODE(0 To 4)   As Byte     '検品担当者ｺｰﾄﾞ

    xKUTI_SU(0 To 1)            As Byte     '口数   '2007.02.01 未使用
    
    KYOSEI_END(0 To 0)          As Byte     '強制完了ﾌﾗｸﾞ

    CANCEL_F(0 To 0)            As Byte     'ｷｬﾝｾﾙﾌﾗｸﾞ
    
    INPUT_BIKOU(0 To 59)        As Byte     '入力備考
    
    INS_BIN(0 To 1)             As Byte     '便
    
    KUTI_SU(0 To 3)             As Byte     '口数   '2007.02.01 桁数変更により新設
    
    
    
    JGYOBU(0 To 0)              As Byte     '事業部     2007.03.14
    NAIGAI(0 To 0)              As Byte     '国内外     2007.03.14
    
    SYU_NO(0 To 11)             As Byte     '出庫表№   2007.03.14
    J_SURYO(0 To 6)             As Byte     '出庫実績数 2007.03.14
    
    
    COL_OKURISAKI_CD(0 To 19)   As Byte     '集約送り先CD   2007.07.07
    OKURISAKI_CD(0 To 8)        As Byte     '送り先CD       2007.07.07
    
    JYUSHO(0 To 159)            As Byte     '住所       2009.11.19
    
    TEL_NO(0 To 19)             As Byte     '電話番号   2010.01.21
    YUBIN_NO(0 To 7)            As Byte     '郵便番号   2010.01.21
        
    JURYO(0 To 5)               As Byte     '重量       2010.01.21
    SAI_SU(0 To 5)              As Byte     '才数       2010.01.21
    
    
    OKURI_NO_SEQ(0 To 2)        As Byte     '送り状№　枝番　2010.01.21
    
    
    KONPOU_F(0 To 0)            As Byte     '梱包区分       2010.01.18
    KUTI_SU_TAN(0 To 5)         As Byte     '口数(単体)     2010.01.21
    SAI_SU_TAN(0 To 5)          As Byte     '才数(単体)     2010.01.21
    
    OKURI_NO_SEQ_TO(0 To 2)     As Byte     '送り状№　枝番　2010.01.21
    
    
    SAI_SU_TAN_SAV(0 To 5)      As Byte     '才数(単体:修正不可)    2010.11.01
    SAI_SU_CALC(0 To 5)         As Byte     '才数計算値(梱包単位)   2010.11.01
    
    
    KUTI_SU_CALC(0 To 5)        As Byte     '口数計算値(梱包単位)   2010.11.9
    
    SEK_KEN_NO(0 To 5)          As Byte     '件管№　　　■管理№(上)   2011.04.30
    SEK_HIN_NO(0 To 5)          As Byte     '品管№　　　■管理№(下)   2011.04.30
    
    SEK_SHOGO_TANTO(0 To 9)     As Byte     '注文ﾃﾞｰﾀ照合担当       2011.05.02
    SEK_SHOGO_DATETIME(0 To 13) As Byte     '注文ﾃﾞｰﾀ照合日時       2011.05.02
    
    
    CNT_BARA_SU(0 To 6)         As Byte     '検品実績　バラ     2012.10.02
    CNT_HAKO_SU(0 To 6)         As Byte     '検品実績　箱       2012.10.02
    
    GAISO_IRI_QTY(0 To 7)       As Byte     '外装入り数         2012.10.02
    
    
    Y_HIN_CHK_CNT(0 To 5)       As Byte     '品番読込み回数     2012.10.02
    J_HIN_CHK_CNT(0 To 5)       As Byte     '品番読込み済み回数 2012.10.02
    
    KEN_HINBAN(0 To 19)         As Byte     '検品中品番         2012.10.02
    
    FILLER(0 To 159)            As Byte     'FILLER             2012.10.02 (157


    INS_TANTO(0 To 9)           As Byte     '追加　担当者       2011.05.06
    Ins_DateTime(0 To 13)       As Byte     '追加　日時         2011.05.06
    UPD_TANTO(0 To 9)           As Byte     '更新　担当者       2011.05.06
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時         2011.05.06
End Type

'データ・バッファ
Public DEL_SYU_HREC               As DEL_SYU_HREC_Tag

'キー定義
Type KEY0_DEL_SYU_H            'ＫＥＹ０
    DEN_NO(0 To 6)              As Byte     '伝票№
    SEQ_NO(0 To 0)              As Byte     '追番
End Type

Type KEY1_DEL_SYU_H            'ＫＥＹ１
    PRINT_NOW(0 To 13)          As Byte     '出荷ﾗﾍﾞﾙ印刷日時
    INS_NOW(0 To 13)            As Byte     '取込み日時
    DATA_CNT(0 To 4)            As Byte     'ﾃﾞｰﾀ発生順
End Type

Type KEY2_DEL_SYU_H            'ＫＥＹ２
    OKURI_NO(0 To 19)           As Byte     '送り状№
End Type

Type KEY3_DEL_SYU_H            'ＫＥＹ３
    SYUKA_YMD(0 To 7)           As Byte     '出荷予定日
End Type

Type KEY4_DEL_SYU_H            'ＫＥＹ４
    ID_NO(0 To 11)              As Byte     'ID_NO(有効 :伝票№ 7 桁+追番 2桁)
End Type



'キー・データ
Public K0_DEL_SYU_H             As KEY0_DEL_SYU_H
Public K1_DEL_SYU_H             As KEY1_DEL_SYU_H
Public K2_DEL_SYU_H             As KEY2_DEL_SYU_H
Public K3_DEL_SYU_H             As KEY3_DEL_SYU_H
Public K4_DEL_SYU_H             As KEY4_DEL_SYU_H

Type DEL_SYU_H_FSpeck
    fs      As BtFileSpeck                  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks7     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private DEL_SYU_H_Speck As DEL_SYU_H_FSpeck

Private Function DEL_SYU_H_Create() As Integer
'********************************************************************
'*
'*              出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    DEL_SYU_H_Create = True
                                            '出荷予定データフルパス取込み
    sts = GetIni("FILE", DEL_SYU_H_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [DEL_SYU_H]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    DEL_SYU_H_Speck.fs.recoleng = Len(DEL_SYU_HREC)     ' レコード長
    DEL_SYU_H_Speck.fs.PageSize = DEL_SYU_H_PG_SIZ      ' ページサイズ
    DEL_SYU_H_Speck.fs.idexnumb = 5                     ' インデックス数
    DEL_SYU_H_Speck.fs.fileflag = 0                     ' ファイルフラグ
    DEL_SYU_H_Speck.fs.reserve = &H0                    ' 予約済み
'---------------------------------------------------' キー０
    DEL_SYU_H_Speck.ks0.keypos = 65                     ' キーポジション
    DEL_SYU_H_Speck.ks0.keyleng = 7                     ' キー長
                                                        ' キーフラグ
    DEL_SYU_H_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    DEL_SYU_H_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    DEL_SYU_H_Speck.ks0.reserve = &H0                   ' 予約済み
    
    DEL_SYU_H_Speck.ks1.keypos = 72                     ' キーポジション
    DEL_SYU_H_Speck.ks1.keyleng = 1                     ' キー長
                                                        ' キーフラグ
    DEL_SYU_H_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfChg
    DEL_SYU_H_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    DEL_SYU_H_Speck.ks1.reserve = &H0                   ' 予約済み
'---------------------------------------------------' キー０
    
'---------------------------------------------------' キー１
    DEL_SYU_H_Speck.ks2.keypos = 312                    ' キーポジション
    DEL_SYU_H_Speck.ks2.keyleng = 14                    ' キー長
                                                        ' キーフラグ
    DEL_SYU_H_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    DEL_SYU_H_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    DEL_SYU_H_Speck.ks2.reserve = &H0                   ' 予約済み

    
    DEL_SYU_H_Speck.ks3.keypos = 298                    ' キーポジション
    DEL_SYU_H_Speck.ks3.keyleng = 14                    ' キー長
                                                        ' キーフラグ
    DEL_SYU_H_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    DEL_SYU_H_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    DEL_SYU_H_Speck.ks3.reserve = &H0                   ' 予約済み
    
    DEL_SYU_H_Speck.ks4.keypos = 326                    ' キーポジション
    DEL_SYU_H_Speck.ks4.keyleng = 5                     ' キー長
                                                        ' キーフラグ
    DEL_SYU_H_Speck.ks4.keyflag = BtKfExt + BtKfDup + BtKfChg
    DEL_SYU_H_Speck.ks4.keytype = Chr(BtKtString)       ' キータイプ
    DEL_SYU_H_Speck.ks4.reserve = &H0                   ' 予約済み
'---------------------------------------------------' キー１
'---------------------------------------------------' キー２
    DEL_SYU_H_Speck.ks5.keypos = 331                    ' キーポジション
    DEL_SYU_H_Speck.ks5.keyleng = 20                    ' キー長
                                                        ' キーフラグ
    DEL_SYU_H_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg
    DEL_SYU_H_Speck.ks5.keytype = Chr(BtKtString)       ' キータイプ
    DEL_SYU_H_Speck.ks5.reserve = &H0                   ' 予約済み
'---------------------------------------------------' キー２
'---------------------------------------------------' キー３
    DEL_SYU_H_Speck.ks6.keypos = 16                      ' キーポジション
    DEL_SYU_H_Speck.ks6.keyleng = 8                     ' キー長
                                                        ' キーフラグ
    DEL_SYU_H_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg
    DEL_SYU_H_Speck.ks6.keytype = Chr(BtKtString)       ' キータイプ
    DEL_SYU_H_Speck.ks6.reserve = &H0                   ' 予約済み


'---------------------------------------------------' キー３
'---------------------------------------------------' キー４
    DEL_SYU_H_Speck.ks7.keypos = 1                    ' キーポジション
    DEL_SYU_H_Speck.ks7.keyleng = 12                  ' キー長
                                                    ' キーフラグ
'    DEL_SYU_H_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfChg
    DEL_SYU_H_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfChg
    DEL_SYU_H_Speck.ks7.keytype = Chr(BtKtString)     ' キータイプ
    DEL_SYU_H_Speck.ks7.reserve = &H0                 ' 予約済み
'---------------------------------------------------' キー３
    sts = BTRV(BtOpCreate, DEL_SYU_H_POS, DEL_SYU_H_Speck, Len(DEL_SYU_H_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
        Exit Function
    End If

    DEL_SYU_H_Create = False

End Function

Function DEL_SYU_H_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    DEL_SYU_H_Open = True
                                            '出荷予定データフルパス取込み
    sts = GetIni("FILE", DEL_SYU_H_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [DEL_SYU_H]読み込みエラー ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = DEL_SYU_H_Create()        '出荷予定データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
                Exit Function
        End Select
    Loop
    DEL_SYU_H_Open = False
End Function
