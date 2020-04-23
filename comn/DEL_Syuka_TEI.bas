Attribute VB_Name = "DEL_Syuka_TEI"
Option Explicit
'********************************************************************
'*
'*              削除済み邸別注文データ  ファイル定義
'*
'*          CREATE 2011.06.23
'********************************************************************
'ファイルＩＤ
Public Const DEL_SYU_TEI_ID$ = "DEL_SYU_TEI"

'ページサイズ
Public Const DEL_SYU_TEI_PG_SIZ% = 4096

'ポジション・ブロック
Public DEL_SYU_TEI_POS            As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type DEL_SYU_TEI_REC_Tag
    SND_YMD(0 To 7)                 As Byte         'データ作成日
    SND_HMS(0 To 5)                 As Byte         'データ作成時刻
    SEQ_NO(0 To 4)                  As Byte         '連番
    JUC_YMD(0 To 7)                 As Byte         '受注日
    NOU_CD(0 To 3)                  As Byte         '納入受入場
    NOU_NM(0 To 19)                 As Byte         '納入受入場名
    TOK_CD(0 To 7)                  As Byte         '得意先ｺｰﾄﾞ
    CHO_CD(0 To 7)                  As Byte         '直納先ｺｰﾄﾞ
    THINB_CD(0 To 19)               As Byte         '得意先品番　■品番(上)
    HINB_CD(0 To 19)                As Byte         '品番　      ■品番(下)
    CHU_CD(0 To 9)                  As Byte         '注文№　    ■指図№(上)
    SYU_JUN(0 To 9)                 As Byte         '出荷順番　  ■指図№(下・左)
    TEI_NM(0 To 29)                 As Byte         '邸名　      ■指図№(下・右)
    JUC_SUU(0 To 7)                 As Byte         '受注数量
    SYU_YMD(0 To 7)                 As Byte         '出荷確定日
    NOU_YMD(0 To 7)                 As Byte         '納入日
    KEN_NO(0 To 5)                  As Byte         '件管№　　　■管理№(上)
    HIN_NO(0 To 5)                  As Byte         '品管№　　　■管理№(下)
    TANP_KB(0 To 0)                 As Byte         '単品区分
    YOBI1_NM(0 To 54)               As Byte         '予備
    GSEQ_NO(0 To 4)                 As Byte         'ﾃﾞｰﾀ総件数
    TEI_LABELID(0 To 12)            As Byte         '邸別ﾗﾍﾞﾙID(注文№■指図№(上)+箱№)    2011.04.25
    HAKO_NO(0 To 2)                 As Byte         '箱№                                   2011.04.25
    JITU_SUU(0 To 7)                As Byte         '実出庫数(梱包場への出庫数 現在未使用)  2011.04.26
    JITU_TANTO(0 To 9)              As Byte         '出庫　担当者(現在未使用)               2011.04.26
    JITU_DATETIME(0 To 13)          As Byte         '出庫　日時(現在未使用)                 2011.04.26
    KONPO_TANTO(0 To 9)             As Byte         '梱包　担当者                           2011.04.26
    KONPO_DATETIME(0 To 13)         As Byte         '梱包　日時                             2011.04.26
    SHOGO_TANTO(0 To 9)             As Byte         '注文ﾃﾞｰﾀ照合担当                       2011.05.02
    SHOGO_DATETIME(0 To 13)         As Byte         '注文ﾃﾞｰﾀ照合日時                       2011.05.02
    
    L_KENKAN(0 To 11)               As Byte         '件管末番 long                          2011.05.06
    L_TEI_NAME(0 To 49)             As Byte         '邸名2 50                               2011.05.06
    L_TOK_NAME(0 To 49)             As Byte         '得意先名 50                            2011.05.06
    L_SOTO_NO(0 To 9)               As Byte         '外箱番号 50 → 10                      2011.05.06
    L_UCHI_NO(0 To 9)               As Byte         '内箱番号 50 → 10                      2011.05.06
    L_WIDTH(0 To 9)                 As Byte         '長さ(幅) 10                            2011.05.06
    L_HEIGHT(0 To 9)                As Byte         '高さ     20                            2011.05.06
    L_CONTENT(0 To 9)               As Byte         '体積     30                            2011.05.06
    L_KNo(0 To 1)                   As Byte         '工場No 2 32                            2011.05.06
    L_SERIES1(0 To 19)              As Byte         '品番シリーズ 20  52                    2011.05.06
    L_SERIES2(0 To 19)              As Byte         '品番シリーズ 2                         2011.05.06
    L_PAGE(0 To 4)                  As Byte         'ページ番号                             2011.05.06
    
    KUTI_SU(0 To 3)                 As Byte         '口数 9999  (邸別ﾗﾍﾞﾙID毎に同じ値)      2011.05.10
    SAI_SU(0 To 5)                  As Byte         '才数 999.99 (邸別ﾗﾍﾞﾙID毎に同じ値)     2011.05.10
    
    KONPO_ID(0 To 19)               As Byte         '梱包ID                                 2011.05.10
    
    
    KENPIN_TANTO(0 To 9)             As Byte        '検品担当者                             2011.05.12
    KENPIN_DATETIME(0 To 13)         As Byte        '検品日時                               2011.05.12
    
    
    SYUGO_KONPO_TANTO(0 To 9)       As Byte         '集合梱包担当者                         2011.05.12
    SYUGO_KONPO_DATETIME(0 To 13)   As Byte         '集合梱包日時                           2011.05.12
    
    
    
    FILLER(0 To 338)                As Byte         'FILLER                                 2011.05.12
    INS_TANTO(0 To 9)               As Byte         '追加　担当者
    Ins_DateTime(0 To 13)           As Byte         '追加　日時
    UPD_TANTO(0 To 9)               As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)           As Byte         '更新　日時



End Type
'データ・バッファ
Public DEL_SYU_TEI_REC                As DEL_SYU_TEI_REC_Tag

'キー定義

Type KEY0_DEL_SYU_TEI                 'ＫＥＹ０
    
    SND_YMD(0 To 7)                 As Byte         'データ作成日
    SND_HMS(0 To 5)                 As Byte         'データ作成時刻
    SEQ_NO(0 To 4)                  As Byte         '連番

End Type


Type KEY1_DEL_SYU_TEI                 'ＫＥＹ１
    
    TEI_LABELID(0 To 12)            As Byte         '邸別ﾗﾍﾞﾙID(注文№■指図№(上)+箱№)

End Type


Type KEY2_DEL_SYU_TEI                 'ＫＥＹ２
    
    KEN_NO(0 To 5)                  As Byte         '件管№　　　■管理№(上)
    HIN_NO(0 To 5)                  As Byte         '件管№　　　■管理№(下)

End Type


Type KEY3_DEL_SYU_TEI                 'ＫＥＹ３
    
    KONPO_ID(0 To 19)               As Byte         '梱包ID     2011.05.10

End Type





'キー・データ
Public K0_DEL_SYU_TEI                 As KEY0_DEL_SYU_TEI
Public K1_DEL_SYU_TEI                 As KEY1_DEL_SYU_TEI
Public K2_DEL_SYU_TEI                 As KEY2_DEL_SYU_TEI

Public K3_DEL_SYU_TEI                 As KEY3_DEL_SYU_TEI   '2011.05.12


Private Type DEL_SYU_TEI_FSpeck
    fs      As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck

    ks6     As BtKeySpeck                               '2011.05.12

End Type

Private DEL_SYU_TEI_Speck  As DEL_SYU_TEI_FSpeck

Private Function DEL_SYU_TEI_Create() As Integer
'********************************************************************
'*
'*              邸別注文データ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    DEL_SYU_TEI_Create = True
                                            '原産マスタフルパス取込み
    sts = GetIni("FILE", DEL_SYU_TEI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [DEL_SYU_TEI]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    DEL_SYU_TEI_Speck.fs.recoleng = Len(DEL_SYU_TEI_REC)    ' レコード長
    DEL_SYU_TEI_Speck.fs.PageSize = DEL_SYU_TEI_PG_SIZ      ' ページサイズ
    DEL_SYU_TEI_Speck.fs.idexnumb = 4                       ' インデックス数
    DEL_SYU_TEI_Speck.fs.fileflag = 0                       ' ファイルフラグ
    DEL_SYU_TEI_Speck.fs.reserve = &H0                      ' 予約済み
'-----------------------------------------------
                                                ' キー０
    DEL_SYU_TEI_Speck.ks0.keypos = 1                        ' キーポジション
    DEL_SYU_TEI_Speck.ks0.keyleng = 8                       ' キー長
                                                            ' キーフラグ
    DEL_SYU_TEI_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfSeg
    DEL_SYU_TEI_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    DEL_SYU_TEI_Speck.ks0.reserve = &H0                     ' 予約済み

    DEL_SYU_TEI_Speck.ks1.keypos = 9                        ' キーポジション
    DEL_SYU_TEI_Speck.ks1.keyleng = 6                       ' キー長
                                                            ' キーフラグ
    DEL_SYU_TEI_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfSeg
    DEL_SYU_TEI_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    DEL_SYU_TEI_Speck.ks1.reserve = &H0                     ' 予約済み

    DEL_SYU_TEI_Speck.ks2.keypos = 15                       ' キーポジション
    DEL_SYU_TEI_Speck.ks2.keyleng = 5                       ' キー長
                                                            ' キーフラグ
    DEL_SYU_TEI_Speck.ks2.keyflag = BtKfExt + BtKfDup
    DEL_SYU_TEI_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    DEL_SYU_TEI_Speck.ks2.reserve = &H0                     ' 予約済み




'-----------------------------------------------
                                                ' キー１
    DEL_SYU_TEI_Speck.ks3.keypos = 255                      ' キーポジション
    DEL_SYU_TEI_Speck.ks3.keyleng = 13                      ' キー長
                                                            ' キーフラグ
    DEL_SYU_TEI_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEL_SYU_TEI_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    DEL_SYU_TEI_Speck.ks3.reserve = &H0                     ' 予約済み



'-----------------------------------------------
                                                ' キー２
    DEL_SYU_TEI_Speck.ks4.keypos = 182                      ' キーポジション
    DEL_SYU_TEI_Speck.ks4.keyleng = 6                       ' キー長
                                                            ' キーフラグ
    DEL_SYU_TEI_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    DEL_SYU_TEI_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    DEL_SYU_TEI_Speck.ks4.reserve = &H0                     ' 予約済み

    DEL_SYU_TEI_Speck.ks5.keypos = 188                      ' キーポジション
    DEL_SYU_TEI_Speck.ks5.keyleng = 6                       ' キー長
                                                            ' キーフラグ
    DEL_SYU_TEI_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEL_SYU_TEI_Speck.ks5.keytype = Chr(BtKtString)         ' キータイプ
    DEL_SYU_TEI_Speck.ks5.reserve = &H0                     ' 予約済み


'-----------------------------------------------
                                                ' キー３
    DEL_SYU_TEI_Speck.ks6.keypos = 570                      ' キーポジション
    DEL_SYU_TEI_Speck.ks6.keyleng = 20                      ' キー長
                                                            ' キーフラグ
    DEL_SYU_TEI_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfDup
    DEL_SYU_TEI_Speck.ks6.keytype = Chr(BtKtString)         ' キータイプ
    DEL_SYU_TEI_Speck.ks6.reserve = &H0                     ' 予約済み



'-----------------------------------------------

    sts = BTRV(BtOpCreate, DEL_SYU_TEI_POS, DEL_SYU_TEI_Speck, Len(DEL_SYU_TEI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "邸別注文データ")
        Exit Function
    End If

    DEL_SYU_TEI_Create = False

End Function

Public Function DEL_SYU_TEI_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              邸別注文データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    DEL_SYU_TEI_Open = True
                                            '邸別注文データ フルパス取込み
    sts = GetIni("FILE", DEL_SYU_TEI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [DEL_SYU_TEI]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, DEL_SYU_TEI_POS, DEL_SYU_TEI_REC, Len(DEL_SYU_TEI_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = DEL_SYU_TEI_Create()        '邸別注文データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, DEL_SYU_TEI_POS, DEL_SYU_TEI_REC, Len(DEL_SYU_TEI_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "邸別注文データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "邸別注文データ")
                Exit Function
        End Select
    Loop

    DEL_SYU_TEI_Open = False

End Function

