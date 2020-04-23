Attribute VB_Name = "O_IDO"
Option Explicit
'********************************************************************
'*
'*              在庫移動歴　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const O_IDO_ID$ = "O_IDO"

'ページサイズ
Public Const O_IDO_PG_SIZ% = 1024

'ポジション・ブロック
Public O_IDO_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type O_IDOREC_Tag
    JITU_DT(0 To 7)                     As Byte     '実績日付
    JITU_TM(0 To 5)                     As Byte     '実績時刻
    JGYOBU(0 To 0)                      As Byte     '事業部区分
    NAIGAI(0 To 0)                      As Byte     '国内外
    HIN_GAI(0 To 19)                    As Byte     '品番（外部）
    RIRK_ID(0 To 1)                     As Byte     '履歴種別
    SUMI_JITU_QTY(0 To 7)               As Byte     '実績数量(商品化済み)
    MI_JITU_QTY(0 To 7)                 As Byte     '実績数量(未商品)
    FROM_SOKO(0 To 1)                   As Byte     'From 倉庫№
    FROM_RETU(0 To 1)                   As Byte     '   　列
    FROM_REN(0 To 1)                    As Byte     '   　連
    FROM_DAN(0 To 1)                    As Byte     '   　段
    TO_SOKO(0 To 1)                     As Byte     'ＴＯ 倉庫№
    TO_RETU(0 To 1)                     As Byte     '   　列
    TO_REN(0 To 1)                      As Byte     '   　連
    TO_DAN(0 To 1)                      As Byte     '   　段
    DEN_DT(0 To 7)                      As Byte     '伝票日付
    DEN_NO(0 To 9)                      As Byte     '伝票№
    PRG_ID(0 To 7)                      As Byte     '出力元プログラム
    HIN_NAI(0 To 19)                    As Byte     '品番（内部）
    NYUKA_DT(0 To 7)                    As Byte     '入荷日付
    NYUKO_DT(0 To 7)                    As Byte     '入庫日付
    WEL_ID(0 To 2)                      As Byte     '対象端末№
    RIRK_NAME(0 To 9)                   As Byte     '履歴種別名称
    HIN_NAME(0 To 39)                   As Byte     '品名
    SUMI_HIN_Zaiko_Qty(0 To 7)          As Byte     '品目別在庫数（商品化済み）
    MI_HIN_Zaiko_Qty(0 To 7)            As Byte     '品目別在庫数（未商品）
    SUMI_FROM_TANA_Zaiko_Qty(0 To 7)    As Byte     'FROM棚別品目別在庫数
    SUMI_TO_TANA_Zaiko_Qty(0 To 7)      As Byte     'TO棚別品目別在庫数
    MI_FROM_TANA_Zaiko_Qty(0 To 7)      As Byte     'FROM棚別品目別在庫数
    MI_TO_TANA_Zaiko_Qty(0 To 7)        As Byte     'TO棚別品目別在庫数
    TOKU_MARK(0 To 0)                   As Byte     '特売りマーク
    MEMO(0 To 59)                       As Byte     'メモ
    TANTO_CODE(0 To 4)                  As Byte     '担当者コード
    TANTO_NAME(0 To 19)                 As Byte     '担当者名称
    MUKE_CODE(0 To 7)                   As Byte     '得意先コード
    MUKE_NAME(0 To 39)                  As Byte     '得意先名称
    SS_CODE(0 To 7)                     As Byte     '直送先コード
    SS_NAME(0 To 39)                    As Byte     '直送先名称
    MUKE_DNAME(0 To 9)                  As Byte     '得意先略称
    MUKE_CHG_CD(0 To 1)                 As Byte     '向け先読替えコード
    SUM_KBN(0 To 0)                     As Byte     '集計区分
    ID_NO(0 To 7)                       As Byte     'ID-NO
    
    Ins_DateTime(0 To 13)               As Byte     '挿入日時2004.12.09
    
    '資材処理の為追加2005.01.05
    SHIIRE_CODE(0 To 4)                 As Byte     '仕入先ｺｰﾄﾞ
    SHIIRE_TANKA(0 To 10)               As Byte     '仕入単価(9(8)V99)
    KEIJYO_YM(0 To 5)                   As Byte     '計上年月(YYYYMM)
    '資材処理の為追加2005.01.05
    
    
    
    FILLER(0 To 167)                     As Byte
    
End Type

'データ・バッファ
Public O_IDOREC   As O_IDOREC_Tag

'キー定義
Type KEY0_O_IDO            'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    JITU_DT(0 To 7)             As Byte     '実績日付
    JITU_TM(0 To 5)             As Byte     '実績時刻
End Type

Type KEY1_O_IDO            'ＫＥＹ１
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
    JITU_DT(0 To 7)             As Byte     '実績日付
    JITU_TM(0 To 5)             As Byte     '実績時刻
End Type




'キー・データ
Public K0_O_IDO                   As KEY0_O_IDO
Public K1_O_IDO                   As KEY1_O_IDO

Type O_IDO_FSpeck
    fs      As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck
    ks8     As BtKeySpeck
End Type

Private O_IDO_Speck               As O_IDO_FSpeck
Private Function O_IDO_Create() As Integer
'********************************************************************
'*
'*              在庫移動歴　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    O_IDO_Create = True
                                            '在庫移動歴フルパス取込み
    sts = GetIni("FILE", O_IDO_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_IDO]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    O_IDO_Speck.fs.recoleng = Len(O_IDOREC)         ' レコード長
    O_IDO_Speck.fs.PageSize = O_IDO_PG_SIZ          ' ページサイズ
    O_IDO_Speck.fs.idexnumb = 2                   ' インデックス数
    O_IDO_Speck.fs.fileflag = 0                   ' ファイルフラグ
    O_IDO_Speck.fs.reserve = &H0                  ' 予約済み
'-----------------------------------------------
                                                ' キー０
    O_IDO_Speck.ks0.keypos = 15                   ' キーポジション
                                                ' キー長
    O_IDO_Speck.ks0.keyleng = 1
                                                ' キーフラグ
    O_IDO_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_IDO_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    O_IDO_Speck.ks0.reserve = &H0                 ' 予約済み
    
    O_IDO_Speck.ks1.keypos = 1                    ' キーポジション
    O_IDO_Speck.ks1.keyleng = 8                   ' キー長
                                                ' キーフラグ
    O_IDO_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_IDO_Speck.ks1.keytype = Chr(BtKtString)     ' キータイプ
    O_IDO_Speck.ks1.reserve = &H0                 ' 予約済み
    
    O_IDO_Speck.ks2.keypos = 9                    ' キーポジション
    O_IDO_Speck.ks2.keyleng = 6                   ' キー長
    O_IDO_Speck.ks2.keyflag = BtKfExt + BtKfDup   ' キーフラグ
    O_IDO_Speck.ks2.keytype = Chr(BtKtString)     ' キータイプ
    O_IDO_Speck.ks2.reserve = &H0                 ' 予約済み
'-----------------------------------------------
                                                ' キー１
    O_IDO_Speck.ks3.keypos = 15                   ' キーポジション
    O_IDO_Speck.ks3.keyleng = 1                   ' キー長
                                                ' キーフラグ
    O_IDO_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_IDO_Speck.ks3.keytype = Chr(BtKtString)     ' キータイプ
    O_IDO_Speck.ks3.reserve = &H0                 ' 予約済み

    O_IDO_Speck.ks4.keypos = 16                   ' キーポジション
    O_IDO_Speck.ks4.keyleng = 1                   ' キー長
                                                ' キーフラグ
    O_IDO_Speck.ks4.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_IDO_Speck.ks4.keytype = Chr(BtKtString)     ' キータイプ
    O_IDO_Speck.ks4.reserve = &H0                 ' 予約済み

    O_IDO_Speck.ks5.keypos = 17                   ' キーポジション
    O_IDO_Speck.ks5.keyleng = 20                  ' キー長
                                                ' キーフラグ
    O_IDO_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_IDO_Speck.ks5.keytype = Chr(BtKtString)     ' キータイプ
    O_IDO_Speck.ks5.reserve = &H0                 ' 予約済み

    O_IDO_Speck.ks6.keypos = 1                    ' キーポジション
    O_IDO_Speck.ks6.keyleng = 8                   ' キー長
                                                ' キーフラグ
    O_IDO_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_IDO_Speck.ks6.keytype = Chr(BtKtString)     ' キータイプ
    O_IDO_Speck.ks6.reserve = &H0                 ' 予約済み

    O_IDO_Speck.ks7.keypos = 9                    ' キーポジション
    O_IDO_Speck.ks7.keyleng = 6                   ' キー長
                                                ' キーフラグ
    O_IDO_Speck.ks7.keyflag = BtKfExt + BtKfDup
    O_IDO_Speck.ks7.keytype = Chr(BtKtString)     ' キータイプ
    O_IDO_Speck.ks7.reserve = &H0                 ' 予約済み
'-----------------------------------------------

    sts = BTRV(BtOpCreate, O_IDO_POS, O_IDO_Speck, Len(O_IDO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "在庫移動歴")
        Exit Function
    End If

    O_IDO_Create = False

End Function

Public Function O_IDO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              在庫移動歴　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    O_IDO_Open = True
                                            '在庫移動歴フルパス取込み
    sts = GetIni("FILE", O_IDO_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_IDO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, O_IDO_POS, O_IDOREC, Len(O_IDOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_IDO_Create()        '在庫移動歴作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_IDO_POS, O_IDOREC, Len(O_IDOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "在庫移動歴")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "在庫移動歴")
                Exit Function
        End Select
    Loop
    O_IDO_Open = False
End Function


