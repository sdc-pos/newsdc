Attribute VB_Name = "ITEM_CATEGORY"
Option Explicit
'********************************************************************
'*
'*              品名カテゴリマスタ  ファイル定義
'*
'*          CREATE 2011.12.07
'********************************************************************
'ファイルＩＤ
Public Const ITEM_CATEGORY_ID$ = "ITEM_CATEGORY"

'ページサイズ
Public Const ITEM_CATEGORY_PG_SIZ% = 4096

'ポジション・ブロック
Public ITEM_CATEGORY_POS            As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************



'レコード定義
Type ITEM_CATEGORYREC_Tag
    JGYOBU(0 To 0)              As Byte     '事業部区分
    CATEGORY_CODE(0 To 7)       As Byte     '品名ｶﾞﾃｺﾞﾘｺｰﾄﾞ
    CATEGORY_NAME(0 To 79)      As Byte     '品名ｶﾞﾃｺﾞﾘ名称
    SEI_LOT(0 To 9)             As Byte     '生産ロット                 少数点可
    KOUSU_LOT(0 To 9)           As Byte     '前後工数(秒/ﾛｯﾄ)　         少数点可
    KOUSU_QTY(0 To 9)           As Byte     '前後工数(秒/個)　          少数点可
    TOKU_TANKA_QTY(0 To 9)      As Byte     '特別単価(作業工数　秒/個)　少数点可
    TOKU_TANKA_KOURYO(0 To 12)  As Byte     '特別単価(工料＠)　         9(10).99
    TOKU_TANKA_HAKO(0 To 12)    As Byte     '特別単価(箱代＠)　         9(10).99
    MEMO(0 To 79)               As Byte     '備考/メモ
    FILLER(0 To 228)            As Byte
    INS_TANTO(0 To 9)           As Byte     '追加　担当者
    Ins_DateTime(0 To 13)       As Byte     '追加　日時
    UPD_TANTO(0 To 9)           As Byte     '更新　担当者
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時

End Type
'データ・バッファ
Public ITEM_CATEGORYREC As ITEM_CATEGORYREC_Tag

'キー定義

Type KEY0_ITEM_CATEGORY                     'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    CATEGORY_CODE(0 To 7)       As Byte     '品名ｶﾞﾃｺﾞﾘｺｰﾄﾞ
End Type




'キー・データ
Public K0_ITEM_CATEGORY         As KEY0_ITEM_CATEGORY

Type ITEM_CATEGORY_FSpeck
    fs      As BtFileSpeck                  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
End Type

Private ITEM_CATEGORY_Speck  As ITEM_CATEGORY_FSpeck

Private Function ITEM_CATEGORY_Create() As Integer
'********************************************************************
'*
'*              品名カテゴリマスタ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************

Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ITEM_CATEGORY_Create = True
                                            '品名カテゴリマスタ フルパス取込み
    sts = GetIni("FILE", ITEM_CATEGORY_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_CATEGORY]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    ITEM_CATEGORY_Speck.fs.recoleng = Len(ITEM_CATEGORYREC)     ' レコード長
    ITEM_CATEGORY_Speck.fs.PageSize = ITEM_CATEGORY_PG_SIZ      ' ページサイズ
    ITEM_CATEGORY_Speck.fs.idexnumb = 1                         ' インデックス数
    ITEM_CATEGORY_Speck.fs.fileflag = 0                         ' ファイルフラグ
    ITEM_CATEGORY_Speck.fs.reserve = &H0                        ' 予約済み
'-----------------------------------------------
                                                ' キー０
    ITEM_CATEGORY_Speck.ks0.keypos = 1                          ' キーポジション
    ITEM_CATEGORY_Speck.ks0.keyleng = 1                         ' キー長
    ITEM_CATEGORY_Speck.ks0.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    ITEM_CATEGORY_Speck.ks0.keytype = Chr(BtKtString)           ' キータイプ
    ITEM_CATEGORY_Speck.ks0.reserve = &H0                       ' 予約済み

    ITEM_CATEGORY_Speck.ks1.keypos = 2                          ' キーポジション
    ITEM_CATEGORY_Speck.ks1.keyleng = 8                         ' キー長
    ITEM_CATEGORY_Speck.ks1.keyflag = BtKfExt                   ' キーフラグ
    ITEM_CATEGORY_Speck.ks1.keytype = Chr(BtKtString)           ' キータイプ
    ITEM_CATEGORY_Speck.ks1.reserve = &H0                       ' 予約済み
'-----------------------------------------------
    sts = BTRV(BtOpCreate, ITEM_CATEGORY_POS, ITEM_CATEGORY_Speck, Len(ITEM_CATEGORY_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "品名カテゴリマスタ")
        Exit Function
    End If

    ITEM_CATEGORY_Create = False

End Function

Public Function ITEM_CATEGORY_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              品名カテゴリマスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_CATEGORY_Open = True
                                            '品名カテゴリマスタ フルパス取込み
    sts = GetIni("FILE", ITEM_CATEGORY_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_CATEGORY]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_CATEGORY_Create()    '品名カテゴリマスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "品名カテゴリマスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "品名カテゴリマスタ")
                Exit Function
        End Select
    Loop

    ITEM_CATEGORY_Open = False

End Function

