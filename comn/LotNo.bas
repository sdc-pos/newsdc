Attribute VB_Name = "LotNo"
Option Explicit
'********************************************************************
'*
'*              床暖管理データ　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const LOTNO_ID$ = "LOTNO"

'ページサイズ
Public Const LOTNO_PG_SIZ% = 512

'ポジション・ブロック
Public LOTNO_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type LOTNOREC_Tag
    Model(0 To 19)      As Byte             '品番
    PLotNo(0 To 19)     As Byte             '製造番号
    IQty(0 To 5)        As Byte             '入荷数
    OQty(0 To 5)        As Byte             '出荷数
    SQty(0 To 5)        As Byte             '在庫数
    EDt(0 To 7)         As Byte             '輸出日(韓国)
    IDt(0 To 7)         As Byte             '入荷日
    ODt(0 To 7)         As Byte             '出荷日
    MemoNo(0 To 19)     As Byte             'No(メモ)
    EntFN(0 To 39)      As Byte             '登録ﾌｧｲﾙ名
    ITantoCode(0 To 4)  As Byte             '入荷担当者ID
    OTantoCode(0 To 4)  As Byte             '出荷担当者ID
    FILLER(0 To 69)     As Byte             '
    EntID(0 To 9)       As Byte             '登録ID
    EntDtm(0 To 13)     As Byte             '登録日時yyyymmddhhmmss
    UpdID(0 To 9)       As Byte             '更新ID
    UpdDtm(0 To 13)     As Byte             '更新日時 yyyymmddhhmmss
End Type

'データ・バッファ
Public LOTNOREC         As LOTNOREC_Tag

'キー定義
Type KEY0_LOTNO         'ＫＥＹ０
    Model(0 To 19)      As Byte             '品番
    PLotNo(0 To 19)     As Byte             '製造番号
End Type

'キー・データ
Public K0_LOTNO         As KEY0_LOTNO

Type LOTNO_FSpeck
    fs      As BtFileSpeck                  'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private LOTNO_Speck     As LOTNO_FSpeck

Private Function LOTNO_Create() As Integer
'********************************************************************
'*
'*              床暖管理データ　Create
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    LOTNO_Create = True
                                            '床暖管理データフルパス取込み
    sts = GetIni("FILE", LOTNO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [LOTNO]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    LOTNO_Speck.fs.recoleng = Len(LOTNOREC)             ' レコード長
    LOTNO_Speck.fs.PageSize = LOTNO_PG_SIZ              ' ページサイズ
    LOTNO_Speck.fs.idexnumb = 1                         ' インデックス数
    LOTNO_Speck.fs.fileflag = 0                         ' ファイルフラグ
    LOTNO_Speck.fs.reserve = &H0                        ' 予約済み
                                                    
'---------------------------------------------------
                                                            ' キー０
    LOTNO_Speck.ks0.keypos = 1                              ' キーポジション
    LOTNO_Speck.ks0.keyleng = 20                            ' キー長
    LOTNO_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' キーフラグ
    LOTNO_Speck.ks0.keytype = Chr(BtKtString)               ' キータイプ
    LOTNO_Speck.ks0.reserve = &H0                           ' 予約済み

    LOTNO_Speck.ks1.keypos = 21                             ' キーポジション
    LOTNO_Speck.ks1.keyleng = 20                            ' キー長
    LOTNO_Speck.ks1.keyflag = BtKfExt + BtKfChg             ' キーフラグ
    LOTNO_Speck.ks1.keytype = Chr(BtKtString)               ' キータイプ
    LOTNO_Speck.ks1.reserve = &H0                           ' 予約済み

'---------------------------------------------------

    sts = BTRV(BtOpCreate, LOTNO_POS, LOTNO_Speck, Len(LOTNO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "床暖管理データ")
        Exit Function
    End If

    LOTNO_Create = False

End Function

Public Function LOTNO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              床暖管理データ　Open
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    LOTNO_Open = True
                                            '床暖管理データ フルパス取込み
    sts = GetIni("FILE", LOTNO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [LOTNO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, LOTNO_POS, LOTNOREC, Len(LOTNOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = LOTNO_Create()        '床暖管理データ 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, LOTNO_POS, LOTNOREC, Len(LOTNOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "床暖管理データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "床暖管理データ")
                Exit Function
        End Select
    Loop

    LOTNO_Open = False

End Function
