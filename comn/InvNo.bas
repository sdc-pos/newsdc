Attribute VB_Name = "InvNo"
Option Explicit
'********************************************************************
'*
'*              床暖送状№データ　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const INVNO_ID$ = "INVNO"

'ページサイズ
Public Const INVNO_PG_SIZ% = 512

'ポジション・ブロック
Public INVNO_POS  As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type INVNOREC_Tag
    INVNO(0 To 19)      As Byte             '送状№
    Model(0 To 19)      As Byte             '品番
    LotNo(0 To 19)      As Byte             '製造番号
    OQty(0 To 5)        As Byte             '出荷数
    ODt(0 To 7)         As Byte             '出荷日
    FILLER(0 To 133)    As Byte             '
    EntID(0 To 9)       As Byte             '登録ID
    EntDtm(0 To 13)     As Byte             '登録日時yyyymmddhhmmss
    UpdID(0 To 9)       As Byte             '更新ID
    UpdDtm(0 To 13)     As Byte             '更新日時 yyyymmddhhmmss
End Type

'データ・バッファ
Public INVNOREC         As INVNOREC_Tag

'キー定義
Type KEY0_INVNO     'ＫＥＹ０
    Model(0 To 19)      As Byte             '品番
    LotNo(0 To 19)      As Byte             '製造番号
End Type

Type KEY1_INVNO     'ＫＥＹ１
    INVNO(0 To 19)      As Byte             '送状№
End Type


'キー・データ
Public K0_INVNO         As KEY0_INVNO

Type INVNO_FSpeck
    fs      As BtFileSpeck                  'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks2     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private INVNO_Speck     As INVNO_FSpeck

Private Function INVNO_Create() As Integer
'********************************************************************
'*
'*              床暖送状№データ　Create
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    INVNO_Create = True
                                            '床暖管理データフルパス取込み
    sts = GetIni("FILE", INVNO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [INVNO]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    INVNO_Speck.fs.recoleng = Len(INVNOREC)             ' レコード長
    INVNO_Speck.fs.PageSize = INVNO_PG_SIZ              ' ページサイズ
    INVNO_Speck.fs.idexnumb = 2                         ' インデックス数
    INVNO_Speck.fs.fileflag = 0                         ' ファイルフラグ
    INVNO_Speck.fs.reserve = &H0                        ' 予約済み
                                                    
'---------------------------------------------------
                                                            ' キー０
    INVNO_Speck.ks0.keypos = 21                             ' キーポジション
    INVNO_Speck.ks0.keyleng = 20                            ' キー長
    INVNO_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' キーフラグ
    INVNO_Speck.ks0.keytype = Chr(BtKtString)               ' キータイプ
    INVNO_Speck.ks0.reserve = &H0                           ' 予約済み

    INVNO_Speck.ks1.keypos = 41                             ' キーポジション
    INVNO_Speck.ks1.keyleng = 20                            ' キー長
    INVNO_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfChg               ' キーフラグ
    INVNO_Speck.ks1.keytype = Chr(BtKtString)               ' キータイプ
    INVNO_Speck.ks1.reserve = &H0                           ' 予約済み

'---------------------------------------------------


'---------------------------------------------------
                                                            ' キー１
    INVNO_Speck.ks2.keypos = 1                              ' キーポジション
    INVNO_Speck.ks2.keyleng = 20                            ' キー長
    INVNO_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg   ' キーフラグ
    INVNO_Speck.ks2.keytype = Chr(BtKtString)               ' キータイプ
    INVNO_Speck.ks2.reserve = &H0                           ' 予約済み


    sts = BTRV(BtOpCreate, INVNO_POS, INVNO_Speck, Len(INVNO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "床暖送状№データ")
        Exit Function
    End If

    INVNO_Create = False

End Function

Public Function INVNO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              床暖送状№データ　Open
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    INVNO_Open = True
                                            '床暖送状№データ フルパス取込み
    sts = GetIni("FILE", INVNO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [INVNO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, INVNO_POS, INVNOREC, Len(INVNOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = INVNO_Create()        '床暖送状№データ 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, INVNO_POS, INVNOREC, Len(INVNOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "床暖送状№データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "床暖送状№データ")
                Exit Function
        End Select
    Loop

    INVNO_Open = False

End Function
