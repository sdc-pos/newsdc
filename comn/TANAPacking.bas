Attribute VB_Name = "TANAPACKING"
Option Explicit
'********************************************************************
'*                                                                  *
'*              棚別個装箱マスタ  ファイル定義                      *
'*                                                                  *
'*          CREATE 2004.02.16                                       *
'********************************************************************
'ファイルＩＤ
Public Const TPACKING_ID$ = "TANAPACKING"

'ページサイズ
Public Const TPACKING_PG_SIZ% = 1024

'ポジション・ブロック
Public TPACKING_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type TPACKINGREC_Tag
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    PACKING_NO(0 To 3)  As Byte     '個装箱№
    RANK(0 To 2)        As Byte     'ランク
    FILLER(0 To 10)     As Byte     'FILLER
End Type
'データ・バッファ
Public TPACKINGREC      As TPACKINGREC_Tag

'キー定義
Type KEY0_TPACKING                  'ＫＥＹ０
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    PACKING_NO(0 To 3)  As Byte     '個装箱№
    RANK(0 To 2)        As Byte     'ランク
End Type

Type KEY1_TPACKING                  'ＫＥＹ１
    PACKING_NO(0 To 3)  As Byte     '個装箱№
    RANK(0 To 2)        As Byte     'ランク
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
End Type
    
'キー・データ
Public K0_TPACKING      As KEY0_TPACKING
Public K1_TPACKING      As KEY1_TPACKING

Type TPACKING_FSpeck
    fs  As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks7 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks8 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks9 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Public TPACKING_Speck   As TPACKING_FSpeck
Private Function TPACKING_Create() As Integer
'********************************************************************
'*                                                                  *
'*              棚別個装箱マスタ  ＣＲＥＡＴＥ                      *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.02.13                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    TPACKING_Create = True
                                            '棚別箱マスタフルパス取込み
    sts = GetIni("FILE", TPACKING_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim$(c)

    TPACKING_Speck.fs.recoleng = Len(TPACKINGREC)       ' レコード長
    TPACKING_Speck.fs.PageSize = TPACKING_PG_SIZ        ' ページサイズ
    TPACKING_Speck.fs.idexnumb = 2                      ' インデックス数
    TPACKING_Speck.fs.fileflag = 0                      ' ファイルフラグ
    TPACKING_Speck.fs.reserve = &H0                     ' 予約済み
'--------------------------------------------------------
                                                        ' キー０
    TPACKING_Speck.ks0.keypos = 1                       ' キーポジション
    TPACKING_Speck.ks0.keyleng = 2                      ' キー長
                                                        ' キーフラグ
    TPACKING_Speck.ks0.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    TPACKING_Speck.ks0.reserve = &H0                    ' 予約済み
                                                        ' キー０
    TPACKING_Speck.ks1.keypos = 3                       ' キーポジション
    TPACKING_Speck.ks1.keyleng = 2                      ' キー長
                                                        ' キーフラグ
    TPACKING_Speck.ks1.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks1.keytype = Chr(BtKtString)        ' キータイプ
    TPACKING_Speck.ks1.reserve = &H0                    ' 予約済み
                                                        ' キー０
    TPACKING_Speck.ks2.keypos = 5                       ' キーポジション
    TPACKING_Speck.ks2.keyleng = 2                      ' キー長
                                                        ' キーフラグ
    TPACKING_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks2.keytype = Chr(BtKtString)        ' キータイプ
    TPACKING_Speck.ks2.reserve = &H0                    ' 予約済み
                                                        ' キー０
    TPACKING_Speck.ks3.keypos = 7                       ' キーポジション
    TPACKING_Speck.ks3.keyleng = 4                      ' キー長
                                                        ' キーフラグ
    TPACKING_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks3.keytype = Chr(BtKtString)        ' キータイプ
    TPACKING_Speck.ks3.reserve = &H0                    ' 予約済み
                                                        ' キー０
    TPACKING_Speck.ks4.keypos = 11                      ' キーポジション
    TPACKING_Speck.ks4.keyleng = 3                      ' キー長
                                                        ' キーフラグ
    TPACKING_Speck.ks4.keyflag = BtKfExt + BtKfChg
    TPACKING_Speck.ks4.keytype = Chr(BtKtString)        ' キータイプ
    TPACKING_Speck.ks4.reserve = &H0                    ' 予約済み
'--------------------------------------------------------
                                                        ' キー１
    TPACKING_Speck.ks5.keypos = 7                       ' キーポジション
    TPACKING_Speck.ks5.keyleng = 4                      ' キー長
                                                        ' キーフラグ
    TPACKING_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks5.keytype = Chr(BtKtString)        ' キータイプ
    TPACKING_Speck.ks5.reserve = &H0                    ' 予約済み
                                                        ' キー１
    TPACKING_Speck.ks6.keypos = 11                      ' キーポジション
    TPACKING_Speck.ks6.keyleng = 3                      ' キー長
                                                        ' キーフラグ
    TPACKING_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks6.keytype = Chr(BtKtString)        ' キータイプ
    TPACKING_Speck.ks6.reserve = &H0                    ' 予約済み
                                                        ' キー１
    TPACKING_Speck.ks7.keypos = 1                       ' キーポジション
    TPACKING_Speck.ks7.keyleng = 2                      ' キー長
                                                        ' キーフラグ
    TPACKING_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks7.keytype = Chr(BtKtString)        ' キータイプ
    TPACKING_Speck.ks7.reserve = &H0                    ' 予約済み
                                                        ' キー１
    TPACKING_Speck.ks8.keypos = 3                       ' キーポジション
    TPACKING_Speck.ks8.keyleng = 2                      ' キー長
                                                        ' キーフラグ
    TPACKING_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TPACKING_Speck.ks8.keytype = Chr(BtKtString)        ' キータイプ
    TPACKING_Speck.ks8.reserve = &H0                    ' 予約済み
                                                        ' キー１
    TPACKING_Speck.ks9.keypos = 5                       ' キーポジション
    TPACKING_Speck.ks9.keyleng = 2                      ' キー長
                                                        ' キーフラグ
    TPACKING_Speck.ks9.keyflag = BtKfExt + BtKfChg
    TPACKING_Speck.ks9.keytype = Chr(BtKtString)        ' キータイプ
    TPACKING_Speck.ks9.reserve = &H0                    ' 予約済み
'--------------------------------------------------------

    sts = BTRV(BtOpCreate, TPACKING_POS, TPACKING_Speck, Len(TPACKING_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "棚別個装箱マスタ")
        Exit Function
    End If

    TPACKING_Create = False

End Function

Function TPACKING_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              棚別個装箱マスタ  ＯＰＥＮ                          *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.02.13                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    TPACKING_Open = True
                                            '棚別個装箱マスタフルパス取込み
    sts = GetIni("FILE", TPACKING_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = TPACKING_Create()         '棚別個装箱マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "棚別個装箱マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "棚別個装箱マスタ")
                Exit Function
        End Select
    Loop
    TPACKING_Open = False
End Function

Function TPACKING_ReCreate() As Integer
'********************************************************************
'*
'*              棚別個装箱マスタ  ファイル再作成
'*
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2004.06.16
'********************************************************************
Dim sts         As Integer

    TPACKING_ReCreate = True

    sts = TPACKING_Create()         '棚別個装箱マスタ作成
    If sts <> False Then
        Exit Function
    End If

    TPACKING_ReCreate = False

End Function

