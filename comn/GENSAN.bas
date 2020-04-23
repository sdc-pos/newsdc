Attribute VB_Name = "GENSAN"
Option Explicit
'********************************************************************
'*
'*              原産国マスタ  ファイル定義
'*
'*          CREATE 2010.07.08
'********************************************************************
'ファイルＩＤ
Public Const GENSAN_ID$ = "GENSAN"

'ページサイズ
Public Const GENSAN_PG_SIZ% = 512

'ポジション・ブロック
Public GENSAN_POS       As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type GENSANREC_Tag
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
    GENSANKOKU(0 To 19)         As Byte     '原産国
    FILLER(0 To 175)            As Byte     'FILLER
    INS_TANTO(0 To 4)           As Byte     '追加　担当者
    Ins_DateTime(0 To 13)       As Byte     '追加　日時
    UPD_TANTO(0 To 4)           As Byte     '更新　担当者
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時

End Type
'データ・バッファ
Public GENSANREC                As GENSANREC_Tag

'キー定義

Type KEY0_GENSAN                'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
    GENSANKOKU(0 To 19)         As Byte     '原産国
End Type




'キー・データ
Public K0_GENSAN                As KEY0_GENSAN

Type GENSAN_FSpeck
    fs      As BtFileSpeck                 ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                 ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
End Type

Private GENSAN_Speck  As GENSAN_FSpeck

Private Function GENSAN_Create() As Integer
'********************************************************************
'*
'*              原産マスタ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    GENSAN_Create = True
                                            '原産マスタフルパス取込み
    sts = GetIni("FILE", GENSAN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [GENSAN]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    GENSAN_Speck.fs.recoleng = Len(GENSANREC)       ' レコード長
    GENSAN_Speck.fs.PageSize = GENSAN_PG_SIZ        ' ページサイズ
    GENSAN_Speck.fs.idexnumb = 1                    ' インデックス数
    GENSAN_Speck.fs.fileflag = 0                    ' ファイルフラグ
    GENSAN_Speck.fs.reserve = &H0                   ' 予約済み
'-----------------------------------------------
                                                ' キー０
    GENSAN_Speck.ks0.keypos = 1                     ' キーポジション
    GENSAN_Speck.ks0.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    GENSAN_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg
    GENSAN_Speck.ks0.keytype = Chr(BtKtString)      ' キータイプ
    GENSAN_Speck.ks0.reserve = &H0                  ' 予約済み

    GENSAN_Speck.ks1.keypos = 2                     ' キーポジション
    GENSAN_Speck.ks1.keyleng = 1                    ' キー長
                                                    ' キーフラグ
    GENSAN_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    GENSAN_Speck.ks1.keytype = Chr(BtKtString)      ' キータイプ
    GENSAN_Speck.ks1.reserve = &H0                  ' 予約済み

    GENSAN_Speck.ks2.keypos = 3                     ' キーポジション
    GENSAN_Speck.ks2.keyleng = 20                   ' キー長
                                                    ' キーフラグ
    GENSAN_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg
    GENSAN_Speck.ks2.keytype = Chr(BtKtString)      ' キータイプ
    GENSAN_Speck.ks2.reserve = &H0                  ' 予約済み

    GENSAN_Speck.ks3.keypos = 23                    ' キーポジション
    GENSAN_Speck.ks3.keyleng = 20                   ' キー長
    GENSAN_Speck.ks3.keyflag = BtKfExt + BtKfChg    ' キーフラグ
    GENSAN_Speck.ks3.keytype = Chr(BtKtString)      ' キータイプ
    GENSAN_Speck.ks3.reserve = &H0                  ' 予約済み
'-----------------------------------------------

    sts = BTRV(BtOpCreate, GENSAN_POS, GENSAN_Speck, Len(GENSAN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "原産マスタ")
        Exit Function
    End If

    GENSAN_Create = False

End Function

Public Function GENSAN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              原産マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    GENSAN_Open = True
                                            '原産マスタフルパス取込み
    sts = GetIni("FILE", GENSAN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [GENSAN]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, GENSAN_POS, GENSANREC, Len(GENSANREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = GENSAN_Create()        '原産マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, GENSAN_POS, GENSANREC, Len(GENSANREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "原産マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "原産マスタ")
                Exit Function
        End Select
    Loop

    GENSAN_Open = False

End Function

