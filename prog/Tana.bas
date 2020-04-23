Attribute VB_Name = "TANA"
Option Explicit
'********************************************************************
'*
'*              棚マスタ  ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const TANA_ID$ = "TANA"
'ページサイズ
Public Const TANA_PG_SIZ% = 1024

'ポジション・ブロック
Public TANA_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type TANAREC_Tag
    SOKO_NO(0 To 1)         As Byte     '倉庫№
    Retu(0 To 1)            As Byte     '棚番　列
    Ren(0 To 1)             As Byte     '棚番　連
    Dan(0 To 1)             As Byte     '棚番　段
    KAHI_KBN(0 To 0)        As Byte     '使用可否
    TANA_COND(0 To 0)       As Byte     '棚状態
    
    ZAIKO_SHOGO_FLG(0 To 0) As Byte     '在庫照合フラグ 2004.02
    
    Tana_Use(0 To 2)        As Byte     '棚の使用状況   2010.12.13
    
    
    FILLER(0 To 6)          As Byte     'FILLER         2010.12.13
End Type
'データ・バッファ
Public TANAREC As TANAREC_Tag


'キー定義
Type KEY0_TANA                 'ＫＥＹ０
    SOKO_NO(0 To 1)         As Byte     '倉庫№
    Retu(0 To 1)            As Byte     '棚番　列
    Ren(0 To 1)             As Byte     '棚番　連
    Dan(0 To 1)             As Byte     '棚番　段
End Type

Type KEY1_TANA                 'ＫＥＹ１
    KAHI_KBN(0 To 0)        As Byte     '使用可否
    SOKO_NO(0 To 1)         As Byte     '倉庫№
    Retu(0 To 1)            As Byte     '棚番　列
    Ren(0 To 1)             As Byte     '棚番　連
    Dan(0 To 1)             As Byte     '棚番　段
End Type

    
'キー・データ
Public K0_TANA              As KEY0_TANA
Public K1_TANA              As KEY1_TANA

Type TANA_FSpeck
    fs              As BtFileSpeck      ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0             As BtKeySpeck       ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1             As BtKeySpeck       ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2             As BtKeySpeck       ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3             As BtKeySpeck       ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4             As BtKeySpeck       ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5             As BtKeySpeck       ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6             As BtKeySpeck       ' ｷｰ ｽﾍﾟｯｸ構造体
    ks7             As BtKeySpeck       ' ｷｰ ｽﾍﾟｯｸ構造体
    ks8             As BtKeySpeck       ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Public TANA_Speck   As TANA_FSpeck
Private Function TANA_Create() As Integer
'********************************************************************
'*                                                                  *
'*              棚マスタ  ＣＲＥＡＴＥ                              *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    TANA_Create = False
                                            '棚マスタフルパス取込み
    sts = GetIni("FILE", TANA_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        TANA_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    TANA_Speck.fs.recoleng = Len(TANAREC)       ' レコード長
    TANA_Speck.fs.PageSize = TANA_PG_SIZ        ' ページサイズ
    TANA_Speck.fs.idexnumb = 2                  ' インデックス数
    TANA_Speck.fs.fileflag = 0                  ' ファイルフラグ
    TANA_Speck.fs.reserve = &H0                 ' 予約済み
'-----------------------------------------------
                                                ' キー０
    TANA_Speck.ks0.keypos = 1                   ' キーポジション
    TANA_Speck.ks0.keyleng = 2                  ' キー長
    TANA_Speck.ks0.keyflag = BtKfExt + BtKfSeg  ' キーフラグ
    TANA_Speck.ks0.keytype = Chr(BtKtString)    ' キータイプ
    TANA_Speck.ks0.reserve = &H0                ' 予約済み

    TANA_Speck.ks1.keypos = 3                   ' キーポジション
    TANA_Speck.ks1.keyleng = 2                  ' キー長
    TANA_Speck.ks1.keyflag = BtKfExt + BtKfSeg  ' キーフラグ
    TANA_Speck.ks1.keytype = Chr(BtKtString)    ' キータイプ
    TANA_Speck.ks1.reserve = &H0                ' 予約済み

    TANA_Speck.ks2.keypos = 5                   ' キーポジション
    TANA_Speck.ks2.keyleng = 2                  ' キー長
    TANA_Speck.ks2.keyflag = BtKfExt + BtKfSeg  ' キーフラグ
    TANA_Speck.ks2.keytype = Chr(BtKtString)    ' キータイプ
    TANA_Speck.ks2.reserve = &H0                ' 予約済み

    TANA_Speck.ks3.keypos = 7                   ' キーポジション
    TANA_Speck.ks3.keyleng = 2                  ' キー長
    TANA_Speck.ks3.keyflag = BtKfExt            ' キーフラグ
    TANA_Speck.ks3.keytype = Chr(BtKtString)    ' キータイプ
    TANA_Speck.ks3.reserve = &H0                ' 予約済み

'-----------------------------------------------
                                                ' キー１
    TANA_Speck.ks4.keypos = 9                   ' キーポジション
    TANA_Speck.ks4.keyleng = 1                  ' キー長
                                                ' キーフラグ
    TANA_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TANA_Speck.ks4.keytype = Chr(BtKtString)    ' キータイプ
    TANA_Speck.ks4.reserve = &H0                ' 予約済み
                                                ' キー１
    TANA_Speck.ks5.keypos = 1                   ' キーポジション
    TANA_Speck.ks5.keyleng = 2                  ' キー長
                                                ' キーフラグ
    TANA_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TANA_Speck.ks5.keytype = Chr(BtKtString)    ' キータイプ
    TANA_Speck.ks5.reserve = &H0                ' 予約済み
                                                ' キー１
    TANA_Speck.ks6.keypos = 3                   ' キーポジション
    TANA_Speck.ks6.keyleng = 2                  ' キー長
                                                ' キーフラグ
    TANA_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TANA_Speck.ks6.keytype = Chr(BtKtString)    ' キータイプ
    TANA_Speck.ks6.reserve = &H0                ' 予約済み
                                                ' キー１
    TANA_Speck.ks7.keypos = 5                   ' キーポジション
    TANA_Speck.ks7.keyleng = 2                  ' キー長
                                                ' キーフラグ
    TANA_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    TANA_Speck.ks7.keytype = Chr(BtKtString)    ' キータイプ
    TANA_Speck.ks7.reserve = &H0                ' 予約済み
                                                ' キー１
    TANA_Speck.ks8.keypos = 7                   ' キーポジション
    TANA_Speck.ks8.keyleng = 2                  ' キー長
                                                ' キーフラグ
    TANA_Speck.ks8.keyflag = BtKfExt + BtKfChg
    TANA_Speck.ks8.keytype = Chr(BtKtString)    ' キータイプ
    TANA_Speck.ks8.reserve = &H0                ' 予約済み

    sts = BTRV(BtOpCreate, TANA_POS, TANA_Speck, Len(TANA_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "棚マスタ")
        TANA_Create = True
    End If
End Function

Function TANA_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              棚マスタ  ＯＰＥＮ                                  *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    TANA_Open = False
                                            '棚マスタフルパス取込み
    sts = GetIni("FILE", TANA_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        TANA_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, TANA_POS, TANAREC, Len(TANAREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = TANA_Create()        '棚マスタ作成
                If sts <> False Then
                    TANA_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, TANA_POS, TANAREC, Len(TANAREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "棚マスタ")
                    TANA_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "棚マスタ")
                TANA_Open = True
                Exit Function
        End Select
    Loop
End Function



