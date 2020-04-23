Attribute VB_Name = "SOKO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              倉庫マスタ  ファイル定義                            *
'*                                                                  *
'*          CREATE 2004.02.16                                       *
'********************************************************************
'ファイルＩＤ
Public Const SOKO_ID$ = "SOKO"

'ページサイズ
Public Const SOKO_PG_SIZ% = 512

'ポジション・ブロック
Public SOKO_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type SOKOREC_Tag
    JGYOBU(0 To 0)      As Byte         '事業部区分
    Soko_No(0 To 1)     As Byte         '倉庫№
    SOKO_NAME(0 To 15)  As Byte         '倉庫名称
    SOKO_BUN(0 To 0)    As Byte         '倉庫分類
    SOKO_KBN(0 To 0)    As Byte         '倉庫区分
    NAIGAI(0 To 0)      As Byte         '国内外
    KAHI_KBN(0 To 0)    As Byte         '使用可否
    KONS_KBN(0 To 0)    As Byte         '混載可否
    RETU_START(0 To 1)  As Byte         '棚番範囲　列　開始
    RETU_END(0 To 1)    As Byte         '棚番範囲　列　終了
    REN_START(0 To 1)   As Byte         '棚番範囲　連　開始
    REN_END(0 To 1)     As Byte         '棚番範囲　連　終了
    DAN_START(0 To 1)   As Byte         '棚番範囲　段　開始
    DAN_END(0 To 1)     As Byte         '棚番範囲　段　終了
    
    ORDER_POINT(0 To 2) As Byte         '発注点 2004.02
    GOODS_ON_F(0 To 0)  As Byte         '商品化倉庫フラグ 2004.02
    
    
    IO_TANKA_No(0 To 1) As Byte         '入出庫単価設定ｺｰﾄﾞ 2008.02.14
    
    FILLER(0 To 13)     As Byte         'FILLER
End Type
'データ・バッファ
Public SOKOREC As SOKOREC_Tag

'キー定義

Type KEY0_SOKO            'ＫＥＹ０
    Soko_No(0 To 1)     As Byte         '倉庫№
End Type
    
'キー・データ
Public K0_SOKO          As KEY0_SOKO

Type SOKO_FSpeck
    fs                  As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                 As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private SOKO_Speck       As SOKO_FSpeck
Private Function SOKO_Create() As Integer
'********************************************************************
'*                                                                  *
'*              倉庫マスタ  ＣＲＥＡＴＥ                            *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SOKO_Create = True
                                            '倉庫マスタフルパス取込み
    sts = GetIni("FILE", SOKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SOKO]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    SOKO_Speck.fs.recoleng = Len(SOKOREC)       ' レコード長
    SOKO_Speck.fs.PageSize = SOKO_PG_SIZ        ' ページサイズ
    SOKO_Speck.fs.idexnumb = 1                  ' インデックス数
    SOKO_Speck.fs.fileflag = 0                  ' ファイルフラグ
    SOKO_Speck.fs.reserve = &H0                 ' 予約済み
                                                ' キー０
    SOKO_Speck.ks0.keypos = 2                   ' キーポジション
    SOKO_Speck.ks0.keyleng = 2                  ' キー長
    SOKO_Speck.ks0.keyflag = BtKfExt            ' キーフラグ
    SOKO_Speck.ks0.keytype = Chr(BtKtString)    ' キータイプ
    SOKO_Speck.ks0.reserve = &H0                ' 予約済み

    sts = BTRV(BtOpCreate, SOKO_POS, SOKO_Speck, Len(SOKO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "倉庫マスタ")
        Exit Function
    End If
    SOKO_Create = False
End Function

Function SOKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              倉庫マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    SOKO_Open = True
                                            '倉庫マスタフルパス取込み
    sts = GetIni("FILE", SOKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, SOKO_POS, SOKOREC, Len(SOKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SOKO_Create()        '倉庫マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SOKO_POS, SOKOREC, Len(SOKOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "倉庫マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "倉庫マスタ")
                Exit Function
        End Select
    Loop
    SOKO_Open = False

End Function
