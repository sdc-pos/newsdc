Attribute VB_Name = "P_SHKENTO_OSAKA"
Option Explicit

'********************************************************************
'*
'*              発注検討　大阪PC向け  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ
Public Const P_SHKENTO_OSAKA_ID$ = "P_SHKENTO_OSAKA"

'ページサイズ
Private Const P_SHKENTO_OSAKA_PG_SIZ% = 512

'ポジション・ブロック
Public P_SHKENTO_OSAKA_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type P_SHKENTO_OSAKA_REC_Tag
    
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    
    SO_SUU(0 To 10)         As Byte         '総必要量(9(8)V99)
    TANKA(0 To 10)          As Byte         '仕入単価(9(8)V99)
    
    ST_SOKO(0 To 1)         As Byte         '標準棚番　倉庫
    ST_RETU(0 To 1)         As Byte         '標準棚番　列
    ST_REN(0 To 1)          As Byte         '標準棚番　連
    ST_DAN(0 To 1)          As Byte         '標準棚番　段
    
    ZAIKO_QTY(0 To 7)       As Byte         '在庫数
    
    SHIJI_Z_QTY(0 To 10)    As Byte         '注文残(9(8)V99)
    
    HIKIATE_Z_QTY(0 To 10)  As Byte         '引当残(9(8)V99)
    
    FUSOKU_QTY(0 To 10)     As Byte         '不足(9(8)V99)
    
    ORDER_QTY(0 To 10)      As Byte         '注文数(9(8)V99)
    
    LOT(0 To 7)             As Byte         '発注ﾛｯﾄ
    
    ORDER_CODE(0 To 4)      As Byte         '注文先ｺｰﾄﾞ
    
    LT(0 To 2)              As Byte         'ﾘｰﾄﾞﾀｲﾑ
    
    
    Y_NOUKI_DT(0 To 7)      As Byte         '予定納期
    
    REC_NO(0 To 3)          As Byte         'ﾚｺｰﾄﾞ№
    
    
    FILLER(0 To 59)         As Byte         'Filler

End Type
'データ・バッファ
Public P_SHKENTO_OSAKA_REC  As P_SHKENTO_OSAKA_REC_Tag

'キー定義

Public Type KEY0_P_SHKENTO_OSAKA            'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
End Type
    
Public Type KEY1_P_SHKENTO_OSAKA            'ＫＥＹ１
    REC_NO(0 To 3)          As Byte         'ﾚｺｰﾄﾞ№
End Type
    
Public Type KEY2_P_SHKENTO_OSAKA            'ＫＥＹ２
    FUSOKU_QTY(0 To 10)     As Byte         '不足(9(8)V99)
End Type
    
    
    
'キー・データ
Public K0_P_SHKENTO_OSAKA   As KEY0_P_SHKENTO_OSAKA
Public K1_P_SHKENTO_OSAKA   As KEY1_P_SHKENTO_OSAKA
Public K2_P_SHKENTO_OSAKA   As KEY2_P_SHKENTO_OSAKA


Type P_SHKENTO_OSAKA_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

End Type

Private P_SHKENTO_OSAKA_Speck   As P_SHKENTO_OSAKA_FSpeck
Private Function P_SHKENTO_OSAKA_Create(Optional F_NAME As String = " ") As Integer
'********************************************************************
'*
'*              発注検討　大阪PC向け  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim Ret             As Integer
    
    P_SHKENTO_OSAKA_Create = True
                                            '発注検討　大阪PC向けフルパス取込み
    sts = GetIni("FILE", P_SHKENTO_OSAKA_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHKENTO_OSAKA]読み込みエラー")
        Exit Function
    End If

    If Trim(F_NAME) = "" Then
        FullPath = RTrim(c)
    Else
        Ret = InStr(1, Trim(c), ".") - 1
        FullPath = Left(Trim(c), Ret) & Trim(F_NAME) & Right(Trim(c), Len(Trim(c)) - Ret)
    End If

    P_SHKENTO_OSAKA_Speck.fs.recoleng = Len(P_SHKENTO_OSAKA_REC)    ' レコード長
    P_SHKENTO_OSAKA_Speck.fs.PageSize = P_SHKENTO_OSAKA_PG_SIZ      ' ページサイズ
    P_SHKENTO_OSAKA_Speck.fs.idexnumb = 3                           ' インデックス数
    P_SHKENTO_OSAKA_Speck.fs.fileflag = 0                           ' ファイルフラグ
    P_SHKENTO_OSAKA_Speck.fs.reserve = &H0                          ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_SHKENTO_OSAKA_Speck.ks0.keypos = 1                            ' キーポジション
    P_SHKENTO_OSAKA_Speck.ks0.keyleng = 1                           ' キー長
    P_SHKENTO_OSAKA_Speck.ks0.keyflag = BtKfExt + BtKfSeg           ' キーフラグ
    P_SHKENTO_OSAKA_Speck.ks0.keytype = Chr(BtKtString)             ' キータイプ
    P_SHKENTO_OSAKA_Speck.ks0.reserve = &H0                         ' 予約済み
    
    P_SHKENTO_OSAKA_Speck.ks1.keypos = 2                            ' キーポジション
    P_SHKENTO_OSAKA_Speck.ks1.keyleng = 1                           ' キー長
    P_SHKENTO_OSAKA_Speck.ks1.keyflag = BtKfExt + BtKfSeg           ' キーフラグ
    P_SHKENTO_OSAKA_Speck.ks1.keytype = Chr(BtKtString)             ' キータイプ
    P_SHKENTO_OSAKA_Speck.ks1.reserve = &H0                         ' 予約済み
    
    P_SHKENTO_OSAKA_Speck.ks2.keypos = 3                            ' キーポジション
    P_SHKENTO_OSAKA_Speck.ks2.keyleng = 20                          ' キー長
    P_SHKENTO_OSAKA_Speck.ks2.keyflag = BtKfExt                     ' キーフラグ
    P_SHKENTO_OSAKA_Speck.ks2.keytype = Chr(BtKtString)             ' キータイプ
    P_SHKENTO_OSAKA_Speck.ks2.reserve = &H0                         ' 予約済み
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    P_SHKENTO_OSAKA_Speck.ks3.keypos = 129                          ' キーポジション
    P_SHKENTO_OSAKA_Speck.ks3.keyleng = 4                           ' キー長
    P_SHKENTO_OSAKA_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfDup ' キーフラグ
    P_SHKENTO_OSAKA_Speck.ks3.keytype = Chr(BtKtString)             ' キータイプ
    P_SHKENTO_OSAKA_Speck.ks3.reserve = &H0                         ' 予約済み
    '--------------------------------------------------- キー１ △
    
    '--------------------------------------------------- キー２ ▽
    P_SHKENTO_OSAKA_Speck.ks4.keypos = 83                          ' キーポジション
    P_SHKENTO_OSAKA_Speck.ks4.keyleng = 11                           ' キー長
    P_SHKENTO_OSAKA_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfDup ' キーフラグ
    P_SHKENTO_OSAKA_Speck.ks4.keytype = Chr(BtKtString)             ' キータイプ
    P_SHKENTO_OSAKA_Speck.ks4.reserve = &H0                         ' 予約済み
    '--------------------------------------------------- キー２ △
    
    
    
    sts = BTRV(BtOpCreate, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_Speck, Len(P_SHKENTO_OSAKA_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "発注検討　大阪PC向けﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_SHKENTO_OSAKA_Create = False

End Function

Public Function P_SHKENTO_OSAKA_Open(mode As Integer, Optional F_NAME As String = " ") As Integer
'********************************************************************
'*
'*              発注検討　大阪PC向け  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim Ret         As Integer

    P_SHKENTO_OSAKA_Open = True
                                                        '発注検討　大阪PC向けフルパス取込み
    sts = GetIni("FILE", P_SHKENTO_OSAKA_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHKENTO_OSAKA]読み込みエラー")
        Exit Function
    End If
    
    If Trim(F_NAME) = "" Then
        FullPath = RTrim(c)
    Else
        Ret = InStr(1, Trim(c), ".") - 1
        FullPath = Left(Trim(c), Ret) & Trim(F_NAME) & Right(Trim(c), Len(Trim(c)) - Ret)
    End If
    
    On Error Resume Next
    Kill (FullPath)
    On Error GoTo 0

    Do
        sts = BTRV(BtOpOpen, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), ByVal FullPath, Len(FullPath), mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHKENTO_OSAKA_Create(F_NAME)          '発注検討　大阪PC向け作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), ByVal FullPath, Len(FullPath), mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "発注検討　大阪PC向け")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "発注検討　大阪PC向け")
                Exit Function
        End Select
    Loop
    
    P_SHKENTO_OSAKA_Open = False

End Function

