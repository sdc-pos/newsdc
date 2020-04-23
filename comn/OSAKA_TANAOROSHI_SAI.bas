Attribute VB_Name = "OSAKA_TANAOROSHI_SAI"
Option Explicit

'********************************************************************
'*
'*              大阪PC　棚卸差異F  ファイル定義
'*
'*          CREATE 2012.04.17
'********************************************************************
'ファイルＩＤ
Public Const OSAKA_TANAOROSHI_SAI_ID$ = "OSAKA_TANAOROSHI_SAI"

'ページサイズ
Private Const OSAKA_TANAOROSHI_SAI_PG_SIZ% = 1024

'ポジション・ブロック
Public OSAKA_TANAOROSHI_SAI_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type OSAKA_TANAOROSHI_SAI_REC_Tag
    
    HIN_GAI(0 To 19)            As Byte         '資材品番
                            
    ST_SOKO(0 To 1)             As Byte         '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)             As Byte         '             列
    ST_REN(0 To 1)              As Byte         '             連
    ST_DAN(0 To 1)              As Byte         '             段
    
    SHIZAI_ZAIKO_QTY(0 To 7)    As Byte         '資材在庫数
    BUZAI_ZAIKO_QTY(0 To 7)     As Byte         '部材センター在庫数

    SAI_SU(0 To 7)              As Byte         '差異数

    FILLER(0 To 75)             As Byte






End Type
'データ・バッファ
Public OSAKA_TANAOROSHI_SAI_REC As OSAKA_TANAOROSHI_SAI_REC_Tag

'キー定義
    
Public Type KEY0_OSAKA_TANAOROSHI_SAI           'ＫＥＹ０
    
    HIN_GAI(0 To 19)            As Byte         '資材品番
    
    
End Type
    
Public Type KEY1_OSAKA_TANAOROSHI_SAI           'ＫＥＹ１
    
    ST_SOKO(0 To 1)             As Byte         '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)             As Byte         '             列
    ST_REN(0 To 1)              As Byte         '             連
    ST_DAN(0 To 1)              As Byte         '             段
    
    HIN_GAI(0 To 19)            As Byte         '資材品番
    
    
End Type
    
    
    
    
    
'キー・データ
Public K0_OSAKA_TANAOROSHI_SAI  As KEY0_OSAKA_TANAOROSHI_SAI
Public K1_OSAKA_TANAOROSHI_SAI  As KEY1_OSAKA_TANAOROSHI_SAI


Type OSAKA_TANAOROSHI_SAI_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

End Type

Private OSAKA_TANAOROSHI_SAI_Speck  As OSAKA_TANAOROSHI_SAI_FSpeck

Private Function OSAKA_TANAOROSHI_SAI_Create() As Integer
'********************************************************************
'*
'*              大阪PC　棚卸差異F  ファイル定義
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim Ret             As Long     '2007.11.13




    OSAKA_TANAOROSHI_SAI_Create = True
                                            '大阪PC　棚卸差異F  フルパス取込み
    sts = GetIni("FILE", OSAKA_TANAOROSHI_SAI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [OSAKA_TANAOROSHI_SAI]　読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)


    OSAKA_TANAOROSHI_SAI_Speck.fs.recoleng = Len(OSAKA_TANAOROSHI_SAI_REC)      ' レコード長
    OSAKA_TANAOROSHI_SAI_Speck.fs.PageSize = OSAKA_TANAOROSHI_SAI_PG_SIZ        ' ページサイズ
    OSAKA_TANAOROSHI_SAI_Speck.fs.idexnumb = 2                                  ' インデックス数
    OSAKA_TANAOROSHI_SAI_Speck.fs.fileflag = 0                                  ' ファイルフラグ
    OSAKA_TANAOROSHI_SAI_Speck.fs.reserve = &H0                                 ' 予約済み
    
    '--------------------------------------------------- キー０ ▽
    
    OSAKA_TANAOROSHI_SAI_Speck.ks0.keypos = 1                                   ' キーポジション
    OSAKA_TANAOROSHI_SAI_Speck.ks0.keyleng = 20                                 ' キー長
    OSAKA_TANAOROSHI_SAI_Speck.ks0.keyflag = BtKfExt + BtKfChg                  ' キーフラグ
    OSAKA_TANAOROSHI_SAI_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    OSAKA_TANAOROSHI_SAI_Speck.ks0.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    OSAKA_TANAOROSHI_SAI_Speck.ks1.keypos = 21                                  ' キーポジション
    OSAKA_TANAOROSHI_SAI_Speck.ks1.keyleng = 2                                  ' キー長
    OSAKA_TANAOROSHI_SAI_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg        ' キーフラグ
    OSAKA_TANAOROSHI_SAI_Speck.ks1.keytype = Chr(BtKtString)                    ' キータイプ
    OSAKA_TANAOROSHI_SAI_Speck.ks1.reserve = &H0                                ' 予約済み
    
    OSAKA_TANAOROSHI_SAI_Speck.ks2.keypos = 23                                  ' キーポジション
    OSAKA_TANAOROSHI_SAI_Speck.ks2.keyleng = 2                                  ' キー長
    OSAKA_TANAOROSHI_SAI_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg        ' キーフラグ
    OSAKA_TANAOROSHI_SAI_Speck.ks2.keytype = Chr(BtKtString)                    ' キータイプ
    OSAKA_TANAOROSHI_SAI_Speck.ks2.reserve = &H0                                ' 予約済み
    
    OSAKA_TANAOROSHI_SAI_Speck.ks3.keypos = 25                                  ' キーポジション
    OSAKA_TANAOROSHI_SAI_Speck.ks3.keyleng = 2                                  ' キー長
    OSAKA_TANAOROSHI_SAI_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg        ' キーフラグ
    OSAKA_TANAOROSHI_SAI_Speck.ks3.keytype = Chr(BtKtString)                    ' キータイプ
    OSAKA_TANAOROSHI_SAI_Speck.ks3.reserve = &H0                                ' 予約済み
    
    OSAKA_TANAOROSHI_SAI_Speck.ks4.keypos = 27                                  ' キーポジション
    OSAKA_TANAOROSHI_SAI_Speck.ks4.keyleng = 2                                  ' キー長
    OSAKA_TANAOROSHI_SAI_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg        ' キーフラグ
    OSAKA_TANAOROSHI_SAI_Speck.ks4.keytype = Chr(BtKtString)                    ' キータイプ
    OSAKA_TANAOROSHI_SAI_Speck.ks4.reserve = &H0                                ' 予約済み
    
    OSAKA_TANAOROSHI_SAI_Speck.ks5.keypos = 1                                   ' キーポジション
    OSAKA_TANAOROSHI_SAI_Speck.ks5.keyleng = 20                                 ' キー長
    OSAKA_TANAOROSHI_SAI_Speck.ks5.keyflag = BtKfExt + BtKfChg                  ' キーフラグ
    OSAKA_TANAOROSHI_SAI_Speck.ks5.keytype = Chr(BtKtString)                    ' キータイプ
    OSAKA_TANAOROSHI_SAI_Speck.ks5.reserve = &H0                                ' 予約済み
    '--------------------------------------------------- キー１ △
    sts = BTRV(BtOpCreate, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_Speck, Len(OSAKA_TANAOROSHI_SAI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "大阪PC　棚卸差異F")
        Exit Function
    End If
    
    OSAKA_TANAOROSHI_SAI_Create = False

End Function

Public Function OSAKA_TANAOROSHI_SAI_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              大阪PC　棚卸差異F  ファイル定義
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim Ret             As Long     '2007.11.13


    OSAKA_TANAOROSHI_SAI_Open = True
                                            '大阪PC　棚卸差異F  フルパス取込み
    sts = GetIni("FILE", OSAKA_TANAOROSHI_SAI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [OSAKA_TANAOROSHI_SAI]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, Len(OSAKA_TANAOROSHI_SAI_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = OSAKA_TANAOROSHI_SAI_Create()   '大阪PC　棚卸差異F ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, Len(OSAKA_TANAOROSHI_SAI_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "大阪PC　棚卸差異F")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "大阪PC　棚卸差異F")
                Exit Function
        End Select
    Loop
    
    OSAKA_TANAOROSHI_SAI_Open = False

End Function

