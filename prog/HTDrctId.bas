Attribute VB_Name = "HTDrctId"
Option Explicit
'********************************************************************
'*
'*              直送先IDﾃﾞｰﾀ　ファイル定義
'*              Create 2016.10.14
'********************************************************************
'ファイルＩＤ
Public Const HTDrctId_ID$ = "HTDrctId"

'ページサイズ
Public Const HTDrctId_PG_SIZ% = 512

'ポジション・ブロック
Public HTDrctId_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type HTDrctIdREC_Tag
    IDNO(0 To 11)       As Byte         '伝票ID
    ChoCode(0 To 8)     As Byte         '直送先ｺｰﾄﾞ
    
    '>>>>>>>>>> 2016.10.27 追加
    ChoName(0 To 39)            As Byte '直送先名
    ChoZip(0 To 6)              As Byte '直送先郵便番号
    ChoTel(0 To 15)             As Byte '直送先電話番号
    ChoAddress(0 To 79)         As Byte '直送先住所
    ChoMemo(0 To 39)            As Byte '直送先メモ
    TMark(0 To 0)               As Byte 'Ｔマーク区分
    '>>>>>>>>>> 2016.10.27 追加
    
    EntID(0 To 11)      As Byte         '登録ID
    EntTm(0 To 13)      As Byte         '登録日時
    UpdID(0 To 11)      As Byte         '更新ID
    UpdTm(0 To 13)      As Byte         '更新日時
End Type

'データ・バッファ
Public HTDrctIdREC      As HTDrctIdREC_Tag

'キー定義
Type KEY0_HTDrctId          'ＫＥＹ０
    IDNO(0 To 19)       As Byte         '伝票ID
End Type

'キー・データ
Public K0_HTDrctId    As KEY0_HTDrctId

Type HTDrctId_FSpeck
    fs                  As BtFileSpeck  'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                 As BtKeySpeck   'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private HTDrctId_Speck  As HTDrctId_FSpeck

Private Function HTDrctId_Create() As Integer
'********************************************************************
'*                                                                  *
'*              直送先IDﾃﾞｰﾀ　ＣＲＥＡＴＥ                        　*
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    HTDrctId_Create = True
                                            '直送先IDﾃﾞｰﾀ   フルパス取込み
    sts = GetIni("FILE", HTDrctId_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTDrctId]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    HTDrctId_Speck.fs.recoleng = Len(HTDrctIdREC)       ' レコード長
    HTDrctId_Speck.fs.PageSize = HTDrctId_PG_SIZ        ' ページサイズ
    HTDrctId_Speck.fs.idexnumb = 1                      ' インデックス数
    HTDrctId_Speck.fs.fileflag = 0                      ' ファイルフラグ
    HTDrctId_Speck.fs.reserve = &H0                     ' 予約済み
'------------------------------------------------
                                                ' キー０
    HTDrctId_Speck.ks0.keypos = 1                       ' キーポジション
    HTDrctId_Speck.ks0.keyleng = 12                     ' キー長
    HTDrctId_Speck.ks0.keyflag = BtKfExt                ' キーフラグ
    HTDrctId_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    HTDrctId_Speck.ks0.reserve = &H0                    ' 予約済み
'------------------------------------------------

    sts = BTRV(BtOpCreate, HTDrctId_POS, HTDrctId_Speck, Len(HTDrctId_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "直送先IDﾃﾞｰﾀ")
        Exit Function
    End If
    
    HTDrctId_Create = False

End Function
Public Function HTDrctId_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              直送先IDﾃﾞｰﾀ　ＯＰＥＮ                            *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    HTDrctId_Open = True
                                        '直送先IDﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", HTDrctId_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTDrctId]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, HTDrctId_POS, HTDrctIdREC, Len(HTDrctIdREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = HTDrctId_Create()        '直送先IDﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, HTDrctId_POS, HTDrctIdREC, Len(HTDrctIdREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "直送先IDﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "直送先IDﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop

    HTDrctId_Open = False

End Function


