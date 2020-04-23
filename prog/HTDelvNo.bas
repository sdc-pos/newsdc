Attribute VB_Name = "HTDelvNo"
Option Explicit
'********************************************************************
'*
'*              送り状№ﾃﾞｰﾀ　ファイル定義
'*              Create 2016.10.14
'********************************************************************
'ファイルＩＤ
Public Const HTDelvNo_ID$ = "HTDelvNo"

'ページサイズ
Public Const HTDelvNo_PG_SIZ% = 512

'ポジション・ブロック
Public HTDelvNo_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type HTDelvNoREC_Tag
    CampName(0 To 19)   As Byte         '運送会社名(ヤマト運輸)
    DelvNo(0 To 19)     As Byte         '送り状№
    ChoCode(0 To 8)     As Byte         '直送先ｺｰﾄﾞ
    ChoName(0 To 19)    As Byte         '直送先名
    EntID(0 To 11)      As Byte         '登録ID
    EntTm(0 To 13)      As Byte         '登録日時
    UpdID(0 To 11)      As Byte         '更新ID
    UpdTm(0 To 13)      As Byte         '更新日時
End Type

'データ・バッファ
Public HTDelvNoREC      As HTDelvNoREC_Tag

'キー定義
Type KEY0_HTDelvNo          'ＫＥＹ０
    DelvNo(0 To 19)     As Byte         '送り状№
End Type

'キー・データ
Public K0_HTDelvNo    As KEY0_HTDelvNo

Type HTDelvNo_FSpeck
    fs                  As BtFileSpeck  'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                 As BtKeySpeck   'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private HTDelvNo_Speck  As HTDelvNo_FSpeck

Private Function HTDelvNo_Create() As Integer
'********************************************************************
'*                                                                  *
'*              送り状№ﾃﾞｰﾀ　ＣＲＥＡＴＥ                          *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    HTDelvNo_Create = True
                                            '送り状№ﾃﾞｰﾀ   フルパス取込み
    sts = GetIni("FILE", HTDelvNo_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTDelvNo]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    HTDelvNo_Speck.fs.recoleng = Len(HTDelvNoREC)       ' レコード長
    HTDelvNo_Speck.fs.PageSize = HTDelvNo_PG_SIZ        ' ページサイズ
    HTDelvNo_Speck.fs.idexnumb = 1                      ' インデックス数
    HTDelvNo_Speck.fs.fileflag = 0                      ' ファイルフラグ
    HTDelvNo_Speck.fs.reserve = &H0                     ' 予約済み
'------------------------------------------------
                                                ' キー０
    HTDelvNo_Speck.ks0.keypos = 21                      ' キーポジション
    HTDelvNo_Speck.ks0.keyleng = 20                     ' キー長
    HTDelvNo_Speck.ks0.keyflag = BtKfExt                ' キーフラグ
    HTDelvNo_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    HTDelvNo_Speck.ks0.reserve = &H0                    ' 予約済み
'------------------------------------------------

    sts = BTRV(BtOpCreate, HTDelvNo_POS, HTDelvNo_Speck, Len(HTDelvNo_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "送り状№ﾃﾞｰﾀ")
        Exit Function
    End If
    
    HTDelvNo_Create = False

End Function
Public Function HTDelvNo_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              送り状№ﾃﾞｰﾀ　ＯＰＥＮ                              *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    HTDelvNo_Open = True
                                        '送り状№ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", HTDelvNo_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTDelvNo]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, HTDelvNo_POS, HTDelvNoREC, Len(HTDelvNoREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = HTDelvNo_Create()        '送り状№ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, HTDelvNo_POS, HTDelvNoREC, Len(HTDelvNoREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "送り状№ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "送り状№ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop

    HTDelvNo_Open = False

End Function


