Attribute VB_Name = "HTIdDelv"
Option Explicit
'********************************************************************
'*
'*              Id送り状№ﾃﾞｰﾀ　ファイル定義
'*              Create 2016.10.14
'********************************************************************
'ファイルＩＤ
Public Const HTIdDelv_ID$ = "HTIdDelv"

'ページサイズ
Public Const HTIdDelv_PG_SIZ% = 512

'ポジション・ブロック
Public HTIdDelv_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type HTIdDelvREC_Tag
    IDNO(0 To 11)       As Byte         '伝票ID
    DelvNo(0 To 19)     As Byte         '送り状№
    EntID(0 To 11)      As Byte         '登録ID
    EntTm(0 To 13)      As Byte         '登録日時
    UpdID(0 To 11)      As Byte         '更新ID
    UpdTm(0 To 13)      As Byte         '更新日時
End Type

'データ・バッファ
Public HTIdDelvREC      As HTIdDelvREC_Tag

'キー定義
Type KEY0_HTIdDelv          'ＫＥＹ０
    IDNO(0 To 11)       As Byte         '伝票ID
    DelvNo(0 To 19)     As Byte         '送り状№
End Type

'キー・データ
Public K0_HTIdDelv    As KEY0_HTIdDelv

Type HTIdDelv_FSpeck
    fs                  As BtFileSpeck  'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                 As BtKeySpeck   'ｷｰ ｽﾍﾟｯｸ構造体
    ks1                 As BtKeySpeck   'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private HTIdDelv_Speck  As HTIdDelv_FSpeck

Private Function HTIdDelv_Create() As Integer
'********************************************************************
'*                                                                  *
'*              Id送り状№ﾃﾞｰﾀ　ＣＲＥＡＴＥ                        *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    HTIdDelv_Create = True
                                            'Id送り状№ﾃﾞｰﾀ   フルパス取込み
    sts = GetIni("FILE", HTIdDelv_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTIdDelv]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    HTIdDelv_Speck.fs.recoleng = Len(HTIdDelvREC)       ' レコード長
    HTIdDelv_Speck.fs.PageSize = HTIdDelv_PG_SIZ        ' ページサイズ
    HTIdDelv_Speck.fs.idexnumb = 1                      ' インデックス数
    HTIdDelv_Speck.fs.fileflag = 0                      ' ファイルフラグ
    HTIdDelv_Speck.fs.reserve = &H0                     ' 予約済み
'------------------------------------------------
                                                ' キー０
    HTIdDelv_Speck.ks0.keypos = 1                       ' キーポジション
    HTIdDelv_Speck.ks0.keyleng = 12                     ' キー長
                                                        ' キーフラグ
    HTIdDelv_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    HTIdDelv_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    HTIdDelv_Speck.ks0.reserve = &H0                    ' 予約済み

                                                ' キー０
    HTIdDelv_Speck.ks1.keypos = 13                      ' キーポジション
    HTIdDelv_Speck.ks1.keyleng = 20                     ' キー長
                                                        ' キーフラグ
    HTIdDelv_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfDup
    HTIdDelv_Speck.ks1.keytype = Chr(BtKtString)        ' キータイプ
    HTIdDelv_Speck.ks1.reserve = &H0                    ' 予約済み


'------------------------------------------------

    sts = BTRV(BtOpCreate, HTIdDelv_POS, HTIdDelv_Speck, Len(HTIdDelv_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "Id送り状№ﾃﾞｰﾀ")
        Exit Function
    End If
    
    HTIdDelv_Create = False

End Function
Public Function HTIdDelv_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              Id送り状№ﾃﾞｰﾀ　ＯＰＥＮ                            *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    HTIdDelv_Open = True
                                        'Id送り状№ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", HTIdDelv_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTIdDelv]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, HTIdDelv_POS, HTIdDelvREC, Len(HTIdDelvREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = HTIdDelv_Create()        'Id送り状№ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, HTIdDelv_POS, HTIdDelvREC, Len(HTIdDelvREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "Id送り状№ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "Id送り状№ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop

    HTIdDelv_Open = False

End Function


