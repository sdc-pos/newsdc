Attribute VB_Name = "SEK_OKURISAKI"
Option Explicit
'********************************************************************
'*
'*              積水送り先　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const SEK_OKURISAKI_ID$ = "SEK_OKURISAKI"

'ページサイズ
Public Const SEK_OKURISAKI_PG_SIZ% = 1024

'ポジション・ブロック
Public SEK_OKURISAKI_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type SEK_OKURISAKIREC_Tag
    OKURISAKI_CD(0 To 7)        As Byte     '得意先コード
    
    MUKE_NAME(0 To 39)          As Byte     '得意先名称


    JYUSHO(0 To 159)            As Byte     '住所       2009.11.19
    
    TEL_NO(0 To 19)             As Byte     '電話番号   2010.01.21
    YUBIN_NO(0 To 7)            As Byte     '郵便番号   2010.01.21



    FILLER(0 To 147)            As Byte     'FILLER





End Type

'データ・バッファ
Public SEK_OKURISAKIREC         As SEK_OKURISAKIREC_Tag

'キー定義
Type KEY0_SEK_OKURISAKI                     'ＫＥＹ０
    OKURISAKI_CD(0 To 7)        As Byte     '得意先コード
End Type


'キー・データ
Public K0_SEK_OKURISAKI         As KEY0_SEK_OKURISAKI

Type SEK_OKURISAKI_FSpeck
    fs      As BtFileSpeck                  'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private SEK_OKURISAKI_Speck     As SEK_OKURISAKI_FSpeck

Private Function SEK_OKURISAKI_Create() As Integer
'********************************************************************
'*
'*              積水送り先　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SEK_OKURISAKI_Create = True
                                            '積水送り先フルパス取込み
    sts = GetIni(App.EXEName, SEK_OKURISAKI_ID, App.EXEName, c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [SEK_OKURISAKI]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    SEK_OKURISAKI_Speck.fs.recoleng = Len(SEK_OKURISAKIREC)     ' レコード長
    SEK_OKURISAKI_Speck.fs.PageSize = SEK_OKURISAKI_PG_SIZ      ' ページサイズ
    SEK_OKURISAKI_Speck.fs.idexnumb = 1                         ' インデックス数
    SEK_OKURISAKI_Speck.fs.fileflag = 0                         ' ファイルフラグ
    SEK_OKURISAKI_Speck.fs.reserve = &H0                        ' 予約済み
                                                    
'---------------------------------------------------
                                                        ' キー０
    SEK_OKURISAKI_Speck.ks0.keypos = 1                          ' キーポジション
    SEK_OKURISAKI_Speck.ks0.keyleng = 8                         ' キー長
    SEK_OKURISAKI_Speck.ks0.keyflag = BtKfExt + BtKfChg         ' キーフラグ
    SEK_OKURISAKI_Speck.ks0.keytype = Chr(BtKtString)           ' キータイプ
    SEK_OKURISAKI_Speck.ks0.reserve = &H0                       ' 予約済み

    sts = BTRV(BtOpCreate, SEK_OKURISAKI_POS, SEK_OKURISAKI_Speck, Len(SEK_OKURISAKI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "積水送り先")
        Exit Function
    End If

    SEK_OKURISAKI_Create = False

End Function

Public Function SEK_OKURISAKI_Open(mode As Integer) As Integer
'********************************************************************
'*
'*              積水送り先　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    SEK_OKURISAKI_Open = True
                                            '積水送り先フルパス取込み
    sts = GetIni(App.EXEName, SEK_OKURISAKI_ID, App.EXEName, c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [積水送り先]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, SEK_OKURISAKI_POS, SEK_OKURISAKIREC, Len(SEK_OKURISAKIREC), ByVal FullPath, Len(FullPath), mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SEK_OKURISAKI_Create()    '積水送り先作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SEK_OKURISAKI_POS, SEK_OKURISAKIREC, Len(SEK_OKURISAKIREC), ByVal FullPath, Len(FullPath), mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "積水送り先")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "積水送り先")
                Exit Function
        End Select
    Loop

    SEK_OKURISAKI_Open = False

End Function
