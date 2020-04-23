Attribute VB_Name = "KEPPINLOG"
Option Explicit
'********************************************************************
'*                                                                  *
'*              在庫集計データ　ファイル定義                        *
'*                                                                  *
'*          CREATE 2004.05.08                                       *
'********************************************************************
'ファイルＩＤ
Public Const KEPPINLOG_ID$ = "KEPPINLOG"

'ページサイズ
Public Const KEPPINLOG_PG_SIZ% = 512

'ポジション・ブロック
Public KEPPINLOG_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                              *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type KEPPINLOGREC_Tag
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    CREATE_DT(0 To 7)       As Byte     '作成日付
    FILLER(0 To 17)         As Byte     'FILLER
End Type

'データ・バッファ
Public KEPPINLOGREC         As KEPPINLOGREC_Tag

'キー定義
Private Type KEY0_KEPPINLOG         'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type


'キー・データ
Public K0_KEPPINLOG As KEY0_KEPPINLOG

Private Type KEPPINLOG_FSpeck
    fs As BtFileSpeck               ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private KEPPINLOG_Speck As KEPPINLOG_FSpeck

Private Function KEPPINLOG_Create() As Integer
'********************************************************************
'*                                                                  *
'*              欠品防止支援ログ　ＣＲＥＡＴＥ                      *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.05.08                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    KEPPINLOG_Create = True
                                            '欠品防止支援ログフルパス取込み
    sts = GetIni("FILE", KEPPINLOG_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[KEPPINLOG] 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    KEPPINLOG_Speck.fs.recoleng = Len(KEPPINLOGREC)         ' レコード長
    KEPPINLOG_Speck.fs.PageSize = KEPPINLOG_PG_SIZ          ' ページサイズ
    KEPPINLOG_Speck.fs.idexnumb = 1                         ' インデックス数
    KEPPINLOG_Speck.fs.fileflag = 0                         ' ファイルフラグ
    KEPPINLOG_Speck.fs.reserve = &H0                        ' 予約済み
'-----------------------------------------------' キー０
    KEPPINLOG_Speck.ks0.keypos = 1                          ' キーポジション
    KEPPINLOG_Speck.ks0.keyleng = 1                         ' キー長
    KEPPINLOG_Speck.ks0.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    KEPPINLOG_Speck.ks0.keytype = Chr(BtKtString)           ' キータイプ
    KEPPINLOG_Speck.ks0.reserve = &H0                       ' 予約済み

    KEPPINLOG_Speck.ks1.keypos = 2                          ' キーポジション
    KEPPINLOG_Speck.ks1.keyleng = 1                         ' キー長
    KEPPINLOG_Speck.ks1.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    KEPPINLOG_Speck.ks1.keytype = Chr(BtKtString)           ' キータイプ
    KEPPINLOG_Speck.ks1.reserve = &H0                       ' 予約済み

    KEPPINLOG_Speck.ks2.keypos = 3                          ' キーポジション
    KEPPINLOG_Speck.ks2.keyleng = 20                        ' キー長
    KEPPINLOG_Speck.ks2.keyflag = BtKfExt                   ' キーフラグ
    KEPPINLOG_Speck.ks2.keytype = Chr(BtKtString)           ' キータイプ
    KEPPINLOG_Speck.ks2.reserve = &H0                       ' 予約済み

    sts = BTRV(BtOpCreate, KEPPINLOG_POS, KEPPINLOG_Speck, Len(KEPPINLOG_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "欠品防止支援ログ")
        Exit Function
    End If
    
    KEPPINLOG_Create = False

End Function

Function KEPPINLOG_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              欠品防止支援ログ　ＯＰＥＮ                          *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.05.08                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    KEPPINLOG_Open = True
                                            '欠品防止支援ログフルパス取込み
    sts = GetIni("FILE", KEPPINLOG_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[KEPPINLOG] 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = KEPPINLOG_Create()    '欠品防止支援ログ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "欠品防止支援ログ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "欠品防止支援ログ")
                Exit Function
        End Select
    Loop

    KEPPINLOG_Open = False

End Function


