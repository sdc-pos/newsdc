Attribute VB_Name = "MTS"
Option Explicit
'********************************************************************
'*
'*              向け先管理マスタ  ファイル定義
'*
'*          CREATE 2004.02.19
'********************************************************************
'ファイルＩＤ
Public Const MTS_ID$ = "MTS"

'ページサイズ
Public Const MTS_PG_SIZ% = 512

'ポジション・ブロック
Public MTS_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type MTSREC_Tag
    NAIGAI(0 To 0)          As Byte         '国内外
    DATA_KBN(0 To 0)        As Byte         'データ区分（未使用）
    MUKE_CODE(0 To 7)       As Byte         '得意先コード
    SS_CODE(0 To 7)         As Byte         '倉庫／ＳＳコード
    MUKE_NAME(0 To 39)      As Byte         '得意先名称
    SS_NAME(0 To 39)        As Byte         'ＳＳ名称
    MUKE_DNAME(0 To 9)      As Byte         '表示略称
    DISPLAY_RANKING(0 To 2) As Byte         '表示順序
    
    SYUKA_KBN(0 To 1)       As Byte         '出荷区分コード 2008.03.12
    FILLER(0 To 14)         As Byte         'FILLER
End Type

'データ・バッファ
Public MTSREC               As MTSREC_Tag

'キー定義

Type KEY0_MTS                 'ＫＥＹ０
    MUKE_CODE(0 To 7)       As Byte         '得意先コード
    SS_CODE(0 To 7)         As Byte         '倉庫／ＳＳコード
End Type

Type KEY1_MTS                 'ＫＥＹ１
    DISPLAY_RANKING(0 To 2) As Byte         '表示順序
    MUKE_CODE(0 To 7)       As Byte         '得意先コード
    SS_CODE(0 To 7)         As Byte         '倉庫／ＳＳコード
End Type

Type KEY2_MTS                 'ＫＥＹ２
    MUKE_CODE(0 To 7)       As Byte         '得意先コード
End Type

Type KEY3_MTS                 'ＫＥＹ３
    SS_CODE(0 To 7)         As Byte         '倉庫／ＳＳコード
End Type

'キー・データ
Public K0_MTS               As KEY0_MTS
Public K1_MTS               As KEY1_MTS
Public K2_MTS               As KEY2_MTS
Public K3_MTS               As KEY3_MTS

Type MTS_FSpeck
    fs  As BtFileSpeck                      'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck                       'ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck
    ks2 As BtKeySpeck
    ks3 As BtKeySpeck
    ks4 As BtKeySpeck
    ks5 As BtKeySpeck
    ks6 As BtKeySpeck
End Type

Private MTS_Speck As MTS_FSpeck
Private Function MTS_Create() As Integer
'********************************************************************
'*
'*              向け先管理マスタ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2004.02.19
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    MTS_Create = True
                                            '向け先管理マスタフルパス取込み
    sts = GetIni("FILE", MTS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [MTS]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    MTS_Speck.fs.recoleng = Len(MTSREC)         ' レコード長
    MTS_Speck.fs.PageSize = MTS_PG_SIZ          ' ページサイズ
    MTS_Speck.fs.idexnumb = 4                   ' インデックス数
    MTS_Speck.fs.fileflag = 0                   ' ファイルフラグ
    MTS_Speck.fs.reserve = &H0                  ' 予約済み
'------------------------------------------------
                                                ' キー０
    MTS_Speck.ks0.keypos = 3                    ' キーポジション
    MTS_Speck.ks0.keyleng = 8                   ' キー長
    MTS_Speck.ks0.keyflag = BtKfExt + BtKfSeg   ' キーフラグ
    MTS_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    MTS_Speck.ks0.reserve = &H0                 ' 予約済み
    
    MTS_Speck.ks1.keypos = 11                   ' キーポジション
    MTS_Speck.ks1.keyleng = 8                   ' キー長
    MTS_Speck.ks1.keyflag = BtKfExt             ' キーフラグ
    MTS_Speck.ks1.keytype = Chr(BtKtString)     ' キータイプ
    MTS_Speck.ks1.reserve = &H0                 ' 予約済み
'------------------------------------------------
                                                ' キー１
    MTS_Speck.ks2.keypos = 109                  ' キーポジション
    MTS_Speck.ks2.keyleng = 3                   ' キー長
                                                ' キーフラグ
    MTS_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfChg
    MTS_Speck.ks2.keytype = Chr(BtKtString)     ' キータイプ
    MTS_Speck.ks2.reserve = &H0                 ' 予約済み
    
    MTS_Speck.ks3.keypos = 3                    ' キーポジション
    MTS_Speck.ks3.keyleng = 8                   ' キー長
                                                ' キーフラグ
    MTS_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg
    MTS_Speck.ks3.keytype = Chr(BtKtString)     ' キータイプ
    MTS_Speck.ks3.reserve = &H0                 ' 予約済み
    
    MTS_Speck.ks4.keypos = 11                   ' キーポジション
    MTS_Speck.ks4.keyleng = 8                   ' キー長
    MTS_Speck.ks4.keyflag = BtKfExt + BtKfChg   ' キーフラグ
    MTS_Speck.ks4.keytype = Chr(BtKtString)     ' キータイプ
    MTS_Speck.ks4.reserve = &H0                 ' 予約済み
'------------------------------------------------
                                                ' キー２
    MTS_Speck.ks5.keypos = 3                    ' キーポジション
    MTS_Speck.ks5.keyleng = 8                   ' キー長
    MTS_Speck.ks5.keyflag = BtKfExt + BtKfDup   ' キーフラグ
    MTS_Speck.ks5.keytype = Chr(BtKtString)     ' キータイプ
    MTS_Speck.ks5.reserve = &H0                 ' 予約済み
'------------------------------------------------
                                                ' キー３
    MTS_Speck.ks6.keypos = 11                   ' キーポジション
    MTS_Speck.ks6.keyleng = 8                   ' キー長
    MTS_Speck.ks6.keyflag = BtKfExt + BtKfDup   ' キーフラグ
    MTS_Speck.ks6.keytype = Chr(BtKtString)     ' キータイプ
    MTS_Speck.ks6.reserve = &H0                 ' 予約済み


    sts = BTRV(BtOpCreate, MTS_POS, MTS_Speck, Len(MTS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "向け先管理マスタ")
        Exit Function
    End If

    MTS_Create = False

End Function

Public Function MTS_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              向け先管理マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    MTS_Open = True
                                            '向け先管理マスタフルパス取込み
    sts = GetIni("FILE", MTS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [MTS]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, MTS_POS, MTSREC, Len(MTSREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = MTS_Create()        '向け先管理マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, MTS_POS, MTSREC, Len(MTSREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "向け先管理マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "向け先管理マスタ")
                Exit Function
        End Select
    Loop
    
    MTS_Open = False

End Function
