Attribute VB_Name = "SE_USOU_HAKO"
Option Explicit
'********************************************************************
'*
'*              輸送箱使用実績  ファイル定義
'*
'*          CREATE 2008.02.25
'********************************************************************
'ファイルＩＤ
Public Const SE_USOU_HAKO_ID$ = "SE_USOU_HAKO"

'ページサイズ
Public Const SE_USOU_HAKO_PG_SIZ% = 1024

'ポジション・ブロック
Public SE_USOU_HAKO_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type SE_USOU_HAKOREC_Tag
    SHIMUKE_CODE(0 To 1)        As Byte         '仕向け先
    JITU_DATE(0 To 7)           As Byte         '実績日付
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '品番
    MTS_CODE(0 To 7)            As Byte         '出荷先
    CYU_KBN(0 To 0)             As Byte         '注文区分(未使用)
    CYOK_KBN(0 To 0)            As Byte         '直送区分(未使用)
    MAISU(0 To 5)               As Byte         '使用枚数
    UPD_TANTO(0 To 4)           As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)       As Byte         '更新　日時
    
    
    SE_USOU_F(0 To 1)           As Byte         '輸送箱　出力ﾌﾗｸﾞ
    
    
    FILLER(0 To 186)            As Byte         'FILLER
End Type

'データ・バッファ
Public SE_USOU_HAKOREC      As SE_USOU_HAKOREC_Tag

'キー定義

Type KEY0_SE_USOU_HAKO      'ＫＥＹ０
    JITU_DATE(0 To 7)           As Byte         '実績日付
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '品番
    MTS_CODE(0 To 7)            As Byte         '出荷先
End Type


Type KEY1_SE_USOU_HAKO      'ＫＥＹ１
    SE_USOU_F(0 To 1)           As Byte         '輸送箱　出力ﾌﾗｸﾞ
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '品番
End Type


Type KEY2_SE_USOU_HAKO      'ＫＥＹ２
    JITU_DATE(0 To 7)           As Byte         '実績日付
    MTS_CODE(0 To 7)            As Byte         '出荷先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '品番
End Type






'キー・データ
Public K0_SE_USOU_HAKO          As KEY0_SE_USOU_HAKO
Public K1_SE_USOU_HAKO          As KEY1_SE_USOU_HAKO
Public K2_SE_USOU_HAKO          As KEY2_SE_USOU_HAKO

Type SE_USOU_HAKO_FSpeck
    fs      As BtFileSpeck                  'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck
    ks8     As BtKeySpeck

    ks9     As BtKeySpeck
    ks10    As BtKeySpeck
    ks11    As BtKeySpeck
    ks12    As BtKeySpeck
    ks13    As BtKeySpeck

End Type

Private SE_USOU_HAKO_Speck As SE_USOU_HAKO_FSpeck
Private Function SE_USOU_HAKO_Create() As Integer
'********************************************************************
'*
'*              輸送箱使用実績  ＣＲＥＡＴＥ
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

    SE_USOU_HAKO_Create = True
                                            '輸送箱使用実績フルパス取込み
    sts = GetIni("FILE", SE_USOU_HAKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_USOU_HAKO]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    SE_USOU_HAKO_Speck.fs.recoleng = Len(SE_USOU_HAKOREC)           ' レコード長
    SE_USOU_HAKO_Speck.fs.PageSize = SE_USOU_HAKO_PG_SIZ            ' ページサイズ
    SE_USOU_HAKO_Speck.fs.idexnumb = 3                              ' インデックス数
    SE_USOU_HAKO_Speck.fs.fileflag = 0                              ' ファイルフラグ
    SE_USOU_HAKO_Speck.fs.reserve = &H0                             ' 予約済み
'------------------------------------------------
    SE_USOU_HAKO_Speck.ks0.keypos = 3                               ' キーポジション
    SE_USOU_HAKO_Speck.ks0.keyleng = 8                              ' キー長
    SE_USOU_HAKO_Speck.ks0.keyflag = BtKfExt + BtKfSeg              ' キーフラグ
    SE_USOU_HAKO_Speck.ks0.keytype = Chr(BtKtString)                ' キータイプ
    SE_USOU_HAKO_Speck.ks0.reserve = &H0                            ' 予約済み

    SE_USOU_HAKO_Speck.ks1.keypos = 11                              ' キーポジション
    SE_USOU_HAKO_Speck.ks1.keyleng = 1                              ' キー長
    SE_USOU_HAKO_Speck.ks1.keyflag = BtKfExt + BtKfSeg              ' キーフラグ
    SE_USOU_HAKO_Speck.ks1.keytype = Chr(BtKtString)                ' キータイプ
    SE_USOU_HAKO_Speck.ks1.reserve = &H0                            ' 予約済み

    SE_USOU_HAKO_Speck.ks2.keypos = 12                              ' キーポジション
    SE_USOU_HAKO_Speck.ks2.keyleng = 1                              ' キー長
    SE_USOU_HAKO_Speck.ks2.keyflag = BtKfExt + BtKfSeg              ' キーフラグ
    SE_USOU_HAKO_Speck.ks2.keytype = Chr(BtKtString)                ' キータイプ
    SE_USOU_HAKO_Speck.ks2.reserve = &H0                            ' 予約済み

    SE_USOU_HAKO_Speck.ks3.keypos = 13                              ' キーポジション
    SE_USOU_HAKO_Speck.ks3.keyleng = 20                             ' キー長
    SE_USOU_HAKO_Speck.ks3.keyflag = BtKfExt + BtKfSeg              ' キーフラグ
    SE_USOU_HAKO_Speck.ks3.keytype = Chr(BtKtString)                ' キータイプ
    SE_USOU_HAKO_Speck.ks3.reserve = &H0                            ' 予約済み

    SE_USOU_HAKO_Speck.ks4.keypos = 33                              ' キーポジション
    SE_USOU_HAKO_Speck.ks4.keyleng = 8                              ' キー長
    SE_USOU_HAKO_Speck.ks4.keyflag = BtKfExt                        ' キーフラグ
    SE_USOU_HAKO_Speck.ks4.keytype = Chr(BtKtString)                ' キータイプ
    SE_USOU_HAKO_Speck.ks4.reserve = &H0                            ' 予約済み
'------------------------------------------------


'------------------------------------------------
    SE_USOU_HAKO_Speck.ks5.keypos = 68                              ' キーポジション
    SE_USOU_HAKO_Speck.ks5.keyleng = 2                              ' キー長
    SE_USOU_HAKO_Speck.ks5.keyflag = BtKfExt + _
                                        BtKfSeg + _
                                        BtKfDup + _
                                        BtKfChg                     ' キーフラグ
    SE_USOU_HAKO_Speck.ks5.keytype = Chr(BtKtString)                ' キータイプ
    SE_USOU_HAKO_Speck.ks5.reserve = &H0                            ' 予約済み

    SE_USOU_HAKO_Speck.ks6.keypos = 11                              ' キーポジション
    SE_USOU_HAKO_Speck.ks6.keyleng = 1                              ' キー長
    SE_USOU_HAKO_Speck.ks6.keyflag = BtKfExt + _
                                        BtKfSeg + _
                                        BtKfDup + _
                                        BtKfChg                     ' キーフラグ
    SE_USOU_HAKO_Speck.ks6.keytype = Chr(BtKtString)                ' キータイプ
    SE_USOU_HAKO_Speck.ks6.reserve = &H0                            ' 予約済み

    SE_USOU_HAKO_Speck.ks7.keypos = 12                              ' キーポジション
    SE_USOU_HAKO_Speck.ks7.keyleng = 1                              ' キー長
    SE_USOU_HAKO_Speck.ks7.keyflag = BtKfExt + _
                                        BtKfSeg + _
                                        BtKfDup + _
                                        BtKfChg                     ' キーフラグ
    SE_USOU_HAKO_Speck.ks7.keytype = Chr(BtKtString)                ' キータイプ
    SE_USOU_HAKO_Speck.ks7.reserve = &H0                            ' 予約済み

    SE_USOU_HAKO_Speck.ks8.keypos = 13                              ' キーポジション
    SE_USOU_HAKO_Speck.ks8.keyleng = 20                             ' キー長
    SE_USOU_HAKO_Speck.ks8.keyflag = BtKfExt + _
                                        BtKfDup + _
                                        BtKfChg                     ' キーフラグ
    SE_USOU_HAKO_Speck.ks8.keytype = Chr(BtKtString)                ' キータイプ
    SE_USOU_HAKO_Speck.ks8.reserve = &H0                            ' 予約済み

'------------------------------------------------
                                                                ' キー１

                                                                ' キー２
'------------------------------------------------
    SE_USOU_HAKO_Speck.ks9.keypos = 3                               ' キーポジション
    SE_USOU_HAKO_Speck.ks9.keyleng = 8                              ' キー長
    SE_USOU_HAKO_Speck.ks9.keyflag = BtKfExt + BtKfSeg              ' キーフラグ
    SE_USOU_HAKO_Speck.ks9.keytype = Chr(BtKtString)                ' キータイプ
    SE_USOU_HAKO_Speck.ks9.reserve = &H0                            ' 予約済み

    SE_USOU_HAKO_Speck.ks10.keypos = 33                             ' キーポジション
    SE_USOU_HAKO_Speck.ks10.keyleng = 8                             ' キー長
    SE_USOU_HAKO_Speck.ks10.keyflag = BtKfExt + BtKfSeg             ' キーフラグ
    SE_USOU_HAKO_Speck.ks10.keytype = Chr(BtKtString)               ' キータイプ
    SE_USOU_HAKO_Speck.ks10.reserve = &H0                           ' 予約済み

    SE_USOU_HAKO_Speck.ks11.keypos = 11                             ' キーポジション
    SE_USOU_HAKO_Speck.ks11.keyleng = 1                             ' キー長
    SE_USOU_HAKO_Speck.ks11.keyflag = BtKfExt + BtKfSeg             ' キーフラグ
    SE_USOU_HAKO_Speck.ks11.keytype = Chr(BtKtString)               ' キータイプ
    SE_USOU_HAKO_Speck.ks11.reserve = &H0                           ' 予約済み

    SE_USOU_HAKO_Speck.ks12.keypos = 12                             ' キーポジション
    SE_USOU_HAKO_Speck.ks12.keyleng = 1                             ' キー長
    SE_USOU_HAKO_Speck.ks12.keyflag = BtKfExt + BtKfSeg             ' キーフラグ
    SE_USOU_HAKO_Speck.ks12.keytype = Chr(BtKtString)               ' キータイプ
    SE_USOU_HAKO_Speck.ks12.reserve = &H0                           ' 予約済み

    SE_USOU_HAKO_Speck.ks13.keypos = 13                             ' キーポジション
    SE_USOU_HAKO_Speck.ks13.keyleng = 20                            ' キー長
    SE_USOU_HAKO_Speck.ks13.keyflag = BtKfExt                       ' キーフラグ
    SE_USOU_HAKO_Speck.ks13.keytype = Chr(BtKtString)               ' キータイプ
    SE_USOU_HAKO_Speck.ks13.reserve = &H0                           ' 予約済み
'------------------------------------------------
                                                                ' キー２
    sts = BTRV(BtOpCreate, SE_USOU_HAKO_POS, SE_USOU_HAKO_Speck, Len(SE_USOU_HAKO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "輸送箱実績")
        Exit Function
    End If

    SE_USOU_HAKO_Create = False

End Function

Public Function SE_USOU_HAKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              輸送箱使用実績  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    SE_USOU_HAKO_Open = True
                                            '輸送箱使用実績 フルパス取込み
    sts = GetIni("FILE", SE_USOU_HAKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_USOU_HAKO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SE_USOU_HAKO_Create()        '輸送箱使用実績作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "輸送箱使用実績")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "輸送箱使用実績")
                Exit Function
        End Select
    Loop
    
    SE_USOU_HAKO_Open = False

End Function
