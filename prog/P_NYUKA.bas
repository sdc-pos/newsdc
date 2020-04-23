Attribute VB_Name = "P_NYU"
Option Explicit
'********************************************************************
'*
'*              資材入荷チェックデータ（前借データ）　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const P_NYU_ID$ = "P_NYU"

'ページサイズ
Public Const P_NYU_PG_SIZ% = 1024

'ポジション・ブロック
Public P_NYU_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type P_NYUREC_Tag
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品番（外部）
    NYUKA_DT(0 To 7)        As Byte         '入荷日
    NYUKA_QTY(0 To 7)       As Byte         '前借数
    SOUSAI_DT(0 To 7)       As Byte         '最新相殺日付
    SOUSAI_QTY(0 To 7)      As Byte         '相殺数量
    WS_ID(0 To 2)           As Byte         '登録端末
    
    SHIIRE_CODE(0 To 4)     As Byte         '仕入先ｺｰﾄﾞ
    SHIIRE_TANKA(0 To 10)   As Byte         '仕入単価(9(8)V99)
    
    FILLER(0 To 40)         As Byte         'FILLER
    UPD_DATETIME(0 To 13)   As Byte         '更新日時
End Type

'データ・バッファ
Public P_NYUREC         As P_NYUREC_Tag

'キー定義
Type KEY0_P_NYU                         'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品番（外部）
    NYUKA_DT(0 To 7)        As Byte         '入荷日
End Type


'キー定義
Type KEY1_P_NYU                         'ＫＥＹ１
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品番（外部）
    SHIIRE_CODE(0 To 4)     As Byte         '仕入先ｺｰﾄﾞ
    SHIIRE_TANKA(0 To 10)   As Byte         '仕入単価(9(8)V99)
End Type


'キー・データ
Public K0_P_NYU         As KEY0_P_NYU
Public K1_P_NYU         As KEY1_P_NYU

Type P_NYU_FSpeck
    fs              As BtFileSpeck      'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
    ks1             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
    ks2             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
    ks3             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
    ks4             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
    ks5             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
    ks6             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
    ks7             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
    ks8             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_NYU_Speck     As P_NYU_FSpeck

Private Function P_NYU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              資材前借データ　ＣＲＥＡＴＥ                        *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    P_NYU_Create = True
                                            '資材前借データフルパス取込み
    sts = GetIni("FILE", P_NYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_NYU]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    P_NYU_Speck.fs.recoleng = Len(P_NYUREC)     ' レコード長
    P_NYU_Speck.fs.PageSize = P_NYU_PG_SIZ      ' ページサイズ
    P_NYU_Speck.fs.idexnumb = 2                 ' インデックス数
    P_NYU_Speck.fs.fileflag = 0                 ' ファイルフラグ
    P_NYU_Speck.fs.reserve = &H0                ' 予約済み
'------------------------------------------------
                                                ' キー０
    P_NYU_Speck.ks0.keypos = 1                  ' キーポジション
    P_NYU_Speck.ks0.keyleng = 1                 ' キー長
    P_NYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    P_NYU_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    P_NYU_Speck.ks0.reserve = &H0               ' 予約済み
                                                ' キー０
    P_NYU_Speck.ks1.keypos = 2                  ' キーポジション
    P_NYU_Speck.ks1.keyleng = 1                 ' キー長
    P_NYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    P_NYU_Speck.ks1.keytype = Chr(BtKtString)   ' キータイプ
    P_NYU_Speck.ks1.reserve = &H0               ' 予約済み
                                                ' キー０
    P_NYU_Speck.ks2.keypos = 3                  ' キーポジション
    P_NYU_Speck.ks2.keyleng = 20                ' キー長
    P_NYU_Speck.ks2.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    P_NYU_Speck.ks2.keytype = Chr(BtKtString)   ' キータイプ
    P_NYU_Speck.ks2.reserve = &H0               ' 予約済み
                                                ' キー０
    P_NYU_Speck.ks3.keypos = 23                 ' キーポジション
    P_NYU_Speck.ks3.keyleng = 8                 ' キー長
    P_NYU_Speck.ks3.keyflag = BtKfExt           ' キーフラグ
    P_NYU_Speck.ks3.keytype = Chr(BtKtString)   ' キータイプ
    P_NYU_Speck.ks3.reserve = &H0               ' 予約済み

'------------------------------------------------
                                                ' キー１
    P_NYU_Speck.ks4.keypos = 1                  ' キーポジション
    P_NYU_Speck.ks4.keyleng = 1                 ' キー長
                                                ' キーフラグ
    P_NYU_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    P_NYU_Speck.ks4.keytype = Chr(BtKtString)   ' キータイプ
    P_NYU_Speck.ks4.reserve = &H0               ' 予約済み
                                                ' キー１
    P_NYU_Speck.ks5.keypos = 2                  ' キーポジション
    P_NYU_Speck.ks5.keyleng = 1                 ' キー長
                                                ' キーフラグ
    P_NYU_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    P_NYU_Speck.ks5.keytype = Chr(BtKtString)   ' キータイプ
    P_NYU_Speck.ks5.reserve = &H0               ' 予約済み
                                                ' キー１
    P_NYU_Speck.ks6.keypos = 3                  ' キーポジション
    P_NYU_Speck.ks6.keyleng = 20                ' キー長
                                                 ' キーフラグ
    P_NYU_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    P_NYU_Speck.ks6.keytype = Chr(BtKtString)   ' キータイプ
    P_NYU_Speck.ks6.reserve = &H0               ' 予約済み
                                                ' キー１
    P_NYU_Speck.ks7.keypos = 58                 ' キーポジション
    P_NYU_Speck.ks7.keyleng = 5                 ' キー長
                                                ' キーフラグ
    P_NYU_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    P_NYU_Speck.ks7.keytype = Chr(BtKtString)   ' キータイプ
    P_NYU_Speck.ks7.reserve = &H0               ' 予約済み

                                                ' キー１
    P_NYU_Speck.ks8.keypos = 63                 ' キーポジション
    P_NYU_Speck.ks8.keyleng = 11                ' キー長
                                                ' キーフラグ
    P_NYU_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_NYU_Speck.ks8.keytype = Chr(BtKtString)   ' キータイプ
    P_NYU_Speck.ks8.reserve = &H0               ' 予約済み


'------------------------------------------------

    sts = BTRV(BtOpCreate, P_NYU_POS, P_NYU_Speck, Len(P_NYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "資材前借データ")
        Exit Function
    End If
    
    P_NYU_Create = False

End Function
Public Function P_NYU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              資材前借データ　ＯＰＥＮ                            *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    P_NYU_Open = True
                                        '資材前借データフルパス取込み
    sts = GetIni("FILE", P_NYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_NYU]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, P_NYU_POS, P_NYUREC, Len(P_NYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_NYU_Create()        '資材前借データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_NYU_POS, P_NYUREC, Len(P_NYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "資材前借データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "資材前借データ")
                Exit Function
        End Select
    Loop

    P_NYU_Open = False

End Function


