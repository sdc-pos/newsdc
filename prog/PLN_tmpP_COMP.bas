Attribute VB_Name = "PLN_tmpP_COMP"
Option Explicit
'********************************************************************
'*
'*              資材所要量中間ファイル ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const PLN_tmpP_COMP_ID$ = "PLN_tmpP_COMP"

'ページサイズ
Public Const PLN_tmpP_COMP_PG_SIZ% = 1024

'ポジション・ブロック
Public PLN_tmpP_COMP_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Type PLN_tmpP_COMP_REC_Tag
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    KO_SYUBETSU(0 To 1)     As Byte     '種別
    KO_JGYOBU(0 To 0)       As Byte     '事業部区分
    KO_NAIGAI(0 To 0)       As Byte     '国内外
    KO_HIN_GAI(0 To 19)     As Byte     '品番（外部）
    YOTEI_DT(0 To 7)        As Byte     '商品化予定日付
    YOTEI_QTY(0 To 7)       As Byte     '商品化予定数
    KO_QTY(0 To 5)          As Byte     '子　員数(999V99)
    USE_QTY(0 To 5)         As Byte     '子　必要数
    DATA_KBN(0 To 0)        As Byte     'ﾃﾞｰﾀ区分
    INS_TANTO(0 To 9)       As Byte     '追加　担当者
    Ins_DateTime(0 To 13)   As Byte     '追加　日時         YYYYMMDDhhmmss

End Type

'データ・バッファ
Public PLN_tmpP_COMP_REC    As PLN_tmpP_COMP_REC_Tag

'キー定義
Type KEY0_PLN_tmpP_COMP                 'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    KO_SYUBETSU(0 To 1)     As Byte     '種別
    KO_JGYOBU(0 To 0)       As Byte     '事業部区分
    KO_NAIGAI(0 To 0)       As Byte     '国内外
    KO_HIN_GAI(0 To 19)     As Byte     '品番（外部）
    YOTEI_DT(0 To 7)        As Byte     '商品化予定日付
End Type

Type KEY1_PLN_tmpP_COMP                 'ＫＥＹ１
    YOTEI_DT(0 To 7)        As Byte     '商品化予定日付
    KO_SYUBETSU(0 To 1)     As Byte     '種別
    KO_JGYOBU(0 To 0)       As Byte     '事業部区分
    KO_NAIGAI(0 To 0)       As Byte     '国内外
    KO_HIN_GAI(0 To 19)     As Byte     '品番（外部）
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type

'キー・データ
Public K0_PLN_tmpP_COMP     As KEY0_PLN_tmpP_COMP
Public K1_PLN_tmpP_COMP     As KEY1_PLN_tmpP_COMP

Type PLN_tmpP_COMP_FSpeck
    fs      As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
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
    ks14    As BtKeySpeck
    ks15    As BtKeySpeck
End Type

Private PLN_tmpP_COMP_Speck  As PLN_tmpP_COMP_FSpeck
Private Function PLN_tmpP_COMP_Create() As Integer
'********************************************************************
'*
'*              資材所要量中間ファイル　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PLN_tmpP_COMP_Create = True
                                            '資材所要量中間ファイル　フルパス取込み
    sts = GetIni("FILE", PLN_tmpP_COMP_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_tmpP_COMP]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    PLN_tmpP_COMP_Speck.fs.recoleng = Len(PLN_tmpP_COMP_REC)    ' レコード長
    PLN_tmpP_COMP_Speck.fs.PageSize = PLN_tmpP_COMP_PG_SIZ      ' ページサイズ
    PLN_tmpP_COMP_Speck.fs.idexnumb = 2                         ' インデックス数
    PLN_tmpP_COMP_Speck.fs.fileflag = 0                         ' ファイルフラグ
    PLN_tmpP_COMP_Speck.fs.reserve = &H0                        ' 予約済み
'---------------------------------------------------'
                                                    ' キー０
    PLN_tmpP_COMP_Speck.ks0.keypos = 1                          ' キーポジション
    PLN_tmpP_COMP_Speck.ks0.keyleng = 1                         ' キー長
    PLN_tmpP_COMP_Speck.ks0.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    PLN_tmpP_COMP_Speck.ks0.keytype = Chr(BtKtString)           ' キータイプ
    PLN_tmpP_COMP_Speck.ks0.reserve = &H0                       ' 予約済み
                                                    
    PLN_tmpP_COMP_Speck.ks1.keypos = 2                          ' キーポジション
    PLN_tmpP_COMP_Speck.ks1.keyleng = 1                         ' キー長
    PLN_tmpP_COMP_Speck.ks1.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    PLN_tmpP_COMP_Speck.ks1.keytype = Chr(BtKtString)           ' キータイプ
    PLN_tmpP_COMP_Speck.ks1.reserve = &H0                       ' 予約済み
                                                    
    PLN_tmpP_COMP_Speck.ks2.keypos = 3                          ' キーポジション
    PLN_tmpP_COMP_Speck.ks2.keyleng = 20                        ' キー長
    PLN_tmpP_COMP_Speck.ks2.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    PLN_tmpP_COMP_Speck.ks2.keytype = Chr(BtKtString)           ' キータイプ
    PLN_tmpP_COMP_Speck.ks2.reserve = &H0                       ' 予約済み
                                                    
                                                    
    PLN_tmpP_COMP_Speck.ks3.keypos = 23                         ' キーポジション
    PLN_tmpP_COMP_Speck.ks3.keyleng = 2                         ' キー長
    PLN_tmpP_COMP_Speck.ks3.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    PLN_tmpP_COMP_Speck.ks3.keytype = Chr(BtKtString)           ' キータイプ
    PLN_tmpP_COMP_Speck.ks3.reserve = &H0                       ' 予約済み
                                                    
    PLN_tmpP_COMP_Speck.ks4.keypos = 25                         ' キーポジション
    PLN_tmpP_COMP_Speck.ks4.keyleng = 1                         ' キー長
    PLN_tmpP_COMP_Speck.ks4.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    PLN_tmpP_COMP_Speck.ks4.keytype = Chr(BtKtString)           ' キータイプ
    PLN_tmpP_COMP_Speck.ks4.reserve = &H0                       ' 予約済み
    
    PLN_tmpP_COMP_Speck.ks5.keypos = 26                         ' キーポジション
    PLN_tmpP_COMP_Speck.ks5.keyleng = 1                         ' キー長
    PLN_tmpP_COMP_Speck.ks5.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    PLN_tmpP_COMP_Speck.ks5.keytype = Chr(BtKtString)           ' キータイプ
    PLN_tmpP_COMP_Speck.ks5.reserve = &H0                       ' 予約済み
    
    PLN_tmpP_COMP_Speck.ks6.keypos = 27                         ' キーポジション
    PLN_tmpP_COMP_Speck.ks6.keyleng = 20                        ' キー長
    PLN_tmpP_COMP_Speck.ks6.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    PLN_tmpP_COMP_Speck.ks6.keytype = Chr(BtKtString)           ' キータイプ
    PLN_tmpP_COMP_Speck.ks6.reserve = &H0                       ' 予約済み
                                                    
    PLN_tmpP_COMP_Speck.ks7.keypos = 47                         ' キーポジション
    PLN_tmpP_COMP_Speck.ks7.keyleng = 8                         ' キー長
    PLN_tmpP_COMP_Speck.ks7.keyflag = BtKfExt                   ' キーフラグ
    PLN_tmpP_COMP_Speck.ks7.keytype = Chr(BtKtString)           ' キータイプ
    PLN_tmpP_COMP_Speck.ks7.reserve = &H0                       ' 予約済み
                                                    
'---------------------------------------------------'
                                                    ' キー１
    PLN_tmpP_COMP_Speck.ks8.keypos = 47                        ' キーポジション
    PLN_tmpP_COMP_Speck.ks8.keyleng = 8                        ' キー長
    PLN_tmpP_COMP_Speck.ks8.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    PLN_tmpP_COMP_Speck.ks8.keytype = Chr(BtKtString)          ' キータイプ
    PLN_tmpP_COMP_Speck.ks8.reserve = &H0                      ' 予約済み
    
    
    PLN_tmpP_COMP_Speck.ks9.keypos = 23                         ' キーポジション
    PLN_tmpP_COMP_Speck.ks9.keyleng = 2                         ' キー長
    PLN_tmpP_COMP_Speck.ks9.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    PLN_tmpP_COMP_Speck.ks9.keytype = Chr(BtKtString)           ' キータイプ
    PLN_tmpP_COMP_Speck.ks9.reserve = &H0                       ' 予約済み
                                                    
    PLN_tmpP_COMP_Speck.ks10.keypos = 25                         ' キーポジション
    PLN_tmpP_COMP_Speck.ks10.keyleng = 1                         ' キー長
    PLN_tmpP_COMP_Speck.ks10.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    PLN_tmpP_COMP_Speck.ks10.keytype = Chr(BtKtString)           ' キータイプ
    PLN_tmpP_COMP_Speck.ks10.reserve = &H0                       ' 予約済み
    
    PLN_tmpP_COMP_Speck.ks11.keypos = 26                        ' キーポジション
    PLN_tmpP_COMP_Speck.ks11.keyleng = 1                        ' キー長
    PLN_tmpP_COMP_Speck.ks11.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    PLN_tmpP_COMP_Speck.ks11.keytype = Chr(BtKtString)          ' キータイプ
    PLN_tmpP_COMP_Speck.ks11.reserve = &H0                      ' 予約済み
    
    PLN_tmpP_COMP_Speck.ks12.keypos = 27                        ' キーポジション
    PLN_tmpP_COMP_Speck.ks12.keyleng = 20                       ' キー長
    PLN_tmpP_COMP_Speck.ks12.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    PLN_tmpP_COMP_Speck.ks12.keytype = Chr(BtKtString)          ' キータイプ
    PLN_tmpP_COMP_Speck.ks12.reserve = &H0                      ' 予約済み
                                                    
    PLN_tmpP_COMP_Speck.ks13.keypos = 1                         ' キーポジション
    PLN_tmpP_COMP_Speck.ks13.keyleng = 1                        ' キー長
    PLN_tmpP_COMP_Speck.ks13.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    PLN_tmpP_COMP_Speck.ks13.keytype = Chr(BtKtString)          ' キータイプ
    PLN_tmpP_COMP_Speck.ks13.reserve = &H0                      ' 予約済み
                                                    
    PLN_tmpP_COMP_Speck.ks14.keypos = 2                         ' キーポジション
    PLN_tmpP_COMP_Speck.ks14.keyleng = 1                        ' キー長
    PLN_tmpP_COMP_Speck.ks14.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    PLN_tmpP_COMP_Speck.ks14.keytype = Chr(BtKtString)          ' キータイプ
    PLN_tmpP_COMP_Speck.ks14.reserve = &H0                      ' 予約済み
                                                    
    PLN_tmpP_COMP_Speck.ks15.keypos = 3                         ' キーポジション
    PLN_tmpP_COMP_Speck.ks15.keyleng = 20                       ' キー長
    PLN_tmpP_COMP_Speck.ks15.keyflag = BtKfExt                  ' キーフラグ
    PLN_tmpP_COMP_Speck.ks15.keytype = Chr(BtKtString)          ' キータイプ
    PLN_tmpP_COMP_Speck.ks15.reserve = &H0                      ' 予約済み
                                                    
                                                    
                                                    
'---------------------------------------------------'
    sts = BTRV(BtOpCreate, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_Speck, Len(PLN_tmpP_COMP_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "資材所要量中間ファイル")
        Exit Function
    End If
    PLN_tmpP_COMP_Create = False
End Function
Public Function PLN_tmpP_COMP_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              資材所要量中間ファイル　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    PLN_tmpP_COMP_Open = True
                                            '資材所要量中間ファイル　フルパス取込み
    sts = GetIni("FILE", PLN_tmpP_COMP_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_P_COMP]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_REC, Len(PLN_tmpP_COMP_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PLN_tmpP_COMP_Create()   '資材所要量中間ファイル　作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_REC, Len(PLN_tmpP_COMP_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "資材所要量中間ファイル")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "資材所要量中間ファイル")
                Exit Function
        End Select
    Loop
    PLN_tmpP_COMP_Open = False

End Function

