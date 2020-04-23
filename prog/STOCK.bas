Attribute VB_Name = "STOCK"
Option Explicit
'********************************************************************
'*                                                                  *
'*              棚卸しデータ  ファイル定義                          *
'*                                                                  *
'********************************************************************
'ファイルＩＤ
Public Const STOCK_ID$ = "STOCK"

'ページサイズ
Public Const STOCK_PG_SIZ% = 1024

'ポジション・ブロック
Public STOCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type STOCKREC_Tag
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    
    ST_SOKO(0 To 1)         As Byte     '標準入庫倉庫
    ST_RETU(0 To 1)         As Byte     '標準入庫倉庫
    ST_REN(0 To 1)          As Byte     '標準入庫倉庫
    ST_DAN(0 To 1)          As Byte     '標準入庫倉庫
    
    
    
    HOST_ZAIKO(0 To 7)      As Byte     '松下理論在庫
    POS_ZAIKO(0 To 7)       As Byte     'ＰＯＳ総在庫
    ST_ZAIKO(0 To 7)        As Byte     '標準棚番在庫
    
    EE1_LOCATION(0 To 7)    As Byte     '別置き１
    EE1_ZAIKO(0 To 7)       As Byte     '在庫
    EE2_LOCATION(0 To 7)    As Byte     '別置き２
    EE2_ZAIKO(0 To 7)       As Byte     '在庫
    EE3_LOCATION(0 To 7)    As Byte     '別置き３
    EE3_ZAIKO(0 To 7)       As Byte     '在庫
    
    ETC_ZAIKO(0 To 7)       As Byte     'その他在庫
    CHECK_MARK(0 To 0)      As Byte     '照合マーク
    PRINT_YMD(0 To 7)       As Byte     '印刷日付
    INPUT_YMD(0 To 7)       As Byte     '入力日付
    
    SAI_QTY(0 To 8)         As Byte     '差異数　2004.06.29
    
    BU_ZAI_QTY(0 To 7)      As Byte     'BU在庫     2007.08.22
    PPSC_ZAI_QTY(0 To 7)    As Byte     'PPSC在庫   2007.08.22
    
    
    
    FILLER(0 To 7)          As Byte
    
End Type
'データ・バッファ
Public STOCKREC As STOCKREC_Tag

'キー定義

Type KEY0_STOCK             'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type

Type KEY1_STOCK             'ＫＥＹ１
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    
    ST_SOKO(0 To 1)         As Byte     '標準入庫倉庫
    ST_RETU(0 To 1)         As Byte     '標準入庫倉庫
    ST_REN(0 To 1)          As Byte     '標準入庫倉庫
    ST_DAN(0 To 1)          As Byte     '標準入庫倉庫
    
    
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type

Type KEY2_STOCK             'ＫＥＹ２
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    ST_SOKO(0 To 1)         As Byte     '標準入庫倉庫   2007.08.22
    CHECK_MARK(0 To 0)      As Byte     '照合マーク
End Type



'全BU用　KEY定義
Type KEY3_STOCK             'ＫＥＹ３
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type

Type KEY4_STOCK             'ＫＥＹ４
    NAIGAI(0 To 0)          As Byte     '国内外
    
    ST_SOKO(0 To 1)         As Byte     '標準入庫倉庫
    ST_RETU(0 To 1)         As Byte     '標準入庫倉庫
    ST_REN(0 To 1)          As Byte     '標準入庫倉庫
    ST_DAN(0 To 1)          As Byte     '標準入庫倉庫
    
    
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type

Type KEY5_STOCK             'ＫＥＹ５
    NAIGAI(0 To 0)          As Byte     '国内外
    ST_SOKO(0 To 1)         As Byte     '標準入庫倉庫
    CHECK_MARK(0 To 0)      As Byte     '照合マーク
End Type




'キー・データ
Public K0_STOCK     As KEY0_STOCK
Public K1_STOCK     As KEY1_STOCK
Public K2_STOCK     As KEY2_STOCK

Public K3_STOCK     As KEY3_STOCK
Public K4_STOCK     As KEY4_STOCK
Public K5_STOCK     As KEY5_STOCK


Private Type STOCK_FSpeck
    
    fs      As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
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
    ks16    As BtKeySpeck
    ks17    As BtKeySpeck
    ks18    As BtKeySpeck

End Type

Private STOCK_Speck As STOCK_FSpeck
Private Function STOCK_Create() As Integer
'********************************************************************
'*                                                                  *
'*              棚卸しデータ  ＣＲＥＡＴＥ                          *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    STOCK_Create = True
                                        '棚卸しデータフルパス取込み
    sts = GetIni("FILE", STOCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [STOCK]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    STOCK_Speck.fs.recoleng = Len(STOCKREC)     ' レコード長
    STOCK_Speck.fs.PageSize = STOCK_PG_SIZ      ' ページサイズ
    
    STOCK_Speck.fs.idexnumb = 6                 ' インデックス数    全BU対応3-->6
    
    STOCK_Speck.fs.fileflag = 0                 ' ファイルフラグ
    STOCK_Speck.fs.reserve = &H0                ' 予約済み
'------------------------------------------------
                                                ' キー０
    STOCK_Speck.ks0.keypos = 1                  ' キーポジション
    STOCK_Speck.ks0.keyleng = 1                 ' キー長
    STOCK_Speck.ks0.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    STOCK_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks0.reserve = &H0               ' 予約済み
                                                
    STOCK_Speck.ks1.keypos = 2                  ' キーポジション
    STOCK_Speck.ks1.keyleng = 1                 ' キー長
    STOCK_Speck.ks1.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    STOCK_Speck.ks1.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks1.reserve = &H0               ' 予約済み
                                                
    STOCK_Speck.ks2.keypos = 3                  ' キーポジション
    STOCK_Speck.ks2.keyleng = 20                ' キー長
    STOCK_Speck.ks2.keyflag = BtKfExt           ' キーフラグ
    STOCK_Speck.ks2.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks2.reserve = &H0               ' 予約済み
'------------------------------------------------
                                                ' キー１
    STOCK_Speck.ks3.keypos = 1                  ' キーポジション
    STOCK_Speck.ks3.keyleng = 1                 ' キー長
    STOCK_Speck.ks3.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    STOCK_Speck.ks3.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks3.reserve = &H0               ' 予約済み
    
    STOCK_Speck.ks4.keypos = 2                  ' キーポジション
    STOCK_Speck.ks4.keyleng = 1                 ' キー長
    STOCK_Speck.ks4.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    STOCK_Speck.ks4.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks4.reserve = &H0               ' 予約済み
                                                
    STOCK_Speck.ks5.keypos = 23                 ' キーポジション
    STOCK_Speck.ks5.keyleng = 8                 ' キー長
    STOCK_Speck.ks5.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    STOCK_Speck.ks5.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks5.reserve = &H0               ' 予約済み
                                                
    STOCK_Speck.ks6.keypos = 3                  ' キーポジション
    STOCK_Speck.ks6.keyleng = 20                ' キー長
    STOCK_Speck.ks6.keyflag = BtKfExt           ' キーフラグ
    STOCK_Speck.ks6.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks6.reserve = &H0               ' 予約済み
'------------------------------------------------
                                                ' キー２
    STOCK_Speck.ks7.keypos = 1                  ' キーポジション
    STOCK_Speck.ks7.keyleng = 1                 ' キー長
    STOCK_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks7.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks7.reserve = &H0               ' 予約済み
    
    STOCK_Speck.ks8.keypos = 2                  ' キーポジション
    STOCK_Speck.ks8.keyleng = 1                 ' キー長
    STOCK_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks8.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks8.reserve = &H0               ' 予約済み
    
                                                
    STOCK_Speck.ks9.keypos = 23                 ' キーポジション
    STOCK_Speck.ks9.keyleng = 2                 ' キー長
    STOCK_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks9.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks9.reserve = &H0               ' 予約済み
    
    
    STOCK_Speck.ks10.keypos = 111                ' キーポジション
    STOCK_Speck.ks10.keyleng = 1                 ' キー長
    STOCK_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup
    STOCK_Speck.ks10.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks10.reserve = &H0               ' 予約済み
'------------------------------------------------
    
    
    
    
    
    
    
    
    
'------------------------------------------------
                                                ' キー３
                                                
    STOCK_Speck.ks11.keypos = 2                 ' キーポジション
    STOCK_Speck.ks11.keyleng = 1                ' キー長
    STOCK_Speck.ks11.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks11.keytype = Chr(BtKtString)  ' キータイプ
    STOCK_Speck.ks11.reserve = &H0              ' 予約済み
                                                
    STOCK_Speck.ks12.keypos = 3                 ' キーポジション
    STOCK_Speck.ks12.keyleng = 20               ' キー長
    STOCK_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfDup
    STOCK_Speck.ks12.keytype = Chr(BtKtString)  ' キータイプ
    STOCK_Speck.ks12.reserve = &H0              ' 予約済み
'------------------------------------------------
                                                ' キー４
    
    STOCK_Speck.ks13.keypos = 2                  ' キーポジション
    STOCK_Speck.ks13.keyleng = 1                 ' キー長
    STOCK_Speck.ks13.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks13.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks13.reserve = &H0               ' 予約済み
                                                
    STOCK_Speck.ks14.keypos = 23                 ' キーポジション
    STOCK_Speck.ks14.keyleng = 8                 ' キー長
    STOCK_Speck.ks14.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks14.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks14.reserve = &H0               ' 予約済み
                                                
    STOCK_Speck.ks15.keypos = 3                  ' キーポジション
    STOCK_Speck.ks15.keyleng = 20                ' キー長
    STOCK_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfDup
    STOCK_Speck.ks15.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks15.reserve = &H0               ' 予約済み
'------------------------------------------------
                                                ' キー５
    
    STOCK_Speck.ks16.keypos = 2                  ' キーポジション
    STOCK_Speck.ks16.keyleng = 1                 ' キー長
    STOCK_Speck.ks16.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks16.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks16.reserve = &H0               ' 予約済み
    
                                                
    STOCK_Speck.ks17.keypos = 23                 ' キーポジション
    STOCK_Speck.ks17.keyleng = 2                 ' キー長
    STOCK_Speck.ks17.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    STOCK_Speck.ks17.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks17.reserve = &H0               ' 予約済み

    
    STOCK_Speck.ks18.keypos = 111                ' キーポジション
    STOCK_Speck.ks18.keyleng = 1                 ' キー長
    STOCK_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfDup
    STOCK_Speck.ks18.keytype = Chr(BtKtString)   ' キータイプ
    STOCK_Speck.ks18.reserve = &H0               ' 予約済み
'------------------------------------------------
    
    
    
    
    
    
    sts = BTRV(BtOpCreate, STOCK_POS, STOCK_Speck, Len(STOCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "棚卸しデータ")
        Exit Function
    End If

    STOCK_Create = False

End Function

Function STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              棚卸しデータ  ＯＰＥＮ                              *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    STOCK_Open = True
                                    '棚卸しデータフルパス取込み
    sts = GetIni("FILE", STOCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [STOCK]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, STOCK_POS, STOCKREC, Len(STOCKREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = STOCK_Create()        '棚卸しデータ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, STOCK_POS, STOCKREC, Len(STOCKREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "棚卸しデータ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "棚卸しデータ")
                Exit Function
        End Select
    Loop
    
    STOCK_Open = False

End Function
