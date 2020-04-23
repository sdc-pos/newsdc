Attribute VB_Name = "GOODS"
Option Explicit
'********************************************************************
'*
'*              商品化集計ファイル（一時ファイル） ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const GOODS_ID$ = "GOODS"

'ページサイズ
Public Const GOODS_PG_SIZ% = 1024

'ポジション・ブロック
Public GOODS_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Type GOODSREC_Tag
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    ST_SOKO(0 To 1)         As Byte     '標準棚番 倉庫
    ST_RETU(0 To 1)         As Byte     '標準棚番 列
    ST_REN(0 To 1)          As Byte     '標準棚番 連
    ST_DAN(0 To 1)          As Byte     '標準棚番 段
    PACKING_NO(0 To 3)      As Byte     '箱№
    Sumi_QTY(0 To 7)        As Byte     '商品化済み在庫数
    Mi_QTY(0 To 7)          As Byte     '未商品在庫数
    AVE_SYUKA(0 To 7)       As Byte     '平均出荷数
    SUMI_PERCENT(0 To 7)    As Byte     '事前商品化状況
End Type

'データ・バッファ
Public GOODSREC             As GOODSREC_Tag

'キー定義
Type KEY0_GOODS                    'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    ST_SOKO(0 To 1)         As Byte     '標準棚番
    SUMI_PERCENT(0 To 7)    As Byte     '事前商品化状況
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type


Type KEY1_GOODS                    'ＫＥＹ１
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    ST_SOKO(0 To 1)         As Byte     '標準棚番
    ST_RETU(0 To 1)         As Byte     '標準棚番 列
    ST_REN(0 To 1)          As Byte     '標準棚番 連
    ST_DAN(0 To 1)          As Byte     '標準棚番 段
    SUMI_PERCENT(0 To 7)    As Byte     '事前商品化状況
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type

Type KEY2_GOODS                    'ＫＥＹ２    2007.11.14
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type



'キー・データ
Public K0_GOODS         As KEY0_GOODS
Public K1_GOODS         As KEY1_GOODS
Public K2_GOODS         As KEY2_GOODS

Type GOODS_FSpeck
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
End Type

Private GOODS_Speck As GOODS_FSpeck
Private Function GOODS_Create() As Integer
'********************************************************************
'*
'*              商品化集計ファイル　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*      2007.11.14  :KEY2(事業部+国内外+品番(外))　追加
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    GOODS_Create = True
                                            '商品化集計ファイル　フルパス取込み
    sts = GetIni("FILE", GOODS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    GOODS_Speck.fs.recoleng = Len(GOODSREC)         ' レコード長
    GOODS_Speck.fs.PageSize = GOODS_PG_SIZ          ' ページサイズ
    GOODS_Speck.fs.idexnumb = 3                     ' インデックス数
    GOODS_Speck.fs.fileflag = 0                     ' ファイルフラグ
    GOODS_Speck.fs.reserve = &H0                    ' 予約済み
'---------------------------------------------------'
                                                    ' キー０
    GOODS_Speck.ks0.keypos = 1                      ' キーポジション
    GOODS_Speck.ks0.keyleng = 1                     ' キー長
    GOODS_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' キーフラグ
    GOODS_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks0.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks1.keypos = 2                      ' キーポジション
    GOODS_Speck.ks1.keyleng = 1                     ' キー長
    GOODS_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' キーフラグ
    GOODS_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks1.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks2.keypos = 23                     ' キーポジション
    GOODS_Speck.ks2.keyleng = 2                     ' キー長
    GOODS_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' キーフラグ
    GOODS_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks2.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks3.keypos = 59                     ' キーポジション
    GOODS_Speck.ks3.keyleng = 8                     ' キー長
    GOODS_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' キーフラグ
    GOODS_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks3.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks4.keypos = 3                      ' キーポジション
    GOODS_Speck.ks4.keyleng = 20                    ' キー長
    GOODS_Speck.ks4.keyflag = BtKfExt + BtKfChg               ' キーフラグ
    GOODS_Speck.ks4.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks4.reserve = &H0                   ' 予約済み


'---------------------------------------------------'
                                                    ' キー１
    GOODS_Speck.ks5.keypos = 1                      ' キーポジション
    GOODS_Speck.ks5.keyleng = 1                     ' キー長
    GOODS_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' キーフラグ
    GOODS_Speck.ks5.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks5.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks6.keypos = 2                      ' キーポジション
    GOODS_Speck.ks6.keyleng = 1                     ' キー長
    GOODS_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' キーフラグ
    GOODS_Speck.ks6.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks6.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks7.keypos = 23                     ' キーポジション
    GOODS_Speck.ks7.keyleng = 8                     ' キー長
    GOODS_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' キーフラグ
    GOODS_Speck.ks7.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks7.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks8.keypos = 59                     ' キーポジション
    GOODS_Speck.ks8.keyleng = 8                     ' キー長
    GOODS_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfSeg     ' キーフラグ
    GOODS_Speck.ks8.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks8.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks9.keypos = 3                      ' キーポジション
    GOODS_Speck.ks9.keyleng = 20                    ' キー長
    GOODS_Speck.ks9.keyflag = BtKfExt + BtKfChg               ' キーフラグ
    GOODS_Speck.ks9.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks9.reserve = &H0                   ' 予約済み

'---------------------------------------------------'
    
    
    
'---------------------------------------------------'
                                                    ' キー１
    GOODS_Speck.ks10.keypos = 1                      ' キーポジション
    GOODS_Speck.ks10.keyleng = 1                     ' キー長
    GOODS_Speck.ks10.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    GOODS_Speck.ks10.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks10.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks11.keypos = 2                      ' キーポジション
    GOODS_Speck.ks11.keyleng = 1                     ' キー長
    GOODS_Speck.ks11.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    GOODS_Speck.ks11.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks11.reserve = &H0                   ' 予約済み
    
    GOODS_Speck.ks12.keypos = 3                      ' キーポジション
    GOODS_Speck.ks12.keyleng = 20                    ' キー長
    GOODS_Speck.ks12.keyflag = BtKfExt               ' キーフラグ
    GOODS_Speck.ks12.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks12.reserve = &H0                   ' 予約済み
'---------------------------------------------------'
    
    
    
    sts = BTRV(BtOpCreate, GOODS_POS, GOODS_Speck, Len(GOODS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "商品化集計ファイル")
        Exit Function
    End If
    
    GOODS_Create = False

End Function
Public Function GOODS_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              商品化集計ファイル　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    GOODS_Open = True
                                            '商品化集計ファイル　フルパス取込み
    sts = GetIni("FILE", GOODS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    
    
    sts = BTRV(BtOpClose, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化集計ファイル")
        End If
    End If
    
    
    On Error Resume Next    '2007.11.14
    Kill (FullPath)         '2007.11.14
    On Error GoTo 0         '2007.11.14
    
    
    Do
        sts = BTRV(BtOpOpen, GOODS_POS, GOODSREC, Len(GOODSREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = GOODS_Create()        '商品化集計ファイル　作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, GOODS_POS, GOODSREC, Len(GOODSREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "商品化集計ファイル")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "商品化集計ファイル")
                Exit Function
        End Select
    Loop
    GOODS_Open = False

End Function

