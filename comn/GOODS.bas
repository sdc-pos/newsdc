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
Public Const GOODS_PG_SIZ% = 4096

'ポジション・ブロック
Public GOODS_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Type GOODSREC_Tag
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
    ST_SOKO(0 To 1)             As Byte     '標準棚番 倉庫
    ST_RETU(0 To 1)             As Byte     '標準棚番 列
    ST_REN(0 To 1)              As Byte     '標準棚番 連
    ST_DAN(0 To 1)              As Byte     '標準棚番 段
    PACKING_NO(0 To 3)          As Byte     '箱№
    Sumi_QTY(0 To 7)            As Byte     '商品化済み在庫数
    Mi_QTY(0 To 7)              As Byte     '未商品在庫数
    AVE_SYUKA(0 To 7)           As Byte     '平均出荷数
    SUMI_PERCENT(0 To 7)        As Byte     '事前商品化状況

    KOSOU(0 To 19)              As Byte     '個装箱 2008.03.03
    GAISOU(0 To 19)             As Byte     '外装箱 2008.03.03


'-------------------------------------  '2011.07.04
    KO_QTY(0 To 5)              As Byte     '子　員数(999V99)
    
    S_AVE_SYUKA_QTY1(0 To 7)    As Byte     '平均生産計画出荷数(1)
    S_AVE_SYUKA_QTY2(0 To 7)    As Byte     '平均生産計画出荷数(2)

    NAI_BUHIN(0 To 0)           As Byte     '国内供給部品区分
    GAI_BUHIN(0 To 0)           As Byte     '海外供給部品区分
'-------------------------------------  '2011.07.04

'-------------------------------------  '2011.09.15
    N_YOTEI_DT(0 To 7)          As Byte     '商品化用入荷予定日
    N_YOTEI_QTY(0 To 7)         As Byte     '商品化用入荷予定数
    N_YOTEI_KEY_NO(0 To 7)      As Byte     '商品化用入荷予定KEY_NO
'-------------------------------------  '2011.09.15



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


Type KEY3_GOODS                    'ＫＥＹ２    2008.03.03
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    ST_SOKO(0 To 1)         As Byte     '標準棚番 倉庫
    AVE_SYUKA(0 To 7)       As Byte     '平均出荷数
    Sumi_QTY(0 To 7)        As Byte     '商品化済み在庫数
    Mi_QTY(0 To 7)          As Byte     '未商品在庫数
    SUMI_PERCENT(0 To 7)    As Byte     '事前商品化状況
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type



'キー・データ
Public K0_GOODS         As KEY0_GOODS
Public K1_GOODS         As KEY1_GOODS
Public K2_GOODS         As KEY2_GOODS
Public K3_GOODS         As KEY3_GOODS

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

    ks13    As BtKeySpeck
    ks14    As BtKeySpeck
    ks15    As BtKeySpeck
    ks16    As BtKeySpeck
    ks17    As BtKeySpeck
    ks18    As BtKeySpeck
    ks19    As BtKeySpeck
    ks20    As BtKeySpeck


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
        Call LOG_OUT(LOG_F, "SYS.INI [GOODS]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    GOODS_Speck.fs.recoleng = Len(GOODSREC)         ' レコード長
    GOODS_Speck.fs.PageSize = GOODS_PG_SIZ          ' ページサイズ
    GOODS_Speck.fs.idexnumb = 4                     ' インデックス数
    GOODS_Speck.fs.fileflag = 0                     ' ファイルフラグ
    GOODS_Speck.fs.reserve = &H0                    ' 予約済み
'---------------------------------------------------'
                                                    ' キー０
    GOODS_Speck.ks0.keypos = 1                      ' キーポジション
    GOODS_Speck.ks0.keyleng = 1                     ' キー長
    GOODS_Speck.ks0.keyflag = BtKfExt + _
                                BtKfChg + _
                                BtKfDup + _
                                BtKfSeg             ' キーフラグ
    GOODS_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks0.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks1.keypos = 2                      ' キーポジション
    GOODS_Speck.ks1.keyleng = 1                     ' キー長
    GOODS_Speck.ks1.keyflag = BtKfExt + _
                                BtKfChg + _
                                BtKfDup + _
                                BtKfSeg             ' キーフラグ
    GOODS_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks1.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks2.keypos = 23                     ' キーポジション
    GOODS_Speck.ks2.keyleng = 2                     ' キー長
    GOODS_Speck.ks2.keyflag = BtKfExt + _
                                BtKfChg + _
                                BtKfDup + _
                                BtKfSeg             ' キーフラグ
    GOODS_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks2.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks3.keypos = 59                     ' キーポジション
    GOODS_Speck.ks3.keyleng = 8                     ' キー長
    GOODS_Speck.ks3.keyflag = BtKfExt + _
                                BtKfChg + _
                                BtKfDup + _
                                BtKfSeg             ' キーフラグ
    GOODS_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks3.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks4.keypos = 3                      ' キーポジション
    GOODS_Speck.ks4.keyleng = 20                    ' キー長
    GOODS_Speck.ks4.keyflag = BtKfExt + _
                                BtKfChg + _
                                BtKfDup             ' キーフラグ
    GOODS_Speck.ks4.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks4.reserve = &H0                   ' 予約済み


'---------------------------------------------------'
                                                    ' キー１
    GOODS_Speck.ks5.keypos = 1                      ' キーポジション
    GOODS_Speck.ks5.keyleng = 1                     ' キー長
    GOODS_Speck.ks5.keyflag = BtKfExt + _
                                BtKfChg + _
                                BtKfDup + _
                                BtKfSeg             ' キーフラグ
    GOODS_Speck.ks5.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks5.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks6.keypos = 2                      ' キーポジション
    GOODS_Speck.ks6.keyleng = 1                     ' キー長
    GOODS_Speck.ks6.keyflag = BtKfExt + _
                                BtKfChg + _
                                BtKfDup + _
                                BtKfSeg             ' キーフラグ
    GOODS_Speck.ks6.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks6.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks7.keypos = 23                     ' キーポジション
    GOODS_Speck.ks7.keyleng = 8                     ' キー長
    GOODS_Speck.ks7.keyflag = BtKfExt + _
                                BtKfChg + _
                                BtKfDup + _
                                BtKfSeg             ' キーフラグ
    GOODS_Speck.ks7.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks7.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks8.keypos = 59                     ' キーポジション
    GOODS_Speck.ks8.keyleng = 8                     ' キー長
    GOODS_Speck.ks8.keyflag = BtKfExt + _
                                BtKfChg + _
                                BtKfDup + _
                                BtKfSeg             ' キーフラグ
    GOODS_Speck.ks8.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks8.reserve = &H0                   ' 予約済み
                                                    
    GOODS_Speck.ks9.keypos = 3                      ' キーポジション
    GOODS_Speck.ks9.keyleng = 20                    ' キー長
    GOODS_Speck.ks9.keyflag = BtKfExt + _
                                BtKfChg + _
                                BtKfDup             ' キーフラグ
    GOODS_Speck.ks9.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_Speck.ks9.reserve = &H0                   ' 予約済み

'---------------------------------------------------'
    
    
    
'---------------------------------------------------'
                                                    ' キー２
    GOODS_Speck.ks10.keypos = 1                     ' キーポジション
    GOODS_Speck.ks10.keyleng = 1                    ' キー長
    GOODS_Speck.ks10.keyflag = BtKfExt + _
                                BtKfDup + _
                                BtKfSeg             ' キーフラグ
    GOODS_Speck.ks10.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks10.reserve = &H0                  ' 予約済み
                                                    
    GOODS_Speck.ks11.keypos = 2                     ' キーポジション
    GOODS_Speck.ks11.keyleng = 1                    ' キー長
    GOODS_Speck.ks11.keyflag = BtKfExt + _
                                BtKfDup + _
                                BtKfSeg             ' キーフラグ
    GOODS_Speck.ks11.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks11.reserve = &H0                  ' 予約済み
    
    GOODS_Speck.ks12.keypos = 3                     ' キーポジション
    GOODS_Speck.ks12.keyleng = 20                   ' キー長
    GOODS_Speck.ks12.keyflag = BtKfExt + _
                                BtKfDup             ' キーフラグ
    GOODS_Speck.ks12.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks12.reserve = &H0                  ' 予約済み
'---------------------------------------------------'
    
'---------------------------------------------------'   2008.03.03
                                                    ' キー３
                                                    
                                                    
    GOODS_Speck.ks13.keypos = 1                     ' キーポジション
    GOODS_Speck.ks13.keyleng = 1                    ' キー長
    GOODS_Speck.ks13.keyflag = BtKfExt + _
                                BtKfSeg + _
                                BtKfDup + _
                                BtKfChg             ' キーフラグ
    GOODS_Speck.ks13.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks13.reserve = &H0                  ' 予約済み
                                                    
    GOODS_Speck.ks14.keypos = 2                     ' キーポジション
    GOODS_Speck.ks14.keyleng = 1                    ' キー長
    GOODS_Speck.ks14.keyflag = BtKfExt + _
                                BtKfSeg + _
                                BtKfDup + _
                                BtKfChg             ' キーフラグ
    GOODS_Speck.ks14.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks14.reserve = &H0                  ' 予約済み
                                                    
    GOODS_Speck.ks15.keypos = 23                    ' キーポジション
    GOODS_Speck.ks15.keyleng = 2                    ' キー長
    GOODS_Speck.ks15.keyflag = BtKfExt + _
                                BtKfSeg + _
                                BtKfDup + _
                                BtKfChg             ' キーフラグ
    GOODS_Speck.ks15.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks15.reserve = &H0                  ' 予約済み
                                                    
                                                    
                                                    
    GOODS_Speck.ks16.keypos = 51                    ' キーポジション
    GOODS_Speck.ks16.keyleng = 8                    ' キー長
    GOODS_Speck.ks16.keyflag = BtKfExt + _
                                BtKfSeg + _
                                BtKfDec + _
                                BtKfDup + _
                                BtKfChg             ' キーフラグ
    GOODS_Speck.ks16.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks16.reserve = &H0                  ' 予約済み
                                                    
    GOODS_Speck.ks17.keypos = 35                    ' キーポジション
    GOODS_Speck.ks17.keyleng = 8                    ' キー長
    GOODS_Speck.ks17.keyflag = BtKfExt + _
                                BtKfSeg + _
                                BtKfDup + _
                                BtKfChg             ' キーフラグ
    GOODS_Speck.ks17.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks17.reserve = &H0                  ' 予約済み
                                                    
    GOODS_Speck.ks18.keypos = 43                    ' キーポジション
    GOODS_Speck.ks18.keyleng = 8                    ' キー長
    GOODS_Speck.ks18.keyflag = BtKfExt + _
                                BtKfSeg + _
                                BtKfDec + _
                                BtKfDup + _
                                BtKfChg             ' キーフラグ
    GOODS_Speck.ks18.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks18.reserve = &H0                  ' 予約済み
                                                    
    GOODS_Speck.ks19.keypos = 59                    ' キーポジション
    GOODS_Speck.ks19.keyleng = 8                    ' キー長
    GOODS_Speck.ks19.keyflag = BtKfExt + _
                                BtKfSeg + _
                                BtKfDup + _
                                BtKfChg             ' キーフラグ
    GOODS_Speck.ks19.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks19.reserve = &H0                  ' 予約済み
    
    GOODS_Speck.ks20.keypos = 3                     ' キーポジション
    GOODS_Speck.ks20.keyleng = 20                   ' キー長
    GOODS_Speck.ks20.keyflag = BtKfExt + _
                                BtKfDup + _
                                BtKfChg             ' キーフラグ
    GOODS_Speck.ks20.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_Speck.ks20.reserve = &H0                  ' 予約済み
    
    
    
    
    sts = BTRV(BtOpCreate, GOODS_POS, GOODS_Speck, Len(GOODS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "商品化集計ファイル")
        Exit Function
    End If
    
    GOODS_Create = False

End Function
Public Function GOODS_Open(Mode As Integer, Optional DEL_F = 0) As Integer
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
        Call LOG_OUT(LOG_F, "SYS.INI [GOODS]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    
''2011.10.01
    If DEL_F = 0 Then
        sts = BTRV(BtOpClose, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "商品化集計ファイル")
            End If
        End If
    
    
        On Error Resume Next    '2007.11.14
        Kill (FullPath)         '2007.11.14
        On Error GoTo 0         '2007.11.14
    End If
''2011.10.01
    
    
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

