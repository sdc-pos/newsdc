Attribute VB_Name = "GOODS_ONO"
Option Explicit
'********************************************************************
'*
'*              商品化集計ファイル（一時ファイル） ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const GOODS_ONO_ID$ = "GOODS_ONO"

'ページサイズ
Public Const GOODS_ONO_PG_SIZ% = 1024

'ポジション・ブロック
Public GOODS_ONO_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Type GOODS_ONOREC_Tag
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
Public GOODS_ONOREC         As GOODS_ONOREC_Tag

'キー定義
Type KEY0_GOODS_ONO                 'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type


Type KEY1_GOODS_ONO                 'ＫＥＹ１
    
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    AVE_SYUKA(0 To 7)       As Byte     '平均出荷数
    Sumi_QTY(0 To 7)        As Byte     '商品化済み在庫数
    Mi_QTY(0 To 7)          As Byte     '未商品在庫数
    SUMI_PERCENT(0 To 7)    As Byte     '事前商品化状況
    ST_SOKO(0 To 1)         As Byte     '標準棚番
    ST_RETU(0 To 1)         As Byte     '標準棚番 列
    ST_REN(0 To 1)          As Byte     '標準棚番 連
    ST_DAN(0 To 1)          As Byte     '標準棚番 段
    HIN_GAI(0 To 19)        As Byte     '品番（外部）

End Type


'キー・データ
Public K0_GOODS_ONO     As KEY0_GOODS_ONO
Public K1_GOODS_ONO     As KEY1_GOODS_ONO

Type GOODS_ONO_FSpeck
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
End Type

Private GOODS_ONO_Speck As GOODS_ONO_FSpeck
Private Function GOODS_ONO_Create() As Integer
'********************************************************************
'*
'*              商品化集計ファイル　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    GOODS_ONO_Create = True
                                            '商品化集計ファイル　フルパス取込み
    sts = GetIni("FILE", GOODS_ONO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS_ONO]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    GOODS_ONO_Speck.fs.recoleng = Len(GOODS_ONOREC) ' レコード長
    GOODS_ONO_Speck.fs.PageSize = GOODS_ONO_PG_SIZ  ' ページサイズ
    GOODS_ONO_Speck.fs.idexnumb = 2                 ' インデックス数
    GOODS_ONO_Speck.fs.fileflag = 0                 ' ファイルフラグ
    GOODS_ONO_Speck.fs.reserve = &H0                ' 予約済み
'---------------------------------------------------'
                                                    ' キー０
    GOODS_ONO_Speck.ks0.keypos = 1                  ' キーポジション
    GOODS_ONO_Speck.ks0.keyleng = 1                 ' キー長
    GOODS_ONO_Speck.ks0.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    GOODS_ONO_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    GOODS_ONO_Speck.ks0.reserve = &H0               ' 予約済み
                                                    
    GOODS_ONO_Speck.ks1.keypos = 2                  ' キーポジション
    GOODS_ONO_Speck.ks1.keyleng = 1                 ' キー長
    GOODS_ONO_Speck.ks1.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    GOODS_ONO_Speck.ks1.keytype = Chr(BtKtString)   ' キータイプ
    GOODS_ONO_Speck.ks1.reserve = &H0               ' 予約済み
                                                    
                                                    
    GOODS_ONO_Speck.ks2.keypos = 3                  ' キーポジション
    GOODS_ONO_Speck.ks2.keyleng = 20                ' キー長
    GOODS_ONO_Speck.ks2.keyflag = BtKfExt           ' キーフラグ
    GOODS_ONO_Speck.ks2.keytype = Chr(BtKtString)   ' キータイプ
    GOODS_ONO_Speck.ks2.reserve = &H0               ' 予約済み


'---------------------------------------------------'
                                                    ' キー１
    GOODS_ONO_Speck.ks3.keypos = 1                      ' キーポジション
    GOODS_ONO_Speck.ks3.keyleng = 1                     ' キー長
    GOODS_ONO_Speck.ks3.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    GOODS_ONO_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_ONO_Speck.ks3.reserve = &H0                   ' 予約済み
                                                    
    GOODS_ONO_Speck.ks4.keypos = 2                      ' キーポジション
    GOODS_ONO_Speck.ks4.keyleng = 1                     ' キー長
    GOODS_ONO_Speck.ks4.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    GOODS_ONO_Speck.ks4.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_ONO_Speck.ks4.reserve = &H0                   ' 予約済み
    
    GOODS_ONO_Speck.ks5.keypos = 51                    ' キーポジション
    GOODS_ONO_Speck.ks5.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    GOODS_ONO_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDec
    GOODS_ONO_Speck.ks5.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_ONO_Speck.ks5.reserve = &H0                  ' 予約済み
                                                    
    GOODS_ONO_Speck.ks6.keypos = 35                    ' キーポジション
    GOODS_ONO_Speck.ks6.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    GOODS_ONO_Speck.ks6.keyflag = BtKfExt + BtKfSeg
    GOODS_ONO_Speck.ks6.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_ONO_Speck.ks6.reserve = &H0                  ' 予約済み
                                                    
    GOODS_ONO_Speck.ks7.keypos = 43                    ' キーポジション
    GOODS_ONO_Speck.ks7.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    GOODS_ONO_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDec
    GOODS_ONO_Speck.ks7.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_ONO_Speck.ks7.reserve = &H0                  ' 予約済み
                                                    
    GOODS_ONO_Speck.ks8.keypos = 59                    ' キーポジション
    GOODS_ONO_Speck.ks8.keyleng = 8                    ' キー長
                                                    ' キーフラグ
    GOODS_ONO_Speck.ks8.keyflag = BtKfExt + BtKfSeg
    GOODS_ONO_Speck.ks8.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_ONO_Speck.ks8.reserve = &H0                  ' 予約済み
    
    GOODS_ONO_Speck.ks9.keypos = 23                     ' キーポジション
    GOODS_ONO_Speck.ks9.keyleng = 2                     ' キー長
    GOODS_ONO_Speck.ks9.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    GOODS_ONO_Speck.ks9.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_ONO_Speck.ks9.reserve = &H0                   ' 予約済み
    
    GOODS_ONO_Speck.ks10.keypos = 25                     ' キーポジション
    GOODS_ONO_Speck.ks10.keyleng = 2                     ' キー長
    GOODS_ONO_Speck.ks10.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    GOODS_ONO_Speck.ks10.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_ONO_Speck.ks10.reserve = &H0                   ' 予約済み
    
    GOODS_ONO_Speck.ks11.keypos = 27                     ' キーポジション
    GOODS_ONO_Speck.ks11.keyleng = 2                     ' キー長
    GOODS_ONO_Speck.ks11.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    GOODS_ONO_Speck.ks11.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_ONO_Speck.ks11.reserve = &H0                   ' 予約済み
    
    GOODS_ONO_Speck.ks12.keypos = 29                    ' キーポジション
    GOODS_ONO_Speck.ks12.keyleng = 2                    ' キー長
    GOODS_ONO_Speck.ks12.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    GOODS_ONO_Speck.ks12.keytype = Chr(BtKtString)      ' キータイプ
    GOODS_ONO_Speck.ks12.reserve = &H0                  ' 予約済み
                                                    
                                                    
    GOODS_ONO_Speck.ks13.keypos = 3                      ' キーポジション
    GOODS_ONO_Speck.ks13.keyleng = 20                    ' キー長
    GOODS_ONO_Speck.ks13.keyflag = BtKfExt               ' キーフラグ
    GOODS_ONO_Speck.ks13.keytype = Chr(BtKtString)       ' キータイプ
    GOODS_ONO_Speck.ks13.reserve = &H0                   ' 予約済み

'---------------------------------------------------'
    sts = BTRV(BtOpCreate, GOODS_ONO_POS, GOODS_ONO_Speck, Len(GOODS_ONO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "商品化集計ファイル")
        Exit Function
    End If
    
    GOODS_ONO_Create = False

End Function
Public Function GOODS_ONO_Open(Mode As Integer) As Integer
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
    
    GOODS_ONO_Open = True
                                            '商品化集計ファイル　フルパス取込み
    sts = GetIni("FILE", GOODS_ONO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS_ONO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, GOODS_ONO_POS, GOODS_ONOREC, Len(GOODS_ONOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = GOODS_ONO_Create()    '商品化集計ファイル　作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, GOODS_ONO_POS, GOODS_ONOREC, Len(GOODS_ONOREC), ByVal FullPath, Len(FullPath), Mode)
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
    GOODS_ONO_Open = False

End Function

