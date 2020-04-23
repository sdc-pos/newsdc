Attribute VB_Name = "GOODS_S"
Option Explicit
'********************************************************************
'*
'*              商品化集計ファイル（一時ファイル） ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const GOODS_S_ID$ = "GOODS_S"

'ページサイズ
Public Const GOODS_S_PG_SIZ% = 512

'ポジション・ブロック
Public GOODS_S_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Private Type GOODS_SREC_Tag
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    Soko_No(0 To 1)         As Byte     '仮想倉庫番号（在庫中）
    ST_SOKO(0 To 1)         As Byte     '標準棚番 倉庫
    ST_RETU(0 To 1)         As Byte     '標準棚番 列
    ST_REN(0 To 1)          As Byte     '標準棚番 連
    ST_DAN(0 To 1)          As Byte     '標準棚番 段
    PACKING_NO(0 To 3)      As Byte     '箱№
    SOKO_QTY(0 To 7)        As Byte     '仮想倉庫分在庫
    Sumi_QTY(0 To 7)        As Byte     '商品化済み在庫数
    Mi_QTY(0 To 7)          As Byte     '未商品在庫数
    AVE_SYUKA(0 To 7)       As Byte     '平均出荷数
    SUMI_PERCENT(0 To 7)    As Byte     '事前商品化状況

    KOSOU(0 To 19)          As Byte     '個装箱 2008.03.03
    GAISOU(0 To 19)         As Byte     '外装箱 2008.03.03


End Type

'データ・バッファ
Public GOODS_SREC             As GOODS_SREC_Tag

'キー定義
Type KEY0_GOODS_S                   'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    Soko_No(0 To 1)         As Byte     '仮想倉庫番号（在庫中）
End Type

Type KEY1_GOODS_S                   'ＫＥＹ１
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    Soko_No(0 To 1)         As Byte     '仮想倉庫番号（在庫中）
    SUMI_PERCENT(0 To 7)    As Byte     '事前商品化状況
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type


'キー・データ
Public K0_GOODS_S         As KEY0_GOODS_S
Public K1_GOODS_S         As KEY1_GOODS_S

Type GOODS_S_FSpeck
    fs      As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
End Type

Private GOODS_S_Speck As GOODS_S_FSpeck
Private Function GOODS_S_Create() As Integer
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

    GOODS_S_Create = True
                                            '商品化集計ファイル　フルパス取込み
    sts = GetIni("FILE", GOODS_S_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS_S]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    GOODS_S_Speck.fs.recoleng = Len(GOODS_SREC)     ' レコード長
    GOODS_S_Speck.fs.PageSize = GOODS_S_PG_SIZ      ' ページサイズ
    GOODS_S_Speck.fs.idexnumb = 2                   ' インデックス数
    GOODS_S_Speck.fs.fileflag = 0                   ' ファイルフラグ
    GOODS_S_Speck.fs.reserve = &H0                  ' 予約済み
'---------------------------------------------------'
                                                    ' キー０
    GOODS_S_Speck.ks0.keypos = 1                    ' キーポジション
    GOODS_S_Speck.ks0.keyleng = 24                  ' キー長
    GOODS_S_Speck.ks0.keyflag = BtKfExt             ' キーフラグ
    GOODS_S_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    GOODS_S_Speck.ks0.reserve = &H0                 ' 予約済み
'---------------------------------------------------'
                                                    ' キー１
    GOODS_S_Speck.ks1.keypos = 1                    ' キーポジション
    GOODS_S_Speck.ks1.keyleng = 1                   ' キー長
    GOODS_S_Speck.ks1.keyflag = BtKfExt + BtKfSeg   ' キーフラグ
    GOODS_S_Speck.ks1.keytype = Chr(BtKtString)     ' キータイプ
    GOODS_S_Speck.ks1.reserve = &H0                 ' 予約済み
                                                    
    GOODS_S_Speck.ks2.keypos = 2                    ' キーポジション
    GOODS_S_Speck.ks2.keyleng = 1                   ' キー長
    GOODS_S_Speck.ks2.keyflag = BtKfExt + BtKfSeg   ' キーフラグ
    GOODS_S_Speck.ks2.keytype = Chr(BtKtString)     ' キータイプ
    GOODS_S_Speck.ks2.reserve = &H0                 ' 予約済み
                                                    
    GOODS_S_Speck.ks3.keypos = 23                   ' キーポジション
    GOODS_S_Speck.ks3.keyleng = 2                   ' キー長
    GOODS_S_Speck.ks3.keyflag = BtKfExt + BtKfSeg   ' キーフラグ
    GOODS_S_Speck.ks3.keytype = Chr(BtKtString)     ' キータイプ
    GOODS_S_Speck.ks3.reserve = &H0                 ' 予約済み
                                                    
    GOODS_S_Speck.ks4.keypos = 69                   ' キーポジション
    GOODS_S_Speck.ks4.keyleng = 8                   ' キー長
    GOODS_S_Speck.ks4.keyflag = BtKfExt + BtKfSeg   ' キーフラグ
    GOODS_S_Speck.ks4.keytype = Chr(BtKtString)     ' キータイプ
    GOODS_S_Speck.ks4.reserve = &H0                 ' 予約済み
                                                    
    GOODS_S_Speck.ks5.keypos = 3                    ' キーポジション
    GOODS_S_Speck.ks5.keyleng = 20                  ' キー長
    GOODS_S_Speck.ks5.keyflag = BtKfExt             ' キーフラグ
    GOODS_S_Speck.ks5.keytype = Chr(BtKtString)     ' キータイプ
    GOODS_S_Speck.ks5.reserve = &H0                 ' 予約済み
'---------------------------------------------------'
    sts = BTRV(BtOpCreate, GOODS_S_POS, GOODS_S_Speck, Len(GOODS_S_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "商品化集計ファイル")
        Exit Function
    End If
    
    GOODS_S_Create = False

End Function
Public Function GOODS_S_Open(Mode As Integer) As Integer
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
    
    GOODS_S_Open = True
                                            '商品化集計ファイル　フルパス取込み
    sts = GetIni("FILE", GOODS_S_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [GOODS_S]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    
    
    sts = BTRV(BtOpClose, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), K0_GOODS_S, Len(K0_GOODS_S), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化集計ファイル")
        End If
    End If
    
    
    On Error Resume Next    '2007.11.14
    Kill (FullPath)         '2007.11.14
    On Error GoTo 0         '2007.11.14
    
    
    
    
    Do
        sts = BTRV(BtOpOpen, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = GOODS_S_Create()      '商品化集計ファイル　作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), ByVal FullPath, Len(FullPath), Mode)
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
    GOODS_S_Open = False

End Function

