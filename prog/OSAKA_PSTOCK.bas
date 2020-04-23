Attribute VB_Name = "OSAKA_PSTOCK"
Option Explicit
'********************************************************************
'*
'*              大阪ＰＣ　循環棚卸Ｆ ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const OSAKA_PSTOCK_ID$ = "OSAKA_PSTOCK"

'ページサイズ
Public Const OSAKA_PSTOCK_PG_SIZ% = 2048

'ポジション・ブロック
Public OSAKA_PSTOCK_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Type OSAKA_PSTOCKREC_Tag
    Soko_No(0 To 1)             As Byte     '倉庫№
    Retu(0 To 1)                As Byte     '棚番　列
    Ren(0 To 1)                 As Byte     '棚番　連
    Dan(0 To 1)                 As Byte     '棚番　段
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
    
    KEIJYO_YM(0 To 5)           As Byte     '計上年月
        
    NYUKO_QTY(0 To 9)           As Byte     '当月入庫数
    SYUKO_QTY(0 To 9)           As Byte     '当月出庫数
    ZAIKO_QTY(0 To 9)           As Byte     '当月在庫残数
    FILLER(0 To 47)             As Byte     'FILLER

    Ins_DateTime(0 To 13)       As Byte     'ﾃﾞｰﾀ作成日時

End Type

'データ・バッファ
Public OSAKA_PSTOCKREC          As OSAKA_PSTOCKREC_Tag

'キー定義
Type KEY0_OSAKA_PSTOCK                      'ＫＥＹ０
    Soko_No(0 To 1)             As Byte     '倉庫№
    Retu(0 To 1)                As Byte     '棚番　列
    Ren(0 To 1)                 As Byte     '棚番　連
    Dan(0 To 1)                 As Byte     '棚番　段
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
End Type



'キー・データ
Public K0_OSAKA_PSTOCK          As KEY0_OSAKA_PSTOCK

Type OSAKA_PSTOCK_FSpeck
    fs      As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
End Type

Private OSAKA_PSTOCK_Speck  As OSAKA_PSTOCK_FSpeck
Private Function OSAKA_PSTOCK_Create() As Integer
'********************************************************************
'*
'*              大阪ＰＣ　循環棚卸Ｆ　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    OSAKA_PSTOCK_Create = True
                                            '在庫データ　フルパス取込み
    sts = GetIni("FILE", OSAKA_PSTOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [OSAKA_PSTOCK]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    OSAKA_PSTOCK_Speck.fs.recoleng = Len(OSAKA_PSTOCKREC)           ' レコード長
    OSAKA_PSTOCK_Speck.fs.PageSize = OSAKA_PSTOCK_PG_SIZ            ' ページサイズ
    OSAKA_PSTOCK_Speck.fs.idexnumb = 1                              ' インデックス数
    OSAKA_PSTOCK_Speck.fs.fileflag = 0                              ' ファイルフラグ
    OSAKA_PSTOCK_Speck.fs.reserve = &H0                             ' 予約済み
'---------------------------------------------------'
                                                    ' キー０
    OSAKA_PSTOCK_Speck.ks0.keypos = 1                               ' キーポジション
    OSAKA_PSTOCK_Speck.ks0.keyleng = 2                              ' キー長
    OSAKA_PSTOCK_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' キーフラグ
    OSAKA_PSTOCK_Speck.ks0.keytype = Chr(BtKtString)                ' キータイプ
    OSAKA_PSTOCK_Speck.ks0.reserve = &H0                            ' 予約済み
                                                    
    OSAKA_PSTOCK_Speck.ks1.keypos = 3                               ' キーポジション
    OSAKA_PSTOCK_Speck.ks1.keyleng = 2                              ' キー長
    OSAKA_PSTOCK_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' キーフラグ
    OSAKA_PSTOCK_Speck.ks1.keytype = Chr(BtKtString)                ' キータイプ
    OSAKA_PSTOCK_Speck.ks1.reserve = &H0                            ' 予約済み
                                                    
    OSAKA_PSTOCK_Speck.ks2.keypos = 5                               ' キーポジション
    OSAKA_PSTOCK_Speck.ks2.keyleng = 2                              ' キー長
    OSAKA_PSTOCK_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' キーフラグ
    OSAKA_PSTOCK_Speck.ks2.keytype = Chr(BtKtString)                ' キータイプ
    OSAKA_PSTOCK_Speck.ks2.reserve = &H0                            ' 予約済み
                                                    
    OSAKA_PSTOCK_Speck.ks3.keypos = 7                               ' キーポジション
    OSAKA_PSTOCK_Speck.ks3.keyleng = 2                              ' キー長
    OSAKA_PSTOCK_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' キーフラグ
    OSAKA_PSTOCK_Speck.ks3.keytype = Chr(BtKtString)                ' キータイプ
    OSAKA_PSTOCK_Speck.ks3.reserve = &H0                            ' 予約済み
                                                    
    OSAKA_PSTOCK_Speck.ks4.keypos = 9                               ' キーポジション
    OSAKA_PSTOCK_Speck.ks4.keyleng = 1                              ' キー長
    OSAKA_PSTOCK_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' キーフラグ
    OSAKA_PSTOCK_Speck.ks4.keytype = Chr(BtKtString)                ' キータイプ
    OSAKA_PSTOCK_Speck.ks4.reserve = &H0                            ' 予約済み
                                                    
    OSAKA_PSTOCK_Speck.ks5.keypos = 10                              ' キーポジション
    OSAKA_PSTOCK_Speck.ks5.keyleng = 1                              ' キー長
    OSAKA_PSTOCK_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' キーフラグ
    OSAKA_PSTOCK_Speck.ks5.keytype = Chr(BtKtString)                ' キータイプ
    OSAKA_PSTOCK_Speck.ks5.reserve = &H0                            ' 予約済み
                                                    
    OSAKA_PSTOCK_Speck.ks6.keypos = 11                              ' キーポジション
    OSAKA_PSTOCK_Speck.ks6.keyleng = 20                             ' キー長
    OSAKA_PSTOCK_Speck.ks6.keyflag = BtKfExt + BtKfChg              ' キーフラグ
    OSAKA_PSTOCK_Speck.ks6.keytype = Chr(BtKtString)                ' キータイプ
    OSAKA_PSTOCK_Speck.ks6.reserve = &H0                            ' 予約済み
                                                    
'---------------------------------------------------'
    sts = BTRV(BtOpCreate, OSAKA_PSTOCK_POS, OSAKA_PSTOCK_Speck, Len(OSAKA_PSTOCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "大阪ＰＣ　循環棚卸Ｆ")
        Exit Function
    End If
    OSAKA_PSTOCK_Create = False
End Function
Public Function OSAKA_PSTOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              大阪ＰＣ　循環棚卸Ｆ　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OSAKA_PSTOCK_Open = True
                                            '大阪ＰＣ　循環棚卸Ｆ　フルパス取込み
    sts = GetIni("FILE", OSAKA_PSTOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [OSAKA_PSTOCK]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = OSAKA_PSTOCK_Create() '大阪ＰＣ　循環棚卸Ｆ　作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "大阪ＰＣ　循環棚卸Ｆ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "大阪ＰＣ　循環棚卸Ｆ")
                Exit Function
        End Select
    Loop
    OSAKA_PSTOCK_Open = False

End Function

