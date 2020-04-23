Attribute VB_Name = "tmpZAIKO_F110010"
Option Explicit
'********************************************************************
'*
'*              在庫データ（一時データ） ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const tmpZAIKO_ID$ = "tmpZAIKO"

'ページサイズ
Public Const tmpZAIKO_PG_SIZ% = 1024

'ポジション・ブロック
Public tmpZAIKO_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Type tmpZAIKOREC_Tag
    SOKO_NO(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.12.05 13-->20
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    GOODS_ON(0 To 0)    As Byte     '商品化／未商品化
    NYUKA_DT(0 To 7)    As Byte     '入荷日付
    NYUKO_DT(0 To 7)    As Byte     '入庫日付
    '2005.12.05 13-->20
    HIN_NAI(0 To 19)    As Byte     '品番（内部）
    YUKO_Z_QTY(0 To 7)  As Byte     '有効在庫数
    LOCK_F(0 To 0)      As Byte     '排他フラグ
    WEL_ID(0 To 2)      As Byte     '使用子機ID
    PRG_ID(0 To 7)      As Byte     '使用中プログラム
    GOODS_YMD(0 To 7)   As Byte     '商品化日付
    
    '2005.12.05 項目追加
    SHIIRE_CODE(0 To 4) As Byte     '仕入先ｺｰﾄﾞ
    SHIIRE_TANKA(0 To 10) As Byte   '仕入単価(9(8)V99)
    KEIJYO_YM(0 To 5)   As Byte     '計上年月
    '2005.12.05 項目追加
    
    '----------------   2010.07.08 ▽
    GENSANKOKU(0 To 19)         As Byte     '原産国名
    SHIIRE_WORK_CENTER(0 To 7)  As Byte     '資材仕入先ﾜｰｸｾﾝﾀｰ
    ID_NO2(0 To 11)             As Byte     'ID_NO
    YOSAN_FROM(0 To 4)          As Byte     '予算単位（元）
    YOSAN_TO(0 To 4)            As Byte     '予算単位（先）
    '----------------   2010.07.08 △
    
    
    FILLER(0 To 24)     As Byte     'FILLER
End Type

'データ・バッファ
Public tmpZAIKOREC      As tmpZAIKOREC_Tag

'キー定義

Type KEY0_tmpZAIKO                  'ＫＥＹ０
    SOKO_NO(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    NYUKA_DT(0 To 7)    As Byte     '入荷日付
End Type

Type KEY1_tmpZAIKO                  'ＫＥＹ１
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    NYUKA_DT(0 To 7)    As Byte     '入荷日付
    SOKO_NO(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
End Type

'キー・データ
Public K0_tmpZAIKO      As KEY0_tmpZAIKO
Public K1_tmpZAIKO      As KEY1_tmpZAIKO

Type tmpZAIKO_FSpeck
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

Private tmpZAIKO_Speck As tmpZAIKO_FSpeck
Private Function tmpZAIKO_Create() As Integer
'********************************************************************
'*
'*              在庫データ（一時データ）　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

Dim Ret         As Integer

    tmpZAIKO_Create = True
                                            '在庫データ　フルパス取込み
    sts = GetIni("FILE", tmpZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpZAIKO]読み込みエラー")
        Exit Function
    End If
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    FullPath = Left(Trim(c), Ret) & "F110010" & Right(Trim(c), Len(Trim(c)) - Ret)
    
    
 



    tmpZAIKO_Speck.fs.recoleng = Len(tmpZAIKOREC)   ' レコード長
    tmpZAIKO_Speck.fs.PageSize = tmpZAIKO_PG_SIZ    ' ページサイズ
    tmpZAIKO_Speck.fs.idexnumb = 2                  ' インデックス数
    tmpZAIKO_Speck.fs.fileflag = 0                  ' ファイルフラグ
    tmpZAIKO_Speck.fs.reserve = &H0                 ' 予約済み
'---------------------------------------------------' キー０
    tmpZAIKO_Speck.ks0.keypos = 1                   ' キーポジション
    tmpZAIKO_Speck.ks0.keyleng = 2                  ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks0.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks0.keytype = Chr(BtKtString)    ' キータイプ
    tmpZAIKO_Speck.ks0.reserve = &H0                ' 予約済み
                                                    
    tmpZAIKO_Speck.ks1.keypos = 3                   ' キーポジション
    tmpZAIKO_Speck.ks1.keyleng = 2                  ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks1.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks1.keytype = Chr(BtKtString)    ' キータイプ
    tmpZAIKO_Speck.ks1.reserve = &H0                ' 予約済み
                                                    
    tmpZAIKO_Speck.ks2.keypos = 5                   ' キーポジション
    tmpZAIKO_Speck.ks2.keyleng = 2                  ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks2.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks2.keytype = Chr(BtKtString)    ' キータイプ
    tmpZAIKO_Speck.ks2.reserve = &H0                ' 予約済み
                                                    
    tmpZAIKO_Speck.ks3.keypos = 7                   ' キーポジション
    tmpZAIKO_Speck.ks3.keyleng = 2                  ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks3.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks3.keytype = Chr(BtKtString)    ' キータイプ
    tmpZAIKO_Speck.ks3.reserve = &H0                ' 予約済み
                                                    
    tmpZAIKO_Speck.ks4.keypos = 9                   ' キーポジション
    tmpZAIKO_Speck.ks4.keyleng = 1                  ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks4.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks4.keytype = Chr(BtKtString)    ' キータイプ
    tmpZAIKO_Speck.ks4.reserve = &H0                ' 予約済み
                                                    
    tmpZAIKO_Speck.ks5.keypos = 10                  ' キーポジション
    tmpZAIKO_Speck.ks5.keyleng = 1                  ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks5.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks5.keytype = Chr(BtKtString)    ' キータイプ
    tmpZAIKO_Speck.ks5.reserve = &H0                ' 予約済み
                                                    
    tmpZAIKO_Speck.ks6.keypos = 11                  ' キーポジション
    tmpZAIKO_Speck.ks6.keyleng = 20                 ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks6.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks6.keytype = Chr(BtKtString)    ' キータイプ
    tmpZAIKO_Speck.ks6.reserve = &H0                ' 予約済み
    
    tmpZAIKO_Speck.ks7.keypos = 32                  ' キーポジション
    tmpZAIKO_Speck.ks7.keyleng = 8                  ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks7.keyflag = BtKfExt
    tmpZAIKO_Speck.ks7.keytype = Chr(BtKtString)    ' キータイプ
    tmpZAIKO_Speck.ks7.reserve = &H0                ' 予約済み
'---------------------------------------------------' キー１
    tmpZAIKO_Speck.ks8.keypos = 9                   ' キーポジション
    tmpZAIKO_Speck.ks8.keyleng = 1                  ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks8.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks8.keytype = Chr(BtKtString)    ' キータイプ
    tmpZAIKO_Speck.ks8.reserve = &H0                ' 予約済み
                                                    
    tmpZAIKO_Speck.ks9.keypos = 10                 ' キーポジション
    tmpZAIKO_Speck.ks9.keyleng = 1                 ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks9.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks9.keytype = Chr(BtKtString)   ' キータイプ
    tmpZAIKO_Speck.ks9.reserve = &H0               ' 予約済み
                                                    
    tmpZAIKO_Speck.ks10.keypos = 11                 ' キーポジション
    tmpZAIKO_Speck.ks10.keyleng = 20                ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks10.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks10.keytype = Chr(BtKtString)   ' キータイプ
    tmpZAIKO_Speck.ks10.reserve = &H0               ' 予約済み
                                                    
    tmpZAIKO_Speck.ks11.keypos = 32                 ' キーポジション
    tmpZAIKO_Speck.ks11.keyleng = 8                 ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks11.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks11.keytype = Chr(BtKtString)   ' キータイプ
    tmpZAIKO_Speck.ks11.reserve = &H0               ' 予約済み
                                                    
    tmpZAIKO_Speck.ks12.keypos = 1                  ' キーポジション
    tmpZAIKO_Speck.ks12.keyleng = 2                 ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks12.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks12.keytype = Chr(BtKtString)   ' キータイプ
    tmpZAIKO_Speck.ks12.reserve = &H0               ' 予約済み
                                                    
    tmpZAIKO_Speck.ks13.keypos = 3                  ' キーポジション
    tmpZAIKO_Speck.ks13.keyleng = 2                 ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks13.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks13.keytype = Chr(BtKtString)   ' キータイプ
    tmpZAIKO_Speck.ks13.reserve = &H0               ' 予約済み
                                                    
    tmpZAIKO_Speck.ks14.keypos = 5                  ' キーポジション
    tmpZAIKO_Speck.ks14.keyleng = 2                 ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks14.keyflag = BtKfExt + BtKfSeg
    tmpZAIKO_Speck.ks14.keytype = Chr(BtKtString)   ' キータイプ
    tmpZAIKO_Speck.ks14.reserve = &H0               ' 予約済み
                                                    
    tmpZAIKO_Speck.ks15.keypos = 7                  ' キーポジション
    tmpZAIKO_Speck.ks15.keyleng = 2                 ' キー長
                                                    ' キーフラグ
    tmpZAIKO_Speck.ks15.keyflag = BtKfExt
    tmpZAIKO_Speck.ks15.keytype = Chr(BtKtString)   ' キータイプ
    tmpZAIKO_Speck.ks15.reserve = &H0               ' 予約済み
'---------------------------------------------------'
    sts = BTRV(BtOpCreate, tmpZAIKO_POS, tmpZAIKO_Speck, Len(tmpZAIKO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "在庫データ（一時データ）")
        Exit Function
    End If
    tmpZAIKO_Create = False
End Function
Public Function tmpZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              在庫データ（一時データ）　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
Dim Ret         As Integer
    
    tmpZAIKO_Open = True
                                            '在庫データ　フルパス取込み
    sts = GetIni("FILE", tmpZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpZAIKO]読み込みエラー")
        Exit Function
    End If
    
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    FullPath = Left(Trim(c), Ret) & "F110010" & Right(Trim(c), Len(Trim(c)) - Ret)
    
    
    Do
        sts = BTRV(BtOpOpen, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpZAIKO_Create()        '在庫データ　作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "在庫データ（一時データ）")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "在庫データ（一時データ）")
                Exit Function
        End Select
    Loop
    tmpZAIKO_Open = False

End Function

