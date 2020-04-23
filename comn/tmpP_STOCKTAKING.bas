Attribute VB_Name = "tmpP_STOCKTAKING"
Option Explicit

'********************************************************************
'*
'*              資材棚卸しﾃﾞｰﾀ  ファイル定義
'*
'*          CREATE 2006.11.22
'********************************************************************
'ファイルＩＤ
Public Const tmpP_STOCK_ID$ = "tmpP_STOCK"

'ページサイズ
Private Const tmpP_STOCK_PG_SIZ% = 1024

'ポジション・ブロック
Public tmpP_STOCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type tmpP_STOCK_REC_Tag
    
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    
    
    CODE(0 To 4)            As Byte         '仕入先ｺｰﾄﾞ
    TANKA(0 To 10)          As Byte         '仕入単価 9(8)V99
    
    INPUT_DATE(0 To 7)      As Byte         '登録日付
    
    G_SYUSHI(0 To 2)        As Byte         '収支単位
    ZEN_ZAIKO_QTY(0 To 7)   As Byte         '前月在庫数量
                            
    NYUKO_QTY(0 To 7)       As Byte         '入庫数
    SYUKO_QTY(0 To 7)       As Byte         '出庫数
    ZAIKO_QTY(0 To 7)       As Byte         '在庫数
    
    
    LAST_SYUKA_DT(0 To 7)   As Byte         '最終出荷日
    LAST_SYUKA_QTY(0 To 7)  As Byte         '最終出荷数量
    
    MOTO_ZAIKO_QTY(0 To 7)  As Byte         '再集計前
    MAEGARI_QTY(0 To 7)     As Byte         '前借数

    
    SYUKA_NON_F(0 To 0)     As Byte         '出荷数計算有無　0:しない　1:する


    ZEN_ZAIKO_KIN(0 To 7)   As Byte         '前月在庫金額

    FILLER(0 To 5)         As Byte          '

End Type
'データ・バッファ
Public tmpP_STOCK_REC       As tmpP_STOCK_REC_Tag

'キー定義
    
Public Type KEY0_tmpP_STOCK                 'ＫＥＹ０
    
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    CODE(0 To 4)            As Byte         '仕入先ｺｰﾄﾞ
    TANKA(0 To 10)          As Byte         '仕入単価 9(8)V99
    
End Type
    
Public Type KEY1_tmpP_STOCK                 'ＫＥＹ１
    
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    
    INPUT_DATE(0 To 7)      As Byte         '登録日付 2006.11.22
    
    
    CODE(0 To 4)            As Byte         '仕入先ｺｰﾄﾞ
    TANKA(0 To 10)          As Byte         '仕入単価 9(8)V99
    
End Type
    
    
Public Type KEY2_tmpP_STOCK                 'ＫＥＹ１
    
    
    G_SYUSHI(0 To 2)        As Byte         '収支単位
    
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    
    INPUT_DATE(0 To 7)      As Byte         '登録日付 2006.11.22
    
    
    CODE(0 To 4)            As Byte         '仕入先ｺｰﾄﾞ
    TANKA(0 To 10)          As Byte         '仕入単価 9(8)V99
    
End Type
    
    
    
'キー・データ
Public K0_tmpP_STOCK        As KEY0_tmpP_STOCK
Public K1_tmpP_STOCK        As KEY1_tmpP_STOCK

Public K2_tmpP_STOCK        As KEY2_tmpP_STOCK


Type tmpP_STOCK_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks7                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks8                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks9                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks10                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

    ks11                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks12                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks13                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks14                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks15                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks16                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks17                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

End Type

Private tmpP_STOCK_Speck    As tmpP_STOCK_FSpeck
Private Function tmpP_STOCK_Create() As Integer
'********************************************************************
'*
'*              資材棚卸しﾃﾞｰﾀ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128





    tmpP_STOCK_Create = True
                                            '資材棚卸しﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", tmpP_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpP_STOCK]読み込みエラー")
        Exit Function
    End If



    FullPath = Trim(c)
    tmpP_STOCK_Speck.fs.recoleng = Len(tmpP_STOCK_REC)  ' レコード長
    tmpP_STOCK_Speck.fs.PageSize = tmpP_STOCK_PG_SIZ    ' ページサイズ
    tmpP_STOCK_Speck.fs.idexnumb = 3                    ' インデックス数
    tmpP_STOCK_Speck.fs.fileflag = 0                    ' ファイルフラグ
    tmpP_STOCK_Speck.fs.reserve = &H0                   ' 予約済み
    
    '--------------------------------------------------- キー０ ▽
    tmpP_STOCK_Speck.ks0.keypos = 1                     ' キーポジション
    tmpP_STOCK_Speck.ks0.keyleng = 1                    ' キー長
    tmpP_STOCK_Speck.ks0.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    tmpP_STOCK_Speck.ks0.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks0.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks1.keypos = 2                     ' キーポジション
    tmpP_STOCK_Speck.ks1.keyleng = 1                    ' キー長
    tmpP_STOCK_Speck.ks1.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    tmpP_STOCK_Speck.ks1.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks1.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks2.keypos = 3                     ' キーポジション
    tmpP_STOCK_Speck.ks2.keyleng = 20                   ' キー長
    tmpP_STOCK_Speck.ks2.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    tmpP_STOCK_Speck.ks2.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks2.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks3.keypos = 23                    ' キーポジション
    tmpP_STOCK_Speck.ks3.keyleng = 5                    ' キー長
    tmpP_STOCK_Speck.ks3.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    tmpP_STOCK_Speck.ks3.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks3.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks4.keypos = 28                    ' キーポジション
    tmpP_STOCK_Speck.ks4.keyleng = 11                   ' キー長
    tmpP_STOCK_Speck.ks4.keyflag = BtKfExt              ' キーフラグ
    tmpP_STOCK_Speck.ks4.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks4.reserve = &H0                  ' 予約済み
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    tmpP_STOCK_Speck.ks5.keypos = 1                     ' キーポジション
    tmpP_STOCK_Speck.ks5.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    tmpP_STOCK_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    tmpP_STOCK_Speck.ks5.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks5.reserve = &H0                  ' 予約済み
    
    
    tmpP_STOCK_Speck.ks6.keypos = 2                     ' キーポジション
    tmpP_STOCK_Speck.ks6.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    tmpP_STOCK_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    tmpP_STOCK_Speck.ks6.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks6.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks7.keypos = 3                     ' キーポジション
    tmpP_STOCK_Speck.ks7.keyleng = 20                   ' キー長
                                                        ' キーフラグ
    tmpP_STOCK_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    tmpP_STOCK_Speck.ks7.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks7.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks8.keypos = 39                    ' キーポジション
    tmpP_STOCK_Speck.ks8.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    tmpP_STOCK_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfChg
    tmpP_STOCK_Speck.ks8.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks8.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks9.keypos = 23                    ' キーポジション
    tmpP_STOCK_Speck.ks9.keyleng = 5                    ' キー長
                                                        ' キーフラグ
    tmpP_STOCK_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfChg
    tmpP_STOCK_Speck.ks9.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks9.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks10.keypos = 28                   ' キーポジション
    tmpP_STOCK_Speck.ks10.keyleng = 11                  ' キー長
    tmpP_STOCK_Speck.ks10.keyflag = BtKfExt + BtKfChg   ' キーフラグ
    tmpP_STOCK_Speck.ks10.keytype = Chr(BtKtString)     ' キータイプ
    tmpP_STOCK_Speck.ks10.reserve = &H0                 ' 予約済み
    
    
    '--------------------------------------------------- キー２ ▽
    
    
    tmpP_STOCK_Speck.ks11.keypos = 47                    ' キーポジション
    tmpP_STOCK_Speck.ks11.keyleng = 3                    ' キー長
    tmpP_STOCK_Speck.ks11.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    tmpP_STOCK_Speck.ks11.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks11.reserve = &H0                  ' 予約済み
    
    
    
    tmpP_STOCK_Speck.ks12.keypos = 1                     ' キーポジション
    tmpP_STOCK_Speck.ks12.keyleng = 1                    ' キー長
    tmpP_STOCK_Speck.ks12.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    tmpP_STOCK_Speck.ks12.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks12.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks13.keypos = 2                     ' キーポジション
    tmpP_STOCK_Speck.ks13.keyleng = 1                    ' キー長
    tmpP_STOCK_Speck.ks13.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    tmpP_STOCK_Speck.ks13.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks13.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks14.keypos = 3                     ' キーポジション
    tmpP_STOCK_Speck.ks14.keyleng = 20                   ' キー長
    tmpP_STOCK_Speck.ks14.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    tmpP_STOCK_Speck.ks14.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks14.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks15.keypos = 39                    ' キーポジション
    tmpP_STOCK_Speck.ks15.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    tmpP_STOCK_Speck.ks15.keyflag = BtKfExt + BtKfSeg
    tmpP_STOCK_Speck.ks15.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks15.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks16.keypos = 23                    ' キーポジション
    tmpP_STOCK_Speck.ks16.keyleng = 5                    ' キー長
    tmpP_STOCK_Speck.ks16.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    tmpP_STOCK_Speck.ks16.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks16.reserve = &H0                  ' 予約済み
    
    tmpP_STOCK_Speck.ks17.keypos = 28                    ' キーポジション
    tmpP_STOCK_Speck.ks17.keyleng = 11                   ' キー長
    tmpP_STOCK_Speck.ks17.keyflag = BtKfExt              ' キーフラグ
    tmpP_STOCK_Speck.ks17.keytype = Chr(BtKtString)      ' キータイプ
    tmpP_STOCK_Speck.ks17.reserve = &H0                  ' 予約済み
    '--------------------------------------------------- キー２ △
    
    
    sts = BTRV(BtOpCreate, tmpP_STOCK_POS, tmpP_STOCK_Speck, Len(tmpP_STOCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "tmp資材棚卸しﾃﾞｰﾀ")
        Exit Function
    End If
    
    tmpP_STOCK_Create = False

End Function

Public Function tmpP_STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              資材棚卸しﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String



    tmpP_STOCK_Open = True
                                            '資材棚卸データフルパス取込み
    sts = GetIni("FILE", tmpP_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpP_STOCK]読み込みエラー")
        Exit Function
    End If
    FullPath = Trim(c)

    Do
        sts = BTRV(BtOpOpen, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpP_STOCK_Create()   '資材棚卸しﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "tmp資材棚卸しﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "tmp資材棚卸しﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    tmpP_STOCK_Open = False

End Function

