Attribute VB_Name = "P_STOCKTAKING"
Option Explicit

'********************************************************************
'*
'*              資材棚卸しﾃﾞｰﾀ  ファイル定義
'*
'*          CREATE 2006.02.15
'********************************************************************
'ファイルＩＤ
Public Const P_STOCK_ID$ = "P_STOCK"

'ページサイズ
Private Const P_STOCK_PG_SIZ% = 1024

'ポジション・ブロック
Public P_STOCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type P_STOCK_REC_Tag
    
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    
    
    CODE(0 To 4)            As Byte         '仕入先ｺｰﾄﾞ
    TANKA(0 To 10)          As Byte         '仕入単価 9(8)V99
    
    INPUT_DATE(0 To 7)      As Byte         '登録日付 2006.11.22
    
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
    
    
    
    ZEN_ZAIKO_KIN(0 To 7)   As Byte         '前月在庫数量
    
    
    FILLER(0 To 5)         As Byte          '


End Type
'データ・バッファ
Public P_STOCK_REC          As P_STOCK_REC_Tag

'キー定義
    
Public Type KEY0_P_STOCK                    'ＫＥＹ０
    
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    CODE(0 To 4)            As Byte         '仕入先ｺｰﾄﾞ
    TANKA(0 To 10)          As Byte         '仕入単価 9(8)V99
    
End Type
    
Public Type KEY1_P_STOCK                    'ＫＥＹ１　2006.11.22
    
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    
    INPUT_DATE(0 To 7)      As Byte         '登録日付 2006.11.22
    
    
    CODE(0 To 4)            As Byte         '仕入先ｺｰﾄﾞ
    TANKA(0 To 10)          As Byte         '仕入単価 9(8)V99
    
End Type
    
    
    
    
    
'キー・データ
Public K0_P_STOCK           As KEY0_P_STOCK
Public K1_P_STOCK           As KEY1_P_STOCK


Type P_STOCK_FSpeck
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

End Type

Private P_STOCK_Speck       As P_STOCK_FSpeck
Private Function P_STOCK_Create() As Integer
'********************************************************************
'*
'*              資材棚卸しﾃﾞｰﾀ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'*      収支毎にファイル名を分ける  2007.11.13
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim ret             As Long     '2007.11.13




    P_STOCK_Create = True
                                            '資材棚卸しﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_STOCK]読み込みエラー")
        Exit Function
    End If


    '2007.11.13
'    FullPath = Trim(c)
    ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - ret)
    '2007.11.13



    P_STOCK_Speck.fs.recoleng = Len(P_STOCK_REC)        ' レコード長
    P_STOCK_Speck.fs.PageSize = P_STOCK_PG_SIZ          ' ページサイズ
    P_STOCK_Speck.fs.idexnumb = 2                       ' インデックス数
    P_STOCK_Speck.fs.fileflag = 0                       ' ファイルフラグ
    P_STOCK_Speck.fs.reserve = &H0                      ' 予約済み
    
    '--------------------------------------------------- キー０ ▽
    P_STOCK_Speck.ks0.keypos = 1                        ' キーポジション
    P_STOCK_Speck.ks0.keyleng = 1                       ' キー長
    P_STOCK_Speck.ks0.keyflag = BtKfExt + BtKfSeg + BtKfChg      ' キーフラグ
    P_STOCK_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCK_Speck.ks0.reserve = &H0                     ' 予約済み
    
    P_STOCK_Speck.ks1.keypos = 2                        ' キーポジション
    P_STOCK_Speck.ks1.keyleng = 1                       ' キー長
    P_STOCK_Speck.ks1.keyflag = BtKfExt + BtKfSeg + BtKfChg      ' キーフラグ
    P_STOCK_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCK_Speck.ks1.reserve = &H0                     ' 予約済み
    
    P_STOCK_Speck.ks2.keypos = 3                        ' キーポジション
    P_STOCK_Speck.ks2.keyleng = 20                      ' キー長
    P_STOCK_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfChg      ' キーフラグ
    P_STOCK_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCK_Speck.ks2.reserve = &H0                     ' 予約済み
    
    P_STOCK_Speck.ks3.keypos = 23                       ' キーポジション
    P_STOCK_Speck.ks3.keyleng = 5                       ' キー長
    P_STOCK_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg      ' キーフラグ
    P_STOCK_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCK_Speck.ks3.reserve = &H0                     ' 予約済み
    
    P_STOCK_Speck.ks4.keypos = 28                       ' キーポジション
    P_STOCK_Speck.ks4.keyleng = 11                      ' キー長
    P_STOCK_Speck.ks4.keyflag = BtKfExt + BtKfChg                ' キーフラグ
    P_STOCK_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCK_Speck.ks4.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    P_STOCK_Speck.ks5.keypos = 1                        ' キーポジション
    P_STOCK_Speck.ks5.keyleng = 1                       ' キー長
                                                        ' キーフラグ
    P_STOCK_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    P_STOCK_Speck.ks5.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCK_Speck.ks5.reserve = &H0                     ' 予約済み
    
    
    P_STOCK_Speck.ks6.keypos = 2                        ' キーポジション
    P_STOCK_Speck.ks6.keyleng = 1                       ' キー長
                                                        ' キーフラグ
    P_STOCK_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    P_STOCK_Speck.ks6.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCK_Speck.ks6.reserve = &H0                     ' 予約済み
    
    P_STOCK_Speck.ks7.keypos = 3                        ' キーポジション
    P_STOCK_Speck.ks7.keyleng = 20                      ' キー長
                                                        ' キーフラグ
    P_STOCK_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    P_STOCK_Speck.ks7.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCK_Speck.ks7.reserve = &H0                     ' 予約済み
    
    P_STOCK_Speck.ks8.keypos = 39                       ' キーポジション
    P_STOCK_Speck.ks8.keyleng = 8                       ' キー長
                                                        ' キーフラグ
    P_STOCK_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfChg
    P_STOCK_Speck.ks8.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCK_Speck.ks8.reserve = &H0                     ' 予約済み
    
    P_STOCK_Speck.ks9.keypos = 23                       ' キーポジション
    P_STOCK_Speck.ks9.keyleng = 5                       ' キー長
                                                        ' キーフラグ
    P_STOCK_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfChg
    P_STOCK_Speck.ks9.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCK_Speck.ks9.reserve = &H0                     ' 予約済み
    
    P_STOCK_Speck.ks10.keypos = 28                      ' キーポジション
    P_STOCK_Speck.ks10.keyleng = 11                     ' キー長
    P_STOCK_Speck.ks10.keyflag = BtKfExt + BtKfChg      ' キーフラグ
    P_STOCK_Speck.ks10.keytype = Chr(BtKtString)        ' キータイプ
    P_STOCK_Speck.ks10.reserve = &H0                     ' 予約済み
    
    
    
    
    
    
    
    
    sts = BTRV(BtOpCreate, P_STOCK_POS, P_STOCK_Speck, Len(P_STOCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "資材棚卸しﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_STOCK_Create = False

End Function

Public Function P_STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              資材棚卸しﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*      収支毎にファイル名を分ける  2007.11.13
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim ret             As Long     '2007.11.13


    P_STOCK_Open = True
                                            '資材棚卸データフルパス取込み
    sts = GetIni("FILE", P_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_STOCK]読み込みエラー")
        Exit Function
    End If

    '2007.11.13
'    FullPath = Trim(c)
    ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - ret)
    '2007.11.13


    Do
        sts = BTRV(BtOpOpen, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_STOCK_Create()   '資材棚卸しﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "資材棚卸しﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "資材棚卸しﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_STOCK_Open = False

End Function

