Attribute VB_Name = "Y_SYU_SUM"
Option Explicit
'********************************************************************
'*
'*              出荷予定（大阪PC出庫表用）データ  ファイル定義
'*              大阪ＰＣ専用    2007.03.14
'*
'********************************************************************
'ファイルＩＤ
Public Const Y_SYU_SUM_ID$ = "Y_SYU_SUM"

'ページサイズ
Public Const Y_SYU_SUM_PG_SIZ% = 4096

'ポジション・ブロック
Public Y_SYU_SUM_POS    As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type Y_SYU_SUMREC_Tag
    SYUKA_YMD(0 To 7)           As Byte     '出荷日付
    INS_BIN(0 To 1)             As Byte     '便
    ST_SOKO(0 To 1)             As Byte     '標準棚番     倉庫
    ST_RETU(0 To 1)             As Byte     '             列
    ST_REN(0 To 1)              As Byte     '             連
    ST_DAN(0 To 1)              As Byte     '             段
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_NO(0 To 19)             As Byte     '品目番号
    
    Y_SURYO(0 To 6)             As Byte     '出荷予定数量
    J_SURYO(0 To 6)             As Byte     '出荷実績数量
    
    SYU_NO(0 To 11)             As Byte     '出庫表№
    DATA_CNT(0 To 3)            As Byte     '件数
    
    ST_ZAIKO_QTY(0 To 7)        As Byte     '標準棚番在庫数
    
    BETU_SOKO(0 To 1)           As Byte     '別置棚番     倉庫
    BETU_RETU(0 To 1)           As Byte     '             列
    BETU_REN(0 To 1)            As Byte     '             連
    BETU_DAN(0 To 1)            As Byte     '             段
    
    BETU_ZAIKO_QTY(0 To 7)      As Byte     '別置在庫数
    
    SYO_ZAIKO_QTY(0 To 7)       As Byte     '商品化室在庫数
    NYU_ZAIKO_QTY(0 To 7)       As Byte     '入荷倉庫在庫数
    
    INS_NOW(0 To 13)            As Byte     'ﾃﾞｰﾀ作成日時
    
    FILLER(0 To 39)             As Byte     'FILLER
End Type

'データ・バッファ
Public Y_SYU_SUMREC             As Y_SYU_SUMREC_Tag

'キー定義
Type KEY0_Y_SYU_SUM         'ＫＥＹ０
    SYUKA_YMD(0 To 7)           As Byte     '出荷日付
    INS_BIN(0 To 1)             As Byte     '便
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_NO(0 To 19)             As Byte     '品目番号
End Type

Type KEY1_Y_SYU_SUM         'ＫＥＹ１
    SYUKA_YMD(0 To 7)           As Byte     '出荷日付
    INS_BIN(0 To 1)             As Byte     '便
    ST_SOKO(0 To 1)             As Byte     '標準棚番     倉庫
    ST_RETU(0 To 1)             As Byte     '             列
    ST_REN(0 To 1)              As Byte     '             連
    ST_DAN(0 To 1)              As Byte     '             段
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_NO(0 To 19)             As Byte     '品目番号
End Type

Type KEY2_Y_SYU_SUM         'ＫＥＹ２
    SYUKA_YMD(0 To 7)           As Byte     '出荷日付
    INS_BIN(0 To 1)             As Byte     '便
End Type

Type KEY3_Y_SYU_SUM         'ＫＥＹ３
    INS_BIN(0 To 1)             As Byte     '便
    SYU_NO(0 To 11)              As Byte     '出庫表№
End Type


'キー・データ
Public K0_Y_SYU_SUM             As KEY0_Y_SYU_SUM
Public K1_Y_SYU_SUM             As KEY1_Y_SYU_SUM
Public K2_Y_SYU_SUM             As KEY2_Y_SYU_SUM
Public K3_Y_SYU_SUM             As KEY3_Y_SYU_SUM

Type Y_SYU_SUM_FSpeck
    fs      As BtFileSpeck                  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    
    ks5     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks7     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks8     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks9     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks10    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks11    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks12    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks13    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    
    ks14    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks15    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    
    ks16    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks17    As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体



End Type

Private Y_SYU_SUM_Speck     As Y_SYU_SUM_FSpeck

Private Function Y_SYU_SUM_Create(Optional F_NAME As String = " ") As Integer
'********************************************************************
'*
'*              出荷予定(大阪PC出庫表用)データ  ＣＲＥＡＴＥ
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

    Y_SYU_SUM_Create = True
                                            '出荷予定データフルパス取込み
    sts = GetIni("FILE", Y_SYU_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [Y_SYU_SUM]読み込みエラー")
        Exit Function
    End If
    
    If Trim(F_NAME) = "" Then
        FullPath = RTrim(c)
    Else
        Ret = InStr(1, Trim(c), ".") - 1
        FullPath = Left(Trim(c), Ret) & Trim(F_NAME) & Right(Trim(c), Len(Trim(c)) - Ret)
    End If

    Y_SYU_SUM_Speck.fs.recoleng = Len(Y_SYU_SUMREC)     ' レコード長
    Y_SYU_SUM_Speck.fs.PageSize = Y_SYU_SUM_PG_SIZ      ' ページサイズ
    Y_SYU_SUM_Speck.fs.idexnumb = 4                     ' インデックス数
    Y_SYU_SUM_Speck.fs.fileflag = 0                     ' ファイルフラグ
    Y_SYU_SUM_Speck.fs.reserve = &H0                    ' 予約済み
'---------------------------------------------------' キー０
    Y_SYU_SUM_Speck.ks0.keypos = 1                      ' キーポジション
    Y_SYU_SUM_Speck.ks0.keyleng = 8                     ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks0.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_SUM_Speck.ks0.reserve = &H0                   ' 予約済み
    
    Y_SYU_SUM_Speck.ks1.keypos = 9                      ' キーポジション
    Y_SYU_SUM_Speck.ks1.keyleng = 2                     ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks1.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_SUM_Speck.ks1.reserve = &H0                   ' 予約済み
    
    Y_SYU_SUM_Speck.ks2.keypos = 19                     ' キーポジション
    Y_SYU_SUM_Speck.ks2.keyleng = 1                     ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks2.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_SUM_Speck.ks2.reserve = &H0                   ' 予約済み
    
    Y_SYU_SUM_Speck.ks3.keypos = 20                     ' キーポジション
    Y_SYU_SUM_Speck.ks3.keyleng = 1                     ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks3.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_SUM_Speck.ks3.reserve = &H0                   ' 予約済み
    
    Y_SYU_SUM_Speck.ks4.keypos = 21                     ' キーポジション
    Y_SYU_SUM_Speck.ks4.keyleng = 20                    ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks4.keyflag = BtKfExt
    Y_SYU_SUM_Speck.ks4.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_SUM_Speck.ks4.reserve = &H0                   ' 予約済み
'---------------------------------------------------' キー０
    
'---------------------------------------------------' キー１
    Y_SYU_SUM_Speck.ks5.keypos = 1                      ' キーポジション
    Y_SYU_SUM_Speck.ks5.keyleng = 8                     ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks5.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks5.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_SUM_Speck.ks5.reserve = &H0                   ' 予約済み
    
    Y_SYU_SUM_Speck.ks6.keypos = 9                      ' キーポジション
    Y_SYU_SUM_Speck.ks6.keyleng = 2                     ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks6.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks6.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_SUM_Speck.ks6.reserve = &H0                   ' 予約済み
    
    Y_SYU_SUM_Speck.ks7.keypos = 11                     ' キーポジション
    Y_SYU_SUM_Speck.ks7.keyleng = 2                     ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks7.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks7.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_SUM_Speck.ks7.reserve = &H0                   ' 予約済み
    
    Y_SYU_SUM_Speck.ks8.keypos = 13                     ' キーポジション
    Y_SYU_SUM_Speck.ks8.keyleng = 2                     ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks8.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks8.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_SUM_Speck.ks8.reserve = &H0                   ' 予約済み
    
    Y_SYU_SUM_Speck.ks9.keypos = 15                     ' キーポジション
    Y_SYU_SUM_Speck.ks9.keyleng = 2                     ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks9.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks9.keytype = Chr(BtKtString)       ' キータイプ
    Y_SYU_SUM_Speck.ks9.reserve = &H0                   ' 予約済み
    
    Y_SYU_SUM_Speck.ks10.keypos = 17                    ' キーポジション
    Y_SYU_SUM_Speck.ks10.keyleng = 2                    ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks10.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks10.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_SUM_Speck.ks10.reserve = &H0                  ' 予約済み
    
    Y_SYU_SUM_Speck.ks11.keypos = 19                    ' キーポジション
    Y_SYU_SUM_Speck.ks11.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks11.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks11.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_SUM_Speck.ks11.reserve = &H0                  ' 予約済み
    
    Y_SYU_SUM_Speck.ks12.keypos = 20                    ' キーポジション
    Y_SYU_SUM_Speck.ks12.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks12.keyflag = BtKfExt + BtKfSeg
    Y_SYU_SUM_Speck.ks12.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_SUM_Speck.ks12.reserve = &H0                  ' 予約済み
    
    Y_SYU_SUM_Speck.ks13.keypos = 21                    ' キーポジション
    Y_SYU_SUM_Speck.ks13.keyleng = 20                   ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks13.keyflag = BtKfExt
    Y_SYU_SUM_Speck.ks13.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_SUM_Speck.ks13.reserve = &H0                  ' 予約済み
'---------------------------------------------------' キー１
    
'---------------------------------------------------' キー２
    Y_SYU_SUM_Speck.ks14.keypos = 1                     ' キーポジション
    Y_SYU_SUM_Speck.ks14.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks14.keyflag = BtKfExt + BtKfDup + BtKfSeg
    Y_SYU_SUM_Speck.ks14.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_SUM_Speck.ks14.reserve = &H0                  ' 予約済み
    
    Y_SYU_SUM_Speck.ks15.keypos = 9                     ' キーポジション
    Y_SYU_SUM_Speck.ks15.keyleng = 2                    ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks15.keyflag = BtKfExt + BtKfDup
    Y_SYU_SUM_Speck.ks15.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_SUM_Speck.ks15.reserve = &H0                  ' 予約済み
'---------------------------------------------------' キー２
'---------------------------------------------------' キー３
    Y_SYU_SUM_Speck.ks16.keypos = 9                     ' キーポジション
    Y_SYU_SUM_Speck.ks16.keyleng = 2                    ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks16.keyflag = BtKfExt + BtKfDup + BtKfSeg
    Y_SYU_SUM_Speck.ks16.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_SUM_Speck.ks16.reserve = &H0                  ' 予約済み

    Y_SYU_SUM_Speck.ks17.keypos = 55                    ' キーポジション
    Y_SYU_SUM_Speck.ks17.keyleng = 12                   ' キー長
                                                        ' キーフラグ
    Y_SYU_SUM_Speck.ks17.keyflag = BtKfExt + BtKfDup
    Y_SYU_SUM_Speck.ks17.keytype = Chr(BtKtString)      ' キータイプ
    Y_SYU_SUM_Speck.ks17.reserve = &H0                  ' 予約済み

'---------------------------------------------------' キー３
    sts = BTRV(BtOpCreate, Y_SYU_SUM_POS, Y_SYU_SUM_Speck, Len(Y_SYU_SUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "出荷予定(大阪PC出庫表用)データ")
        Exit Function
    End If

    Y_SYU_SUM_Create = False

End Function

Function Y_SYU_SUM_Open(Mode As Integer, Optional F_NAME As String = " ") As Integer
'********************************************************************
'*
'*              出荷予定(大阪PC出庫表用)データ  ＯＰＥＮ
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
    
    Y_SYU_SUM_Open = True
                                            '出荷予定データフルパス取込み
    sts = GetIni("FILE", Y_SYU_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [Y_SYU_SUM]読み込みエラー ")
        Exit Function
    End If
    
    If Trim(F_NAME) = "" Then
        FullPath = RTrim(c)
    Else
        Ret = InStr(1, Trim(c), ".") - 1
        FullPath = Left(Trim(c), Ret) & Trim(F_NAME) & Right(Trim(c), Len(Trim(c)) - Ret)
    End If
    
''    On Error Resume Next
''    Kill (FullPath)
''    On Error GoTo 0
    
    Do
        sts = BTRV(BtOpOpen, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Y_SYU_SUM_Create(F_NAME)      '出荷予定データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "出荷予定(大阪PC出庫表用)データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "出荷予定(大阪PC出庫表用)データ")
                Exit Function
        End Select
    Loop
    Y_SYU_SUM_Open = False
End Function
