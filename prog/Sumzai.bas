Attribute VB_Name = "SUMZ"
Option Explicit
'********************************************************************
'*                                                                  *
'*              在庫集計データ　ファイル定義                          *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'ファイルＩＤ
Public Const SUMZ_ID$ = "SUMZ"

'ページサイズ
Public Const SUMZ_PG_SIZ% = 2048

'ポジション・ブロック
Public SUMZ_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                              *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type SUMZREC_Tag
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    ST_SOKO(0 To 1)         As Byte     '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)         As Byte     '             列
    ST_REN(0 To 1)          As Byte     '             連
    ST_DAN(0 To 1)          As Byte     '             段
    T_Zai_Qty(0 To 7)       As Byte     '在庫総数(当日)
    ZEN_Zai_Qty(0 To 7)     As Byte     '在庫総数(前日)
    SYK_E_QTY(0 To 7)       As Byte     '出庫済み数
    NYUKA_YQTY(0 To 7)      As Byte     '入荷予定数
    HS_ZAIQTY(0 To 7)       As Byte     'ﾎｽﾄ在庫数(当日)
    ZEN_HS_ZAIQTY(0 To 7)   As Byte     'ﾎｽﾄ在庫数(前日)
    SAI_QTY(0 To 7)         As Byte     '差異数
    SUM_DT(0 To 7)          As Byte     '集計日付
    
    BU_ZAI_QTY(0 To 7)      As Byte     'BU在庫
    PPSC_ZAI_QTY(0 To 7)    As Byte     'PPSC在庫
    
    
    ZEN_SAI_QTY(0 To 7)     As Byte     '前日差異数 2009.02.09
    SAI_YMD(0 To 7)         As Byte     '差異発生日 2009.02.09
    FILLER(0 To 1)          As Byte     'FILLER     2009.02.09
End Type

'データ・バッファ
Public SUMZREC As SUMZREC_Tag

'キー定義
Private Type KEY0_SUMZ            'ＫＥＹ０
    JGYOBU(0 To 0) As Byte          '事業部区分
    NAIGAI(0 To 0) As Byte          '国内外
    HIN_GAI(0 To 19) As Byte        '品番（外部）
End Type

Private Type KEY1_SUMZ            'ＫＥＹ１
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    ST_SOKO(0 To 1)     As Byte     '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)     As Byte     '             列
    ST_REN(0 To 1)      As Byte     '             連
    ST_DAN(0 To 1)      As Byte     '             段
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type

'キー・データ
Public K0_SUMZ As KEY0_SUMZ
Public K1_SUMZ As KEY1_SUMZ

Private Type SUMZ_FSpeck
    fs As BtFileSpeck               ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks7 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks8 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks9 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private SUMZ_Speck As SUMZ_FSpeck

Private Function SUMZ_Create() As Integer
'********************************************************************
'*                                                                  *
'*              在庫集計データ　ＣＲＥＡＴＥ                        *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SUMZ_Create = True
                                            '在庫集計データフルパス取込み
    sts = GetIni("FILE", SUMZ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[SUMZ] 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    SUMZ_Speck.fs.recoleng = Len(SUMZREC)       ' レコード長
    SUMZ_Speck.fs.PageSize = SUMZ_PG_SIZ        ' ページサイズ
    SUMZ_Speck.fs.idexnumb = 2                  ' インデックス数
    SUMZ_Speck.fs.fileflag = 0                  ' ファイルフラグ
    SUMZ_Speck.fs.reserve = &H0                 ' 予約済み
'-----------------------------------------------' キー０
    SUMZ_Speck.ks0.keypos = 1                   ' キーポジション
    SUMZ_Speck.ks0.keyleng = 1                  ' キー長
    SUMZ_Speck.ks0.keyflag = BtKfExt + BtKfSeg  ' キーフラグ
    SUMZ_Speck.ks0.keytype = Chr(BtKtString)    ' キータイプ
    SUMZ_Speck.ks0.reserve = &H0                ' 予約済み

    SUMZ_Speck.ks1.keypos = 2                   ' キーポジション
    SUMZ_Speck.ks1.keyleng = 1                  ' キー長
    SUMZ_Speck.ks1.keyflag = BtKfExt + BtKfSeg  ' キーフラグ
    SUMZ_Speck.ks1.keytype = Chr(BtKtString)    ' キータイプ
    SUMZ_Speck.ks1.reserve = &H0                ' 予約済み

    SUMZ_Speck.ks2.keypos = 3                   ' キーポジション
    SUMZ_Speck.ks2.keyleng = 20                 ' キー長
    SUMZ_Speck.ks2.keyflag = BtKfExt            ' キーフラグ
    SUMZ_Speck.ks2.keytype = Chr(BtKtString)    ' キータイプ
    SUMZ_Speck.ks2.reserve = &H0                ' 予約済み
'-----------------------------------------------' キー１
    SUMZ_Speck.ks3.keypos = 1                   ' キーポジション
    SUMZ_Speck.ks3.keyleng = 1                  ' キー長
                                                ' キーフラグ
    SUMZ_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks3.keytype = Chr(BtKtString)    ' キータイプ
    SUMZ_Speck.ks3.reserve = &H0                ' 予約済み
    
    SUMZ_Speck.ks4.keypos = 2                   ' キーポジション
    SUMZ_Speck.ks4.keyleng = 1                  ' キー長
                                                ' キーフラグ
    SUMZ_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks4.keytype = Chr(BtKtString)    ' キータイプ
    SUMZ_Speck.ks4.reserve = &H0                ' 予約済み
    
    SUMZ_Speck.ks5.keypos = 23                  ' キーポジション
    SUMZ_Speck.ks5.keyleng = 2                  ' キー長
                                                ' キーフラグ
    SUMZ_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks5.keytype = Chr(BtKtString)    ' キータイプ
    SUMZ_Speck.ks5.reserve = &H0                ' 予約済み
    
    SUMZ_Speck.ks6.keypos = 25                  ' キーポジション
    SUMZ_Speck.ks6.keyleng = 2                  ' キー長
                                                ' キーフラグ
    SUMZ_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks6.keytype = Chr(BtKtString)    ' キータイプ
    SUMZ_Speck.ks6.reserve = &H0                ' 予約済み
    
    SUMZ_Speck.ks7.keypos = 27                  ' キーポジション
    SUMZ_Speck.ks7.keyleng = 2                  ' キー長
                                                ' キーフラグ
    SUMZ_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks7.keytype = Chr(BtKtString)    ' キータイプ
    SUMZ_Speck.ks7.reserve = &H0                ' 予約済み
    
    SUMZ_Speck.ks8.keypos = 29                  ' キーポジション
    SUMZ_Speck.ks8.keyleng = 2                  ' キー長
                                                ' キーフラグ
    SUMZ_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfSeg
    SUMZ_Speck.ks8.keytype = Chr(BtKtString)    ' キータイプ
    SUMZ_Speck.ks8.reserve = &H0                ' 予約済み
    
    SUMZ_Speck.ks9.keypos = 3                   ' キーポジション
    SUMZ_Speck.ks9.keyleng = 20                 ' キー長
    SUMZ_Speck.ks9.keyflag = BtKfExt + BtKfChg  ' キーフラグ
    SUMZ_Speck.ks9.keytype = Chr(BtKtString)    ' キータイプ
    SUMZ_Speck.ks9.reserve = &H0                ' 予約済み

    sts = BTRV(BtOpCreate, SUMZ_POS, SUMZ_Speck, Len(SUMZ_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "在庫集計データ")
        Exit Function
    End If
    
    SUMZ_Create = False

End Function

Function SUMZ_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              在庫集計データ　ＯＰＥＮ                            *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    SUMZ_Open = True
                                            '在庫集計データフルパス取込み
    sts = GetIni("FILE", SUMZ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[SUMZ] 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, SUMZ_POS, SUMZREC, Len(SUMZREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SUMZ_Create()        '在庫集計データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SUMZ_POS, SUMZREC, Len(SUMZREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "在庫集計データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "在庫集計データ")
                Exit Function
        End Select
    Loop

    SUMZ_Open = False
End Function


