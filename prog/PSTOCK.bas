Attribute VB_Name = "PSTOCK"
Option Explicit
'********************************************************************
'*                                                                  *
'*              範囲内移動分在庫一覧　ファイル定義                  *
'*                                                                  *
'*          CREATE 2004.04.27                                       *
'********************************************************************
'ファイルＩＤ
Public Const PSTOCK_ID = "PSTOCK"

'ページサイズ
Public Const PSTOCK_PG_SIZ% = 512

'ポジション・ブロック
Public PSTOCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Private Type PSTOCKREC_Tag
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    ST_Location(0 To 7) As Byte     '標準棚番
    T_Zai_Qty(0 To 7)   As Byte     '在庫総数
    HS_ZAIQTY(0 To 7)   As Byte     'ﾎｽﾄ在庫数
    Plus_QTY(0 To 7)    As Byte     '在庫＋
    Minus_QTY(0 To 7)   As Byte     '在庫－
End Type

'データ・バッファ
Public PSTOCKREC        As PSTOCKREC_Tag

'キー定義
Private Type KEY0_PSTOCK            'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type

Private Type KEY1_PSTOCK            'ＫＥＹ１
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    ST_Location(0 To 7) As Byte     '標準棚番
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type

'キー・データ
Public K0_PSTOCK        As KEY0_PSTOCK
Public K1_PSTOCK        As KEY1_PSTOCK

Private Type PSTOCK_FSpeck
    fs As BtFileSpeck               ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Public PSTOCK_Speck As PSTOCK_FSpeck

Private Function PSTOCK_Create() As Integer
'********************************************************************
'*                                                                  *
'*              範囲内移動分在庫一覧データ　ＣＲＥＡＴＥ            *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.04.27                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PSTOCK_Create = True
                                            '範囲内移動分在庫一覧データフルパス取込み
    sts = GetIni("FILE", PSTOCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[PSTOCK] 読み込みエラー ")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    PSTOCK_Speck.fs.recoleng = Len(PSTOCKREC)       ' レコード長
    PSTOCK_Speck.fs.PageSize = PSTOCK_PG_SIZ        ' ページサイズ
    PSTOCK_Speck.fs.idexnumb = 2                    ' インデックス数
    PSTOCK_Speck.fs.fileflag = 0                    ' ファイルフラグ
    PSTOCK_Speck.fs.reserve = &H0                   ' 予約済み
'-----------------------------------------------    ' キー０
    PSTOCK_Speck.ks0.keypos = 1                     ' キーポジション
    PSTOCK_Speck.ks0.keyleng = 1                    ' キー長
    PSTOCK_Speck.ks0.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    PSTOCK_Speck.ks0.keytype = Chr(BtKtString)      ' キータイプ
    PSTOCK_Speck.ks0.reserve = &H0                  ' 予約済み

    PSTOCK_Speck.ks1.keypos = 2                     ' キーポジション
    PSTOCK_Speck.ks1.keyleng = 1                    ' キー長
    PSTOCK_Speck.ks1.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    PSTOCK_Speck.ks1.keytype = Chr(BtKtString)      ' キータイプ
    PSTOCK_Speck.ks1.reserve = &H0                  ' 予約済み

    PSTOCK_Speck.ks2.keypos = 3                     ' キーポジション
    PSTOCK_Speck.ks2.keyleng = 20                   ' キー長
    PSTOCK_Speck.ks2.keyflag = BtKfExt              ' キーフラグ
    PSTOCK_Speck.ks2.keytype = Chr(BtKtString)      ' キータイプ
    PSTOCK_Speck.ks2.reserve = &H0                  ' 予約済み
'-----------------------------------------------    ' キー１
    PSTOCK_Speck.ks3.keypos = 1                     ' キーポジション
    PSTOCK_Speck.ks3.keyleng = 1                    ' キー長
    PSTOCK_Speck.ks3.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    PSTOCK_Speck.ks3.keytype = Chr(BtKtString)      ' キータイプ
    PSTOCK_Speck.ks3.reserve = &H0                  ' 予約済み

    PSTOCK_Speck.ks4.keypos = 2                     ' キーポジション
    PSTOCK_Speck.ks4.keyleng = 1                    ' キー長
    PSTOCK_Speck.ks4.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    PSTOCK_Speck.ks4.keytype = Chr(BtKtString)      ' キータイプ
    PSTOCK_Speck.ks4.reserve = &H0                  ' 予約済み

    PSTOCK_Speck.ks5.keypos = 23                    ' キーポジション
    PSTOCK_Speck.ks5.keyleng = 8                    ' キー長
    PSTOCK_Speck.ks5.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    PSTOCK_Speck.ks5.keytype = Chr(BtKtString)      ' キータイプ
    PSTOCK_Speck.ks5.reserve = &H0                  ' 予約済み

    PSTOCK_Speck.ks6.keypos = 3                     ' キーポジション
    PSTOCK_Speck.ks6.keyleng = 20                   ' キー長
    PSTOCK_Speck.ks6.keyflag = BtKfExt              ' キーフラグ
    PSTOCK_Speck.ks6.keytype = Chr(BtKtString)      ' キータイプ
    PSTOCK_Speck.ks6.reserve = &H0                  ' 予約済み


    sts = BTRV(BtOpCreate, PSTOCK_POS, PSTOCK_Speck, Len(PSTOCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "範囲内移動分在庫データ")
        Exit Function
    End If
    
    PSTOCK_Create = False

End Function
Public Function PSTOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              範囲内移動分在庫一覧データ　ＯＰＥＮ                *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.04.27                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    PSTOCK_Open = True
                                            '範囲内移動分在庫一覧データフルパス取込み
    sts = GetIni("FILE", PSTOCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[PSTOCK] 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, PSTOCK_POS, PSTOCKREC, Len(PSTOCKREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PSTOCK_Create()               '範囲内移動分在庫一覧データ作成
                If sts Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PSTOCK_POS, PSTOCKREC, Len(SUMZREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "範囲内移動分在庫一覧データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "範囲内移動分在庫一覧データ")
                Exit Function
        End Select
    Loop

    PSTOCK_Open = False

End Function


