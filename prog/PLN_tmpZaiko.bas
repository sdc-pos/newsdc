Attribute VB_Name = "PLN_tmpZaiko"
Option Explicit
'********************************************************************
'*
'*              資材所要量確認画面中間ファイル ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const PLN_tmpZaiko_ID$ = "PLN_tmpZaiko"

'ページサイズ
Public Const PLN_tmpZaiko_PG_SIZ% = 1024

'ポジション・ブロック
Public PLN_tmpZaiko_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Type PLN_tmpZaikoREC_Tag
    SYUBETSU(0 To 1)        As Byte     '種別
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    RIREKI_DT(0 To 7)       As Byte     '年月日
    DATA_KBN(0 To 0)        As Byte     'ﾃﾞｰﾀ区分
    ST_ZAIKO_QTY(0 To 5)    As Byte     '開始時在庫数
    SYOUHI_QTY(0 To 5)      As Byte     '消費
    NYUKA_QTY(0 To 5)       As Byte     '入荷
    ZAIKO_QTY(0 To 5)       As Byte     '在庫残
    INS_TANTO(0 To 9)       As Byte     '追加　担当者
    Ins_DateTime(0 To 13)   As Byte     '追加　日時         YYYYMMDDhhmmss

End Type

'データ・バッファ
Public PLN_tmpZaikoREC      As PLN_tmpZaikoREC_Tag

'キー定義
Type KEY0_PLN_tmpZaiko              'ＫＥＹ０
    SYUBETSU(0 To 1)        As Byte     '種別
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
    RIREKI_DT(0 To 7)       As Byte     '年月日
End Type

Type KEY1_PLN_tmpZaiko              'ＫＥＹ１
    RIREKI_DT(0 To 7)       As Byte     '年月日
    SYUBETSU(0 To 1)        As Byte     '種別
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type

'キー・データ
Public K0_PLN_tmpZaiko      As KEY0_PLN_tmpZaiko
Public K1_PLN_tmpZaiko      As KEY1_PLN_tmpZaiko

Type PLN_tmpZaiko_FSpeck
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
End Type

Private PLN_tmpZaiko_Speck  As PLN_tmpZaiko_FSpeck
Private Function PLN_tmpZaiko_Create() As Integer
'********************************************************************
'*
'*              資材所要量確認画面中間ファイル　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PLN_tmpZaiko_Create = True
                                            '在庫データ　フルパス取込み
    sts = GetIni("FILE", PLN_tmpZaiko_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_tmpZaiko]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    PLN_tmpZaiko_Speck.fs.recoleng = Len(PLN_tmpZaikoREC)   ' レコード長
    PLN_tmpZaiko_Speck.fs.PageSize = PLN_tmpZaiko_PG_SIZ    ' ページサイズ
    PLN_tmpZaiko_Speck.fs.idexnumb = 2                      ' インデックス数
    PLN_tmpZaiko_Speck.fs.fileflag = 0                      ' ファイルフラグ
    PLN_tmpZaiko_Speck.fs.reserve = &H0                     ' 予約済み
'---------------------------------------------------'
                                                    ' キー０
    PLN_tmpZaiko_Speck.ks0.keypos = 1                       ' キーポジション
    PLN_tmpZaiko_Speck.ks0.keyleng = 2                      ' キー長
    PLN_tmpZaiko_Speck.ks0.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    PLN_tmpZaiko_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    PLN_tmpZaiko_Speck.ks0.reserve = &H0                    ' 予約済み
                                                    
    PLN_tmpZaiko_Speck.ks1.keypos = 3                       ' キーポジション
    PLN_tmpZaiko_Speck.ks1.keyleng = 1                      ' キー長
    PLN_tmpZaiko_Speck.ks1.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    PLN_tmpZaiko_Speck.ks1.keytype = Chr(BtKtString)        ' キータイプ
    PLN_tmpZaiko_Speck.ks1.reserve = &H0                    ' 予約済み
                                                    
    PLN_tmpZaiko_Speck.ks2.keypos = 4                       ' キーポジション
    PLN_tmpZaiko_Speck.ks2.keyleng = 1                      ' キー長
    PLN_tmpZaiko_Speck.ks2.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    PLN_tmpZaiko_Speck.ks2.keytype = Chr(BtKtString)        ' キータイプ
    PLN_tmpZaiko_Speck.ks2.reserve = &H0                    ' 予約済み
                                                    
    PLN_tmpZaiko_Speck.ks3.keypos = 5                       ' キーポジション
    PLN_tmpZaiko_Speck.ks3.keyleng = 20                     ' キー長
    PLN_tmpZaiko_Speck.ks3.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    PLN_tmpZaiko_Speck.ks3.keytype = Chr(BtKtString)        ' キータイプ
    PLN_tmpZaiko_Speck.ks3.reserve = &H0                    ' 予約済み
                                                    
    PLN_tmpZaiko_Speck.ks4.keypos = 25                      ' キーポジション
    PLN_tmpZaiko_Speck.ks4.keyleng = 8                      ' キー長
    PLN_tmpZaiko_Speck.ks4.keyflag = BtKfExt                ' キーフラグ
    PLN_tmpZaiko_Speck.ks4.keytype = Chr(BtKtString)        ' キータイプ
    PLN_tmpZaiko_Speck.ks4.reserve = &H0                    ' 予約済み
                                                    
'---------------------------------------------------'
                                                    ' キー１
    PLN_tmpZaiko_Speck.ks5.keypos = 25                      ' キーポジション
    PLN_tmpZaiko_Speck.ks5.keyleng = 8                      ' キー長
                                                            ' キーフラグ
    PLN_tmpZaiko_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup
    PLN_tmpZaiko_Speck.ks5.keytype = Chr(BtKtString)        ' キータイプ
    PLN_tmpZaiko_Speck.ks5.reserve = &H0                    ' 予約済み
    
    PLN_tmpZaiko_Speck.ks6.keypos = 1                       ' キーポジション
    PLN_tmpZaiko_Speck.ks6.keyleng = 2                      ' キー長
                                                            ' キーフラグ
    PLN_tmpZaiko_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfDup
    PLN_tmpZaiko_Speck.ks6.keytype = Chr(BtKtString)        ' キータイプ
    PLN_tmpZaiko_Speck.ks6.reserve = &H0                    ' 予約済み
    
    
    PLN_tmpZaiko_Speck.ks7.keypos = 3                       ' キーポジション
    PLN_tmpZaiko_Speck.ks7.keyleng = 1                      ' キー長
                                                            ' キーフラグ
    PLN_tmpZaiko_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup
    PLN_tmpZaiko_Speck.ks7.keytype = Chr(BtKtString)        ' キータイプ
    PLN_tmpZaiko_Speck.ks7.reserve = &H0                    ' 予約済み
                                                    
    PLN_tmpZaiko_Speck.ks8.keypos = 4                       ' キーポジション
    PLN_tmpZaiko_Speck.ks8.keyleng = 1                      ' キー長
                                                            ' キーフラグ
    PLN_tmpZaiko_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfDup
    PLN_tmpZaiko_Speck.ks8.keytype = Chr(BtKtString)        ' キータイプ
    PLN_tmpZaiko_Speck.ks8.reserve = &H0                    ' 予約済み
                                                    
    PLN_tmpZaiko_Speck.ks9.keypos = 5                       ' キーポジション
    PLN_tmpZaiko_Speck.ks9.keyleng = 20                     ' キー長
                                                            ' キーフラグ
    PLN_tmpZaiko_Speck.ks9.keyflag = BtKfExt + BtKfDup
    PLN_tmpZaiko_Speck.ks9.keytype = Chr(BtKtString)        ' キータイプ
    PLN_tmpZaiko_Speck.ks9.reserve = &H0                    ' 予約済み
'---------------------------------------------------'
    sts = BTRV(BtOpCreate, PLN_tmpZaiko_POS, PLN_tmpZaiko_Speck, Len(PLN_tmpZaiko_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "資材所要量確認画面中間ファイル")
        Exit Function
    End If
    PLN_tmpZaiko_Create = False
End Function
Public Function PLN_tmpZaiko_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              資材所要量確認画面中間ファイル　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    PLN_tmpZaiko_Open = True
                                            '資材所要量確認画面中間ファイル　フルパス取込み
    sts = GetIni("FILE", PLN_tmpZaiko_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_tmpZaiko]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PLN_tmpZaiko_Create() '資材所要量確認画面中間ファイル　作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "資材所要量確認画面中間ファイル")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "資材所要量確認画面中間ファイル")
                Exit Function
        End Select
    Loop
    PLN_tmpZaiko_Open = False

End Function

