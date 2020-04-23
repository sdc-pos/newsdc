Attribute VB_Name = "ODR_BUHIN_ORDER"
Option Explicit
'********************************************************************
'*                                                                  *
'*              子部品　注文Ｆ ファイル定義                         *
'*                                                                  *
'*          CREATE 2008.02.19                                       *
'********************************************************************
'ファイルＩＤ
Public Const ODR_BUHIN_ORDER_ID$ = "ODR_BUHIN_ORDER"

'ページサイズ
Private Const ODR_BUHIN_ORDER_PG_SIZ% = 1024

'ポジション・ブロック
Public ODR_BUHIN_ORDER_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type ODR_BUHIN_ORDER_REC_Tag
    SEL_DATE(0 To 7)            As Byte         '対象日付
    
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    DATA_KBN(0 To 0)            As Byte         'データ区分 1:予定 2:実績
    USE_YM(0 To 5)              As Byte         '使用月（YYYYMM)
    NYUKO_QTY(0 To 7)           As Byte         '注文数

End Type
'データ・バッファ
Public ODR_BUHIN_ORDER_REC            As ODR_BUHIN_ORDER_REC_Tag



'キー定義

Type KEY0_ODR_BUHIN_ORDER                           'ＫＥＹ０
    SEL_DATE(0 To 7)            As Byte         '対象日付
    
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番

    DATA_KBN(0 To 0)            As Byte         'データ区分 1:予定 2:実績

End Type

Type KEY1_ODR_BUHIN_ORDER                           'ＫＥＹ１
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番

    SEL_DATE(0 To 7)            As Byte         '対象日付

    DATA_KBN(0 To 0)            As Byte         'データ区分 1:予定 2:実績

End Type




'キー・データ
Public K0_ODR_BUHIN_ORDER           As KEY0_ODR_BUHIN_ORDER
Public K1_ODR_BUHIN_ORDER           As KEY1_ODR_BUHIN_ORDER

Type ODR_BUHIN_ORDER_FSpeck
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
    

End Type

Private ODR_BUHIN_ORDER_Speck       As ODR_BUHIN_ORDER_FSpeck
Private Function ODR_BUHIN_ORDER_Create() As Integer
'********************************************************************
'*                                                                  *
'*              子部品＿注文Ｆ  ＣＲＥＡＴＥ                        *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim sBuffer         As String * 255
Dim com             As String


Dim Ret             As Integer


    ODR_BUHIN_ORDER_Create = True
                                            '子部品＿注文Ｆフルパス取込み
    sts = GetIni("FILE", ODR_BUHIN_ORDER_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_ORDER]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)


    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)


    ODR_BUHIN_ORDER_Speck.fs.recoleng = Len(ODR_BUHIN_ORDER_REC)      ' レコード長
    ODR_BUHIN_ORDER_Speck.fs.PageSize = ODR_BUHIN_ORDER_PG_SIZ        ' ページサイズ
    ODR_BUHIN_ORDER_Speck.fs.idexnumb = 2                       ' インデックス数
    ODR_BUHIN_ORDER_Speck.fs.fileflag = 0                       ' ファイルフラグ
    ODR_BUHIN_ORDER_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    ODR_BUHIN_ORDER_Speck.ks0.keypos = 1                        ' キーポジション
    ODR_BUHIN_ORDER_Speck.ks0.keyleng = 8                       ' キー長
                                                                ' キーフラグ
    ODR_BUHIN_ORDER_Speck.ks0.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    ODR_BUHIN_ORDER_Speck.ks0.reserve = &H0                     ' 予約済み
    
    ODR_BUHIN_ORDER_Speck.ks1.keypos = 9                        ' キーポジション
    ODR_BUHIN_ORDER_Speck.ks1.keyleng = 1                       ' キー長
                                                                ' キーフラグ
    ODR_BUHIN_ORDER_Speck.ks1.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    ODR_BUHIN_ORDER_Speck.ks1.reserve = &H0                     ' 予約済み
    
    ODR_BUHIN_ORDER_Speck.ks2.keypos = 10                       ' キーポジション
    ODR_BUHIN_ORDER_Speck.ks2.keyleng = 1                       ' キー長
                                                                ' キーフラグ
    ODR_BUHIN_ORDER_Speck.ks2.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    ODR_BUHIN_ORDER_Speck.ks2.reserve = &H0                     ' 予約済み
    
    ODR_BUHIN_ORDER_Speck.ks3.keypos = 11                       ' キーポジション
    ODR_BUHIN_ORDER_Speck.ks3.keyleng = 20                      ' キー長
                                                                ' キーフラグ
    ODR_BUHIN_ORDER_Speck.ks3.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    ODR_BUHIN_ORDER_Speck.ks3.reserve = &H0                     ' 予約済み
    
    ODR_BUHIN_ORDER_Speck.ks4.keypos = 31                       ' キーポジション
    ODR_BUHIN_ORDER_Speck.ks4.keyleng = 1                       ' キー長
                                                                ' キーフラグ
    ODR_BUHIN_ORDER_Speck.ks4.keyflag = BtKfExt
    ODR_BUHIN_ORDER_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    ODR_BUHIN_ORDER_Speck.ks4.reserve = &H0                     ' 予約済み
    
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    ODR_BUHIN_ORDER_Speck.ks5.keypos = 9                        ' キーポジション
    ODR_BUHIN_ORDER_Speck.ks5.keyleng = 1                       ' キー長
                                                                ' キーフラグ
    ODR_BUHIN_ORDER_Speck.ks5.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks5.keytype = Chr(BtKtString)         ' キータイプ
    ODR_BUHIN_ORDER_Speck.ks5.reserve = &H0                     ' 予約済み
    
    ODR_BUHIN_ORDER_Speck.ks6.keypos = 10                       ' キーポジション
    ODR_BUHIN_ORDER_Speck.ks6.keyleng = 1                       ' キー長
                                                                ' キーフラグ
    ODR_BUHIN_ORDER_Speck.ks6.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks6.keytype = Chr(BtKtString)         ' キータイプ
    ODR_BUHIN_ORDER_Speck.ks6.reserve = &H0                     ' 予約済み
    
    ODR_BUHIN_ORDER_Speck.ks7.keypos = 11                       ' キーポジション
    ODR_BUHIN_ORDER_Speck.ks7.keyleng = 20                      ' キー長
                                                                ' キーフラグ
    ODR_BUHIN_ORDER_Speck.ks7.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks7.keytype = Chr(BtKtString)         ' キータイプ
    ODR_BUHIN_ORDER_Speck.ks7.reserve = &H0                     ' 予約済み
    
    ODR_BUHIN_ORDER_Speck.ks8.keypos = 1                        ' キーポジション
    ODR_BUHIN_ORDER_Speck.ks8.keyleng = 8                       ' キー長
                                                                ' キーフラグ
    ODR_BUHIN_ORDER_Speck.ks8.keyflag = BtKfExt + BtKfSeg
    ODR_BUHIN_ORDER_Speck.ks8.keytype = Chr(BtKtString)         ' キータイプ
    ODR_BUHIN_ORDER_Speck.ks8.reserve = &H0                     ' 予約済み
    
    
    ODR_BUHIN_ORDER_Speck.ks9.keypos = 31                       ' キーポジション
    ODR_BUHIN_ORDER_Speck.ks9.keyleng = 1                       ' キー長
                                                                ' キーフラグ
    ODR_BUHIN_ORDER_Speck.ks9.keyflag = BtKfExt
    ODR_BUHIN_ORDER_Speck.ks9.keytype = Chr(BtKtString)         ' キータイプ
    '--------------------------------------------------- キー１ △
    
    
    
    
    sts = BTRV(BtOpCreate, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_Speck, Len(ODR_BUHIN_ORDER_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "子部品＿注文Ｆ")
        Exit Function
    End If
    
    ODR_BUHIN_ORDER_Create = False

End Function

Public Function ODR_BUHIN_ORDER_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              子部品＿注文Ｆ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim sBuffer     As String * 255
Dim com         As String


Dim Ret         As Integer


    ODR_BUHIN_ORDER_Open = True
                                            '子部品＿注文Ｆフルパス取込み
    sts = GetIni("FILE", ODR_BUHIN_ORDER_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_ORDER]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)


    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)


    Do
        sts = BTRV(BtOpOpen, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ODR_BUHIN_ORDER_Create()      '子部品＿注文Ｆ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "子部品 注文Ｆ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "子部品 注文Ｆ")
                Exit Function
        End Select
    Loop
    
    ODR_BUHIN_ORDER_Open = False
    
End Function
