Attribute VB_Name = "P_SHKENTO"
Option Explicit

'********************************************************************
'*
'*              発注検討ﾌｧｲﾙ  ファイル定義
'*
'*          CREATE 2006.11.17
'********************************************************************
'ファイルＩＤ
Public Const P_SHKENTO_ID$ = "P_SHKENTO"

'ページサイズ
Private Const P_SHKENTO_PG_SIZ% = 1024

'ポジション・ブロック
Public P_SHKENTO_POS       As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義

Private Type JITU_TBL_Tag
    JITU_YM(0 To 6)        As Byte
    JITU_QTY(0 To 7)        As Byte
End Type





Public Type P_SHKENTO_REC_Tag
    
    
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
                                        '実績消費数量
    JITU_TBL(0 To 2)        As JITU_TBL_Tag
    
    LT_CODE(0 To 0)         As Byte     'ﾘｰﾄﾞﾀｲﾑ　ｺｰﾄﾞ
    LT_DAYS(0 To 2)         As Byte     'ﾘｰﾄﾞﾀｲﾑ　日数
    
    SYUSHI_CODE(0 To 2)     As Byte     '収支単位
    
    ZAIKO_STANDARD(0 To 7)  As Byte     '基準在庫
    ZAIKO_QTY(0 To 7)       As Byte     '現在庫
    
    LOT(0 To 7)             As Byte     '発注ﾛｯﾄ
    ORDER_CODE(0 To 4)      As Byte     '注文先ｺｰﾄﾞ
        
    SHIJI_Z_QTY(0 To 7)     As Byte     '発注残
    SHIJI_Z_CODE(0 To 0)    As Byte     '発注残ｺｰﾄﾞ
        
    SHIJI_QTY_R(0 To 7)     As Byte     '発注数理論
    SHIJI_QTY_K(0 To 7)     As Byte     '発注数確定
    SHIJI_CODE(0 To 0)      As Byte     '発注ｺｰﾄﾞ

    TANKA(0 To 10)          As Byte     '受入単価(9(8)V99)
    KINGAKU(0 To 9)         As Byte     '受入金額(S9(9))

    SORT_KEY(0 To 9)        As Byte
    
    S_YMD(0 To 7)           As Byte     '指定　開始年月日
    E_YMD(0 To 7)           As Byte
    
    
    FILLER(0 To 15)         As Byte     '指定　終了年月日

End Type
'データ・バッファ
Public P_SHKENTO_REC        As P_SHKENTO_REC_Tag

'キー定義

Public Type KEY0_P_SHKENTO                     'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type

Public Type KEY1_P_SHKENTO                     'ＫＥＹ１
    SORT_KEY(0 To 9)        As Byte
End Type
    
'キー・データ
Public K0_P_SHKENTO         As KEY0_P_SHKENTO
Public K1_P_SHKENTO         As KEY1_P_SHKENTO

Type P_SHKENTO_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_SHKENTO_Speck     As P_SHKENTO_FSpeck
Private Function P_SHKENTO_Create() As Integer
'********************************************************************
'*
'*              資材発注検討ﾌｧｲﾙ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim sBuffer         As String * 255
Dim com             As String

Dim Ret             As Integer


    P_SHKENTO_Create = True
                                            '資材受入履歴ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SHKENTO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHKENTO]読み込みエラー")
        Exit Function
    End If


    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If

    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)


    P_SHKENTO_Speck.fs.recoleng = Len(P_SHKENTO_REC)    ' レコード長
    P_SHKENTO_Speck.fs.PageSize = P_SHKENTO_PG_SIZ      ' ページサイズ
    P_SHKENTO_Speck.fs.idexnumb = 2                     ' インデックス数
    P_SHKENTO_Speck.fs.fileflag = 0                     ' ファイルフラグ
    P_SHKENTO_Speck.fs.reserve = &H0                    ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_SHKENTO_Speck.ks0.keypos = 1                      ' キーポジション
    P_SHKENTO_Speck.ks0.keyleng = 1                     ' キー長
    P_SHKENTO_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    P_SHKENTO_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    P_SHKENTO_Speck.ks0.reserve = &H0                   ' 予約済み
    
    P_SHKENTO_Speck.ks1.keypos = 2                      ' キーポジション
    P_SHKENTO_Speck.ks1.keyleng = 1                     ' キー長
    P_SHKENTO_Speck.ks1.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    P_SHKENTO_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    P_SHKENTO_Speck.ks1.reserve = &H0                   ' 予約済み
    
    P_SHKENTO_Speck.ks2.keypos = 3                      ' キーポジション
    P_SHKENTO_Speck.ks2.keyleng = 20                    ' キー長
    P_SHKENTO_Speck.ks2.keyflag = BtKfExt               ' キーフラグ
    P_SHKENTO_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    P_SHKENTO_Speck.ks2.reserve = &H0                   ' 予約済み
    
    
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    P_SHKENTO_Speck.ks3.keypos = 151                    ' キーポジション
    P_SHKENTO_Speck.ks3.keyleng = 10                    ' キー長
                                                        ' キーフラグ
    P_SHKENTO_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_SHKENTO_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    P_SHKENTO_Speck.ks3.reserve = &H0                   ' 予約済み
    
    '--------------------------------------------------- キー１ ▽
    
    sts = BTRV(BtOpCreate, P_SHKENTO_POS, P_SHKENTO_Speck, Len(P_SHKENTO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "資材発注検討ﾌｧｲﾙ")
        Exit Function
    End If
    
    P_SHKENTO_Create = False

End Function

Public Function P_SHKENTO_Open(mode As Integer) As Integer
'********************************************************************
'*
'*              資材発注検討ﾌｧｲﾙ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim sBuffer         As String * 255
Dim com             As String

Dim Ret             As Integer

    P_SHKENTO_Open = True
                                            '資材発注検討ﾌｧｲﾙフルパス取込み
    sts = GetIni("FILE", P_SHKENTO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHKENTO]読み込みエラー")
        Exit Function
    End If
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If

    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)
    
    

    Do
        sts = BTRV(BtOpOpen, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), ByVal FullPath, Len(FullPath), mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHKENTO_Create()   '資材発注検討ﾌｧｲﾙ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), ByVal FullPath, Len(FullPath), mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "資材発注検討ﾌｧｲﾙ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "資材発注検討ﾌｧｲﾙ")
                Exit Function
        End Select
    Loop
    
    P_SHKENTO_Open = False

End Function

