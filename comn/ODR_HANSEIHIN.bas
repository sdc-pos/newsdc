Attribute VB_Name = "ODR_HANSEIHIN"
Option Explicit
'********************************************************************
'*                                                                  *
'*              半製品管理データ  ファイル定義                      *
'*                                                                  *
'*          CREATE 2008.04.26                                       *
'********************************************************************
'ファイルＩＤ
Public Const ODR_HANSEIHIN_ID$ = "ODR_HANSEIHIN"

'ページサイズ
Private Const ODR_HANSEIHIN_PG_SIZ% = 512

'ポジション・ブロック
Public ODR_HANSEIHIN_POS        As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type ODR_HANSEIHIN_O_REC_Tag                 '親ﾚｺｰﾄﾞ
    
    
    USE_YM(0 To 5)          As Byte         '使用月
    INPUT_NO(0 To 3)        As Byte         '入力順
    USE_YMD(0 To 7)         As Byte         '使用日付
    SEQNO(0 To 2)           As Byte         '追番(000)
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    SHIJI_QTY(0 To 11)      As Byte         '指示数 S9(8)V99
    UPD_TANTO(0 To 4)       As Byte         '更新　担当者
    UPD_DATE(0 To 7)        As Byte         '更新　日付
    UPD_TIME(0 To 5)        As Byte         '更新　時刻
    
    
    IO_FLG(0 To 0)          As Byte         '入出庫ﾌﾗｸﾞ     '2008.05.14
    
    
    
    
    FILLER(0 To 180)        As Byte
    
    

End Type
'データ・バッファ
Public ODR_HANSEIHIN_O_REC  As ODR_HANSEIHIN_O_REC_Tag


Public Type ODR_HANSEIHIN_K_REC_Tag                 '子ﾚｺｰﾄﾞ
    
    USE_YM(0 To 5)          As Byte         '使用月
    INPUT_NO(0 To 3)        As Byte         '入力順
    USE_YMD(0 To 7)         As Byte         '使用日付
    SEQNO(0 To 2)           As Byte         '追番(000)
    KO_JGYOBU(0 To 0)       As Byte         '事業部
    KO_NAIGAI(0 To 0)       As Byte         '国内外
    KO_HIN_GAI(0 To 19)     As Byte         '親品番
    KO_QTY(0 To 7)          As Byte         '指示数 9(5)V99
    USE_QTY(0 To 11)        As Byte         '指示数 S9(8)V99
    ZAITEI_F(0 To 0)        As Byte         '在訂マーク
    UPD_TANTO(0 To 4)       As Byte         '更新　担当者
    UPD_DATE(0 To 7)        As Byte         '更新　日付
    UPD_TIME(0 To 5)        As Byte         '更新　時刻
    
    
    IO_FLG(0 To 0)          As Byte         '入出庫ﾌﾗｸﾞ     '2008.05.14
    
    
    FILLER(0 To 171)        As Byte
    

End Type
'データ・バッファ
Public ODR_HANSEIHIN_K_REC  As ODR_HANSEIHIN_K_REC_Tag

'キー定義

Type KEY0_ODR_HANSEIHIN                           'ＫＥＹ０
'    USE_YM(0 To 5)          As Byte         '使用月            2008.05.13
    INPUT_NO(0 To 3)        As Byte         '入力順
    SEQNO(0 To 2)           As Byte         '追番
End Type
    
Type KEY1_ODR_HANSEIHIN                           'ＫＥＹ１
    KO_JGYOBU(0 To 0)       As Byte         '事業部
    KO_NAIGAI(0 To 0)       As Byte         '国内外
    KO_HIN_GAI(0 To 19)     As Byte         '親品番
End Type
    
Type KEY2_ODR_HANSEIHIN                           'ＫＥＹ２
    USE_YMD(0 To 7)         As Byte         '使用日付
End Type
    
    
    
'キー・データ
Public K0_ODR_HANSEIHIN     As KEY0_ODR_HANSEIHIN
Public K1_ODR_HANSEIHIN     As KEY1_ODR_HANSEIHIN
Public K2_ODR_HANSEIHIN     As KEY2_ODR_HANSEIHIN

Type ODR_HANSEIHIN_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体


End Type

Private ODR_HANSEIHIN_Speck As ODR_HANSEIHIN_FSpeck
Private Function ODR_HANSEIHIN_Create() As Integer
'********************************************************************
'*                                                                  *
'*              半製品管理データ  ＣＲＥＡＴＥ                      *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ODR_HANSEIHIN_Create = True
                                            '半製品管理データフルパス取込み
    sts = GetIni("FILE", ODR_HANSEIHIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_HANSEIHIN]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    ODR_HANSEIHIN_Speck.fs.recoleng = Len(ODR_HANSEIHIN_O_REC)      ' レコード長
    ODR_HANSEIHIN_Speck.fs.PageSize = ODR_HANSEIHIN_PG_SIZ          ' ページサイズ
    ODR_HANSEIHIN_Speck.fs.idexnumb = 3                             ' インデックス数
    ODR_HANSEIHIN_Speck.fs.fileflag = 0                             ' ファイルフラグ
    ODR_HANSEIHIN_Speck.fs.reserve = &H0                            ' 予約済み
    '--------------------------------------------------- キー０ ▽
    
    ODR_HANSEIHIN_Speck.ks0.keypos = 7                      ' キーポジション
    ODR_HANSEIHIN_Speck.ks0.keyleng = 4                     ' キー長
    ODR_HANSEIHIN_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    ODR_HANSEIHIN_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    ODR_HANSEIHIN_Speck.ks0.reserve = &H0                   ' 予約済み
    
    ODR_HANSEIHIN_Speck.ks1.keypos = 19                     ' キーポジション
    ODR_HANSEIHIN_Speck.ks1.keyleng = 3                     ' キー長
    ODR_HANSEIHIN_Speck.ks1.keyflag = BtKfExt               ' キーフラグ
    ODR_HANSEIHIN_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    ODR_HANSEIHIN_Speck.ks1.reserve = &H0                   ' 予約済み
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    ODR_HANSEIHIN_Speck.ks2.keypos = 22                     ' キーポジション
    ODR_HANSEIHIN_Speck.ks2.keyleng = 1                     ' キー長
                                                            ' キーフラグ
    ODR_HANSEIHIN_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ODR_HANSEIHIN_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    ODR_HANSEIHIN_Speck.ks2.reserve = &H0                   ' 予約済み
    
    ODR_HANSEIHIN_Speck.ks3.keypos = 23                     ' キーポジション
    ODR_HANSEIHIN_Speck.ks3.keyleng = 1                     ' キー長
                                                            ' キーフラグ
    ODR_HANSEIHIN_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ODR_HANSEIHIN_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    ODR_HANSEIHIN_Speck.ks3.reserve = &H0                   ' 予約済み
    
    ODR_HANSEIHIN_Speck.ks4.keypos = 24                     ' キーポジション
    ODR_HANSEIHIN_Speck.ks4.keyleng = 20                    ' キー長
                                                            ' キーフラグ
    ODR_HANSEIHIN_Speck.ks4.keyflag = BtKfExt + BtKfDup + BtKfChg
    ODR_HANSEIHIN_Speck.ks4.keytype = Chr(BtKtString)       ' キータイプ
    ODR_HANSEIHIN_Speck.ks4.reserve = &H0                   ' 予約済み
    '--------------------------------------------------- キー１ △
    
    '--------------------------------------------------- キー２ ▽
    ODR_HANSEIHIN_Speck.ks5.keypos = 11                     ' キーポジション
    ODR_HANSEIHIN_Speck.ks5.keyleng = 8                    ' キー長
                                                            ' キーフラグ
    ODR_HANSEIHIN_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg
    ODR_HANSEIHIN_Speck.ks5.keytype = Chr(BtKtString)       ' キータイプ
    ODR_HANSEIHIN_Speck.ks5.reserve = &H0                   ' 予約済み
    '--------------------------------------------------- キー２ △
    
    
    sts = BTRV(BtOpCreate, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_Speck, Len(ODR_HANSEIHIN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "半製品管理データ")
        Exit Function
    End If
    
    ODR_HANSEIHIN_Create = False

End Function

Public Function ODR_HANSEIHIN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              半製品管理データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ODR_HANSEIHIN_Open = True
                                            '半製品管理データフルパス取込み
    sts = GetIni("FILE", ODR_HANSEIHIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_HANSEIHIN]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ODR_HANSEIHIN_Create()    '半製品管理データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "半製品管理データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "半製品管理データ")
                Exit Function
        End Select
    Loop
    
    ODR_HANSEIHIN_Open = False

End Function
