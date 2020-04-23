Attribute VB_Name = "PLN_O_HOURS"
Option Explicit
'********************************************************************
'*
'*              担当者別勤務時間データ  ファイル定義
'*
'*          CREATE 2011.09.13
'********************************************************************
'ファイルＩＤ
Public Const PLN_O_HOURS_ID$ = "PLN_O_HOURS"

'ページサイズ
Public Const PLN_O_HOURS_PG_SIZ% = 512

'ポジション・ブロック
Public PLN_O_HOURS_POS            As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type PLN_O_HOURS_REC_Tag
    TANTO_CODE(0 To 4)              As Byte         'データ作成日
    O_DATE(0 To 7)                  As Byte         '年月日
    O_Time(0 To 3)                  As Byte         '勤務時間 99.9
    FILLER(0 To 62)                 As Byte
    INS_TANTO(0 To 9)               As Byte         '追加　担当者
    Ins_DateTime(0 To 13)           As Byte         '追加　日時
    UPD_TANTO(0 To 9)               As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)           As Byte         '更新　日時



End Type
'データ・バッファ
Public PLN_O_HOURS_REC              As PLN_O_HOURS_REC_Tag

'キー定義

Type KEY0_PLN_O_HOURS               'ＫＥＹ０
    
    TANTO_CODE(0 To 4)              As Byte         '担当者ｺｰﾄﾞ
    O_DATE(0 To 7)                  As Byte         '年月日

End Type

Type KEY1_PLN_O_HOURS               'ＫＥＹ１
    
    O_DATE(0 To 7)                  As Byte         '年月日
    TANTO_CODE(0 To 4)              As Byte         '担当者ｺｰﾄﾞ

End Type






'キー・データ
Public K0_PLN_O_HOURS               As KEY0_PLN_O_HOURS
Public K1_PLN_O_HOURS               As KEY1_PLN_O_HOURS



Private Type PLN_O_HOURS_FSpeck
    fs      As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck

End Type

Private PLN_O_HOURS_Speck           As PLN_O_HOURS_FSpeck

Private Function PLN_O_HOURS_Create() As Integer
'********************************************************************
'*
'*              担当者別勤務時間データ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PLN_O_HOURS_Create = True
                                            '担当者別勤務時間データ フルパス取込み
    sts = GetIni("FILE", PLN_O_HOURS_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_O_HOURS]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    PLN_O_HOURS_Speck.fs.recoleng = Len(PLN_O_HOURS_REC)    ' レコード長
    PLN_O_HOURS_Speck.fs.PageSize = PLN_O_HOURS_PG_SIZ      ' ページサイズ
    PLN_O_HOURS_Speck.fs.idexnumb = 2                       ' インデックス数
    PLN_O_HOURS_Speck.fs.fileflag = 0                       ' ファイルフラグ
    PLN_O_HOURS_Speck.fs.reserve = &H0                      ' 予約済み
'-----------------------------------------------
                                                ' キー０
    PLN_O_HOURS_Speck.ks0.keypos = 1                        ' キーポジション
    PLN_O_HOURS_Speck.ks0.keyleng = 5                       ' キー長
    PLN_O_HOURS_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    PLN_O_HOURS_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    PLN_O_HOURS_Speck.ks0.reserve = &H0                     ' 予約済み

    PLN_O_HOURS_Speck.ks1.keypos = 6                        ' キーポジション
    PLN_O_HOURS_Speck.ks1.keyleng = 8                       ' キー長
    PLN_O_HOURS_Speck.ks1.keyflag = BtKfExt                 ' キーフラグ
    PLN_O_HOURS_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    PLN_O_HOURS_Speck.ks1.reserve = &H0                     ' 予約済み

'-----------------------------------------------
                                                ' キー１
    PLN_O_HOURS_Speck.ks2.keypos = 6                        ' キーポジション
    PLN_O_HOURS_Speck.ks2.keyleng = 8                       ' キー長
    PLN_O_HOURS_Speck.ks2.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    PLN_O_HOURS_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    PLN_O_HOURS_Speck.ks2.reserve = &H0                     ' 予約済み

    PLN_O_HOURS_Speck.ks3.keypos = 1                        ' キーポジション
    PLN_O_HOURS_Speck.ks3.keyleng = 5                       ' キー長
    PLN_O_HOURS_Speck.ks3.keyflag = BtKfExt                 ' キーフラグ
    PLN_O_HOURS_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    PLN_O_HOURS_Speck.ks3.reserve = &H0                     ' 予約済み

'-----------------------------------------------

    sts = BTRV(BtOpCreate, PLN_O_HOURS_POS, PLN_O_HOURS_Speck, Len(PLN_O_HOURS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "担当者別勤務時間データ")
        Exit Function
    End If

    PLN_O_HOURS_Create = False

End Function

Public Function PLN_O_HOURS_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              担当者別勤務時間データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    PLN_O_HOURS_Open = True
                                            '担当者別勤務時間データ フルパス取込み
    sts = GetIni("FILE", PLN_O_HOURS_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_O_HOURS]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PLN_O_HOURS_Create()  '担当者別勤務時間データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "担当者別勤務時間データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "担当者別勤務時間データ")
                Exit Function
        End Select
    Loop

    PLN_O_HOURS_Open = False

End Function

