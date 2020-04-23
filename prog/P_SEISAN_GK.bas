Attribute VB_Name = "P_SEISAN_GK"
Option Explicit

'********************************************************************
'*
'*              生産実績明細集計ﾃﾞｰﾀ  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ
Public Const P_SEISAN_GK_ID$ = "P_SEISAN_GK"

'ページサイズ
Private Const P_SEISAN_GK_PG_SIZ% = 1024

'ポジション・ブロック
Public P_SEISAN_GK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
Private Type UCHIWAKE_TBL_Tag
    KIN(0 To 10)            As Byte         '完了分金額
End Type

'レコード定義
Public Type P_SEISAN_GK_REC_Tag
    
    TORI_KBN(0 To 0)        As Byte         '取引先区分
    TORI_CODE(0 To 4)       As Byte         '取引先ｺｰﾄﾞ
    
    UCHIWAKE_TBL(0 To 9)    As UCHIWAKE_TBL_Tag

    CNT(0 To 10)            As Byte         '件数
    QTY(0 To 10)            As Byte         '数量
    KAZEI(0 To 10)          As Byte         '課税対象額

End Type
'データ・バッファ
Public P_SEISAN_GK_REC      As P_SEISAN_GK_REC_Tag

'キー定義
Public Type KEY0_P_SEISAN_GK                'ＫＥＹ０
    TORI_CODE(0 To 4)       As Byte         '取引先ｺｰﾄﾞ
End Type
    
Public Type KEY1_P_SEISAN_GK                'ＫＥＹ１
    TORI_KBN(0 To 0)        As Byte         '取引先区分
    TORI_CODE(0 To 4)       As Byte         '取引先ｺｰﾄﾞ
End Type
    
'キー・データ
Public K0_P_SEISAN_GK       As KEY0_P_SEISAN_GK
Public K1_P_SEISAN_GK       As KEY1_P_SEISAN_GK

Type P_SEISAN_GK_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_SEISAN_GK_Speck   As P_SEISAN_GK_FSpeck
Private Function P_SEISAN_GK_Create() As Integer
'********************************************************************
'*
'*              生産実績明細集計ﾃﾞｰﾀ  ＣＲＥＡＴＥ
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


    P_SEISAN_GK_Create = True
                                            '生産実績明細集計ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SEISAN_GK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_GK]読み込みエラー")
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

    P_SEISAN_GK_Speck.fs.recoleng = Len(P_SEISAN_GK_REC)    ' レコード長
    P_SEISAN_GK_Speck.fs.PageSize = P_SEISAN_GK_PG_SIZ      ' ページサイズ
    P_SEISAN_GK_Speck.fs.idexnumb = 2                       ' インデックス数
    P_SEISAN_GK_Speck.fs.fileflag = 0                       ' ファイルフラグ
    P_SEISAN_GK_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_SEISAN_GK_Speck.ks0.keypos = 2                        ' キーポジション
    P_SEISAN_GK_Speck.ks0.keyleng = 4                       ' キー長
    P_SEISAN_GK_Speck.ks0.keyflag = BtKfExt                 ' キーフラグ
    P_SEISAN_GK_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    P_SEISAN_GK_Speck.ks0.reserve = &H0                     ' 予約済み
    
    
    
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    P_SEISAN_GK_Speck.ks1.keypos = 1                        ' キーポジション
    P_SEISAN_GK_Speck.ks1.keyleng = 1                       ' キー長
    P_SEISAN_GK_Speck.ks1.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    P_SEISAN_GK_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    P_SEISAN_GK_Speck.ks1.reserve = &H0                     ' 予約済み
    
    P_SEISAN_GK_Speck.ks2.keypos = 2                        ' キーポジション
    P_SEISAN_GK_Speck.ks2.keyleng = 5                       ' キー長
    P_SEISAN_GK_Speck.ks2.keyflag = BtKfExt                 ' キーフラグ
    P_SEISAN_GK_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    P_SEISAN_GK_Speck.ks2.reserve = &H0                     ' 予約済み
    
    
    
    '--------------------------------------------------- キー１ △
    
    
    
    sts = BTRV(BtOpCreate, P_SEISAN_GK_POS, P_SEISAN_GK_Speck, Len(P_SEISAN_GK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "生産実績明細集計ﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_SEISAN_GK_Create = False

End Function

Public Function P_SEISAN_GK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              生産実績明細集計ﾃﾞｰﾀ  ＯＰＥＮ
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

    P_SEISAN_GK_Open = True
                                            '生産実績明細集計ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SEISAN_GK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_GK]読み込みエラー")
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
        sts = BTRV(BtOpOpen, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SEISAN_GK_Create()  '生産実績明細集計ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "生産実績明細集計ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "生産実績明細集計ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_SEISAN_GK_Open = False

End Function

