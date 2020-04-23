Attribute VB_Name = "P_SEISAN_DET"
Option Explicit

'********************************************************************
'*
'*              生産実績明細ﾃﾞｰﾀ  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ
Public Const P_SEISAN_DET_ID$ = "P_SEISAN_DET"

'ページサイズ
Private Const P_SEISAN_DET_PG_SIZ% = 1024

'ポジション・ブロック
Public P_SEISAN_DET_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************

'レコード定義
Public Type P_SEISAN_DET_REC_Tag
    
    TORI_KBN(0 To 0)        As Byte         '取引先区分
    TORI_CODE(0 To 4)       As Byte         '取引先ｺｰﾄﾞ
    UKEIRE_DT(0 To 7)       As Byte         '受入日
    SHIJI_NO(0 To 4)        As Byte         '指図票№
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    HIN_GAI(0 To 19)        As Byte         '親品番
    UKEIRE_QTY(0 To 10)      As Byte         '受入数(9(8)V99)
    S_CLASS_CODE(0 To 19)   As Byte         '商品化ｸﾗｽ
    F_CLASS_CODE(0 To 19)   As Byte         '付加ｸﾗｽ
    N_CLASS_CODE(0 To 19)   As Byte         '内職ｸﾗｽ
    KOURYOU(0 To 10)        As Byte         '単価 9(8)V99
    KIN(0 To 10)            As Byte         '金額


End Type
'データ・バッファ
Public P_SEISAN_DET_REC     As P_SEISAN_DET_REC_Tag

'キー定義
Public Type KEY0_P_SEISAN_DET               'ＫＥＹ０
    TORI_CODE(0 To 4)       As Byte         '取引先ｺｰﾄﾞ
End Type
    
    
'キー・データ
Public K0_P_SEISAN_DET      As KEY0_P_SEISAN_DET

Type P_SEISAN_DET_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_SEISAN_DET_Speck  As P_SEISAN_DET_FSpeck
Private Function P_SEISAN_DET_Create() As Integer
'********************************************************************
'*
'*              生産実績明細ﾃﾞｰﾀ  ＣＲＥＡＴＥ
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

    P_SEISAN_DET_Create = True
                                            '生産実績明細ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SEISAN_DET_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_DET]読み込みエラー")
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

    P_SEISAN_DET_Speck.fs.recoleng = Len(P_SEISAN_DET_REC)  ' レコード長
    P_SEISAN_DET_Speck.fs.PageSize = P_SEISAN_DET_PG_SIZ    ' ページサイズ
    P_SEISAN_DET_Speck.fs.idexnumb = 1                      ' インデックス数
    P_SEISAN_DET_Speck.fs.fileflag = 0                      ' ファイルフラグ
    P_SEISAN_DET_Speck.fs.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ ▽
    
    P_SEISAN_DET_Speck.ks0.keypos = 2                       ' キーポジション
    P_SEISAN_DET_Speck.ks0.keyleng = 5                      ' キー長
    P_SEISAN_DET_Speck.ks0.keyflag = BtKfExt + BtKfDup      ' キーフラグ
    P_SEISAN_DET_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    P_SEISAN_DET_Speck.ks0.reserve = &H0                    ' 予約済み
    
    
    
    '--------------------------------------------------- キー０ △
    
    
    
    
    sts = BTRV(BtOpCreate, P_SEISAN_DET_POS, P_SEISAN_DET_Speck, Len(P_SEISAN_DET_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "生産実績明細ﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_SEISAN_DET_Create = False

End Function

Public Function P_SEISAN_DET_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              生産実績明細ﾃﾞｰﾀ  ＯＰＥＮ
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


    P_SEISAN_DET_Open = True
                                            '生産実績明細データフルパス取込み
    sts = GetIni("FILE", P_SEISAN_DET_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_DET]読み込みエラー")
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
        sts = BTRV(BtOpOpen, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SEISAN_DET_Create()     '生産実績明細ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "生産実績明細ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "生産実績明細ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_SEISAN_DET_Open = False

End Function

