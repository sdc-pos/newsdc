Attribute VB_Name = "P_SH_SHI_SUM"
Option Explicit

'********************************************************************
'*
'*              資材仕入集計ﾃﾞｰﾀ  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ
Public Const P_SHSHI_SUM_ID$ = "P_SHSHI_SUM"

'ページサイズ
Private Const P_SHSHI_SUM_PG_SIZ% = 1024

'ポジション・ブロック
Public P_SHSHI_SUM_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
Private Type SHIIRE_TBL_Tag
    SHIIRE(0 To 9)          As Byte
End Type

'レコード定義
Public Type P_SHSHI_SUM_REC_Tag
    
    SHIIRE_CODE(0 To 4)     As Byte             '仕入先ｺｰﾄﾞ
    TORI_KBN(0 To 0)        As Byte             '取引先区分
    SHIIRE_TBL(0 To 6)      As SHIIRE_TBL_Tag

End Type
'データ・バッファ
Public P_SHSHI_SUM_REC      As P_SHSHI_SUM_REC_Tag

'キー定義
Public Type KEY0_P_SHSHI_SUM            'ＫＥＹ０
    SHIIRE_CODE(0 To 4)     As Byte             '仕入先ｺｰﾄﾞ
    TORI_KBN(0 To 0)        As Byte             '取引先区分
End Type
    
'キー・データ
Public K0_P_SHSHI_SUM       As KEY0_P_SHSHI_SUM

Type P_SHSHI_SUM_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_SHSHI_SUM_Speck   As P_SHSHI_SUM_FSpeck
Private Function P_SHSHI_SUM_Create() As Integer
'********************************************************************
'*
'*              資材仕入集計ﾃﾞｰﾀ    ＣＲＥＡＴＥ
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


    P_SHSHI_SUM_Create = True
                                            '資材仕入集計ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SHSHI_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHSHI_SUM]読み込みエラー")
        Exit Function
    End If

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


'    Ret = InStr(1, Trim(c), ".") - 1
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    
    
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)

    P_SHSHI_SUM_Speck.fs.recoleng = Len(P_SHSHI_SUM_REC)    ' レコード長
    P_SHSHI_SUM_Speck.fs.PageSize = P_SHSHI_SUM_PG_SIZ      ' ページサイズ
    P_SHSHI_SUM_Speck.fs.idexnumb = 1                       ' インデックス数
    P_SHSHI_SUM_Speck.fs.fileflag = 0                       ' ファイルフラグ
    P_SHSHI_SUM_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_SHSHI_SUM_Speck.ks0.keypos = 1                       ' キーポジション
    P_SHSHI_SUM_Speck.ks0.keyleng = 5                      ' キー長
    P_SHSHI_SUM_Speck.ks0.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    P_SHSHI_SUM_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    P_SHSHI_SUM_Speck.ks0.reserve = &H0                    ' 予約済み
    
    
    P_SHSHI_SUM_Speck.ks1.keypos = 6                       ' キーポジション
    P_SHSHI_SUM_Speck.ks1.keyleng = 1                      ' キー長
    P_SHSHI_SUM_Speck.ks1.keyflag = BtKfExt                ' キーフラグ
    P_SHSHI_SUM_Speck.ks1.keytype = Chr(BtKtString)        ' キータイプ
    P_SHSHI_SUM_Speck.ks1.reserve = &H0                    ' 予約済み
    
    '--------------------------------------------------- キー０ △
    
    sts = BTRV(BtOpCreate, P_SHSHI_SUM_POS, P_SHSHI_SUM_Speck, Len(P_SHSHI_SUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "資材仕入集計ﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_SHSHI_SUM_Create = False

End Function

Public Function P_SHSHI_SUM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              資材仕入集計ﾃﾞｰﾀ  ＯＰＥＮ
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

    P_SHSHI_SUM_Open = True
                                            '資材仕入集計ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SHSHI_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHSHI_SUM]読み込みエラー")
        Exit Function
    End If
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


'    Ret = InStr(1, Trim(c), ".") - 1
    
    
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)

    Do
        sts = BTRV(BtOpOpen, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHSHI_SUM_Create()  '資材仕入集計ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "資材仕入集計ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "資材仕入集計ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_SHSHI_SUM_Open = False

End Function

