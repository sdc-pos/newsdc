Attribute VB_Name = "P_SEISAN_SUM"
Option Explicit

'********************************************************************
'*
'*              生産実績集計ﾃﾞｰﾀ  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ
Public Const P_SEISAN_SUM_ID$ = "P_SEISAN_SUM"

'ページサイズ
Private Const P_SEISAN_SUM_PG_SIZ% = 1024

'ポジション・ブロック
Public P_SEISAN_SUM_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
Private Type UCHIWAKE_Tag
    CNT(0 To 4)             As Byte         '外部生産　件数
    SURYO(0 To 10)          As Byte         '外部生産  数量 9(8)V99
                                               
    TANKA(0 To 10)          As Byte         '商品化価格     9(8)V99
    KINGAKU(0 To 9)         As Byte         '商品化金額     9(10)
    
    SH_TANKA(0 To 10)       As Byte         '資材価格       9(8)V99
    SH_KINGAKU(0 To 9)      As Byte         '資材金額       9(10)
    
    KO_TANKA(0 To 10)       As Byte         '工料価格       9(8)V99
    KO_KINGAKU(0 To 9)      As Byte         '工料金額       9(10)
    
    ETC_TANKA(0 To 10)      As Byte         'その他価格     9(8)V99
    ETC_KINGAKU(0 To 9)     As Byte         'その他金額     9(10)


End Type
'レコード定義
Public Type P_SEISAN_SUM_REC_Tag
    
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先
    CLASS_CODE(0 To 19)     As Byte         'クラス（品番）
        
    UCHIWAKE(0 To 1)        As UCHIWAKE_Tag
    
    KO_GENKA(0 To 9)        As Byte         '個装　原価      9(10)
    GA_GENKA(0 To 9)        As Byte         '外装　原価      9(10)
    GK_GENKA(0 To 9)        As Byte         '外注工料        9(10)

    Filler(0 To 3)         As Byte
End Type
'データ・バッファ
Public P_SEISAN_SUM_REC     As P_SEISAN_SUM_REC_Tag

'キー定義
Public Type KEY0_P_SEISAN_SUM               'ＫＥＹ０
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先
    CLASS_CODE(0 To 19)     As Byte         'クラス（品番）
End Type
    
'キー・データ
Public K0_P_SEISAN_SUM      As KEY0_P_SEISAN_SUM

Type P_SEISAN_SUM_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_SEISAN_SUM_Speck  As P_SEISAN_SUM_FSpeck
Private Function P_SEISAN_SUM_Create() As Integer
'********************************************************************
'*
'*              生産実績集計ﾃﾞｰﾀ  ＣＲＥＡＴＥ
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

    P_SEISAN_SUM_Create = True
                                            '生産実績集計ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SEISAN_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_SUM]読み込みエラー")
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




    P_SEISAN_SUM_Speck.fs.recoleng = Len(P_SEISAN_SUM_REC)  ' レコード長
    P_SEISAN_SUM_Speck.fs.PageSize = P_SEISAN_SUM_PG_SIZ    ' ページサイズ
    P_SEISAN_SUM_Speck.fs.idexnumb = 1                      ' インデックス数
    P_SEISAN_SUM_Speck.fs.fileflag = 0                      ' ファイルフラグ
    P_SEISAN_SUM_Speck.fs.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_SEISAN_SUM_Speck.ks0.keypos = 1                       ' キーポジション
    P_SEISAN_SUM_Speck.ks0.keyleng = 2                      ' キー長
    P_SEISAN_SUM_Speck.ks0.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    P_SEISAN_SUM_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    P_SEISAN_SUM_Speck.ks0.reserve = &H0                    ' 予約済み
    
    
    P_SEISAN_SUM_Speck.ks1.keypos = 3                       ' キーポジション
    P_SEISAN_SUM_Speck.ks1.keyleng = 20                     ' キー長
    P_SEISAN_SUM_Speck.ks1.keyflag = BtKfExt                ' キーフラグ
    P_SEISAN_SUM_Speck.ks1.keytype = Chr(BtKtString)        ' キータイプ
    P_SEISAN_SUM_Speck.ks1.reserve = &H0                    ' 予約済み
    
    '--------------------------------------------------- キー０ △
    
    
    
    sts = BTRV(BtOpCreate, P_SEISAN_SUM_POS, P_SEISAN_SUM_Speck, Len(P_SEISAN_SUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "生産実績集計ﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_SEISAN_SUM_Create = False

End Function

Public Function P_SEISAN_SUM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              生産実績集計ﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim sBuffer As String * 255
Dim com     As String

Dim Ret             As Integer


    P_SEISAN_SUM_Open = True
                                            '生産実績集計ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SEISAN_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_SUM]読み込みエラー")
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
        sts = BTRV(BtOpOpen, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SEISAN_SUM_Create() '生産実績集計ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "生産実績集計ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "生産実績集計ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_SEISAN_SUM_Open = False

End Function

