Attribute VB_Name = "P_CODE"
Option Explicit
'********************************************************************
'*                                                                  *
'*              商品ラベルコントロール  ファイル定義                *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'ファイルＩＤ
Public Const P_CODE_ID$ = "P_CODE"

'ページサイズ
Private Const P_CODE_PG_SIZ% = 512

'ポジション・ブロック
Public P_CODE_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type P_CODEREC_Tag
    
    DATA_KBN(0 To 1)        As Byte         'ﾃﾞｰﾀ区分
    C_Code(0 To 9)          As Byte         'ｺｰﾄﾞ
    C_NAME(0 To 59)         As Byte         'ﾛﾝｸﾞﾈｰﾑ名称
    C_RNAME(0 To 19)        As Byte         'ｼｮｰﾄﾈｰﾑ名称
    OPTION1(0 To 9)         As Byte         'ｵﾌﾟｼｮﾝ1
    OPTION2(0 To 9)         As Byte         'ｵﾌﾟｼｮﾝ2
    FILLER(0 To 60)         As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_CODEREC           As P_CODEREC_Tag

'キー定義

Type KEY0_P_CODE                            'ＫＥＹ０
    DATA_KBN(0 To 1)        As Byte         'ﾃﾞｰﾀ区分
    C_Code(0 To 9)          As Byte         'ｺｰﾄﾞ
End Type
    
'キー・データ
Public K0_P_CODE            As KEY0_P_CODE

Type P_CODE_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_CODE_Speck        As P_CODE_FSpeck
Private Function P_CODE_Create() As Integer
'********************************************************************
'*                                                                  *
'*              コードマスタ  ＣＲＥＡＴＥ                          *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_CODE_Create = True
                                            'コードマスタフルパス取込み
    sts = GetIni("FILE", P_CODE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_CODE]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_CODE_Speck.fs.recoleng = Len(P_CODEREC)          ' レコード長
    P_CODE_Speck.fs.PageSize = P_CODE_PG_SIZ           ' ページサイズ
    P_CODE_Speck.fs.idexnumb = 1                        ' インデックス数
    P_CODE_Speck.fs.fileflag = 0                        ' ファイルフラグ
    P_CODE_Speck.fs.reserve = &H0                       ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_CODE_Speck.ks0.keypos = 1                         ' キーポジション
    P_CODE_Speck.ks0.keyleng = 2                        ' キー長
    P_CODE_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    P_CODE_Speck.ks0.keytype = Chr(BtKtString)          ' キータイプ
    P_CODE_Speck.ks0.reserve = &H0                      ' 予約済み
    
    P_CODE_Speck.ks1.keypos = 3                         ' キーポジション
    P_CODE_Speck.ks1.keyleng = 10                       ' キー長
    P_CODE_Speck.ks1.keyflag = BtKfExt                  ' キーフラグ
    P_CODE_Speck.ks1.keytype = Chr(BtKtString)          ' キータイプ
    P_CODE_Speck.ks1.reserve = &H0                      ' 予約済み
    
    '--------------------------------------------------- キー０ △
    sts = BTRV(BtOpCreate, P_CODE_POS, P_CODE_Speck, Len(P_CODE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "コードマスタ")
        Exit Function
    End If
    
    P_CODE_Create = False

End Function

Public Function P_CODE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              コードマスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_CODE_Open = True
                                            'コードマスタフルパス取込み
    sts = GetIni("FILE", P_CODE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_CODE]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_CODE_POS, P_CODEREC, Len(P_CODEREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_CODE_Create()      'コードマスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_CODE_POS, P_CODEREC, Len(P_CODEREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "コードマスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "コードマスタ")
                Exit Function
        End Select
    Loop
    
    P_CODE_Open = False

End Function
