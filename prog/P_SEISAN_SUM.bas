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
Private Type UCHIWAKE_TBL_Tag
    NAI_TANKA(0 To 10)      As Byte         '内部　単価
    GAI_TANKA(0 To 10)      As Byte         '外部　単価
End Type

'レコード定義
Public Type P_SEISAN_SUM_REC_Tag
    
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先
    CLASS_CODE(0 To 19)     As Byte         'クラス（品番）
        
    GK_NAI_CNT(0 To 4)      As Byte         '内部生産　件数
    GK_NAI_SURYO(0 To 10)   As Byte         '内部生産  数量
    GK_GAI_CNT(0 To 4)      As Byte         '外部生産　件数
    GK_GAI_SURYO(0 To 10)   As Byte         '外部生産  数量
                                               
    GK_TANKA(0 To 10)       As Byte         '合計単価
    
                                            '生産　内訳
    UCHIWAKE_TBL(0 To 2)    As UCHIWAKE_TBL_Tag

    KO_GENKA(0 To 10)       As Byte         '個装　原価
    GA_GENKA(0 To 10)       As Byte         '外装　原価
    GK_GENKA(0 To 10)       As Byte         '外注工料

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

    P_SEISAN_SUM_Create = True
                                            '生産実績集計ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SEISAN_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_SUM]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

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

    P_SEISAN_SUM_Open = True
                                            '生産実績集計ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SEISAN_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_SUM]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

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

