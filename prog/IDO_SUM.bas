Attribute VB_Name = "IDO_SUM"
Option Explicit
'********************************************************************
'*
'*              在庫移動歴集計ファイル  ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const IDO_SUM_ID$ = "IDO_SUM"

'ページサイズ
Public Const IDO_SUM_PG_SIZ% = 4096

'ポジション・ブロック
Public IDO_SUM_POS    As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type IDO_SUMREC_Tag
    JGYOBU(0 To 0)              As Byte     '事業部
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番外部
    ZAIKO_QTY(0 To 7)           As Byte     '在庫数
    LAST_DATE(0 To 7)           As Byte     '最終実績日付
    LAST_TIME(0 To 5)           As Byte     '最終実績時刻
    
    J_PLUS_CNT(0 To 7)          As Byte     '在庫+
    J_MAINA_CNT(0 To 7)         As Byte     '在庫-
    J_SYUKA_CNT(0 To 7)         As Byte     '出荷
    J_IDO_CNT(0 To 7)           As Byte     '移動
    FILLER(0 To 51)             As Byte     'FILLER
End Type

'データ・バッファ
Public IDO_SUMREC               As IDO_SUMREC_Tag

'キー定義
Type KEY0_IDO_SUM               'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番外部
End Type


'キー・データ
Public K0_IDO_SUM               As KEY0_IDO_SUM

Type IDO_SUM_FSpeck
    fs      As BtFileSpeck                  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private IDO_SUM_Speck As IDO_SUM_FSpeck

Private Function IDO_SUM_Create() As Integer
'********************************************************************
'*
'*              在庫移動歴集計ファイル  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    IDO_SUM_Create = True
                                            '在庫移動歴集計ファイルフルパス取込み
    sts = GetIni("FILE", IDO_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [IDO_SUM]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    IDO_SUM_Speck.fs.recoleng = Len(Y_SYUREC)       ' レコード長
    IDO_SUM_Speck.fs.PageSize = IDO_SUM_PG_SIZ      ' ページサイズ
    IDO_SUM_Speck.fs.idexnumb = 1                   ' インデックス数
    IDO_SUM_Speck.fs.fileflag = 0                   ' ファイルフラグ
    IDO_SUM_Speck.fs.reserve = &H0                  ' 予約済み
'---------------------------------------------------' キー０
    IDO_SUM_Speck.ks0.keypos = 1                    ' キーポジション
    IDO_SUM_Speck.ks0.keyleng = 1                   ' キー長
                                                    ' キーフラグ
    IDO_SUM_Speck.ks0.keyflag = BtKfExt + BtKfSeg
    IDO_SUM_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    IDO_SUM_Speck.ks0.reserve = &H0                 ' 予約済み
    
    IDO_SUM_Speck.ks1.keypos = 2                    ' キーポジション
    IDO_SUM_Speck.ks1.keyleng = 1                   ' キー長
    IDO_SUM_Speck.ks1.keyflag = BtKfExt + BtKfSeg   ' キーフラグ
    IDO_SUM_Speck.ks1.keytype = Chr(BtKtString)     ' キータイプ
    IDO_SUM_Speck.ks1.reserve = &H0                 ' 予約済み
    
    IDO_SUM_Speck.ks2.keypos = 3                    ' キーポジション
    IDO_SUM_Speck.ks2.keyleng = 20                  ' キー長
    IDO_SUM_Speck.ks2.keyflag = BtKfExt             ' キーフラグ
    IDO_SUM_Speck.ks2.keytype = Chr(BtKtString)     ' キータイプ
    IDO_SUM_Speck.ks2.reserve = &H0                 ' 予約済み

'---------------------------------------------------' キー０
    
    sts = BTRV(BtOpCreate, IDO_SUM_POS, IDO_SUM_Speck, Len(IDO_SUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "在庫移動歴集計ファイル")
        Exit Function
    End If

    IDO_SUM_Create = False

End Function

Function IDO_SUM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              在庫移動歴集計ファイル  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    IDO_SUM_Open = True
                                            '在庫移動歴集計ファイルフルパス取込み
    sts = GetIni("FILE", IDO_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [IDO_SUM]読み込みエラー ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, IDO_SUM_POS, IDO_SUMREC, Len(IDO_SUMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = IDO_SUM_Create()      '在庫移動歴集計ファイル作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "在庫移動歴集計ファイル")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "在庫移動歴集計ファイル")
                Exit Function
        End Select
    Loop
    Y_SYU_Open = False
End Function
