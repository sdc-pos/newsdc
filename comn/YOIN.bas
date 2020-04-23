Attribute VB_Name = "YOIN"
Option Explicit
'********************************************************************
'*
'*              要因マスタ  ファイル定義
'*
'*          CREATE 2001.02.14
'********************************************************************
'ファイルＩＤ
Public Const YOIN_ID$ = "YOIN"

'ページサイズ
Public Const YOIN_PG_SIZ% = 512

'ポジション・ブロック
Public YOIN_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type YOINREC_Tag
    CODE_TYPE(0 To 0)       As Byte     '主バーコードタイプ
    YOIN_CODE(0 To 0)       As Byte     '要因
    YOIN_DNAME(0 To 9)      As Byte     '作業表示略称
    SUM_KBN(0 To 0)         As Byte     '集計区分
    SYSTEM_F(0 To 0)        As Byte     'システム予約ﾌﾗｸﾞ2004.02
    REGI_F(0 To 0)          As Byte     '登録可否ﾌﾗｸﾞ
    PARAM_F(0 To 0)         As Byte     '付加ﾊﾟﾗﾒｰﾀ(0:なし 1:向け先 2:倉庫)
    Soko_No(0 To 1)         As Byte     '倉庫№（仮想）
    DSP_No(0 To 1)          As Byte     '表示順　2007.12.10
    FILLER(0 To 3)          As Byte
End Type

'データ・バッファ
Public YOINREC As YOINREC_Tag

'キー定義

Type KEY0_YOIN                 'ＫＥＹ０
    CODE_TYPE(0 To 0)       As Byte     '主バーコードタイプ
    YOIN_CODE(0 To 0)       As Byte     '要因
End Type

Type KEY1_YOIN                 'ＫＥＹ０    2007.12.10
    DSP_No(0 To 1)          As Byte     '表示順　2007.12.10
    CODE_TYPE(0 To 0)       As Byte     '主バーコードタイプ
End Type

'キー・データ
Public K0_YOIN As KEY0_YOIN
Public K1_YOIN As KEY1_YOIN             '2007.12.10

Type YOIN_FSpeck
    fs          As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0         As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1         As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2         As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体    2007.12.10
    ks3         As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体    2007.12.10
End Type

Private YOIN_Speck As YOIN_FSpeck

Private Function YOIN_Create() As Integer
'********************************************************************
'*
'*              要因マスタ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2001.02.14
'*          UPDATE 2007.12.10   KEY1追加
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    YOIN_Create = True
                                            '要因マスタフルパス取込み
    sts = GetIni("FILE", YOIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    YOIN_Speck.fs.recoleng = Len(YOINREC)           ' レコード長
    YOIN_Speck.fs.PageSize = YOIN_PG_SIZ            ' ページサイズ
    YOIN_Speck.fs.idexnumb = 2                      ' インデックス数    2007.12.10
    YOIN_Speck.fs.fileflag = 0                      ' ファイルフラグ
    YOIN_Speck.fs.reserve = &H0                     ' 予約済み
'----------------------------------------------------
                                                    ' キー０
    YOIN_Speck.ks0.keypos = 1                       ' キーポジション
    YOIN_Speck.ks0.keyleng = 1                      ' キー長
    YOIN_Speck.ks0.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    YOIN_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    YOIN_Speck.ks0.reserve = &H0                    ' 予約済み
                                                    ' キー０
    YOIN_Speck.ks1.keypos = 2                       ' キーポジション
    YOIN_Speck.ks1.keyleng = 1                      ' キー長
    YOIN_Speck.ks1.keyflag = BtKfExt                ' キーフラグ
    YOIN_Speck.ks1.keytype = Chr(BtKtString)        ' キータイプ
    YOIN_Speck.ks1.reserve = &H0                    ' 予約済み
    
'----------------------------------------------------
    
'----------------------------------------------------   2007.12.10
                                                    ' キー１
    YOIN_Speck.ks2.keypos = 19                      ' キーポジション
    YOIN_Speck.ks2.keyleng = 2                      ' キー長
                                                    ' キーフラグ
    YOIN_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    YOIN_Speck.ks2.keytype = Chr(BtKtString)        ' キータイプ
    YOIN_Speck.ks2.reserve = &H0                    ' 予約済み
                                                    ' キー１
    YOIN_Speck.ks3.keypos = 1                       ' キーポジション
    YOIN_Speck.ks3.keyleng = 1                      ' キー長
    YOIN_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    YOIN_Speck.ks3.keytype = Chr(BtKtString)        ' キータイプ
    YOIN_Speck.ks3.reserve = &H0                    ' 予約済み
    
'----------------------------------------------------   2007.12.10
    
    
    
    
    sts = BTRV(BtOpCreate, YOIN_POS, YOIN_Speck, Len(YOIN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "要因マスタ")
        Exit Function
    End If
    YOIN_Create = False
End Function

Function YOIN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              要因マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2001.02.14
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    YOIN_Open = True
                                            '要因マスタフルパス取込み
    sts = GetIni("FILE", YOIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, YOIN_POS, YOINREC, Len(YOINREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = YOIN_Create()        '要因マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, YOIN_POS, YOINREC, Len(YOINREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "要因マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "要因マスタ")
                Exit Function
        End Select
    Loop
    
    YOIN_Open = False

End Function



