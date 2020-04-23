Attribute VB_Name = "ITEM_CHG"
Option Explicit
'********************************************************************
'*
'*              品目読み替え  ファイル定義
'*
'*          CREATE 2018.02.03
'********************************************************************
'ファイルＩＤ
Public Const ITEM_CHG_ID$ = "ITEM_CHG"

'ページサイズ
Public Const ITEM_CHG_PG_SIZ% = 4096

'ポジション・ブロック
Public ITEM_CHG_POS             As POSBLK
'=
'=
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************

'レコード定義
Type ITEM_CHG_REC_Tag
    N_JGYOBU(0 To 0)            As Byte     '新　事業部区分
    N_NAIGAI(0 To 0)            As Byte     '新　国内外
    N_HIN_GAI(0 To 19)          As Byte     '新　品番（外部）
    HIN_NAME(0 To 39)           As Byte     '品名
    O_HIN_GAI(0 To 39)          As Byte     '旧　品番（外部）（備考）

End Type
'データ・バッファ
Public ITEM_CHG_REC As ITEM_CHG_REC_Tag

'キー定義

Type KEY0_ITEM_CHG            'ＫＥＹ０
    N_JGYOBU(0 To 0)            As Byte     '新　事業部区分
    N_NAIGAI(0 To 0)            As Byte     '新　国内外
    N_HIN_GAI(0 To 19)          As Byte     '新　品番（外部）
End Type




'キー・データ
Public K0_ITEM_CHG  As KEY0_ITEM_CHG

Type ITEM_CHG_FSpeck
    fs      As BtFileSpeck                 ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                 ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
End Type

Private ITEM_CHG_Speck  As ITEM_CHG_FSpeck

Private Function ITEM_CHG_Create() As Integer
'********************************************************************
'*
'*              品目読み替え  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ITEM_CHG_Create = True
                                            '品目マスタフルパス取込み
    sts = GetIni("FILE", ITEM_CHG_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_CHG]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    ITEM_CHG_Speck.fs.recoleng = Len(ITEM_CHG_REC)  ' レコード長
    ITEM_CHG_Speck.fs.PageSize = ITEM_PG_SIZ       ' ページサイズ
    ITEM_CHG_Speck.fs.idexnumb = 1                 ' インデックス数
    ITEM_CHG_Speck.fs.fileflag = 0                 ' ファイルフラグ
    ITEM_CHG_Speck.fs.reserve = &H0                ' 予約済み
'-----------------------------------------------
                                                ' キー０
    ITEM_CHG_Speck.ks0.keypos = 1                              ' キーポジション
    ITEM_CHG_Speck.ks0.keyleng = 1                             ' キー長
    ITEM_CHG_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' キーフラグ
    ITEM_CHG_Speck.ks0.keytype = Chr(BtKtString)               ' キータイプ
    ITEM_CHG_Speck.ks0.reserve = &H0                           ' 予約済み

    ITEM_CHG_Speck.ks1.keypos = 2                              ' キーポジション
    ITEM_CHG_Speck.ks1.keyleng = 1                             ' キー長
    ITEM_CHG_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' キーフラグ
    ITEM_CHG_Speck.ks1.keytype = Chr(BtKtString)               ' キータイプ
    ITEM_CHG_Speck.ks1.reserve = &H0                           ' 予約済み

    ITEM_CHG_Speck.ks2.keypos = 3                              ' キーポジション
    ITEM_CHG_Speck.ks2.keyleng = 20                            ' キー長
    ITEM_CHG_Speck.ks2.keyflag = BtKfExt + BtKfChg             ' キーフラグ
    ITEM_CHG_Speck.ks2.keytype = Chr(BtKtString)               ' キータイプ
    ITEM_CHG_Speck.ks2.reserve = &H0                           ' 予約済み
'-----------------------------------------------





    sts = BTRV(BtOpCreate, ITEM_CHG_POS, ITEM_CHG_Speck, Len(ITEM_CHG_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "品目読み替え")
        Exit Function
    End If

    ITEM_CHG_Create = False

End Function

Public Function ITEM_CHG_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              品目読み替え  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_CHG_Open = True
                                            '品目マスタフルパス取込み
    sts = GetIni("FILE", ITEM_CHG_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_CHG]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_CHG_POS, ITEM_CHG_REC, Len(ITEM_CHG_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_CHG_Create()        '品目マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_CHG_POS, ITEM_CHG_REC, Len(ITEM_CHG_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "品目読み替え")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "品目読み替え")
                Exit Function
        End Select
    Loop

    ITEM_CHG_Open = False

End Function

