Attribute VB_Name = "KEPPIN"
Option Explicit
'********************************************************************
'*
'*              欠品データ  ファイル定義
'*
'*          CREATE 2013.08.23
'********************************************************************
'ファイルＩＤ
Public Const KEPPIN_ID$ = "KEPPIN"

'ページサイズ
Public Const KEPPIN_PG_SIZ% = 512

'ポジション・ブロック
Public KEPPIN_POS               As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type KEPPINREC_Tag
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
    KEPPIN_CNT(0 To 7)          As Byte     '欠品　件数
    KEPPIN_QTY(0 To 7)          As Byte     '欠品　個数
End Type
'データ・バッファ
Public KEPPINREC                As KEPPINREC_Tag

'キー定義

Type KEY0_KEPPIN                'ＫＥＹ０
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
End Type
'キー・データ
Public K0_KEPPIN                As KEY0_KEPPIN

Type KEPPIN_FSpeck
    fs      As BtFileSpeck                  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private KEPPIN_Speck            As KEPPIN_FSpeck

Private Function KEPPIN_Create() As Integer
'********************************************************************
'*
'*              欠品データ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    KEPPIN_Create = True
                                            '欠品データ フルパス取込み
    sts = GetIni("FILE", KEPPIN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [KEPPIN]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    KEPPIN_Speck.fs.recoleng = Len(KEPPINREC)       ' レコード長
    KEPPIN_Speck.fs.PageSize = KEPPIN_PG_SIZ        ' ページサイズ
    KEPPIN_Speck.fs.idexnumb = 1                    ' インデックス数
    KEPPIN_Speck.fs.fileflag = 0                    ' ファイルフラグ
    KEPPIN_Speck.fs.reserve = &H0                   ' 予約済み
'-----------------------------------------------
                                                ' キー０
    KEPPIN_Speck.ks0.keypos = 1                     ' キーポジション
    KEPPIN_Speck.ks0.keyleng = 20                   ' キー長
                                                    ' キーフラグ
    KEPPIN_Speck.ks0.keyflag = BtKfExt + BtKfChg
    KEPPIN_Speck.ks0.keytype = Chr(BtKtString)      ' キータイプ
    KEPPIN_Speck.ks0.reserve = &H0                  ' 予約済み
'-----------------------------------------------

    sts = BTRV(BtOpCreate, KEPPIN_POS, KEPPIN_Speck, Len(KEPPIN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "欠品データ")
        Exit Function
    End If

    KEPPIN_Create = False

End Function

Public Function KEPPIN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              欠品データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    KEPPIN_Open = True
                                            '欠品データ フルパス取込み
    sts = GetIni("FILE", KEPPIN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [KEPPIN]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, KEPPIN_POS, KEPPINREC, Len(KEPPINREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = KEPPIN_Create()        '欠品データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, KEPPIN_POS, KEPPINREC, Len(KEPPINREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "欠品データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "欠品データ")
                Exit Function
        End Select
    Loop

    KEPPIN_Open = False

End Function

