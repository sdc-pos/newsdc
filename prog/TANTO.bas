Attribute VB_Name = "TANTO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              担当者マスタ  ファイル定義                          *
'*                                                                  *
'*          CREATE 2004.02.14                                       *
'********************************************************************
'ファイルＩＤ
Public Const TANTO_ID$ = "TANTO"

'ページサイズ
Public Const TANTO_PG_SIZ% = 512

'ポジション・ブロック
Public TANTO_POS            As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type TANTOREC_Tag
    TANTO_CODE(0 To 4)      As Byte         '担当者コード
    TANTO_NAME(0 To 19)     As Byte         '担当者名称
    POST_CODE(0 To 1)       As Byte         '部署
    KUBUN(0 To 1)           As Byte         '区分 空白：対象外 2011.09.30
    FILLER(0 To 18)         As Byte         'FILLER 20-->19-->18 2011.09.30
End Type

'データ・バッファ
Public TANTOREC As TANTOREC_Tag

'キー定義

Type KEY0_TANTO                 'ＫＥＹ０
    TANTO_CODE(0 To 4)      As Byte         '担当者コード
End Type

'キー・データ
Public K0_TANTO             As KEY0_TANTO

Type TANTO_FSpeck
    fs  As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Public TANTO_Speck As TANTO_FSpeck
 
Private Function TANTO_Create() As Integer
'********************************************************************
'*                                                                  *
'*              担当者マスタ  ＣＲＥＡＴＥ                      　  *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.02.14                                       *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    TANTO_Create = True
                                            '担当者マスタフルパス取込み
    sts = GetIni("FILE", TANTO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    TANTO_Speck.fs.recoleng = Len(TANTOREC)             ' レコード長
    TANTO_Speck.fs.PageSize = TANTO_PG_SIZ%             ' ページサイズ
    TANTO_Speck.fs.idexnumb = 1                         ' インデックス数
    TANTO_Speck.fs.fileflag = 0                         ' ファイルフラグ
    TANTO_Speck.fs.reserve = &H0                        ' 予約済み
                                                        ' キー０
    TANTO_Speck.ks0.keypos = 1                          ' キーポジション
    TANTO_Speck.ks0.keyleng = 5                         ' キー長
    TANTO_Speck.ks0.keyflag = BtKfExt                   ' キーフラグ
    TANTO_Speck.ks0.keytype = Chr(BtKtString)           ' キータイプ
    TANTO_Speck.ks0.reserve = &H0                       ' 予約済み

    sts = BTRV(BtOpCreate, TANTO_POS, TANTO_Speck, Len(TANTO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "担当者マスタ")
    End If

    TANTO_Create = False

End Function

Function TANTO_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              担当者マスタ  ＯＰＥＮ                          　  *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.02.14                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    TANTO_Open = True
                                            '担当者マスタフルパス取込み
    sts = GetIni("FILE", TANTO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, TANTO_POS, TANTOREC, Len(TANTOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = TANTO_Create()        '担当者マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, TANTO_POS, TANTOREC, Len(TANTOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "担当者マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "担当者マスタ")
                Exit Function
        End Select
    Loop

    TANTO_Open = False
    
End Function
