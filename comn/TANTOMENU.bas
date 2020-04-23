Attribute VB_Name = "TANTOMENU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              担当者別メニュー  ファイル定義                      *
'*                                                                  *
'*          CREATE 2004.02.20                                       *
'********************************************************************
'ファイルＩＤ
Public Const TMENU_ID$ = "TANTOMENU"

'ページサイズ
Public Const TMENU_PG_SIZ% = 512

'ポジション・ブロック
Public TMENU_POS            As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type TMENUREC_Tag
    
    TANTO_CODE(0 To 4)      As Byte         '担当者コード
    MENU_GRP_NO(0 To 1)     As Byte         'メニューグループ
    FILLER(0 To 16)         As Byte         'FILLER

End Type

'データ・バッファ
Public TMENUREC             As TMENUREC_Tag

'キー定義

Type KEY0_TMENU                         'ＫＥＹ０
    TANTO_CODE(0 To 4)      As Byte         '担当者コード
End Type

'キー・データ
Public K0_TMENU             As KEY0_TMENU

Type TMENU_FSpeck
    fs  As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private TMENU_Speck          As TMENU_FSpeck
 
Private Function TMENU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              担当者別メニュー  ＣＲＥＡＴＥ                      *
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

    TMENU_Create = False
                                            '担当者別メニューフルパス取込み
    sts = GetIni("FILE", TMENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        TMENU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    TMENU_Speck.fs.recoleng = Len(TMENUREC)             ' レコード長
    TMENU_Speck.fs.PageSize = TMENU_PG_SIZ%             ' ページサイズ
    TMENU_Speck.fs.idexnumb = 1                         ' インデックス数
    TMENU_Speck.fs.fileflag = 0                         ' ファイルフラグ
    TMENU_Speck.fs.reserve = &H0                        ' 予約済み
'--------------------------------------------------------
                                                        ' キー０
    TMENU_Speck.ks0.keypos = 1                          ' キーポジション
    TMENU_Speck.ks0.keyleng = 5                         ' キー長
    TMENU_Speck.ks0.keyflag = BtKfExt                   ' キーフラグ
    TMENU_Speck.ks0.keytype = Chr(BtKtString)           ' キータイプ
    TMENU_Speck.ks0.reserve = &H0                       ' 予約済み
    
'--------------------------------------------------------

    sts = BTRV(BtOpCreate, TMENU_POS, TMENU_Speck, Len(TMENU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "担当者別メニュー")
        TMENU_Create = True
    End If

End Function

Function TMENU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              担当者別メニュー  ＯＰＥＮ                          *
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
    
    TMENU_Open = False
                                            '担当者別メニューフルパス取込み
    sts = GetIni("FILE", TMENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        TMENU_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, TMENU_POS, TMENUREC, Len(TMENUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = TMENU_Create()        '担当者別メニュー作成
                If sts <> False Then
                    TMENU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, TMENU_POS, TMENUREC, Len(TMENUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "メニュー管理マスタ")
                    TMENU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "メニュー管理マスタ")
                TMENU_Open = True
                Exit Function
        End Select
    Loop
End Function
