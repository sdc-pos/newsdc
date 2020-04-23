Attribute VB_Name = "P_TANTOMENU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              新担当者別メニュー  ファイル定義                    *
'*                                                                  *
'*          CREATE 2004.02.20                                       *
'********************************************************************
'ファイルＩＤ
Public Const P_TMENU_ID$ = "P_TANTOMENU"

'ページサイズ
Public Const P_TMENU_PG_SIZ% = 2048

'ポジション・ブロック
Public P_TMENU_POS            As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義

Private Type MENU_NO_Tag
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    MENU_NO(0 To 1)         As Byte         'ﾒﾆｭｰ№

End Type

Type P_TMENUREC_Tag
    TANTO_CODE(0 To 4)      As Byte         '担当者コード
    
    MENU_T(0 To 179)        As MENU_NO_Tag  'ﾒﾆｭｰ№     29-->179 2006.10.11
    
    FILLER(0 To 298)        As Byte         'FILLER

End Type

'データ・バッファ
Public P_TMENUREC           As P_TMENUREC_Tag

'キー定義

Type KEY0_P_TMENU                           'ＫＥＹ０
    TANTO_CODE(0 To 4)      As Byte         '担当者コード
End Type

'キー・データ
Public K0_P_TMENU           As KEY0_P_TMENU

Type P_TMENU_FSpeck
    fs  As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_TMENU_Speck       As P_TMENU_FSpeck
 
Private Function P_TMENU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              新担当者別メニュー  ＣＲＥＡＴＥ                    *
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

    P_TMENU_Create = False
                                            '担当者別メニューフルパス取込み
    sts = GetIni("FILE", P_TMENU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_TMENU_ID]読み込みエラー")
        P_TMENU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    P_TMENU_Speck.fs.recoleng = Len(P_TMENUREC)         ' レコード長
    P_TMENU_Speck.fs.PageSize = P_TMENU_PG_SIZ%         ' ページサイズ
    P_TMENU_Speck.fs.idexnumb = 1                       ' インデックス数
    P_TMENU_Speck.fs.fileflag = 0                       ' ファイルフラグ
    P_TMENU_Speck.fs.reserve = &H0                      ' 予約済み
'--------------------------------------------------------
                                                        ' キー０
    P_TMENU_Speck.ks0.keypos = 1                        ' キーポジション
    P_TMENU_Speck.ks0.keyleng = 5                       ' キー長
    P_TMENU_Speck.ks0.keyflag = BtKfExt                 ' キーフラグ
    P_TMENU_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    P_TMENU_Speck.ks0.reserve = &H0                     ' 予約済み
    
'--------------------------------------------------------

    sts = BTRV(BtOpCreate, P_TMENU_POS, P_TMENU_Speck, Len(P_TMENU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "新担当者別メニュー")
        P_TMENU_Create = True
    End If

End Function

Function P_TMENU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              新担当者別メニュー  ＯＰＥＮ                        *
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
    
    P_TMENU_Open = False
                                            '担当者別メニューフルパス取込み
    sts = GetIni("FILE", P_TMENU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_TMENU_ID]読み込みエラー")
        P_TMENU_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_TMENU_Create()        '担当者別メニュー作成
                If sts <> False Then
                    P_TMENU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "新メニュー管理マスタ")
                    P_TMENU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "新メニュー管理マスタ")
                P_TMENU_Open = True
                Exit Function
        End Select
    Loop
End Function
