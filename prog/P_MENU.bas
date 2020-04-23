Attribute VB_Name = "P_MENU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              新メニュー管理マスタ    ファイル定義                *
'*                                                                  *
'*          CREATE 2004.02.20                                       *
'********************************************************************
'ファイルＩＤ
Public Const P_MENU_ID$ = "P_MENU"

'ページサイズ
Public Const P_MENU_PG_SIZ% = 1024

'ポジション・ブロック
Public P_MENU_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義

Private Type SAGYO_Tag
    YOIN(0 To 1)            As Byte         '要因
    PARAM(0 To 15)          As Byte         'ﾊﾟﾗﾒｰﾀ(向け先)
    Disp(0 To 19)           As Byte         'ﾊﾟﾗﾒｰﾀ(向け先)
    LOG_OUT(0 To 0)         As Byte         'ﾛｸﾞ出力 0:出力なし 1:あり
End Type


Type P_MENUREC_Tag
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    MENU_NO(0 To 1)         As Byte         'メニューグループ№
    MENU_DSP(0 To 19)       As Byte         '表示内容
    SAGYO(0 To 19)          As SAGYO_Tag    '作業内容
    FILLER(0 To 175)        As Byte         '作業内容
End Type

'データ・バッファ
Public P_MENUREC            As P_MENUREC_Tag

'キー定義

Type KEY0_P_MENU                'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    MENU_NO(0 To 1)         As Byte         'メニューグループ№
End Type
'キー・データ
Public K0_P_MENU            As KEY0_P_MENU

Type P_MENU_FSpeck
    fs  As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Public P_MENU_Speck         As P_MENU_FSpeck
 
Private Function P_MENU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              新メニュー管理マスタ  ＣＲＥＡＴＥ                    *
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

    P_MENU_Create = False
                                            'メニュー管理マスタフルパス取込み
    sts = GetIni("FILE", P_MENU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        P_MENU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    P_MENU_Speck.fs.recoleng = Len(P_MENUREC)           ' レコード長
    P_MENU_Speck.fs.PageSize = P_MENU_PG_SIZ%           ' ページサイズ
    P_MENU_Speck.fs.idexnumb = 1                        ' インデックス数
    P_MENU_Speck.fs.fileflag = 0                        ' ファイルフラグ
    P_MENU_Speck.fs.reserve = &H0                       ' 予約済み
'-------------------------------------------------------
                                                        ' キー０
    P_MENU_Speck.ks0.keypos = 1                         ' キーポジション
    P_MENU_Speck.ks0.keyleng = 1                        ' キー長
    P_MENU_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    P_MENU_Speck.ks0.keytype = Chr(BtKtString)          ' キータイプ
    P_MENU_Speck.ks0.reserve = &H0                      ' 予約済み
    
    P_MENU_Speck.ks1.keypos = 2                         ' キーポジション
    P_MENU_Speck.ks1.keyleng = 1                        ' キー長
    P_MENU_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    P_MENU_Speck.ks1.keytype = Chr(BtKtString)          ' キータイプ
    P_MENU_Speck.ks1.reserve = &H0                      ' 予約済み
    
    P_MENU_Speck.ks2.keypos = 3                         ' キーポジション
    P_MENU_Speck.ks2.keyleng = 2                        ' キー長
    P_MENU_Speck.ks2.keyflag = BtKfExt                  ' キーフラグ
    P_MENU_Speck.ks2.keytype = Chr(BtKtString)          ' キータイプ
    P_MENU_Speck.ks2.reserve = &H0                      ' 予約済み
    
'-------------------------------------------------------

    sts = BTRV(BtOpCreate, P_MENU_POS, P_MENU_Speck, Len(P_MENU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "メニュー管理マスタ")
        P_MENU_Create = True
    End If

End Function

Function P_MENU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              新メニュー管理マスタ  ＯＰＥＮ                      *
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
    
    P_MENU_Open = False
                                            'メニュー管理マスタフルパス取込み
    sts = GetIni("FILE", P_MENU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        P_MENU_Open = True
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, P_MENU_POS, P_MENUREC, Len(P_MENUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_MENU_Create()         'メニュー管理マスタ作成
                If sts <> False Then
                    P_MENU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_MENU_POS, P_MENUREC, Len(P_MENUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "新メニュー管理マスタ")
                    P_MENU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "新メニュー管理マスタ")
                P_MENU_Open = True
                Exit Function
        End Select
    Loop
End Function
