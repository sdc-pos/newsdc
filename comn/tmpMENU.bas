Attribute VB_Name = "tmpMENU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              メニュー管理マスタ（一時ファイル）    ファイル定義  *
'*                                                                  *
'*          CREATE 2004.02.26                                       *
'********************************************************************
'ファイルＩＤ
Public Const tmpMENU_ID$ = "tmpMENU"

'ページサイズ
Public Const tmpMENU_PG_SIZ% = 512

'ポジション・ブロック
Public tmpMENU_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type tmpMENUREC_Tag
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    MENU_LV1(0 To 2)        As Byte         'メニューレベル１
    MENU_LV2(0 To 2)        As Byte         'メニューレベル２
    MENU_LV3(0 To 2)        As Byte         'メニューレベル３
    DEL_FLG(0 To 0)         As Byte         '削除フラグ
    MENU_KBN(0 To 0)        As Byte         'メニュ―区分
    DISPLAY_ITEM(0 To 19)   As Byte         '表示項目
    CODE_TYPE(0 To 0)       As Byte         '主バーコードタイプ
    YOIN_CODE(0 To 0)       As Byte         '要因
    PARAM_F(0 To 0)         As Byte         '付加ﾊﾟﾗﾒｰﾀ(0:なし 1:向け先)
    PARAM(0 To 15)          As Byte         'パラメータ

End Type

'データ・バッファ
Public tmpMENUREC As tmpMENUREC_Tag

'キー定義

Type KEY0_tmpMENU               'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    MENU_LV1(0 To 2)        As Byte         'メニューレベル１
    MENU_LV2(0 To 2)        As Byte         'メニューレベル２
    MENU_LV3(0 To 2)        As Byte         'メニューレベル３
End Type

'キー・データ
Public K0_tmpMENU           As KEY0_tmpMENU

Type tmpMENU_FSpeck
    fs  As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Public tmpMENU_Speck        As tmpMENU_FSpeck
 
Private Function tmpMENU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              メニュー管理マスタ  ＣＲＥＡＴＥ                    *
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

    tmpMENU_Create = False
                                            'メニュー管理マスタフルパス取込み
    sts = GetIni("FILE", tmpMENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        tmpMENU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    tmpMENU_Speck.fs.recoleng = Len(tmpMENUREC)         ' レコード長
    tmpMENU_Speck.fs.PageSize = tmpMENU_PG_SIZ%         ' ページサイズ
    tmpMENU_Speck.fs.idexnumb = 1                       ' インデックス数
    tmpMENU_Speck.fs.fileflag = 0                       ' ファイルフラグ
    tmpMENU_Speck.fs.reserve = &H0                      ' 予約済み
'-------------------------------------------------------
                                                        ' キー０
    tmpMENU_Speck.ks0.keypos = 1                        ' キーポジション
    tmpMENU_Speck.ks0.keyleng = 1                       ' キー長
    tmpMENU_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    tmpMENU_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    tmpMENU_Speck.ks0.reserve = &H0                     ' 予約済み
    
    tmpMENU_Speck.ks1.keypos = 2                        ' キーポジション
    tmpMENU_Speck.ks1.keyleng = 1                       ' キー長
    tmpMENU_Speck.ks1.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    tmpMENU_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    tmpMENU_Speck.ks1.reserve = &H0                     ' 予約済み
    
    tmpMENU_Speck.ks2.keypos = 3                        ' キーポジション
    tmpMENU_Speck.ks2.keyleng = 3                       ' キー長
    tmpMENU_Speck.ks2.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    tmpMENU_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    tmpMENU_Speck.ks2.reserve = &H0                     ' 予約済み
    
    tmpMENU_Speck.ks3.keypos = 6                        ' キーポジション
    tmpMENU_Speck.ks3.keyleng = 3                       ' キー長
    tmpMENU_Speck.ks3.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    tmpMENU_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    tmpMENU_Speck.ks3.reserve = &H0                     ' 予約済み
    
    tmpMENU_Speck.ks4.keypos = 9                        ' キーポジション
    tmpMENU_Speck.ks4.keyleng = 3                       ' キー長
    tmpMENU_Speck.ks4.keyflag = BtKfExt                 ' キーフラグ
    tmpMENU_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    tmpMENU_Speck.ks4.reserve = &H0                     ' 予約済み
'-------------------------------------------------------


    sts = BTRV(BtOpCreate, tmpMENU_POS, tmpMENU_Speck, Len(tmpMENU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "メニュー管理マスタ(一時ファイル)")
        tmpMENU_Create = True
    End If

End Function

Function tmpMENU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              メニュー管理マスタ  ＯＰＥＮ                        *
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
    
    tmpMENU_Open = False
                                            'メニュー管理マスタフルパス取込み
    sts = GetIni("FILE", tmpMENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        tmpMENU_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpMENU_Create()      'メニュー管理マスタ作成
                If sts <> False Then
                    tmpMENU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "メニュー管理マスタ（一時ファイル）")
                    tmpMENU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "メニュー管理マスタ（一時ファイル）")
                tmpMENU_Open = True
                Exit Function
        End Select
    Loop
End Function
