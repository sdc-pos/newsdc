Attribute VB_Name = "MENU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              メニュー管理マスタ    ファイル定義                  *
'*                                                                  *
'*          CREATE 2004.02.20                                       *
'********************************************************************
'ファイルＩＤ
Public Const MENU_ID$ = "MENU"

'ページサイズ
Public Const MENU_PG_SIZ% = 1024

'ポジション・ブロック
Public MENU_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type MENUREC_Tag
    MENU_GRP_NO(0 To 1)     As Byte         'メニューグループ№
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    MENU_LV1(0 To 2)        As Byte         'メニューレベル１
    MENU_LV2(0 To 2)        As Byte         'メニューレベル２
    MENU_LV3(0 To 2)        As Byte         'メニューレベル３
    MENU_GRP(0 To 19)       As Byte         'メニューグループ
    MENU_KBN(0 To 0)        As Byte         'メニュ―区分
    DISPLAY_ITEM(0 To 19)   As Byte         '表示項目
    CODE_TYPE(0 To 0)       As Byte         '主バーコードタイプ
    YOIN_CODE(0 To 0)       As Byte         '要因
    PARAM(0 To 15)          As Byte         'パラメータ
    FILLER(0 To 23)         As Byte         'FILLER

End Type

'データ・バッファ
Public MENUREC As MENUREC_Tag

'キー定義

Type KEY0_MENU                  'ＫＥＹ０
    MENU_GRP_NO(0 To 1)     As Byte         'メニューグループ№
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    MENU_LV1(0 To 2)        As Byte         'メニューレベル１
    MENU_LV2(0 To 2)        As Byte         'メニューレベル２
    MENU_LV3(0 To 2)        As Byte         'メニューレベル３
End Type

Type KEY1_MENU                  'ＫＥＹ１
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    MENU_GRP_NO(0 To 1)     As Byte         'メニューグループ№
    MENU_LV1(0 To 2)        As Byte         'メニューレベル１
    MENU_LV2(0 To 2)        As Byte         'メニューレベル２
    MENU_LV3(0 To 2)        As Byte         'メニューレベル３
End Type

'キー・データ
Public K0_MENU              As KEY0_MENU
Public K1_MENU              As KEY1_MENU

Type MENU_FSpeck
    fs  As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks7 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks8 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks9 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks10 As BtKeySpeck              ' ｷｰ ｽﾍﾟｯｸ構造体
    ks11 As BtKeySpeck              ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Public MENU_Speck           As MENU_FSpeck
 
Private Function MENU_Create() As Integer
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

    MENU_Create = False
                                            'メニュー管理マスタフルパス取込み
    sts = GetIni("FILE", MENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        MENU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    MENU_Speck.fs.recoleng = Len(MENUREC)               ' レコード長
    MENU_Speck.fs.PageSize = MENU_PG_SIZ%               ' ページサイズ
    MENU_Speck.fs.idexnumb = 2                          ' インデックス数
    MENU_Speck.fs.fileflag = 0                          ' ファイルフラグ
    MENU_Speck.fs.reserve = &H0                         ' 予約済み
'-------------------------------------------------------
                                                        ' キー０
    MENU_Speck.ks0.keypos = 1                           ' キーポジション
    MENU_Speck.ks0.keyleng = 2                          ' キー長
    MENU_Speck.ks0.keyflag = BtKfExt + BtKfSeg          ' キーフラグ
    MENU_Speck.ks0.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks0.reserve = &H0                        ' 予約済み
    
    MENU_Speck.ks1.keypos = 3                           ' キーポジション
    MENU_Speck.ks1.keyleng = 1                          ' キー長
    MENU_Speck.ks1.keyflag = BtKfExt + BtKfSeg          ' キーフラグ
    MENU_Speck.ks1.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks1.reserve = &H0                        ' 予約済み
    
    MENU_Speck.ks2.keypos = 4                           ' キーポジション
    MENU_Speck.ks2.keyleng = 1                          ' キー長
    MENU_Speck.ks2.keyflag = BtKfExt + BtKfSeg          ' キーフラグ
    MENU_Speck.ks2.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks2.reserve = &H0                        ' 予約済み
    
    MENU_Speck.ks3.keypos = 5                           ' キーポジション
    MENU_Speck.ks3.keyleng = 3                          ' キー長
    MENU_Speck.ks3.keyflag = BtKfExt + BtKfSeg          ' キーフラグ
    MENU_Speck.ks3.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks3.reserve = &H0                        ' 予約済み
    
    MENU_Speck.ks4.keypos = 8                           ' キーポジション
    MENU_Speck.ks4.keyleng = 3                          ' キー長
    MENU_Speck.ks4.keyflag = BtKfExt + BtKfSeg          ' キーフラグ
    MENU_Speck.ks4.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks4.reserve = &H0                        ' 予約済み
    
    MENU_Speck.ks5.keypos = 11                          ' キーポジション
    MENU_Speck.ks5.keyleng = 3                          ' キー長
    MENU_Speck.ks5.keyflag = BtKfExt                    ' キーフラグ
    MENU_Speck.ks5.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks5.reserve = &H0                        ' 予約済み
'-------------------------------------------------------
                                                        ' キー０
    MENU_Speck.ks6.keypos = 3                           ' キーポジション
    MENU_Speck.ks6.keyleng = 1                          ' キー長
    MENU_Speck.ks6.keyflag = BtKfExt + BtKfSeg          ' キーフラグ
    MENU_Speck.ks6.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks6.reserve = &H0                        ' 予約済み
    
    MENU_Speck.ks7.keypos = 4                           ' キーポジション
    MENU_Speck.ks7.keyleng = 1                          ' キー長
    MENU_Speck.ks7.keyflag = BtKfExt + BtKfSeg          ' キーフラグ
    MENU_Speck.ks7.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks7.reserve = &H0                        ' 予約済み
    
    MENU_Speck.ks8.keypos = 1                           ' キーポジション
    MENU_Speck.ks8.keyleng = 2                          ' キー長
    MENU_Speck.ks8.keyflag = BtKfExt + BtKfSeg          ' キーフラグ
    MENU_Speck.ks8.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks8.reserve = &H0                        ' 予約済み
    
    MENU_Speck.ks9.keypos = 5                           ' キーポジション
    MENU_Speck.ks9.keyleng = 3                          ' キー長
    MENU_Speck.ks9.keyflag = BtKfExt + BtKfSeg          ' キーフラグ
    MENU_Speck.ks9.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks9.reserve = &H0                        ' 予約済み
    
    MENU_Speck.ks10.keypos = 8                           ' キーポジション
    MENU_Speck.ks10.keyleng = 3                          ' キー長
    MENU_Speck.ks10.keyflag = BtKfExt + BtKfSeg          ' キーフラグ
    MENU_Speck.ks10.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks10.reserve = &H0                        ' 予約済み
    
    MENU_Speck.ks11.keypos = 11                          ' キーポジション
    MENU_Speck.ks11.keyleng = 3                          ' キー長
    MENU_Speck.ks11.keyflag = BtKfExt                    ' キーフラグ
    MENU_Speck.ks11.keytype = Chr(BtKtString)            ' キータイプ
    MENU_Speck.ks11.reserve = &H0                        ' 予約済み
'-------------------------------------------------------

    sts = BTRV(BtOpCreate, MENU_POS, MENU_Speck, Len(MENU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "メニュー管理マスタ")
        MENU_Create = True
    End If

End Function

Function MENU_Open(Mode As Integer) As Integer
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
    
    MENU_Open = False
                                            'メニュー管理マスタフルパス取込み
    sts = GetIni("FILE", MENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        MENU_Open = True
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, MENU_POS, MENUREC, Len(MENUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = MENU_Create()         'メニュー管理マスタ作成
                If sts <> False Then
                    MENU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, MENU_POS, MENUREC, Len(MENUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "メニュー管理マスタ")
                    MENU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "メニュー管理マスタ")
                MENU_Open = True
                Exit Function
        End Select
    Loop
End Function
