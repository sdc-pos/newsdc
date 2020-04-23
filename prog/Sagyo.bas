Attribute VB_Name = "SAGYO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              作業管理マスタ  ファイル定義                        *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'*          UPDATE 2001.02.14  作業正式名称の削除
'*                             向け先表示項目の変更
'*                             国内外有無，向け先有無の削除
'********************************************************************
'ファイルＩＤ
Global Const SAGYO_ID = "SAGYO"

'ページサイズ
Global Const SAGYO_PG_SIZ% = 512

'ポジション・ブロック
Global SAGYO_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type SAGYOREC_Tag
    BAR_TYPE(0 To 2)    As Byte     '主バーコード体系
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    PARM(0 To 3)        As Byte     'パラメータ
    SAGYO_DNAME(0 To 15) As Byte    '表示略称
    LCD1_TYPE(0 To 0)   As Byte     'LCD1行目制御
    LCD2_TYPE(0 To 0)   As Byte     'LCD2行目制御
    LCD3_TYPE(0 To 0)   As Byte     'LCD3行目制御
    LCD4_TYPE(0 To 0)   As Byte     'LCD4行目制御
    LCD2_DSP(0 To 15)   As Byte     'LCD2行目表示内容
    LCD3_DSP(0 To 15)   As Byte     'LCD3行目表示内容
    LCD4_DSP(0 To 15)   As Byte     'LCD4行目表示内容
    LOCK_F(0 To 0)      As Byte     '排他フラグ
    FILLER(0 To 3)      As Byte     'FILLER
End Type
'データ・バッファ
Global SAGYOREC As SAGYOREC_Tag

'キー定義

Type KEY0_SAGYO            'ＫＥＹ０
    BAR_TYPE(0 To 2)    As Byte     '主バーコード体系
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    PARM(0 To 3)        As Byte     'パラメータ
End Type

Type KEY0_SAGY1            'ＫＥＹ１
    JGYOBU(0 To 0)      As Byte     '事業部区分
    BAR_TYPE(0 To 2)    As Byte     '主バーコード体系
    NAIGAI(0 To 0)      As Byte     '国内外
    PARM(0 To 3)        As Byte     'パラメータ
End Type

'キー・データ
Global K0_SAGYO As KEY0_SAGYO
Global K1_SAGYO As KEY0_SAGY1

Type SAGYO_FSpeck
    fs As BtFileSpeck               ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck
    ks2 As BtKeySpeck
    ks3 As BtKeySpeck
End Type

Global SAGYO_Speck As SAGYO_FSpeck
 

Private Function SAGYO_Create() As Integer
'********************************************************************
'*                                                                  *
'*              作業管理マスタ  ＣＲＥＡＴＥ                        *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim FullPath As String
Dim c As String * 128

    SAGYO_Create = False
                                            '作業管理マスタフルパス取込み
    sts = GetIni("FILE", SAGYO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        SAGYO_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    SAGYO_Speck.fs.recoleng = Len(SAGYOREC)     ' レコード長
    SAGYO_Speck.fs.PageSize = SAGYO_PG_SIZ      ' ページサイズ
    SAGYO_Speck.fs.idexnumb = 2                 ' インデックス数
    SAGYO_Speck.fs.fileflag = 0                 ' ファイルフラグ
    SAGYO_Speck.fs.reserve = &H0                ' 予約済み
                                                ' キー０
    SAGYO_Speck.ks0.keypos = 1                  ' キーポジション
                                                ' キー長
    SAGYO_Speck.ks0.keyleng = 3 + 1 + 1 + 4
    SAGYO_Speck.ks0.keyflag = BtKfExt           ' キーフラグ
    SAGYO_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    SAGYO_Speck.ks0.reserve = &H0               ' 予約済み
                                                ' キー１
    SAGYO_Speck.ks1.keypos = 4                  ' キーポジション
    SAGYO_Speck.ks1.keyleng = 1                 ' キー長
    SAGYO_Speck.ks1.keyflag = BtKfSeg 'BtKfExt           ' キーフラグ
    SAGYO_Speck.ks1.keytype = Chr(BtKtString)   ' キータイプ
    SAGYO_Speck.ks1.reserve = &H0               ' 予約済み
    SAGYO_Speck.ks2.keypos = 1                  ' キーポジション
    SAGYO_Speck.ks2.keyleng = 3                 ' キー長
    SAGYO_Speck.ks2.keyflag = BtKfSeg 'BtKfExt           ' キーフラグ
    SAGYO_Speck.ks2.keytype = Chr(BtKtString)   ' キータイプ
    SAGYO_Speck.ks2.reserve = &H0               ' 予約済み
    SAGYO_Speck.ks3.keypos = 5                  ' キーポジション
    SAGYO_Speck.ks3.keyleng = 5                 ' キー長
    SAGYO_Speck.ks3.keyflag = BtKfExt           ' キーフラグ
    SAGYO_Speck.ks3.keytype = Chr(BtKtString)   ' キータイプ
    SAGYO_Speck.ks3.reserve = &H0               ' 予約済み
    
    sts = BTRV(BtOpCreate, SAGYO_POS, SAGYO_Speck, Len(SAGYO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "作業管理マスタ")
        SAGYO_Create = True
    End If
End Function

Function SAGYO_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              作業管理マスタ  ＯＰＥＮ                            *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    SAGYO_Open = False
                                            '作業管理マスタフルパス取込み
    sts = GetIni("FILE", SAGYO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        SAGYO_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, SAGYO_POS, SAGYOREC, Len(SAGYOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SAGYO_Create()        '作業管理マスタ作成
                If sts <> False Then
                    SAGYO_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SAGYO_POS, SAGYOREC, Len(SAGYOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "作業管理マスタ")
                    SAGYO_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "作業管理マスタ")
                SAGYO_Open = True
                Exit Function
        End Select
    Loop
End Function
