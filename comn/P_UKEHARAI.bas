Attribute VB_Name = "P_UKEHARAI"
Option Explicit
'********************************************************************
'*                                                                  *
'*              受払先マスタ  ファイル定義                          *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'ファイルＩＤ
Public Const P_UKEHARAI_ID$ = "P_UKEHARAI"

'ページサイズ
Private Const P_UKEHARAI_PG_SIZ% = 512

'ポジション・ブロック
Public P_UKEHARAI_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type P_UKEHARAIREC_Tag
    
    
    
    UKEHARAI_CODE(0 To 4)   As Byte         '受払先ｺｰﾄﾞ
    SYUSHI_CODE(0 To 2)     As Byte         '収支ｺｰﾄﾞ
    UKEHARAI_NAME(0 To 49)  As Byte         '受払先名称
    UKEHARAI_RNAME(0 To 29) As Byte         '受払先略称
    BUSHO_NAME(0 To 39)     As Byte         '部署名／営業所名
    TEL_NO(0 To 14)         As Byte         '電話番号
    FAX_NO(0 To 14)         As Byte         'FAX番号
    YUBIN_NO(0 To 7)        As Byte         '郵便番号
    ADDR1(0 To 39)          As Byte         '住所1
    ADDR2(0 To 39)          As Byte         '住所2
    TORI_KBN(0 To 0)        As Byte         '取引先区分
    FILLER(0 To 117)        As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_UKEHARAIREC        As P_UKEHARAIREC_Tag

'キー定義

Type KEY0_P_UKEHARAI                        'ＫＥＹ０
    UKEHARAI_CODE(0 To 4)   As Byte         '受払先ｺｰﾄﾞ
End Type
    
Type KEY1_P_UKEHARAI                        'ＫＥＹ１
    TORI_KBN(0 To 0)        As Byte         '取引先区分
    UKEHARAI_CODE(0 To 4)   As Byte         '受払先ｺｰﾄﾞ
End Type
    
'キー・データ
Public K0_P_UKEHARAI        As KEY0_P_UKEHARAI
Public K1_P_UKEHARAI        As KEY1_P_UKEHARAI

Type P_UKEHARAI_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_UKEHARAI_Speck    As P_UKEHARAI_FSpeck
Private Function P_UKEHARAI_Create() As Integer
'********************************************************************
'*                                                                  *
'*              受払先マスタ  ＣＲＥＡＴＥ                          *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_UKEHARAI_Create = True
                                            '受払先マスタフルパス取込み
    sts = GetIni("FILE", P_UKEHARAI_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_UKEHARAI]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_UKEHARAI_Speck.fs.recoleng = Len(P_UKEHARAIREC)   ' レコード長
    P_UKEHARAI_Speck.fs.PageSize = P_UKEHARAI_PG_SIZ     ' ページサイズ
    P_UKEHARAI_Speck.fs.idexnumb = 2                    ' インデックス数
    P_UKEHARAI_Speck.fs.fileflag = 0                    ' ファイルフラグ
    P_UKEHARAI_Speck.fs.reserve = &H0                   ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_UKEHARAI_Speck.ks0.keypos = 1                     ' キーポジション
    P_UKEHARAI_Speck.ks0.keyleng = 5                    ' キー長
    P_UKEHARAI_Speck.ks0.keyflag = BtKfExt              ' キーフラグ
    P_UKEHARAI_Speck.ks0.keytype = Chr(BtKtString)      ' キータイプ
    P_UKEHARAI_Speck.ks0.reserve = &H0                  ' 予約済み
    
    '--------------------------------------------------- キー０ △
    '--------------------------------------------------- キー０ ▽
    P_UKEHARAI_Speck.ks1.keypos = 247                   ' キーポジション
    P_UKEHARAI_Speck.ks1.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_UKEHARAI_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_UKEHARAI_Speck.ks1.keytype = Chr(BtKtString)      ' キータイプ
    P_UKEHARAI_Speck.ks1.reserve = &H0                  ' 予約済み
    
    
    P_UKEHARAI_Speck.ks2.keypos = 1                     ' キーポジション
    P_UKEHARAI_Speck.ks2.keyleng = 5                    ' キー長
    P_UKEHARAI_Speck.ks2.keyflag = BtKfExt + BtKfChg    ' キーフラグ
    P_UKEHARAI_Speck.ks2.keytype = Chr(BtKtString)      ' キータイプ
    P_UKEHARAI_Speck.ks2.reserve = &H0                  ' 予約済み
    
    '--------------------------------------------------- キー０ △
    
    sts = BTRV(BtOpCreate, P_UKEHARAI_POS, P_UKEHARAI_Speck, Len(P_UKEHARAI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "受払先マスタ")
        Exit Function
    End If
    
    P_UKEHARAI_Create = False

End Function

Public Function P_UKEHARAI_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              受払先マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_UKEHARAI_Open = True
                                            '受払先マスタフルパス取込み
    sts = GetIni("FILE", P_UKEHARAI_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_UKEHARAI]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_UKEHARAI_Create()   '受払先マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "受払先マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "受払先マスタ")
                Exit Function
        End Select
    Loop
    
    P_UKEHARAI_Open = False

End Function
