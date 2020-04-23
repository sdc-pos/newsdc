Attribute VB_Name = "P_COMPO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              構成マスタ  ファイル定義                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'ファイルＩＤ
Public Const P_COMPO_ID$ = "P_COMPO"

'ページサイズ
Private Const P_COMPO_PG_SIZ% = 512

'ポジション・ブロック
Public P_COMPO_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type P_COMPO_O_REC_Tag                '親ﾚｺｰﾄﾞ
    
    
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
    CLASS_CODE(0 To 19)     As Byte         '基本ｸﾗｽ
    BIKOU(0 To 119)         As Byte         '備考
    F_CLASS_CODE(0 To 19)   As Byte         '付加ｸﾗｽ
    N_CLASS_CODE(0 To 19)   As Byte         '内職ｸﾗｽ
    FILLER(0 To 28)         As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_COMPO_O_REC        As P_COMPO_O_REC_Tag


Public Type P_COMPOREC_K_Tag                '子ﾚｺｰﾄﾞ
    
    
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
    KO_SYUBETSU(0 To 1)     As Byte         '子　種別
    KO_JGYOBU(0 To 0)       As Byte         '子　事業部
    KO_NAIGAI(0 To 0)       As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)     As Byte         '子　品番
    KO_QTY(0 To 5)          As Byte         '子　員数(999V99)
    KO_BIKOU(0 To 39)       As Byte         '子　備考
    CLASS_CODE(0 To 19)     As Byte         '基本ｸﾗｽ
    FILLER(0 To 118)        As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_COMPO_K_REC        As P_COMPOREC_K_Tag

'キー定義

Type KEY0_P_COMPO                           'ＫＥＹ０
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
End Type
    
'キー・データ
Public K0_P_COMPO           As KEY0_P_COMPO

Type P_COMPO_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_COMPO_Speck       As P_COMPO_FSpeck
Private Function P_COMPO_Create() As Integer
'********************************************************************
'*                                                                  *
'*              構成マスタ  ＣＲＥＡＴＥ                            *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_COMPO_Create = True
                                            '構成マスタフルパス取込み
    sts = GetIni("FILE", P_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_COMPO]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_COMPO_Speck.fs.recoleng = Len(P_COMPO_O_REC)      ' レコード長
    P_COMPO_Speck.fs.PageSize = P_COMPO_PG_SIZ          ' ページサイズ
    P_COMPO_Speck.fs.idexnumb = 1                       ' インデックス数
    P_COMPO_Speck.fs.fileflag = 0                       ' ファイルフラグ
    P_COMPO_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_COMPO_Speck.ks0.keypos = 1                        ' キーポジション
    P_COMPO_Speck.ks0.keyleng = 2                       ' キー長
    P_COMPO_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    P_COMPO_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    P_COMPO_Speck.ks0.reserve = &H0                     ' 予約済み
    
    P_COMPO_Speck.ks1.keypos = 3                        ' キーポジション
    P_COMPO_Speck.ks1.keyleng = 1                       ' キー長
    P_COMPO_Speck.ks1.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    P_COMPO_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    P_COMPO_Speck.ks1.reserve = &H0                     ' 予約済み
    
    P_COMPO_Speck.ks2.keypos = 4                        ' キーポジション
    P_COMPO_Speck.ks2.keyleng = 1                       ' キー長
    P_COMPO_Speck.ks2.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    P_COMPO_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    P_COMPO_Speck.ks2.reserve = &H0                     ' 予約済み
    
    P_COMPO_Speck.ks3.keypos = 5                        ' キーポジション
    P_COMPO_Speck.ks3.keyleng = 20                      ' キー長
    P_COMPO_Speck.ks3.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    P_COMPO_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    P_COMPO_Speck.ks3.reserve = &H0                     ' 予約済み
    
    P_COMPO_Speck.ks4.keypos = 25                       ' キーポジション
    P_COMPO_Speck.ks4.keyleng = 1                       ' キー長
    P_COMPO_Speck.ks4.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    P_COMPO_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    P_COMPO_Speck.ks4.reserve = &H0                     ' 予約済み
    
    P_COMPO_Speck.ks5.keypos = 26                       ' キーポジション
    P_COMPO_Speck.ks5.keyleng = 3                       ' キー長
    P_COMPO_Speck.ks5.keyflag = BtKfExt                 ' キーフラグ
    P_COMPO_Speck.ks5.keytype = Chr(BtKtString)         ' キータイプ
    P_COMPO_Speck.ks5.reserve = &H0                     ' 予約済み
    
    '--------------------------------------------------- キー０ △
    sts = BTRV(BtOpCreate, P_COMPO_POS, P_COMPO_Speck, Len(P_COMPO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "構成マスタ")
        Exit Function
    End If
    
    P_COMPO_Create = False

End Function

Public Function P_COMPO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              構成マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_COMPO_Open = True
                                            '構成マスタフルパス取込み
    sts = GetIni("FILE", P_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_COMPO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_COMPO_Create()      '構成マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "構成マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "構成マスタ")
                Exit Function
        End Select
    Loop
    
    P_COMPO_Open = False

End Function
