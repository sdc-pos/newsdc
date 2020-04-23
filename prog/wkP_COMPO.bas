Attribute VB_Name = "wkP_COMPO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              構成マスタ  ファイル定義                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'ファイルＩＤ
Public Const wkP_COMPO_ID$ = "wkP_COMPO"

'ページサイズ
Private Const wkP_COMPO_PG_SIZ% = 512

'ポジション・ブロック
Public wkP_COMPO_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type wkP_COMPOREC_Tag
    
    
    SHIMUKE(0 To 2)         As Byte         '仕向け先
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
    FILLER(0 To 137)        As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public wkP_COMPOREC         As wkP_COMPOREC_Tag

'キー定義

Type KEY0_wkP_COMPO                         'ＫＥＹ０
    SHIMUKE(0 To 2)         As Byte         '仕向け先
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
End Type
    
'キー・データ
Public K0_wkP_COMPO         As KEY0_wkP_COMPO

Type wkP_COMPO_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private wkP_COMPO_Speck     As wkP_COMPO_FSpeck

Public Function wkP_COMPO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              構成マスタ（別ポジショニング）  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    wkP_COMPO_Open = True
                                            '構成マスタフルパス取込み
    sts = GetIni("FILE", wkP_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [wkP_COMPO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wkP_COMPO_POS, wkP_COMPOREC, Len(wkP_COMPOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "構成マスタ")
                Exit Function
        End Select
    Loop
    
    wkP_COMPO_Open = False

End Function
