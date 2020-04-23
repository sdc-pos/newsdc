Attribute VB_Name = "OLD_P_SUKEIRE"
Option Explicit

'********************************************************************
'*
'*              商品化指図受入履歴データ  ファイル定義
'*
'*          CREATE 2005.12.14
'********************************************************************
'ファイルＩＤ
Public Const OLD_P_SUKEIRE_ID$ = "OLD_P_SUKEIRE"

'ページサイズ
Private Const OLD_P_SUKEIRE_PG_SIZ% = 1024

'ポジション・ブロック
Public OLD_P_SUKEIRE_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************

Private Type GENKA_TBL_Tag          '原価情報のﾃｰﾌﾞﾙ
    NIN(0 To 2)             As Byte         '人数
    TIMES(0 To 5)           As Byte         '時間
End Type




'レコード定義
Public Type OLD_P_SUKEIRE_REC_Tag
    
    SHIJI_NO(0 To 4)        As Byte         '指図票№
    SEQNO(0 To 2)           As Byte         '追番
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    UKEIRE_DT(0 To 7)       As Byte         '受入日
    UKEIRE_QTY(0 To 10)     As Byte         '受入数量(9(8)V999)
                                            '原価項目
    GENKA_TBL(0 To 9)       As GENKA_TBL_Tag
    
    JISEKI_NAME(0 To 19)    As Byte         '自責要因名
    JISEKI_NIN(0 To 2)      As Byte         '自責  人
    JISEKI_TIMES(0 To 5)    As Byte         '自責  分
    TASEKI_NAME(0 To 19)    As Byte         '他責要因名
    TASEKI_NIN(0 To 2)      As Byte         '他責  人
    TASEKI_TIMES(0 To 5)    As Byte         '他責  分
    
    LAST_F(0 To 0)          As Byte         '最終受入ﾌﾗｸﾞ 0:継続 1:最終
    TORI_CODE(0 To 4)       As Byte         '取引先
    FILLER(0 To 94)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public OLD_P_SUKEIRE_REC        As OLD_P_SUKEIRE_REC_Tag

'キー定義

Type KEY0_OLD_P_SUKEIRE                         'ＫＥＹ０
    SHIJI_NO(0 To 4)        As Byte         '指図票№
    SEQNO(0 To 2)           As Byte         '追番
End Type
    
Type KEY1_OLD_P_SUKEIRE                         'ＫＥＹ１
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    UKEIRE_DT(0 To 7)       As Byte         '受入日
End Type
    
Type KEY2_OLD_P_SUKEIRE                         'ＫＥＹ２
    TORI_CODE(0 To 4)       As Byte         '取引先
    UKEIRE_DT(0 To 7)       As Byte         '受入日
End Type
    
'キー・データ
Public K0_OLD_P_SUKEIRE         As KEY0_OLD_P_SUKEIRE
Public K1_OLD_P_SUKEIRE         As KEY1_OLD_P_SUKEIRE
Public K2_OLD_P_SUKEIRE         As KEY2_OLD_P_SUKEIRE


Public Function OLD_P_SUKEIRE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              商品化指図受入履歴ﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    OLD_P_SUKEIRE_Open = True
                                            '商品化指図受入履歴ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", OLD_P_SUKEIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [OLD_P_SUKEIRE]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, OLD_P_SUKEIRE_POS, OLD_P_SUKEIRE_REC, Len(OLD_P_SUKEIRE_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "商品化指図受入履歴ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    OLD_P_SUKEIRE_Open = False

End Function

