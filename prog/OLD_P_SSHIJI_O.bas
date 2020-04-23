Attribute VB_Name = "OLD_P_SSHIJI_O"
Option Explicit

'********************************************************************
'*
'*              商品化指図データ（親）  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ
Public Const OLD_P_SSHIJI_O_ID$ = "OLD_P_SSHIJI_O"

'ページサイズ
Private Const OLD_P_SSHIJI_O_PG_SIZ% = 1024

'ポジション・ブロック
Public OLD_P_SSHIJI_O_POS As POSBLK
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
Public Type OLD_P_SSHIJI_O_REC_Tag
    
    SHIJI_NO(0 To 4)        As Byte         '指図票№
    HAKKO_DT(0 To 7)        As Byte         '発行日
    Print_datetime(0 To 13) As Byte         '発行日時
    TANTO_CODE(0 To 4)      As Byte         '担当者ｺｰﾄﾞ
    SHONIN_CODE(0 To 4)     As Byte         '承認者ｺｰﾄﾞ
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    SHIJI_QTY(0 To 10)      As Byte         '指示数(9(8)V99)
    UKEHARAI_CODE(0 To 4)   As Byte         '手配先ｺｰﾄﾞ
    S_CLASS_CODE(0 To 19)   As Byte         '商品化ｸﾗｽ
    F_CLASS_CODE(0 To 19)   As Byte         '付加ｸﾗｽ
    N_CLASS_CODE(0 To 19)   As Byte         '内職ｸﾗｽ
    S_TANTO(0 To 1)         As Byte         '収単／担当者コード
    SAMPLE_F(0 To 0)        As Byte         '見本作成
    SHIJI_F(0 To 0)         As Byte         '指示形態 0:通常　1:ｽﾎﾟｯﾄ　2：欠品解除 3:再梱包(2007.11.09)
    TORI_KBN(0 To 0)        As Byte
    
    PRI_SHIJI(0 To 0)       As Byte         '出力対象 指図票
    PRI_PARTS(0 To 0)       As Byte         '出力対象 ﾊﾟｰﾂﾗﾍﾞﾙ
    PRI_GAISOU(0 To 0)      As Byte         '出力対象 外装ﾗﾍﾞﾙ
    PRI_KISHU(0 To 0)       As Byte         '出力対象 機種ﾗﾍﾞﾙ
    
    BIKOU(0 To 119)         As Byte         '備考
    
    
    KAN_F(0 To 0)           As Byte         '完了F
    KAN_DT(0 To 7)          As Byte         '完了日
    BUNNOU_CNT(0 To 1)      As Byte         '分納回数
    UKEIRE_QTY(0 To 10)     As Byte         '受入数（合計）
                                            '原価項目
    GENKA_TBL(0 To 9)       As GENKA_TBL_Tag
    
    JISEKI_NAME(0 To 19)    As Byte         '自責要因名
    JISEKI_NIN(0 To 2)      As Byte         '自責  人
    JISEKI_TIMES(0 To 5)    As Byte         '自責  分
    TASEKI_NAME(0 To 19)    As Byte         '他責要因名
    TASEKI_NIN(0 To 2)      As Byte         '他責  人
    TASEKI_TIMES(0 To 5)    As Byte         '他責  分
    
    
    CANCEL_F(0 To 0)        As Byte         'ｷｬﾝｾﾙF
    CANCEL_DATETIME(0 To 13) As Byte        'ｷｬﾝｾﾙ日時
    
    ORDER_DT(0 To 7)        As Byte         '受注日 2007.02.20
    
    FILLER(0 To 38)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public OLD_P_SSHIJI_O_REC       As OLD_P_SSHIJI_O_REC_Tag

'キー定義

Type KEY0_OLD_P_SSHIJI_O                        'ＫＥＹ０
    SHIJI_NO(0 To 4)        As Byte         '指図票№
End Type

Type KEY1_OLD_P_SSHIJI_O                        'ＫＥＹ１
    KAN_F(0 To 0)           As Byte         '完了F
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    KAN_DT(0 To 7)          As Byte         '完了日
    SHIJI_NO(0 To 4)        As Byte         '指図票№
End Type
    
Type KEY2_OLD_P_SSHIJI_O                        'ＫＥＹ２
    ORDER_DT(0 To 7)        As Byte         '受注日 2007.02.20
End Type
    
    
    
    
    
    
'キー・データ
Public K0_OLD_P_SSHIJI_O        As KEY0_OLD_P_SSHIJI_O
Public K1_OLD_P_SSHIJI_O        As KEY1_OLD_P_SSHIJI_O
Public K2_OLD_P_SSHIJI_O        As KEY2_OLD_P_SSHIJI_O


Public Function OLD_P_SSHIJI_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              商品化指図(親)ﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    OLD_P_SSHIJI_O_Open = True
                                            '商品化指図(親)ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", OLD_P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [OLD_P_SSHIJI_O]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, OLD_P_SSHIJI_O_POS, OLD_P_SSHIJI_O_REC, Len(OLD_P_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "商品化指図(親)ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    OLD_P_SSHIJI_O_Open = False

End Function

