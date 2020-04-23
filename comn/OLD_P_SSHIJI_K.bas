Attribute VB_Name = "OLD_P_SSHIJI_K"
Option Explicit

'********************************************************************
'*
'*              商品化指図データ（子）  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ


Public Const OLD_P_SSHIJI_K_ID$ = "OLD_P_SSHIJI_K"

'ページサイズ
Private Const OLD_P_SSHIJI_K_PG_SIZ% = 512

'ポジション・ブロック
Public OLD_P_SSHIJI_K_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************

'レコード定義
Public Type OLD_P_SSHIJI_K_REC_Tag
    
    SHIJI_NO(0 To 4)        As Byte         '指図票№
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
    KO_SYUBETSU(0 To 1)     As Byte         '子　種別
    KO_JGYOBU(0 To 0)       As Byte         '子　事業部
    KO_NAIGAI(0 To 0)       As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)     As Byte         '子　品番
    KO_QTY(0 To 5)          As Byte         '子　員数(999V99)
    KO_SHIJI_QTY(0 To 10)   As Byte         '指示数(9(8)V99)
    KO_BIKOU(0 To 39)       As Byte         '子　備考
'    KO_ID_NO(0 To 7)        As Byte         '子 ＩＤ＿ＮＯ
    KO_ID_NO(0 To 11)       As Byte         '子 ＩＤ＿ＮＯ (8桁→12桁)  2006/05/24
    CALCEL_F(0 To 0)        As Byte         'ｷｬﾝｾﾙF
    CANCEL_DATETIME(0 To 13) As Byte        'ｷｬﾝｾﾙ日時
'    FILLER(0 To 64)         As Byte         'Filler
    FILLER(0 To 60)         As Byte         'Filler                    2006/05/24
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public OLD_P_SSHIJI_K_REC       As OLD_P_SSHIJI_K_REC_Tag

'キー定義

Type KEY0_OLD_P_SSHIJI_K                        'ＫＥＹ０
    SHIJI_NO(0 To 4)        As Byte         '指図票№
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
End Type
    
    
Type KEY1_OLD_P_SSHIJI_K                        'ＫＥＹ１
    KO_JGYOBU(0 To 0)       As Byte         '子　事業部
'    KO_ID_NO(0 To 7)        As Byte         '子 ＩＤ＿ＮＯ
    KO_ID_NO(0 To 11)       As Byte         '子 ＩＤ＿ＮＯ (8桁→12桁)  2006/05/24
End Type
    
    
'キー・データ
Public K0_OLD_P_SSHIJI_K        As KEY0_OLD_P_SSHIJI_K
Public K1_OLD_P_SSHIJI_K        As KEY1_OLD_P_SSHIJI_K


Public Function OLD_P_SSHIJI_K_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              商品化指図データ（子）  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    OLD_P_SSHIJI_K_Open = True
                                            '手配指図データ（子）フルパス取込み
    sts = GetIni("FILE", OLD_P_SSHIJI_K_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [OLD_P_SSHIJI_K]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, OLD_P_SSHIJI_K_POS, OLD_P_SSHIJI_K_REC, Len(OLD_P_SSHIJI_K_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "手配指図データ（子）マスタ")
                Exit Function
        End Select
    Loop
    
    OLD_P_SSHIJI_K_Open = False

End Function
