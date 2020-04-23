Attribute VB_Name = "OLD_KEPPINLOG"
Option Explicit
'********************************************************************
'*
'*              欠品防止支援ログ　ファイル定義
'*
'*          CREATE 2004.05.08
'********************************************************************
'ファイルＩＤ
Public Const OLD_KEPPINLOG_ID$ = "OLD_KEPPINLOG"

'ページサイズ
Public Const OLD_KEPPINLOG_PG_SIZ% = 512

'ポジション・ブロック
Public OLD_KEPPINLOG_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                              *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type OLD_KEPPINLOGREC_Tag
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 12)        As Byte     '品番（外部）
    CREATE_DT(0 To 7)       As Byte     '作成日付
    FILLER(0 To 16)         As Byte     'FILLER
End Type

'データ・バッファ
Public OLD_KEPPINLOGREC     As OLD_KEPPINLOGREC_Tag

'キー定義
Private Type KEY0_OLD_KEPPINLOG     'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 12)    As Byte     '品番（外部）
End Type


'キー・データ
Public K0_OLD_KEPPINLOG As KEY0_OLD_KEPPINLOG



Function OLD_KEPPINLOG_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              欠品防止支援ログ　ＯＰＥＮ                          *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.05.08                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_KEPPINLOG_Open = True
                                            '欠品防止支援ログフルパス取込み
    sts = GetIni("FILE", OLD_KEPPINLOG_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI[OLD_KEPPINLOG] 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_KEPPINLOG_POS, OLD_KEPPINLOGREC, Len(OLD_KEPPINLOGREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_KEPPINLOG_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "（旧）欠品防止支援ログ")
                Exit Function
        End Select
    Loop

    OLD_KEPPINLOG_Open = False

End Function


