Attribute VB_Name = "OLD_J_NYU"
Option Explicit
'********************************************************************
'*
'*              （旧）入荷チェックデータ　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const OLD_J_NYU_ID$ = "OLD_J_NYU"

'ページサイズ
Public Const OLD_J_NYU_PG_SIZ% = 512

'ポジション・ブロック
Public OLD_J_NYU_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type OLD_J_NYUREC_Tag
    JGYOBU(0 To 0)      As Byte         '事業部区分
    NAIGAI(0 To 0)      As Byte         '国内外
    HIN_GAI(0 To 12)    As Byte         '品番（外部）
    JITU_QTY(0 To 7)    As Byte         '実績数量
    FILLER(0 To 12)     As Byte         'FILLER
End Type

'データ・バッファ
Public OLD_J_NYUREC     As OLD_J_NYUREC_Tag

'キー定義
Type KEY0_OLD_J_NYU            'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte         '事業部区分
    NAIGAI(0 To 0)      As Byte         '国内外
    HIN_GAI(0 To 12)    As Byte         '品番（外部）
End Type

'キー・データ
Public K0_OLD_J_NYU     As KEY0_OLD_J_NYU
Public Function OLD_J_NYU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              （旧）入荷チェックデータ　ＯＰＥＮ                  *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_J_NYU_Open = True
                                        '入荷チェックデータフルパス取込み
    sts = GetIni("FILE", OLD_J_NYU_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI [OLD_J_NYU]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_J_NYU_POS, OLD_J_NYUREC, Len(OLD_J_NYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_J_NYU_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "(旧)入荷チェックデータ")
                Exit Function
        End Select
    Loop

    OLD_J_NYU_Open = False

End Function


