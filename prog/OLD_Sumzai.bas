Attribute VB_Name = "OLD_SUMZ"
Option Explicit
'********************************************************************
'*
'*              （旧）在庫集計データ　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const OLD_SUMZ_ID$ = "OLD_SUMZ"

'ページサイズ
Public Const OLD_SUMZ_PG_SIZ% = 2048

'ポジション・ブロック
Public OLD_SUMZ_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type OLD_SUMZREC_Tag
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 12)        As Byte     '品番（外部）
    ST_SOKO(0 To 1)         As Byte     '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)         As Byte     '             列
    ST_REN(0 To 1)          As Byte     '             連
    ST_DAN(0 To 1)          As Byte     '             段
    T_Zai_Qty(0 To 7)       As Byte     '在庫総数(当日)
    ZEN_Zai_Qty(0 To 7)     As Byte     '在庫総数(前日)
    SYK_E_QTY(0 To 7)       As Byte     '出庫済み数
    NYUKA_YQTY(0 To 7)      As Byte     '入荷予定数
    HS_ZAIQTY(0 To 7)       As Byte     'ﾎｽﾄ在庫数(当日)
    ZEN_HS_ZAIQTY(0 To 7)   As Byte     'ﾎｽﾄ在庫数(前日)
    SAI_QTY(0 To 7)         As Byte     '差異数
    SUM_DT(0 To 7)          As Byte     '集計日付
    FILLER(0 To 8)          As Byte     'FILLER
End Type

'データ・バッファ
Public OLD_SUMZREC As OLD_SUMZREC_Tag

'キー定義
Private Type KEY0_OLD_SUMZ          'ＫＥＹ０
    JGYOBU(0 To 0) As Byte          '事業部区分
    NAIGAI(0 To 0) As Byte          '国内外
    HIN_GAI(0 To 12) As Byte        '品番（外部）
End Type


'キー・データ
Public K0_OLD_SUMZ As KEY0_OLD_SUMZ
Function OLD_SUMZ_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              在庫集計データ　ＯＰＥＮ                            *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_SUMZ_Open = True
                                            '在庫集計データフルパス取込み
    sts = GetIni("FILE", OLD_SUMZ_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI[OLD_SUMZ] 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_SUMZ_POS, OLD_SUMZREC, Len(OLD_SUMZREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_SUMZ_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "（旧）在庫集計データ")
                Exit Function
        End Select
    Loop

    OLD_SUMZ_Open = False
End Function


