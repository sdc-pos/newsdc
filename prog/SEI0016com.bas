Attribute VB_Name = "SEI0016com"

Option Explicit


Public KOUSEI      As New XArrayDB

'Public KO_KOUSEI    As New XArrayDB



'********************************************************************
'*                                                                  *
'*              構成マスタ  ファイル定義                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'ファイルＩＤ
Public Const wP_COMPO_ID$ = "P_COMPO"

'ポジション・ブロック
Public wP_COMPO_POS         As POSBLK
'データ・バッファ
Public wP_COMPO_O_REC        As P_COMPO_O_REC_Tag
'データ・バッファ
Public wP_COMPO_K_REC        As P_COMPOREC_K_Tag
    
'キー・データ
Public K0_wP_COMPO           As KEY0_P_COMPO




Public Function wP_COMPO_Open(Mode As Integer) As Integer
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

    wP_COMPO_Open = True
                                            '構成マスタフルパス取込み
    sts = GetIni("FILE", wP_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_COMPO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wP_COMPO_POS, wP_COMPO_O_REC, Len(wP_COMPO_O_REC), ByVal FullPath, Len(FullPath), Mode)
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
    
    wP_COMPO_Open = False

End Function

