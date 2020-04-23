Attribute VB_Name = "ODR3020G"
Option Explicit
'********************************************************************
'*
'*              ＯＤＲ３０２０用　共通変数
'*
'********************************************************************
Public NAIGAI_CODE()   As String * 1
Public NAIGAI_NAME()   As String

'---------------------------------------------- *部品推移データ別ポインタ
'ポジショニング
Public wODR_BUHIN_SUII_POS  As POSBLK
'データ・バッファ
Public wODR_BUHIN_SUII_REC  As ODR_BUHIN_SUII_REC_Tag
'キー・データ
Public K0_wODR_BUHIN_SUII   As KEY0_ODR_BUHIN_SUII
Public K1_wODR_BUHIN_SUII   As KEY1_ODR_BUHIN_SUII


'---------------------------------------------- *半製品管理データ別ポインタ
'ポジショニング
Public wODR_HANSEIHIN_POS   As POSBLK
'データ・バッファ
Public wODR_HANSEIHIN_O_REC As ODR_HANSEIHIN_O_REC_Tag
Public wODR_HANSEIHIN_K_REC As ODR_HANSEIHIN_K_REC_Tag
'キー・データ
Public K0_wODR_HANSEIHIN    As KEY0_ODR_HANSEIHIN
Public K1_wODR_HANSEIHIN    As KEY1_ODR_HANSEIHIN
Public K2_wODR_HANSEIHIN    As KEY2_ODR_HANSEIHIN


Public Function wODR_BUHIN_SUII_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              部品推移データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim sBuffer         As String * 255
Dim com             As String


Dim Ret             As Integer


    wODR_BUHIN_SUII_Open = True
                                            '部品推移データフルパス取込み
    sts = GetIni("FILE", ODR_BUHIN_SUII_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_SUII]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)


    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)


    Do
        sts = BTRV(BtOpOpen, wODR_BUHIN_SUII_POS, wODR_BUHIN_SUII_REC, Len(wODR_BUHIN_SUII_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "部品推移データ")
                Exit Function
        End Select
    Loop
    
    wODR_BUHIN_SUII_Open = False

End Function

Public Function wODR_HANSEIHIN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              半製品管理データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    wODR_HANSEIHIN_Open = True
                                            '半製品管理データフルパス取込み
    sts = GetIni("FILE", ODR_HANSEIHIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_HANSEIHIN]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wODR_HANSEIHIN_POS, wODR_HANSEIHIN_O_REC, Len(wODR_HANSEIHIN_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "半製品管理データ")
                Exit Function
        End Select
    Loop
    
    wODR_HANSEIHIN_Open = False

End Function

