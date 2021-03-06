Attribute VB_Name = "F900070bas"
Option Explicit

Public OutFno   As Integer
Public DataNo   As Integer

Public ZaikoData   As Variant


'R[hθ`
Private Type OutREC_Tag
    TEXT_NO(0 To 8) As Byte         'Γ·½Δ
    JGYOBU(0 To 0) As Byte          'Ζζͺ
    CYOK_KBN(0 To 0) As Byte        'Όζͺ
    DEN_DT(0 To 7) As Byte          '`[ϊt
    IO_KBN(0 To 0) As Byte          'όoΙζͺ
    PM_KBN(0 To 0) As Byte          'Τζͺ
    DEN_SYU(0 To 0) As Byte         '`[νΚ
    DEN_NO(0 To 5) As Byte          '`[
    CYU_KBN(0 To 0) As Byte         'Άζͺ
    HIN_GAI(0 To 12) As Byte        'iΤiOj
    HIN_NAI(0 To 12) As Byte        'iΤiΰj
    HIN_NAME(0 To 24) As Byte       'iΌ
    YOTEI_QTY(0 To 5) As Byte       'Κ
    YOSAN_FROM(0 To 4) As Byte      '\ZPΚi³j
    YOSAN_TO(0 To 4) As Byte        '\ZPΚiζj
    HOST_SOKO(0 To 1) As Byte       'qΙζͺiΞ½Δj
    HOST_TANA(0 To 7) As Byte       'IΤiΞ½Δj
    SYUK_CODE(0 To 4) As Byte       'xζ^oΧζ
    SYUK_NAME(0 To 19) As Byte      'xζ^oΧζΌ
    REC_END(0 To 0) As Byte         'ΪΊ°ΔήI[Ο°Έ(@)
    CR_LF(0 To 1) As Byte           'CR.LF
End Type

'f[^Eobt@
Public OutREC As OutREC_Tag

Function OutREC_Open_Proc() As Integer
Dim ans As Integer
Dim sts As Integer

    
    On Error Resume Next    '΄Χ°ΔΧ―ΜίON
    
    
    Kill "c:\zaiko\shiji79.dat"
    OutREC_Open_Proc = False
    OutFno = FreeFile
    
    On Error GoTo OutREC_Open_Err    '΄Χ°ΔΧ―ΜίON

    Open "c:\zaiko\shiji79.dat" For Binary As #OutFno

    Exit Function

OutREC_Open_Err:     '΄Χ°Ω°Αέ
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Beep
    ans = MsgBox("G[ [έΙΪΗf[^ : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    OutREC_Open_Proc = True
End Function

Public Function Data_Put_Proc() As Integer
    
Dim wk13    As String * 13
Dim wk25    As String * 25
    
    Data_Put_Proc = True
    
    
    On Error GoTo Error_Proc
    
    DataNo = DataNo + 1
                                'Γ·½Δ
    Call UniCode_Conv(OutREC.TEXT_NO, (Format(DataNo, "000000000")))
                                'Ζ
    Call UniCode_Conv(OutREC.JGYOBU, "7")
                                'Όζͺ
    Call UniCode_Conv(OutREC.CYOK_KBN, "0")
                                '`[ϊt
    Call UniCode_Conv(OutREC.DEN_DT, Format(Now, "YYYYMMDD"))
                                'όoΙζͺ
    Call UniCode_Conv(OutREC.IO_KBN, "1")
                                'Τζͺ
    Call UniCode_Conv(OutREC.PM_KBN, "")
                                '`[νΚ
    Call UniCode_Conv(OutREC.DEN_SYU, "")
                                '`[
    Call UniCode_Conv(OutREC.DEN_NO, Format(DataNo, "000000"))
                                'Άζͺ
    Call UniCode_Conv(OutREC.CYU_KBN, " ")
                                'iΤO
    
    wk13 = ZaikoData(0)
    Call UniCode_Conv(OutREC.HIN_GAI, wk13)
                                'iΤΰ
    wk13 = ZaikoData(2)
    Call UniCode_Conv(OutREC.HIN_NAI, wk13)
                                'iΌ
    wk25 = ZaikoData(1)
    Call UniCode_Conv(OutREC.HIN_NAME, wk25)
                                'Κ
    If Not IsNumeric(ZaikoData(3)) Then
        Call UniCode_Conv(OutREC.YOTEI_QTY, "000000")
    Else
        Call UniCode_Conv(OutREC.YOTEI_QTY, Format(CInt(ZaikoData(3)), "000000"))
    End If
                                '\ZPΚi³j
    Call UniCode_Conv(OutREC.YOSAN_FROM, "")
                                '\ZPΚiζj
    Call UniCode_Conv(OutREC.YOSAN_TO, "")
                                'qΙζͺizXgj
    Call UniCode_Conv(OutREC.HOST_SOKO, "K ")
                                'IΤizXgj
    Call UniCode_Conv(OutREC.HOST_TANA, "")
                                'xζ^oΧζ
    Call UniCode_Conv(OutREC.SYUK_CODE, "")
                                'xζ^oΧζΌ
    Call UniCode_Conv(OutREC.SYUK_NAME, "")
                                'R[hI[}[Ni@j
    Call UniCode_Conv(OutREC.REC_END, "@")
    Call UniCode_Conv(OutREC.CR_LF, vbCrLf)
    
    Put #OutFno, , OutREC

    Data_Put_Proc = False
    Exit Function

Error_Proc:

End Function
