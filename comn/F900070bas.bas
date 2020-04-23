Attribute VB_Name = "F900070bas"
Option Explicit

Public OutFno   As Integer
Public DataNo   As Integer

Public ZaikoData   As Variant


'���R�[�h��`
Private Type OutREC_Tag
    TEXT_NO(0 To 8) As Byte         '÷�ć�
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    CYOK_KBN(0 To 0) As Byte        '�����敪
    DEN_DT(0 To 7) As Byte          '�`�[���t
    IO_KBN(0 To 0) As Byte          '���o�ɋ敪
    PM_KBN(0 To 0) As Byte          '�ԍ��敪
    DEN_SYU(0 To 0) As Byte         '�`�[���
    DEN_NO(0 To 5) As Byte          '�`�[��
    CYU_KBN(0 To 0) As Byte         '�����敪
    HIN_GAI(0 To 12) As Byte        '�i�ԁi�O���j
    HIN_NAI(0 To 12) As Byte        '�i�ԁi�����j
    HIN_NAME(0 To 24) As Byte       '�i��
    YOTEI_QTY(0 To 5) As Byte       '����
    YOSAN_FROM(0 To 4) As Byte      '�\�Z�P�ʁi���j
    YOSAN_TO(0 To 4) As Byte        '�\�Z�P�ʁi��j
    HOST_SOKO(0 To 1) As Byte       '�q�ɋ敪�iνāj
    HOST_TANA(0 To 7) As Byte       '�I�ԁiνāj
    SYUK_CODE(0 To 4) As Byte       '�x����^�o�א�
    SYUK_NAME(0 To 19) As Byte      '�x����^�o�א於
    REC_END(0 To 0) As Byte         'ں��ޏI�[ϰ�(@)
    CR_LF(0 To 1) As Byte           'CR.LF
End Type

'�f�[�^�E�o�b�t�@
Public OutREC As OutREC_Tag

Function OutREC_Open_Proc() As Integer
Dim ans As Integer
Dim sts As Integer

    
    On Error Resume Next    '�װ�ׯ��ON
    
    
    Kill "c:\zaiko\shiji79.dat"
    OutREC_Open_Proc = False
    OutFno = FreeFile
    
    On Error GoTo OutREC_Open_Err    '�װ�ׯ��ON

    Open "c:\zaiko\shiji79.dat" For Binary As #OutFno

    Exit Function

OutREC_Open_Err:     '�װ����ٰ��
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Beep
    ans = MsgBox("�G���[ [�݌Ɉڊǃf�[�^ : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    OutREC_Open_Proc = True
End Function

Public Function Data_Put_Proc() As Integer
    
Dim wk13    As String * 13
Dim wk25    As String * 25
    
    Data_Put_Proc = True
    
    
    On Error GoTo Error_Proc
    
    DataNo = DataNo + 1
                                '÷�ć�
    Call UniCode_Conv(OutREC.TEXT_NO, (Format(DataNo, "000000000")))
                                '���ƕ�
    Call UniCode_Conv(OutREC.JGYOBU, "7")
                                '�����敪
    Call UniCode_Conv(OutREC.CYOK_KBN, "0")
                                '�`�[���t
    Call UniCode_Conv(OutREC.DEN_DT, Format(Now, "YYYYMMDD"))
                                '���o�ɋ敪
    Call UniCode_Conv(OutREC.IO_KBN, "1")
                                '�ԍ��敪
    Call UniCode_Conv(OutREC.PM_KBN, "")
                                '�`�[���
    Call UniCode_Conv(OutREC.DEN_SYU, "")
                                '�`�[��
    Call UniCode_Conv(OutREC.DEN_NO, Format(DataNo, "000000"))
                                '�����敪
    Call UniCode_Conv(OutREC.CYU_KBN, " ")
                                '�i�ԊO��
    
    wk13 = ZaikoData(0)
    Call UniCode_Conv(OutREC.HIN_GAI, wk13)
                                '�i�ԓ���
    wk13 = ZaikoData(2)
    Call UniCode_Conv(OutREC.HIN_NAI, wk13)
                                '�i��
    wk25 = ZaikoData(1)
    Call UniCode_Conv(OutREC.HIN_NAME, wk25)
                                '����
    If Not IsNumeric(ZaikoData(3)) Then
        Call UniCode_Conv(OutREC.YOTEI_QTY, "000000")
    Else
        Call UniCode_Conv(OutREC.YOTEI_QTY, Format(CInt(ZaikoData(3)), "000000"))
    End If
                                '�\�Z�P�ʁi���j
    Call UniCode_Conv(OutREC.YOSAN_FROM, "")
                                '�\�Z�P�ʁi��j
    Call UniCode_Conv(OutREC.YOSAN_TO, "")
                                '�q�ɋ敪�i�z�X�g�j
    Call UniCode_Conv(OutREC.HOST_SOKO, "K ")
                                '�I�ԁi�z�X�g�j
    Call UniCode_Conv(OutREC.HOST_TANA, "")
                                '�x����^�o�א�
    Call UniCode_Conv(OutREC.SYUK_CODE, "")
                                '�x����^�o�א於
    Call UniCode_Conv(OutREC.SYUK_NAME, "")
                                '���R�[�h�I�[�}�[�N�i@�j
    Call UniCode_Conv(OutREC.REC_END, "@")
    Call UniCode_Conv(OutREC.CR_LF, vbCrLf)
    
    Put #OutFno, , OutREC

    Data_Put_Proc = False
    Exit Function

Error_Proc:

End Function
