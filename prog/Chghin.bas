Attribute VB_Name = "CHGH"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �O���i�ԕύX  �t�@�C����`                            *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'�t�@�C���h�c
Global Const CHGH_ID = "CHGH"

'�t�@�C����
Global CHGH_No As Integer
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type CHGHREC_Tag
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
    CR(0 To 0) As Byte              '��د������
    LF(0 To 0) As Byte              'ײ�̨���
End Type

'�f�[�^�E�o�b�t�@
Global CHGHREC As CHGHREC_Tag
Function CHGH_Open() As Integer
'********************************************************************
'*                                                                  *
'*              �O���i�ԕύX�ۗ��ް�  �n�o�d�m                        *
'*                                                                  *
'*      �߂�l:false ����                                            *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.05.28  S.Shibano                            *
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo CHGH_Op_Err    '�װ�ׯ��ON

    CHGH_Open = False
                            '�O���i�ԕύX�ۗ��ް��t���p�X�捞��
    sts = GetIni("FILE", CHGH_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        CHGH_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)

    CHGH_No = FreeFile

    Open FullPath For Binary As #CHGH_No

    Exit Function

CHGH_Op_Err:     '�װ����ٰ��
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case ErrDiskNotReady
            Beep
            ans = MsgBox("�h���C�u���m�F���ĉ�����", vbYesNo _
                            + vbExclamation + vbDefaultButton1, "�m�F����")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable
            Beep
            ans = MsgBox("�h���C�u�܂��̓p�X��������܂���" & FullPath, vbExclamation)
        Case ErrNotFound
            Beep
            ans = MsgBox("�t�@�C����������܂���" & FullPath, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[ [CHGH Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select
    CHGH_Open = True
    Exit Function
End Function
Function CHGH_Get() As Integer
'********************************************************************
'*                                                                  *
'*              �O���i�ԕύX�ۗ��ް�  �f�d�s�@                        *
'*                                                                  *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.05.28  S.Shibano                            *
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo CHGH_Get_Err    '�װ�ׯ��ON

    CHGH_Get = False

    Get #CHGH_No, , CHGHREC

    Exit Function

CHGH_Get_Err:     '�װ����ٰ��
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68
    Select Case Err.Number
        Case ErrDiskNotReady        '��ײ�ނ���������������Ă��Ȃ�
            Beep
            ans = MsgBox("�h���C�u���m�F���ĉ�����", vbYesNo _
                  + vbExclamation + vbDefaultButton1, "�m�F����")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable   '��ײ��or�߽��������Ȃ�
            Beep
            ans = MsgBox("�h���C�u�܂��̓p�X��������܂���" & CHGH_ID _
                  , vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[ [CHGH Put : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
    CHGH_Get = True
End Function
Function CHGH_Put(Put_Kbn As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �O���i�ԕύX�ۗ��ް�  �o�t�s�@                        *
'*                                                                  *
'*�@�@�@�����@�F�u�O�v �ۗ��f�[�^�ւo�t�s                              *
'*�@�@�@�@�@�@�@�u�P�v �捞�f�[�^�ւo�t�s                              *
'*                                                                  *
'*      �߂�l:false ����                                            *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.05.28  S.Shibano                            *
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo CHGH_Put_Err    '�װ�ׯ��ON

    CHGH_Put = False

    If Put_Kbn = 0 Then
        Put #CHGH_No, , CHGHREC
    Else
        Put #XX_SIJ_No, , CHGHREC
    End If

    Exit Function

CHGH_Put_Err:     '�װ����ٰ��
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68
    Select Case Err.Number
        Case ErrDiskNotReady        '��ײ�ނ���������������Ă��Ȃ�
            Beep
            ans = MsgBox("�h���C�u���m�F���ĉ�����", vbYesNo _
                  + vbExclamation + vbDefaultButton1, "�m�F����")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable   '��ײ��or�߽��������Ȃ�
            Beep
            ans = MsgBox("�h���C�u�܂��̓p�X��������܂���" & CHGH_ID _
                  , vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[ [CHGH Put : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
    CHGH_Put = True
    Exit Function
End Function


