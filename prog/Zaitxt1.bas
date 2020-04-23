Attribute VB_Name = "HS_ZAI1"
Option Explicit
'********************************************************************
'*
'*              �݌ɐݒ�f�[�^ �t�@�C����`
'*
'*          CREATE 2001.05.18
'********************************************************************
'�t�@�C���h�c
Global Const HS_ZAI_ID1 = "HS_ZAI1"         '����@���ƕ�
'�t�@�C����
Global HS_ZAI_No As Integer
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type HS_ZAIREC_Tag
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    HOST_SOKO(0 To 1) As Byte       '�q�ɋ敪�iνāj
    HIN_GAI(0 To 12) As Byte        '�i�ԁi�O���j
    HIN_NAI(0 To 12) As Byte        '�i�ԁi�����j
    HIN_NAME(0 To 24) As Byte       '�i��
    HOST_TANA(0 To 7) As Byte       '�I�ԁiνāj
    QTY_SIGN(0 To 0) As Byte        '���ʃT�C��
    ZEN_Z_QTY(0 To 6) As Byte       '�O���݌ɐ�
    FILLER(0 To 8) As Byte          'FILLER
    REC_END(0 To 0) As Byte         'ں��ޏI�[ϰ�(@)
    CR_LF(0 To 1) As Byte           'CR.LF
End Type

'�f�[�^�E�o�b�t�@
Global HS_ZAIREC As HS_ZAIREC_Tag
Function HS_ZAI_Open1(Mode As Integer, FPass As String) As Integer
'********************************************************************
'*
'*       ����@���ƕ�  �݌ɐݒ�f�[�^  �n�o�d�m
'*
'*      �����@:OPEN���[�h�i0:�Q�Ɓ@1:�X�V�j
'*
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2001.05.18
'*
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo HS_ZAI_Op_Err    '�װ�ׯ��ON

    HS_ZAI_Open1 = False
                            '�z�X�g��M�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", HS_ZAI_ID1, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        HS_ZAI_Open1 = True
        Exit Function
    End If
    FullPath = RTrim(c)
    FPass = FullPath

    HS_ZAI_No = FreeFile

    If Mode = ZERO Then
        Open FullPath For Input As #HS_ZAI_No
    Else
        Open FullPath For Binary As #HS_ZAI_No
    End If

    Exit Function

HS_ZAI_Op_Err:     '�װ����ٰ��
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case ErrDiskNotReady
            If Mode = 1 Then
                Beep
                ans = MsgBox("�h���C�u���m�F���ĉ�����", vbYesNo + vbExclamation + vbDefaultButton1, "�m�F����")
                If ans = vbYes Then
                    Resume
                End If
            End If
        Case ErrDeviceUnavailable
            If Mode = 1 Then
                Beep
                ans = MsgBox("�h���C�u�܂��̓p�X��������܂���" & FullPath, vbExclamation)
            End If
        Case ErrNotFound
            If Mode = 1 Then
                Beep
                ans = MsgBox("�t�@�C����������܂���" & FullPath, vbExclamation)
            End If
        Case Else
            If Mode = 1 Then
                Beep
                ans = MsgBox("�G���[ [HS_ZAI Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
    HS_ZAI_Open1 = True
    Exit Function
End Function
Function HS_ZAI_Get1() As Integer
'********************************************************************
'*
'*              �݌ɐݒ�f�[�^  �f�d�s
'*
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2001.05.18
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo HS_ZAI_Put_Err    '�װ�ׯ��ON

    HS_ZAI_Get1 = False

    Get #HS_ZAI_No, , HS_ZAIREC

Exit Function

HS_ZAI_Put_Err:     '�װ����ٰ��
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
            ans = MsgBox("�h���C�u�܂��̓p�X��������܂���", vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[ [HS_ZAI Put : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
    HS_ZAI_Get1 = True
    Exit Function
End Function
