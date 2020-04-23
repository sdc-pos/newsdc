Attribute VB_Name = "SYUDUP"
Option Explicit
'********************************************************************
'*
'*              �o�ח\��d���f�[�^  �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const SYUDUP_ID = "SYUDUP"

'�t�@�C����
Global SYUDUP_No As Integer
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'-------------------------------------------'
'���R�[�h��`
Type SYUDUPREC_Tag
    JGYOBU(0 To 7)              As Byte     '���Ə�
    DATA_KBN(0 To 0)            As Byte     '�f�[�^�敪
    TORI_KBN(0 To 1)            As Byte     '����敪
    ID_NO(0 To 7)               As Byte     'ID-NO
    HIN_NO(0 To 19)             As Byte     '�i�ڔԍ�
    DEN_NO(0 To 9)              As Byte     '�`�[�ԍ�
    SURYO(0 To 6)               As Byte     '�o�ɐ���
    MUKE_CODE(0 To 7)           As Byte     '���Ӑ�R�[�h
    SYUKO_SYUSI(0 To 1)         As Byte     '�o�Ɏ��x
    SYUKA_YMD(0 To 7)           As Byte     '�o�ד��t
    ODER_NO(0 To 11)            As Byte     '�I�[�_�[�ԍ�
    ITEM_NO(0 To 4)             As Byte     '�A�C�e���ԍ�
    MUKE_NAME(0 To 23)          As Byte     '���Ӑ於��
    CHU_KBN(0 To 0)             As Byte     '�����敪
    CHU_KBN_NAME(0 To 9)        As Byte     '�����敪����
    EXPORT_KBN(0 To 0)          As Byte     '�A�o�o�׌����敪
    LABEL_ISSUE_KBN(0 To 0)     As Byte     '�����x�����s�敪
    LABEL_ISSUE_UNIT(0 To 4)    As Byte     '�����x�����s�P�ʐ�
    LABEL_TANKA_KBN(0 To 0)     As Byte     '�����x���P���\���敪
    TANKA(0 To 9)               As Byte     '�P��
    KINGAKU(0 To 9)             As Byte     '���z
    BIKOU2(0 To 19)             As Byte     '���l�Q
    REBATE_KBN(0 To 0)          As Byte     '���x�[�g�敪
    CHOHA_KBN(0 To 0)           As Byte     '���[�敪
    ATAISA_KBN(0 To 0)          As Byte     '�l���敪
    REP_KISHU(0 To 19)          As Byte     '��\�@��
    NS__KANRI_NO(0 To 8)        As Byte     '�m�r�Ǘ��ԍ�
    MTS_HIN_CODE(0 To 10)       As Byte     '�l�s�r���i�R�[�h
    BIKOU1(0 To 39)             As Byte     '���l�P
    CHOKU_KBN(0 To 0)           As Byte     '�����敪
    REBATE_RATE(0 To 4)         As Byte     '���x�[�g��
    HIN_NAME(0 To 19)           As Byte     '�i��
    JGYOBU_GAI(0 To 7)          As Byte     '�ΊO���Ə�
    SS_CODE(0 To 7)             As Byte     '������R�[�h
    CRLF(0 To 1)                As Byte     'CRLF
End Type

'�f�[�^�E�o�b�t�@
Public SYUDUPREC As SYUDUPREC_Tag
Function SYUDUP_Open() As Integer
'********************************************************************
'*
'*              �o�ח\��d���ް�  �n�o�d�m
'*
'*      �����@:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************

Dim ans         As Integer
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    On Error GoTo SYUDUP_Op_Err     '�װ�ׯ��ON

    SYUDUP_Open = True
                                    
    If GetIni("FILE", SYUDUP_ID, "SYS", c) Then
        Call Log_Out(LOG_F, "SYS.INI [SYUDUP]�ǂݍ��݃G���[")
        Exit Function
    End If
                                    
    FullPath = RTrim(c)
    
    SYUDUP_No = FreeFile

    Open FullPath For Binary As #SYUDUP_No
    
    SYUDUP_Open = False

    Exit Function

HS_SIJ_Op_Err:     '�װ����ٰ��
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
                ans = MsgBox("�G���[ [HS_SIJ Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
End Function
Function SYUDUP_Get() As Integer
'********************************************************************
'*
'*              �o�ח\��d���ް�  �f�d�s
'*
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim ans         As Integer
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    On Error GoTo SYUDUP_Put_Err    '�װ�ׯ��ON

    SYUDUP_Get = True

    Get #SYUDUP_No, , SYUDUPREC

    SYUDUP_Get = False
    
    Exit Function

SYUDUP_Put_Err:     '�װ����ٰ��
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
            ans = MsgBox("�h���C�u�܂��̓p�X��������܂���" & SYUDUP_ID _
                  , vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[ [SYUDUP Put : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
End Function
Function SYUDUP_Put(Put_Kbn As Integer) As Integer
'********************************************************************
'*
'*              �o�ח\��d���ް�  �o�t�s
'*
'*�@�@�@�����@�F�u�O�v �ۗ��f�[�^�ւo�t�s
'*�@�@�@�@�@�@�@�u�P�v �捞�f�[�^�ւo�t�s
'*
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim ans         As Integer
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    On Error GoTo SYUDUP_Put_Err    '�װ�ׯ��ON

    SYUDUP_Put = True

    If Put_Kbn = 0 Then
        Put #SYUDUP_No, , SYUDUPREC
    Else
        Put #XX_SIJ_No, , SYUDUPREC
    End If

    SYUDUP_Put = False
    
    Exit Function

SYUDUP_Put_Err:     '�װ����ٰ��
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
            ans = MsgBox("�h���C�u�܂��̓p�X��������܂���" & SYUDUP_ID _
                  , vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[ [SYUDUP Put : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
End Function


