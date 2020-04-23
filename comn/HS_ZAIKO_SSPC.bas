Attribute VB_Name = "HS_ZAIKO_SSPC"
Option Explicit
'********************************************************************
'*
'*              �z�X�g��M�f�[�^ �t�@�C����`�i�r�r�o�b�j
'*
'*          CREATE 2006.05.26
'********************************************************************
'�t�@�C���h�c
Public Const HS_ZAI_SSPC$ = "HS_ZAI_SSPC"
'�t�@�C����
Public HS_ZAI_SSPC_No As Integer
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`(����)
Type HS_ZAI_SSPCREC_Tag

    HS_JIGYOBA_K(0 To 0)    As Byte     '���Ə�敪
    HS_JIGYOBA(0 To 7)      As Byte     '���Y�Ǘ����Ə�R�[�h
    HS_SHUSI(0 To 7)        As Byte     '�݌Ɏ��x�R�[�h
    HS_FIL1(0 To 7)         As Byte
    HS_FIL2(0 To 7)         As Byte
    HS_HIN_GAI(0 To 19)     As Byte     '�i�ڔԍ�           '2016.03.07
'    HS_HIN_GAI(0 To 12)     As Byte     '�i�ڔԍ�          '2016.03.07
    HS_HIN_NAI(0 To 12)     As Byte     '�H��i�ڔԍ�
    HS_HIN_NAME(0 To 24)    As Byte     '�i�ږ�
    HS_TANA(0 To 7)         As Byte     '۹���ݔԍ��P
    HS_SURYO(0 To 7)        As Byte     '�I�݌ɐ�
    HS_atmark(0 To 0)       As Byte     '�I�[����
    HS_CRLF(0 To 1)         As Byte     'CR + LF

End Type

'�f�[�^�E�o�b�t�@
Public HS_ZAI_SSPCREC As HS_ZAI_SSPCREC_Tag
'-------------------------------------------'
Public Function HS_ZAI_SSPC_Open(Mode As Integer, Optional JGYOBU As String = "") As Integer
'********************************************************************
'*
'*      �z�X�g��M�f�[�^  �n�o�d�m
'*
'*      �����@:OPEN���[�h�i0:�Q�Ɓ@1:�X�V�j
'*             �ް�����   (1:���Ɂ@2:�o��)
'*
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2004.03.05
'********************************************************************

Dim ans         As Integer
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim Ret

    On Error GoTo HS_ZAI_SSPC_Op_Err     '�װ�ׯ��ON

    HS_ZAI_SSPC_Open = True
                                    
    If GetIni("FILE", HS_ZAI_SSPC, "SYS", c) Then
        Call LOG_OUT(LOG_F, "SYS.INI [HS_ZAI_SSPC_SIJ]�ǂݍ��݃G���[")
        Exit Function
    End If
                                    
    FullPath = RTrim(c)
    
    
    If JGYOBU <> "" Then        '���ƕ��w��L�莞�͎��ƕ��R�[�h��t������
        Ret = InStr(1, Trim(FullPath), ".") - 1
        FullPath = Left(Trim(FullPath), Ret) & "_" & JGYOBU & Right(Trim(FullPath), Len(Trim(FullPath)) - Ret)
    End If
        
    
    
    HS_ZAI_SSPC_No = FreeFile

    Open FullPath For Binary As #HS_ZAI_SSPC_No
    
    HS_ZAI_SSPC_Open = False

    Exit Function

HS_ZAI_SSPC_Op_Err:     '�װ����ٰ��
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
                ans = MsgBox("�G���[ [HS_ZAI_SSPC Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
End Function
