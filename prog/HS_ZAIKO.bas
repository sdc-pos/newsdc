Attribute VB_Name = "HS_ZAIKO"
Option Explicit
'********************************************************************
'*
'*              �z�X�g��M�f�[�^ �t�@�C����`
'*
'*          CREATE 2004.03.04
'********************************************************************
'�t�@�C���h�c
Public Const HS_ZAIKO$ = "HS_ZAI"
'�t�@�C����
Public HS_Zaiko_No As Integer
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`(����)
Type HS_ZAIKOREC_Tag
    
    
    
    HS_JIGYOBA(0 To 7)  As Byte
    HS_HIN_GAI(0 To 19) As Byte
    HS_SHUSI(0 To 1)    As Byte
    HS_SURYO(0 To 7)    As Byte
    HS_TANA1(0 To 9)    As Byte
    HS_TANA2(0 To 9)    As Byte
    HS_TANA3(0 To 9)    As Byte
    HS_FIL(0 To 11)     As Byte
    HS_CRLF(0 To 1)     As Byte
    
    
    
'    HIN_GAI(0 To 12)    As Byte
'    HIN_TANA(0 To 9)    As Byte
'    HIN_SURYO(0 To 4)   As Byte
'    HIN_NAME1(0 To 49)  As Byte
'    HIN_NAME2(0 To 49)  As Byte
    
    
'    CR_LF(0 To 1)       As Byte           'CR.LF
    
    
    
End Type

'�f�[�^�E�o�b�t�@
Public HS_ZAIKOREC As HS_ZAIKOREC_Tag
'-------------------------------------------'
Public Function HS_ZAIKO_Open(Mode As Integer, Optional JGYOBU As String = "") As Integer
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

Dim ret

    On Error GoTo HS_Zaiko_Op_Err     '�װ�ׯ��ON

    HS_ZAIKO_Open = True
                                    
    If GetIni("FILE", HS_ZAIKO, "SYS", c) Then
        Call Log_Out(LOG_F, "SYS.INI [HS_ZAIKO_SIJ]�ǂݍ��݃G���[")
        Exit Function
    End If
                                    
    FullPath = RTrim(c)
    
    
    If JGYOBU <> "" Then        '���ƕ��w��L�莞�͎��ƕ��R�[�h��t������
        ret = InStr(1, Trim(FullPath), ".") - 1
        FullPath = Left(Trim(FullPath), ret) & "_" & JGYOBU & Right(Trim(FullPath), Len(Trim(FullPath)) - ret)
    End If
        
    
    
    HS_Zaiko_No = FreeFile

    Open FullPath For Input As #HS_Zaiko_No
    
    HS_ZAIKO_Open = False

    Exit Function

HS_Zaiko_Op_Err:     '�װ����ٰ��
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
                ans = MsgBox("�G���[ [HS_ZAIKO Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
End Function
