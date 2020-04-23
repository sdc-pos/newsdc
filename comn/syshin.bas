Attribute VB_Name = "SYSHIN"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���V�X�e���i�ڃf�[�^�i�捞�݃��[�N�j �t�@�C����`            *
'*                                                                  *
'*          CREATE 1997.08.27  M.Yoshizawa                            *
'********************************************************************
'�t�@�C���h�c
Global Const SYS_HIN_ID = "SYS_HIN"
'�t�@�C����
Global SYS_HIN_No As Integer
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SYS_HINREC_Tag
    No(0 To 3) As Byte               '
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    NAIGAI(0 To 0) As Byte          '�����O
    HIN_GAI(0 To 12) As Byte        '�i�ԁi�O���j
    HIN_NAME(0 To 24) As Byte       '�i��
    ST_SET_DT(0 To 7) As Byte       '�W���q�ɐݒ���t
    ST_SOKO(0 To 1) As Byte         '�W�����ɑq�� �q��
    ST_RETU(0 To 1) As Byte         '             ��
    ST_REN(0 To 1) As Byte          '             �A
    ST_DAN(0 To 1) As Byte          '             �i
    BEF_SOKO(0 To 1) As Byte        '�O����ɑq�� �q��
    BEF_RETU(0 To 1) As Byte        '             ��
    BEF_REN(0 To 1) As Byte         '             �A
    BEF_DAN(0 To 1) As Byte         '             �i
    LAST_NYU_DT(0 To 7) As Byte     '�ŏI���ɓ��t
    LAST_SYU_DT(0 To 7) As Byte     '�ŏI�o�ɓ��t
    HIN_NAI(0 To 12) As Byte        '�i�ԁi�����j
    BIKOU_SOKO(0 To 1) As Byte      '���l �z�X�g�q��
    BIKOU_TANA(0 To 7) As Byte      '���l �z�X�g�I��
    SIZAI_CD(0 To 4) As Byte        '���ރR�[�h
    HOJYU_P(0 To 7) As Byte         '��[�_
    AVE_SYUKA(0 To 7) As Byte       '�����Ϗo�א�
    SAMPLE_QTY(0 To 0) As Byte       '�T���v����
    LAST_INP_DT(0 To 7) As Byte     '�ŏI���ד��t
    FILLER(0 To 12) As Byte         'FILLER
End Type

'�f�[�^�E�o�b�t�@
Global SYS_HINREC As SYS_HINREC_Tag
Function SYS_HIN_Open(Mode As Integer, FPass As String) As Integer
'********************************************************************
'*                                                                  *
'*      ���i�ڃ}�X�^�f�[�^  �n�o�d�m         �@                       *
'*                                                                  *
'*      �����@:OPEN���[�h�i0:�Q�Ɓ@1:�X�V�j                           *
'*                                                                  *
'*      �߂�l:false ����                                            *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.08.27  M.Yoshizawa                          *
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo SYS_HIN_Op_Err    '�װ�ׯ��ON

    SYS_HIN_Open = False
                            '�z�X�g��M�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", SYS_HIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        SYS_HIN_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    FPass = FullPath

    SYS_HIN_No = FreeFile

    If Mode = 0 Then
        Open FullPath For Input As #SYS_HIN_No
    Else
        Open FullPath For Binary As #SYS_HIN_No
    End If

    Exit Function

SYS_HIN_Op_Err:     '�װ����ٰ��
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
                ans = MsgBox("�G���[ [WK_ZAI Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
    SYS_HIN_Open = True
    Exit Function
End Function
Function SYS_HIN_Get() As Integer
'********************************************************************
'*                                                                  *
'*              ���i�ڃ}�X�^�f�[�^�i�捞�݃��[�N�j  �f�d�s�@           *
'*                                                                  *
'*      �߂�l:false ����                                            *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.08.26  M.Yoshizawa                          *
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo SYS_HIN_Get_Err    '�װ�ׯ��ON

    SYS_HIN_Get = False

    Get #SYS_HIN_No, , SYS_HINREC

Exit Function

SYS_HIN_Get_Err:     '�װ����ٰ��
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
            ans = MsgBox("�h���C�u�܂��̓p�X��������܂���" & SYS_HIN_ID _
                  , vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[ [WK_ZAI Get : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
    SYS_HIN_Get = True
    Exit Function
End Function
