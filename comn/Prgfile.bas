Attribute VB_Name = "PRGF"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �v���O�����`�F�b�N�t�@�C����`                      *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'�t�@�C���h�c
Global Const PRGF_ID = "PRGF"

'�y�[�W�T�C�Y
Global Const PRGF_PG_SIZ% = 512
'�t�@�C���p�X
Global PRGFPath As String
'���R�[�h��`
Type PRGF
    PROG_ID(0 To 7) As Byte         '�v���O�����h�c
    END_CTL(0 To 0) As Byte         '�I������L��
    START_DT(0 To 7) As Byte        '�J�n���t
    START_TM(0 To 5) As Byte        '�J�n����
    FILLER(0 To 6) As Byte          'FILLER
End Type
'�j�d�x�O
Type KEY0_PRGF_Tag
    PROG_ID(0 To 7) As Byte         '�v���O�����h�c
End Type
'�j�d�x�P
Type KEY1_PRGF_Tag
    END_CTL(0 To 0) As Byte         '�I������L��
    PROG_ID(0 To 7) As Byte         '�v���O�����h�c
End Type

Global PRGFRec As PRGF

Global K0_PRGF As KEY0_PRGF_Tag
Global K1_PRGF As KEY1_PRGF_Tag

Global PRGF_Pos As POSBLK
    
Type PRGF_FSpeck
    fs As BtFileSpeck                 ' ̧�� ��߯��\����
    ks0 As BtKeySpeck                 ' �� ��߯��\����
    ks1 As BtKeySpeck
    ks2 As BtKeySpeck
End Type

Global PRGF_Speck As PRGF_FSpeck



'****************************************************
'*      ����v���O�������s�`�F�b�N�t�@�C���쐬      *
'*  ��  ��: �Ȃ�                                    *
'*                                                  *
'*  �߂�l: false   ����I��                        *
'*          true    �ُ�I��                        *
'*          CREATE 1997.06.06  S.Shibano            *
'****************************************************
Private Function PRGF_Create() As Integer
Dim sts As Integer
Dim c As String * 128
Dim messge As String
    
    PRGF_Create = False
    
    sts = GetIni("FILE", PRGF_ID, "SYS", c)
    If sts <> False Then
        messge = "SYS.INI �Ǎ��݃G���["
        Call Log_Out(LOG_F, messge)
        MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        PRGF_Create = True
        Exit Function
    End If
    PRGFPath = RTrim(c)
    PRGF_Speck.fs.recoleng = Len(PRGFRec)           ' ���R�[�h��
    PRGF_Speck.fs.PageSize = PRGF_PG_SIZ            ' �y�[�W�T�C�Y
    PRGF_Speck.fs.idexnumb = 2                      ' �C���f�b�N�X��
    PRGF_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    PRGF_Speck.fs.reserve = &H0                     ' �\��ς�
                                                    ' �L�[�O
    PRGF_Speck.ks0.keypos = 1                       ' �L�[�|�W�V����
    PRGF_Speck.ks0.keyleng = 8                      ' �L�[��
    PRGF_Speck.ks0.keyflag = BtKfExt                ' �L�[�t���O
    PRGF_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PRGF_Speck.ks0.reserve = &H0                    ' �\��ς�
                                                    ' �L�[�P
    PRGF_Speck.ks1.keypos = 9                       ' �L�[�|�W�V����
    PRGF_Speck.ks1.keyleng = 1                      ' �L�[��
    PRGF_Speck.ks1.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    PRGF_Speck.ks1.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PRGF_Speck.ks1.reserve = &H0                    ' �\��ς�
                                                    ' �L�[�P
    PRGF_Speck.ks2.keypos = 1                       ' �L�[�|�W�V����
    PRGF_Speck.ks2.keyleng = 8                      ' �L�[��
    PRGF_Speck.ks2.keyflag = BtKfExt                ' �L�[�t���O
    PRGF_Speck.ks2.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    PRGF_Speck.ks2.reserve = &H0                    ' �\��ς�

    sts = BTRV(BtOpCreate, PRGF_Pos, PRGF_Speck, Len(PRGF_Speck), ByVal PRGFPath, Len(PRGFPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�N���v���Z�X�`�F�b�N")
        PRGF_Create = True
        Exit Function
    End If
    sts = BTRV(BtOpOpen, PRGF_Pos, PRGFRec, Len(PRGFRec), ByVal PRGFPath, Len(PRGFPath), 0)
    If sts Then
        Call File_Error(sts, BtOpOpen, "�N���v���Z�X�`�F�b�N")
        PRGF_Create = True
        Exit Function
    End If
End Function

'****************************************************
'*      ����v���O�������s�`�F�b�N�t�@�C���n�o�d�m  *
'*  ��  ��: Open Mode(Btrieve�Q��)                  *
'*                                                  *
'*  �߂�l: false   ����I��                        *
'*          true    �ُ�I��                        *
'*          CREATE 1997.06.06  S.Shibano            *
'****************************************************
Function PRGF_Open(Mode As Integer) As Integer

Dim c As String * 128
Dim messge As String
Dim sts As Integer

    PRGF_Open = False

    sts = GetIni("FILE", PRGF_ID, "SYS", c)
    If sts <> False Then
        messge = "SYS.INI �Ǎ��݃G���["
        Call Log_Out(LOG_F, messge)
        PRGF_Open = True
        Exit Function
    End If
    PRGFPath = RTrim(c)
    sts = BTRV(BtOpOpen, PRGF_Pos, PRGFRec, Len(PRGFRec), ByVal PRGFPath, Len(PRGFPath), 0)
    If sts Then
        If sts = BtErrFileNotFound Then
            sts = PRGF_Create()
            If sts Then
                PRGF_Open = True
                Exit Function
            End If
        Else
            Call File_Error(sts, BtOpOpen, "�N���v���Z�X�`�F�b�N")
            PRGF_Open = True
            Exit Function
        End If
    End If
End Function
