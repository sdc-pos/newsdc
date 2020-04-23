Attribute VB_Name = "SEQCK"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �\��捞�݃`�F�b�N �t�@�C����`                       *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'�t�@�C���h�c
Global Const SEQCK_ID = "SEQCK"

'�y�[�W�T�C�Y
Global Const SEQCK_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Global SEQCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SEQCKREC_Tag
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    SEQ_MODE(0 To 0)    As Byte     '��荞�݋敪
    LAST_TXTNO(0 To 8)  As Byte     '�ŏI�e�L�X�g��
    LAST_GET_DT(0 To 7) As Byte     '�ŏI�捞�ݓ��t
    LAST_GET_TM(0 To 5) As Byte     '�ŏI�捞�ݎ���
End Type

'�f�[�^�E�o�b�t�@
Global SEQCKREC         As SEQCKREC_Tag
'�L�[��`

Type KEY0_SEQCK            '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    SEQ_MODE(0 To 0)    As Byte     '��荞�݋敪
End Type
    
'�L�[�E�f�[�^
Global K0_SEQCK         As KEY0_SEQCK

Type SEQCK_FSpeck
    fs      As BtFileSpeck          ' ̧�� ��߯��\����
    ks0     As BtKeySpeck           ' �� ��߯��\����
End Type

Global SEQCK_Speck As SEQCK_FSpeck
Private Function SEQCK_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �\��捞�݃`�F�b�N  �b�q�d�`�s�d                    *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim FullPath As String
Dim c As String * 128

    SEQCK_Create = False
                                            '�\��捞�݃`�F�b�N�t���p�X�捞��
    sts = GetIni("FILE", SEQCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        SEQCK_Create = True
        Exit Function
    End If

    FullPath = RTrim$(c)

    SEQCK_Speck.fs.recoleng = Len(SEQCKREC)     ' ���R�[�h��
    SEQCK_Speck.fs.PageSize = SEQCK_PG_SIZ      ' �y�[�W�T�C�Y
    SEQCK_Speck.fs.idexnumb = 1                 ' �C���f�b�N�X��
    SEQCK_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    SEQCK_Speck.fs.reserve = &H0                ' �\��ς�
                                                ' �L�[�O
    SEQCK_Speck.ks0.keypos = 1                  ' �L�[�|�W�V����
    SEQCK_Speck.ks0.keyleng = 1 + 1             ' �L�[��
    SEQCK_Speck.ks0.keyflag = BtKfExt           ' �L�[�t���O
    SEQCK_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    SEQCK_Speck.ks0.reserve = &H0               ' �\��ς�

    sts = BTRV(BtOpCreate, SEQCK_POS, SEQCK_Speck, Len(SEQCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�\��捞�݃`�F�b�N")
        SEQCK_Create = True
    End If
End Function
Function SEQCK_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �\��捞�݃`�F�b�N  �n�o�d�m                        *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    SEQCK_Open = False
                                            '�\��捞�݃`�F�b�N�t���p�X�捞��
    sts = GetIni("FILE", SEQCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        SEQCK_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, SEQCK_POS, SEQCKREC, Len(SEQCKREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SEQCK_Create()        '�\��捞�݃`�F�b�N�쐬
                If sts <> False Then
                    SEQCK_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SEQCK_POS, SEQCKREC, Len(SEQCKREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�\��捞�݃`�F�b�N")
                    SEQCK_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�\��捞�݃`�F�b�N")
                SEQCK_Open = True
                Exit Function
        End Select
    Loop
End Function

