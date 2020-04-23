Attribute VB_Name = "P_CODE"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���i���x���R���g���[��  �t�@�C����`                *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�t�@�C���h�c
Public Const P_CODE_ID$ = "P_CODE"

'�y�[�W�T�C�Y
Private Const P_CODE_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public P_CODE_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type P_CODEREC_Tag
    
    DATA_KBN(0 To 1)        As Byte         '�ް��敪
    C_Code(0 To 9)          As Byte         '����
    C_NAME(0 To 59)         As Byte         '�ݸ�Ȱі���
    C_RNAME(0 To 19)        As Byte         '����Ȱі���
    OPTION1(0 To 9)         As Byte         '��߼��1
    OPTION2(0 To 9)         As Byte         '��߼��2
    FILLER(0 To 60)         As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_CODEREC           As P_CODEREC_Tag

'�L�[��`

Type KEY0_P_CODE                            '�j�d�x�O
    DATA_KBN(0 To 1)        As Byte         '�ް��敪
    C_Code(0 To 9)          As Byte         '����
End Type
    
'�L�[�E�f�[�^
Public K0_P_CODE            As KEY0_P_CODE

Type P_CODE_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_CODE_Speck        As P_CODE_FSpeck
Private Function P_CODE_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �R�[�h�}�X�^  �b�q�d�`�s�d                          *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_CODE_Create = True
                                            '�R�[�h�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_CODE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_CODE]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_CODE_Speck.fs.recoleng = Len(P_CODEREC)          ' ���R�[�h��
    P_CODE_Speck.fs.PageSize = P_CODE_PG_SIZ           ' �y�[�W�T�C�Y
    P_CODE_Speck.fs.idexnumb = 1                        ' �C���f�b�N�X��
    P_CODE_Speck.fs.fileflag = 0                        ' �t�@�C���t���O
    P_CODE_Speck.fs.reserve = &H0                       ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_CODE_Speck.ks0.keypos = 1                         ' �L�[�|�W�V����
    P_CODE_Speck.ks0.keyleng = 2                        ' �L�[��
    P_CODE_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    P_CODE_Speck.ks0.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    P_CODE_Speck.ks0.reserve = &H0                      ' �\��ς�
    
    P_CODE_Speck.ks1.keypos = 3                         ' �L�[�|�W�V����
    P_CODE_Speck.ks1.keyleng = 10                       ' �L�[��
    P_CODE_Speck.ks1.keyflag = BtKfExt                  ' �L�[�t���O
    P_CODE_Speck.ks1.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    P_CODE_Speck.ks1.reserve = &H0                      ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    sts = BTRV(BtOpCreate, P_CODE_POS, P_CODE_Speck, Len(P_CODE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�R�[�h�}�X�^")
        Exit Function
    End If
    
    P_CODE_Create = False

End Function

Public Function P_CODE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �R�[�h�}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_CODE_Open = True
                                            '�R�[�h�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_CODE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_CODE]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_CODE_POS, P_CODEREC, Len(P_CODEREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_CODE_Create()      '�R�[�h�}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_CODE_POS, P_CODEREC, Len(P_CODEREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�R�[�h�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�R�[�h�}�X�^")
                Exit Function
        End Select
    Loop
    
    P_CODE_Open = False

End Function
