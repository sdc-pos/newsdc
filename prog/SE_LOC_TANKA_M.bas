Attribute VB_Name = "SE_LOC_TANKA_M"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���o�ɒP���ݒ�}�X�^  �t�@�C����`                  *
'*                                                                  *
'*          CREATE 2008.02.05                                       *
'********************************************************************
'�t�@�C���h�c
Public Const SE_LOC_TANKA_M_ID$ = "SE_LOC_TANKA_M"

'�y�[�W�T�C�Y
Public Const SE_LOC_TANKA_M_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public SE_LOC_TANKA_M_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SE_LOC_TANKA_M_REC_Tag
    
    SE_IO_TANKA_No(0 To 1)      As Byte     '�j�d�x
    SE_Name(0 To 39)            As Byte     '����
    
    SE_IN_KOUSU(0 To 5)         As Byte     '���Ɂ@�H�� 9(3)V99
    SE_IN_TANKA(0 To 10)        As Byte     '���Ɂ@�P�� 9(8)V99
    SE_IN_SET_DATE(0 To 7)      As Byte     '���Ɂ@�P���ݒ��

    SE_OUT_KOUSU(0 To 5)        As Byte     '�o�Ɂ@�H�� 9(3)V99
    SE_OUT_TANKA(0 To 10)       As Byte     '�o�Ɂ@�P�� 9(8)V99
    SE_OUT_SET_DATE(0 To 7)     As Byte     '�o�Ɂ@�P���ݒ��

    SE_S_IN_KOUSU(0 To 5)       As Byte     '�����@�H�� 9(3)V99
    SE_S_IN_TANKA(0 To 10)      As Byte     '�����@�P�� 9(8)V99     ���ݖ��g�p
    SE_S_IN_SET_DATE(0 To 7)    As Byte     '�����@�P���ݒ��       ���ݖ��g�p

    SE_S_OUT_KOUSU(0 To 5)      As Byte     '���o�@�H�� 9(3)V99
    SE_S_OUT_TANKA(0 To 10)     As Byte     '���o�@�P�� 9(8)V99     ���ݖ��g�p
    SE_S_OUT_SET_DATE(0 To 7)   As Byte     '���o�@�P���ݒ��       ���ݖ��g�p


    UPD_TANTO(0 To 4)           As Byte     '�X�V�@�S����
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����



    FILLER(0 To 94)             As Byte
    
End Type
'�f�[�^�E�o�b�t�@
Public SE_LOC_TANKA_M_REC       As SE_LOC_TANKA_M_REC_Tag

'�L�[��`

Type KEY0_SE_LOC_TANKA_M                    '�j�d�x�O
    SE_IO_TANKA_No(0 To 1)      As Byte     '�j�d�x
End Type
    
'�L�[�E�f�[�^
Public K0_SE_LOC_TANKA_M        As KEY0_SE_LOC_TANKA_M

Type SE_LOC_TANKA_M_FSpeck
    fs                  As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                 As BtKeySpeck   ' �� ��߯��\����
End Type

Private SE_LOC_TANKA_M_Speck    As SE_LOC_TANKA_M_FSpeck
Private Function SE_LOC_TANKA_M_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���o�ɒP���ݒ�}�X�^  �b�q�d�`�s�d                  *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SE_LOC_TANKA_M_Create = True
                                            '���o�ɒP���ݒ�}�X�^   �t���p�X�捞��
    sts = GetIni("FILE", SE_LOC_TANKA_M_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_LOC_TANKA_M]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    SE_LOC_TANKA_M_Speck.fs.recoleng = Len(SE_LOC_TANKA_M_REC)  ' ���R�[�h��
    SE_LOC_TANKA_M_Speck.fs.PageSize = SE_LOC_TANKA_M_PG_SIZ    ' �y�[�W�T�C�Y
    SE_LOC_TANKA_M_Speck.fs.idexnumb = 1                        ' �C���f�b�N�X��
    SE_LOC_TANKA_M_Speck.fs.fileflag = 0                        ' �t�@�C���t���O
    SE_LOC_TANKA_M_Speck.fs.reserve = &H0                       ' �\��ς�
    
    
    '-------------------------------------------'   �j�d�x�O
    SE_LOC_TANKA_M_Speck.ks0.keypos = 1                 ' �L�[�|�W�V����
    SE_LOC_TANKA_M_Speck.ks0.keyleng = 2                ' �L�[��
    SE_LOC_TANKA_M_Speck.ks0.keyflag = BtKfExt          ' �L�[�t���O
    SE_LOC_TANKA_M_Speck.ks0.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    SE_LOC_TANKA_M_Speck.ks0.reserve = &H0              ' �\��ς�
    '-------------------------------------------'   �j�d�x�O

    sts = BTRV(BtOpCreate, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_Speck, Len(SE_LOC_TANKA_M_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���o�ɒP���ݒ�}�X�^")
        Exit Function
    End If
    
    SE_LOC_TANKA_M_Create = False

End Function

Function SE_LOC_TANKA_M_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���o�ɒP���ݒ�}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    SE_LOC_TANKA_M_Open = True
                                                '���o�ɒP���ݒ�}�X�^   �t���p�X�捞��
    sts = GetIni("FILE", SE_LOC_TANKA_M_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_LOC_TANKA_M]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SE_LOC_TANKA_M_Create()   '���o�ɒP���ݒ�}�X�^ �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���o�ɒP���ݒ�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���o�ɒP���ݒ�}�X�^")
                Exit Function
        End Select
    Loop
    SE_LOC_TANKA_M_Open = False

End Function