Attribute VB_Name = "SE_SHOHIN_TANKA_M"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���i���P���}�X�^  �t�@�C����`                      *
'*                                                                  *
'*          CREATE 2008.02.05                                       *
'********************************************************************
'�t�@�C���h�c
Public Const SE_SHOHIN_TANKA_M_ID$ = "SE_SHOHIN_TANKA_M"

'�y�[�W�T�C�Y
Public Const SE_SHOHIN_TANKA_M_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public SE_SHOHIN_TANKA_M_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SE_SHOHIN_TANKA_M_REC_Tag
    
    SE_HIN_GAI(0 To 19)         As Byte     '�j�d�x
    
    SE_KOU_TANKA(0 To 10)       As Byte     '�H���@�P�� 9(8)V99
    SE_KOU_SET_DATE(0 To 7)     As Byte     '�H���@�P���ݒ��

    SE_SIZ_TANKA(0 To 10)       As Byte     '���ށ@�P�� 9(8)V99
    SE_SIZ_SET_DATE(0 To 7)     As Byte     '���ށ@�P���ݒ��


    FILLER(0 To 198)            As Byte
    
End Type
'�f�[�^�E�o�b�t�@
Public SE_SHOHIN_TANKA_M_REC    As SE_SHOHIN_TANKA_M_REC_Tag

'�L�[��`

Type KEY0_SE_SHOHIN_TANKA_M                 '�j�d�x�O
    SE_HIN_GAI(0 To 19)         As Byte     '�j�d�x
End Type
    
'�L�[�E�f�[�^
Public K0_SE_SHOHIN_TANKA_M     As KEY0_SE_SHOHIN_TANKA_M

Type SE_SHOHIN_TANKA_M_FSpeck
    fs                  As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                 As BtKeySpeck   ' �� ��߯��\����
End Type

Private SE_SHOHIN_TANKA_M_Speck As SE_SHOHIN_TANKA_M_FSpeck
Private Function SE_SHOHIN_TANKA_M_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���i���P���}�X�^  �b�q�d�`�s�d                      *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SE_SHOHIN_TANKA_M_Create = True
                                            '���i���P���}�X�^   �t���p�X�捞��
    sts = GetIni("FILE", SE_SHOHIN_TANKA_M_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_SHOHIN_TANKA_M]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    SE_SHOHIN_TANKA_M_Speck.fs.recoleng = Len(SE_SHOHIN_TANKA_M_REC)    ' ���R�[�h��
    SE_SHOHIN_TANKA_M_Speck.fs.PageSize = SE_SHOHIN_TANKA_M_PG_SIZ      ' �y�[�W�T�C�Y
    SE_SHOHIN_TANKA_M_Speck.fs.idexnumb = 1                             ' �C���f�b�N�X��
    SE_SHOHIN_TANKA_M_Speck.fs.fileflag = 0                             ' �t�@�C���t���O
    SE_SHOHIN_TANKA_M_Speck.fs.reserve = &H0                            ' �\��ς�
    
    
    '-------------------------------------------'   �j�d�x�O
    SE_SHOHIN_TANKA_M_Speck.ks0.keypos = 1                 ' �L�[�|�W�V����
    SE_SHOHIN_TANKA_M_Speck.ks0.keyleng = 2                ' �L�[��
    SE_SHOHIN_TANKA_M_Speck.ks0.keyflag = BtKfExt          ' �L�[�t���O
    SE_SHOHIN_TANKA_M_Speck.ks0.keytype = Chr(BtKtString)  ' �L�[�^�C�v
    SE_SHOHIN_TANKA_M_Speck.ks0.reserve = &H0              ' �\��ς�
    '-------------------------------------------'   �j�d�x�O

    sts = BTRV(BtOpCreate, SE_SHOHIN_TANKA_M_POS, SE_SHOHIN_TANKA_M_Speck, Len(SE_SHOHIN_TANKA_M_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���P�[�V�����ʒP���ݒ�}�X�^")
        Exit Function
    End If
    
    SE_SHOHIN_TANKA_M_Create = False

End Function

Function SE_SHOHIN_TANKA_M_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���P�[�V�����ʒP���ݒ�}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    SE_SHOHIN_TANKA_M_Open = True
                                                '���P�[�V�����ʒP���ݒ�}�X�^   �t���p�X�捞��
    sts = GetIni("FILE", SE_SHOHIN_TANKA_M_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_SHOHIN_TANKA_M]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, SE_SHOHIN_TANKA_M_POS, SE_SHOHIN_TANKA_M_REC, Len(SE_SHOHIN_TANKA_M_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SE_SHOHIN_TANKA_M_Create()   '���P�[�V�����ʒP���ݒ�}�X�^ �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SE_SHOHIN_TANKA_M_POS, SE_SHOHIN_TANKA_M_REC, Len(SE_SHOHIN_TANKA_M_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���P�[�V�����ʒP���ݒ�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���P�[�V�����ʒP���ݒ�}�X�^")
                Exit Function
        End Select
    Loop
    SE_SHOHIN_TANKA_M_Open = False

End Function
