Attribute VB_Name = "ABC"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �`�a�b�Ǘ��W�v�t�@�C���i�ꎞ�t�@�C���j �t�@�C����` *
'*                                                                  *
'*          CREATE 2004.04.22                                       *
'********************************************************************
'�t�@�C���h�c
Public Const ABC_ID = "ABC"

'�y�[�W�T�C�Y
Public Const ABC_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public ABC_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type ABCREC_Tag
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    ST_LOCATION(0 To 7) As Byte     '�W���I��
    PACKING_NO(0 To 3)  As Byte     '������
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    RANK_NOW(0 To 2)    As Byte     '���ݐݒ胉���N
    RANK_NEW(0 To 2)    As Byte     '�V�����N

End Type
'�f�[�^�E�o�b�t�@
Public ABCREC           As ABCREC_Tag


'�L�[��`
Type KEY0_ABC                       '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    ST_LOCATION(0 To 7) As Byte     '�W���I��
    PACKING_NO(0 To 3)  As Byte     '������
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type
    
'�L�[�E�f�[�^
Public K0_ABC           As KEY0_ABC

Private Type ABC_FSpeck
    fs  As BtFileSpeck              ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck               ' �� ��߯��\����
    ks2 As BtKeySpeck               ' �� ��߯��\����
    ks3 As BtKeySpeck               ' �� ��߯��\����
    ks4 As BtKeySpeck               ' �� ��߯��\����
End Type

Private ABC_Speck    As ABC_FSpeck
Private Function ABC_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ABC�Ǘ��W�v�t�@�C��  �b�q�d�`�s�d                   *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.04.22                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ABC_Create = True
                                            'ABC�Ǘ��W�v�t�@�C���t���p�X�捞��
    sts = GetIni("FILE", ABC_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[ABC] �ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim$(c)

    ABC_Speck.fs.recoleng = Len(ABCREC)             ' ���R�[�h��
    ABC_Speck.fs.PageSize = ABC_PG_SIZ              ' �y�[�W�T�C�Y
    ABC_Speck.fs.idexnumb = 1                       ' �C���f�b�N�X��
    ABC_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    ABC_Speck.fs.reserve = &H0                      ' �\��ς�
                                                    
'---------------------------------------------------' �L�[�O
    ABC_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    ABC_Speck.ks0.keyleng = 1                       ' �L�[��
    ABC_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    ABC_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ABC_Speck.ks0.reserve = &H0                     ' �\��ς�

    ABC_Speck.ks1.keypos = 2                        ' �L�[�|�W�V����
    ABC_Speck.ks1.keyleng = 1                       ' �L�[��
    ABC_Speck.ks1.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    ABC_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ABC_Speck.ks1.reserve = &H0                     ' �\��ς�

    ABC_Speck.ks2.keypos = 3                        ' �L�[�|�W�V����
    ABC_Speck.ks2.keyleng = 8                       ' �L�[��
    ABC_Speck.ks2.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    ABC_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ABC_Speck.ks2.reserve = &H0                     ' �\��ς�

    ABC_Speck.ks3.keypos = 11                       ' �L�[�|�W�V����
    ABC_Speck.ks3.keyleng = 4                       ' �L�[��
    ABC_Speck.ks3.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    ABC_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ABC_Speck.ks3.reserve = &H0                     ' �\��ς�

    ABC_Speck.ks4.keypos = 15                       ' �L�[�|�W�V����
    ABC_Speck.ks4.keyleng = 20                      ' �L�[��
    ABC_Speck.ks4.keyflag = BtKfExt                 ' �L�[�t���O
    ABC_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ABC_Speck.ks4.reserve = &H0                     ' �\��ς�

    sts = BTRV(BtOpCreate, ABC_POS, ABC_Speck, Len(ABC_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�`�a�b�Ǘ��W�v�t�@�C��")
        Exit Function
    End If
    
    ABC_Create = False

End Function

Function ABC_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �`�a�b�Ǘ��W�v�t�@�C��  �n�o�d�m                    *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.04.22                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ABC_Open = True
                                            '�`�a�b�Ǘ��W�v�t�@�C���t���p�X�捞��
    sts = GetIni("FILE", ABC_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, ABC_POS, ABCREC, Len(ABCREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ABC_Create()        '�`�a�b�Ǘ��W�v�t�@�C���쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ABC_POS, ABCREC, Len(ABCREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�`�a�b�Ǘ��W�v�t�@�C��")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�`�a�b�Ǘ��W�v�t�@�C��")
                Exit Function
        End Select
    Loop
    ABC_Open = False
End Function
