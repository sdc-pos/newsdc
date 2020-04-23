Attribute VB_Name = "J_NYU"
Option Explicit
'********************************************************************
'*
'*              ���׃`�F�b�N�f�[�^�@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const J_NYU_ID$ = "J_NYU"

'�y�[�W�T�C�Y
Public Const J_NYU_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public J_NYU_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type J_NYUREC_Tag
    JGYOBU(0 To 0)      As Byte         '���ƕ��敪
    NAIGAI(0 To 0)      As Byte         '�����O
    HIN_GAI(0 To 19)    As Byte         '�i�ԁi�O���j
    JITU_QTY(0 To 7)    As Byte         '���ѐ���
    INS_DATE(0 To 7)    As Byte         '�o�^��
    FILLER(0 To 25)     As Byte         'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public J_NYUREC         As J_NYUREC_Tag

'�L�[��`
Type KEY0_J_NYU            '�j�d�x�O
    JGYOBU(0 To 0)      As Byte         '���ƕ��敪
    NAIGAI(0 To 0)      As Byte         '�����O
    HIN_GAI(0 To 19)    As Byte         '�i�ԁi�O���j
End Type

'�L�[�E�f�[�^
Public K0_J_NYU         As KEY0_J_NYU

Type J_NYU_FSpeck
    fs              As BtFileSpeck      '̧�� ��߯��\����
    ks0             As BtKeySpeck       '�� ��߯��\����
    ks1             As BtKeySpeck       '�� ��߯��\����
    ks2             As BtKeySpeck       '�� ��߯��\����
End Type

Private J_NYU_Speck     As J_NYU_FSpeck

Private Function J_NYU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���׃`�F�b�N�f�[�^�@�b�q�d�`�s�d                    *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    J_NYU_Create = True
                                            '���׃`�F�b�N�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", J_NYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [J_NYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    J_NYU_Speck.fs.recoleng = Len(J_NYUREC)     ' ���R�[�h��
    J_NYU_Speck.fs.PageSize = J_NYU_PG_SIZ      ' �y�[�W�T�C�Y
    J_NYU_Speck.fs.idexnumb = 1                 ' �C���f�b�N�X��
    J_NYU_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    J_NYU_Speck.fs.reserve = &H0                ' �\��ς�
'------------------------------------------------
                                                ' �L�[�O
    J_NYU_Speck.ks0.keypos = 1                  ' �L�[�|�W�V����
    J_NYU_Speck.ks0.keyleng = 1                 ' �L�[��
    J_NYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    J_NYU_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    J_NYU_Speck.ks0.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    J_NYU_Speck.ks1.keypos = 2                  ' �L�[�|�W�V����
    J_NYU_Speck.ks1.keyleng = 1                 ' �L�[��
    J_NYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg ' �L�[�t���O
    J_NYU_Speck.ks1.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    J_NYU_Speck.ks1.reserve = &H0               ' �\��ς�
                                                ' �L�[�O
    J_NYU_Speck.ks2.keypos = 3                  ' �L�[�|�W�V����
    J_NYU_Speck.ks2.keyleng = 20                ' �L�[��
    J_NYU_Speck.ks2.keyflag = BtKfExt           ' �L�[�t���O
    J_NYU_Speck.ks2.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    J_NYU_Speck.ks2.reserve = &H0               ' �\��ς�
'------------------------------------------------

    sts = BTRV(BtOpCreate, J_NYU_POS, J_NYU_Speck, Len(J_NYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���׃`�F�b�N�f�[�^")
        Exit Function
    End If
    
    J_NYU_Create = False

End Function
Public Function J_NYU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ���׃`�F�b�N�f�[�^�@�n�o�d�m                        *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    J_NYU_Open = True
                                        '���׃`�F�b�N�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", J_NYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [J_NYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, J_NYU_POS, J_NYUREC, Len(J_NYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = J_NYU_Create()        '���׃`�F�b�N�f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, J_NYU_POS, J_NYUREC, Len(J_NYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���׃`�F�b�N�f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���׃`�F�b�N�f�[�^")
                Exit Function
        End Select
    Loop

    J_NYU_Open = False

End Function


