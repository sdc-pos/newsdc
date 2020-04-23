Attribute VB_Name = "P_COMPO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �\���}�X�^  �t�@�C����`                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�t�@�C���h�c
Public Const P_COMPO_ID$ = "P_COMPO"

'�y�[�W�T�C�Y
Private Const P_COMPO_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_COMPO_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type P_COMPO_O_REC_Tag                '�eں���
    
    
    SHIMUKE_CODE(0 To 1)    As Byte         '�d������
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
    CLASS_CODE(0 To 19)     As Byte         '��{�׽
    BIKOU(0 To 119)         As Byte         '���l
    F_CLASS_CODE(0 To 19)   As Byte         '�t���׽
    N_CLASS_CODE(0 To 19)   As Byte         '���E�׽
    FILLER(0 To 28)         As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_COMPO_O_REC        As P_COMPO_O_REC_Tag


Public Type P_COMPOREC_K_Tag                '�qں���
    
    
    SHIMUKE_CODE(0 To 1)    As Byte         '�d������
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
    KO_SYUBETSU(0 To 1)     As Byte         '�q�@���
    KO_JGYOBU(0 To 0)       As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�q�@�i��
    KO_QTY(0 To 5)          As Byte         '�q�@����(999V99)
    KO_BIKOU(0 To 39)       As Byte         '�q�@���l
    CLASS_CODE(0 To 19)     As Byte         '��{�׽
    FILLER(0 To 118)        As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_COMPO_K_REC        As P_COMPOREC_K_Tag

'�L�[��`

Public Type KEY0_P_COMPO                    '�j�d�x�O
    SHIMUKE_CODE(0 To 1)    As Byte         '�d������
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
    
    
Public Type KEY1_P_COMPO                    '�j�d�x�P   '2014.06.23
    KO_HIN_GAI(0 To 19)        As Byte         '�e�i��
End Type
    
Public Type KEY2_P_COMPO                    '�j�d�x�Q   '2018.02.20
    SHIMUKE_CODE(0 To 1)    As Byte         '�d������
    KO_JGYOBU(0 To 0)       As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�q�@�i��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
    
    
'�L�[�E�f�[�^
Public K0_P_COMPO           As KEY0_P_COMPO

Public K1_P_COMPO           As KEY1_P_COMPO             '2014.06.23

Public K2_P_COMPO           As KEY2_P_COMPO             '2018.0.220



Type P_COMPO_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����

    ks6                     As BtKeySpeck   ' �� ��߯��\����    '2014.06.23

    ks7                     As BtKeySpeck   ' �� ��߯��\����    '2018.02.20
    ks8                     As BtKeySpeck   ' �� ��߯��\����    '2018.02.20
    ks9                     As BtKeySpeck   ' �� ��߯��\����    '2018.02.20
    ks10                    As BtKeySpeck   ' �� ��߯��\����    '2018.02.20
    ks11                    As BtKeySpeck   ' �� ��߯��\����    '2018.02.20
    ks12                    As BtKeySpeck   ' �� ��߯��\����    '2018.02.20

End Type

Private P_COMPO_Speck       As P_COMPO_FSpeck
Private Function P_COMPO_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �\���}�X�^  �b�q�d�`�s�d                            *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_COMPO_Create = True
                                            '�\���}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_COMPO]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_COMPO_Speck.fs.recoleng = Len(P_COMPO_O_REC)      ' ���R�[�h��
    P_COMPO_Speck.fs.PageSize = P_COMPO_PG_SIZ          ' �y�[�W�T�C�Y
'    P_COMPO_Speck.fs.idexnumb = 1                      ' �C���f�b�N�X��    2014.06.24
'    P_COMPO_Speck.fs.idexnumb = 2                       ' �C���f�b�N�X��    2014.06.24
    P_COMPO_Speck.fs.idexnumb = 3                       ' �C���f�b�N�X��    2014.06.24
    P_COMPO_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    P_COMPO_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_COMPO_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    P_COMPO_Speck.ks0.keyleng = 2                       ' �L�[��
    P_COMPO_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    P_COMPO_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks0.reserve = &H0                     ' �\��ς�
    
    P_COMPO_Speck.ks1.keypos = 3                        ' �L�[�|�W�V����
    P_COMPO_Speck.ks1.keyleng = 1                       ' �L�[��
    P_COMPO_Speck.ks1.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    P_COMPO_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    P_COMPO_Speck.ks2.keypos = 4                        ' �L�[�|�W�V����
    P_COMPO_Speck.ks2.keyleng = 1                       ' �L�[��
    P_COMPO_Speck.ks2.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    P_COMPO_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    P_COMPO_Speck.ks3.keypos = 5                        ' �L�[�|�W�V����
    P_COMPO_Speck.ks3.keyleng = 20                      ' �L�[��
    P_COMPO_Speck.ks3.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    P_COMPO_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks3.reserve = &H0                     ' �\��ς�
    
    P_COMPO_Speck.ks4.keypos = 25                       ' �L�[�|�W�V����
    P_COMPO_Speck.ks4.keyleng = 1                       ' �L�[��
    P_COMPO_Speck.ks4.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    P_COMPO_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks4.reserve = &H0                     ' �\��ς�
    
    P_COMPO_Speck.ks5.keypos = 26                       ' �L�[�|�W�V����
    P_COMPO_Speck.ks5.keyleng = 3                       ' �L�[��
    P_COMPO_Speck.ks5.keyflag = BtKfExt                 ' �L�[�t���O
    P_COMPO_Speck.ks5.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks5.reserve = &H0                     ' �\��ς�
    
    
    
    '--------------------------------------------------- �L�[�P ��  2014.06.23
    P_COMPO_Speck.ks6.keypos = 33                       ' �L�[�|�W�V����
    P_COMPO_Speck.ks6.keyleng = 20                      ' �L�[��
    P_COMPO_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg ' �L�[�t���O
    P_COMPO_Speck.ks6.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks6.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�Q ��  2018.02.20
    P_COMPO_Speck.ks7.keypos = 1                        ' �L�[�|�W�V����
    P_COMPO_Speck.ks7.keyleng = 2                       ' �L�[��
                                                        ' �L�[�t���O
    P_COMPO_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_COMPO_Speck.ks7.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks7.reserve = &H0                     ' �\��ς�
    
    P_COMPO_Speck.ks8.keypos = 31                       ' �L�[�|�W�V����
    P_COMPO_Speck.ks8.keyleng = 1                       ' �L�[��
                                                        ' �L�[�t���O
    P_COMPO_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_COMPO_Speck.ks8.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks8.reserve = &H0                     ' �\��ς�
    
    P_COMPO_Speck.ks9.keypos = 32                       ' �L�[�|�W�V����
    P_COMPO_Speck.ks9.keyleng = 1                       ' �L�[��
                                                        ' �L�[�t���O
    P_COMPO_Speck.ks9.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_COMPO_Speck.ks9.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks9.reserve = &H0                     ' �\��ς�
    
    P_COMPO_Speck.ks10.keypos = 33                       ' �L�[�|�W�V����
    P_COMPO_Speck.ks10.keyleng = 20                       ' �L�[��
                                                        ' �L�[�t���O
    P_COMPO_Speck.ks10.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_COMPO_Speck.ks10.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks10.reserve = &H0                     ' �\��ς�
    
    P_COMPO_Speck.ks11.keypos = 25                       ' �L�[�|�W�V����
    P_COMPO_Speck.ks11.keyleng = 1                       ' �L�[��
                                                        ' �L�[�t���O
    P_COMPO_Speck.ks11.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_COMPO_Speck.ks11.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks11.reserve = &H0                     ' �\��ς�
    
    P_COMPO_Speck.ks12.keypos = 26                       ' �L�[�|�W�V����
    P_COMPO_Speck.ks12.keyleng = 3                       ' �L�[��
                                                        ' �L�[�t���O
    P_COMPO_Speck.ks12.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_COMPO_Speck.ks12.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_COMPO_Speck.ks12.reserve = &H0                     ' �\��ς�
    
    
    
    
    
    sts = BTRV(BtOpCreate, P_COMPO_POS, P_COMPO_Speck, Len(P_COMPO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�\���}�X�^")
        Exit Function
    End If
    
    P_COMPO_Create = False

End Function

Public Function P_COMPO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �\���}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_COMPO_Open = True
                                            '�\���}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_COMPO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_COMPO_Create()      '�\���}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�\���}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�\���}�X�^")
                Exit Function
        End Select
    Loop
    
    P_COMPO_Open = False

End Function
