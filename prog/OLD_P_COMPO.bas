Attribute VB_Name = "OLD_P_COMPO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �\���}�X�^  �t�@�C����`                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�t�@�C���h�c
Public Const OLD_P_COMPO_ID$ = "OLD_P_COMPO"

'�y�[�W�T�C�Y
Private Const OLD_P_COMPO_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public OLD_P_COMPO_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type OLD_P_COMPO_O_REC_Tag                '�eں���
    
    
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
Public OLD_P_COMPO_O_REC        As OLD_P_COMPO_O_REC_Tag


Public Type OLD_P_COMPOREC_K_Tag                '�qں���
    
    
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
Public OLD_P_COMPO_K_REC        As OLD_P_COMPOREC_K_Tag

'�L�[��`

Type KEY0_OLD_P_COMPO                           '�j�d�x�O
    SHIMUKE_CODE(0 To 1)    As Byte         '�d������
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
'�L�[�E�f�[�^
Public K0_OLD_P_COMPO           As KEY0_OLD_P_COMPO

Type OLD_P_COMPO_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private OLD_P_COMPO_Speck       As OLD_P_COMPO_FSpeck
Private Function OLD_P_COMPO_Create() As Integer
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

    OLD_P_COMPO_Create = True
                                            '�\���}�X�^�t���p�X�捞��
    sts = GetIni("FILE", OLD_P_COMPO_ID, "CONV2008", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [OLD_P_COMPO]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    OLD_P_COMPO_Speck.fs.recoleng = Len(OLD_P_COMPO_O_REC)      ' ���R�[�h��
    OLD_P_COMPO_Speck.fs.PageSize = OLD_P_COMPO_PG_SIZ          ' �y�[�W�T�C�Y
    OLD_P_COMPO_Speck.fs.idexnumb = 1                       ' �C���f�b�N�X��
    OLD_P_COMPO_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    OLD_P_COMPO_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    OLD_P_COMPO_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    OLD_P_COMPO_Speck.ks0.keyleng = 2                       ' �L�[��
    OLD_P_COMPO_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    OLD_P_COMPO_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    OLD_P_COMPO_Speck.ks0.reserve = &H0                     ' �\��ς�
    
    OLD_P_COMPO_Speck.ks1.keypos = 3                        ' �L�[�|�W�V����
    OLD_P_COMPO_Speck.ks1.keyleng = 1                       ' �L�[��
    OLD_P_COMPO_Speck.ks1.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    OLD_P_COMPO_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    OLD_P_COMPO_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    OLD_P_COMPO_Speck.ks2.keypos = 4                        ' �L�[�|�W�V����
    OLD_P_COMPO_Speck.ks2.keyleng = 1                       ' �L�[��
    OLD_P_COMPO_Speck.ks2.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    OLD_P_COMPO_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    OLD_P_COMPO_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    OLD_P_COMPO_Speck.ks3.keypos = 5                        ' �L�[�|�W�V����
    OLD_P_COMPO_Speck.ks3.keyleng = 20                      ' �L�[��
    OLD_P_COMPO_Speck.ks3.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    OLD_P_COMPO_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    OLD_P_COMPO_Speck.ks3.reserve = &H0                     ' �\��ς�
    
    OLD_P_COMPO_Speck.ks4.keypos = 25                       ' �L�[�|�W�V����
    OLD_P_COMPO_Speck.ks4.keyleng = 1                       ' �L�[��
    OLD_P_COMPO_Speck.ks4.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    OLD_P_COMPO_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    OLD_P_COMPO_Speck.ks4.reserve = &H0                     ' �\��ς�
    
    OLD_P_COMPO_Speck.ks5.keypos = 26                       ' �L�[�|�W�V����
    OLD_P_COMPO_Speck.ks5.keyleng = 3                       ' �L�[��
    OLD_P_COMPO_Speck.ks5.keyflag = BtKfExt                 ' �L�[�t���O
    OLD_P_COMPO_Speck.ks5.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    OLD_P_COMPO_Speck.ks5.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    sts = BTRV(BtOpCreate, OLD_P_COMPO_POS, OLD_P_COMPO_Speck, Len(OLD_P_COMPO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�\���}�X�^")
        Exit Function
    End If
    
    OLD_P_COMPO_Create = False

End Function

Public Function OLD_P_COMPO_Open(Mode As Integer) As Integer
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

    OLD_P_COMPO_Open = True
                                            '�\���}�X�^�t���p�X�捞��
    sts = GetIni("FILE", OLD_P_COMPO_ID, "CONV2008", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [OLD_P_COMPO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, OLD_P_COMPO_POS, OLD_P_COMPO_O_REC, Len(OLD_P_COMPO_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = OLD_P_COMPO_Create()      '�\���}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, OLD_P_COMPO_POS, OLD_P_COMPO_O_REC, Len(OLD_P_COMPO_O_REC), ByVal FullPath, Len(FullPath), Mode)
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
    
    OLD_P_COMPO_Open = False

End Function
