Attribute VB_Name = "ODR_HANSEIHIN"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �����i�Ǘ��f�[�^  �t�@�C����`                      *
'*                                                                  *
'*          CREATE 2008.04.26                                       *
'********************************************************************
'�t�@�C���h�c
Public Const ODR_HANSEIHIN_ID$ = "ODR_HANSEIHIN"

'�y�[�W�T�C�Y
Private Const ODR_HANSEIHIN_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public ODR_HANSEIHIN_POS        As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type ODR_HANSEIHIN_O_REC_Tag                 '�eں���
    
    
    USE_YM(0 To 5)          As Byte         '�g�p��
    INPUT_NO(0 To 3)        As Byte         '���͏�
    USE_YMD(0 To 7)         As Byte         '�g�p���t
    SEQNO(0 To 2)           As Byte         '�ǔ�(000)
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    SHIJI_QTY(0 To 11)      As Byte         '�w���� S9(8)V99
    UPD_TANTO(0 To 4)       As Byte         '�X�V�@�S����
    UPD_DATE(0 To 7)        As Byte         '�X�V�@���t
    UPD_TIME(0 To 5)        As Byte         '�X�V�@����
    
    
    IO_FLG(0 To 0)          As Byte         '���o���׸�     '2008.05.14
    
    
    
    
    FILLER(0 To 180)        As Byte
    
    

End Type
'�f�[�^�E�o�b�t�@
Public ODR_HANSEIHIN_O_REC  As ODR_HANSEIHIN_O_REC_Tag


Public Type ODR_HANSEIHIN_K_REC_Tag                 '�qں���
    
    USE_YM(0 To 5)          As Byte         '�g�p��
    INPUT_NO(0 To 3)        As Byte         '���͏�
    USE_YMD(0 To 7)         As Byte         '�g�p���t
    SEQNO(0 To 2)           As Byte         '�ǔ�(000)
    KO_JGYOBU(0 To 0)       As Byte         '���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�e�i��
    KO_QTY(0 To 7)          As Byte         '�w���� 9(5)V99
    USE_QTY(0 To 11)        As Byte         '�w���� S9(8)V99
    ZAITEI_F(0 To 0)        As Byte         '�ݒ��}�[�N
    UPD_TANTO(0 To 4)       As Byte         '�X�V�@�S����
    UPD_DATE(0 To 7)        As Byte         '�X�V�@���t
    UPD_TIME(0 To 5)        As Byte         '�X�V�@����
    
    
    IO_FLG(0 To 0)          As Byte         '���o���׸�     '2008.05.14
    
    
    FILLER(0 To 171)        As Byte
    

End Type
'�f�[�^�E�o�b�t�@
Public ODR_HANSEIHIN_K_REC  As ODR_HANSEIHIN_K_REC_Tag

'�L�[��`

Type KEY0_ODR_HANSEIHIN                           '�j�d�x�O
'    USE_YM(0 To 5)          As Byte         '�g�p��            2008.05.13
    INPUT_NO(0 To 3)        As Byte         '���͏�
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
Type KEY1_ODR_HANSEIHIN                           '�j�d�x�P
    KO_JGYOBU(0 To 0)       As Byte         '���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�e�i��
End Type
    
Type KEY2_ODR_HANSEIHIN                           '�j�d�x�Q
    USE_YMD(0 To 7)         As Byte         '�g�p���t
End Type
    
    
    
'�L�[�E�f�[�^
Public K0_ODR_HANSEIHIN     As KEY0_ODR_HANSEIHIN
Public K1_ODR_HANSEIHIN     As KEY1_ODR_HANSEIHIN
Public K2_ODR_HANSEIHIN     As KEY2_ODR_HANSEIHIN

Type ODR_HANSEIHIN_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����

    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����


End Type

Private ODR_HANSEIHIN_Speck As ODR_HANSEIHIN_FSpeck
Private Function ODR_HANSEIHIN_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �����i�Ǘ��f�[�^  �b�q�d�`�s�d                      *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ODR_HANSEIHIN_Create = True
                                            '�����i�Ǘ��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", ODR_HANSEIHIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_HANSEIHIN]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    ODR_HANSEIHIN_Speck.fs.recoleng = Len(ODR_HANSEIHIN_O_REC)      ' ���R�[�h��
    ODR_HANSEIHIN_Speck.fs.PageSize = ODR_HANSEIHIN_PG_SIZ          ' �y�[�W�T�C�Y
    ODR_HANSEIHIN_Speck.fs.idexnumb = 3                             ' �C���f�b�N�X��
    ODR_HANSEIHIN_Speck.fs.fileflag = 0                             ' �t�@�C���t���O
    ODR_HANSEIHIN_Speck.fs.reserve = &H0                            ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    ODR_HANSEIHIN_Speck.ks0.keypos = 7                      ' �L�[�|�W�V����
    ODR_HANSEIHIN_Speck.ks0.keyleng = 4                     ' �L�[��
    ODR_HANSEIHIN_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    ODR_HANSEIHIN_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ODR_HANSEIHIN_Speck.ks0.reserve = &H0                   ' �\��ς�
    
    ODR_HANSEIHIN_Speck.ks1.keypos = 19                     ' �L�[�|�W�V����
    ODR_HANSEIHIN_Speck.ks1.keyleng = 3                     ' �L�[��
    ODR_HANSEIHIN_Speck.ks1.keyflag = BtKfExt               ' �L�[�t���O
    ODR_HANSEIHIN_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ODR_HANSEIHIN_Speck.ks1.reserve = &H0                   ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    ODR_HANSEIHIN_Speck.ks2.keypos = 22                     ' �L�[�|�W�V����
    ODR_HANSEIHIN_Speck.ks2.keyleng = 1                     ' �L�[��
                                                            ' �L�[�t���O
    ODR_HANSEIHIN_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ODR_HANSEIHIN_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ODR_HANSEIHIN_Speck.ks2.reserve = &H0                   ' �\��ς�
    
    ODR_HANSEIHIN_Speck.ks3.keypos = 23                     ' �L�[�|�W�V����
    ODR_HANSEIHIN_Speck.ks3.keyleng = 1                     ' �L�[��
                                                            ' �L�[�t���O
    ODR_HANSEIHIN_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ODR_HANSEIHIN_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ODR_HANSEIHIN_Speck.ks3.reserve = &H0                   ' �\��ς�
    
    ODR_HANSEIHIN_Speck.ks4.keypos = 24                     ' �L�[�|�W�V����
    ODR_HANSEIHIN_Speck.ks4.keyleng = 20                    ' �L�[��
                                                            ' �L�[�t���O
    ODR_HANSEIHIN_Speck.ks4.keyflag = BtKfExt + BtKfDup + BtKfChg
    ODR_HANSEIHIN_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ODR_HANSEIHIN_Speck.ks4.reserve = &H0                   ' �\��ς�
    '--------------------------------------------------- �L�[�P ��
    
    '--------------------------------------------------- �L�[�Q ��
    ODR_HANSEIHIN_Speck.ks5.keypos = 11                     ' �L�[�|�W�V����
    ODR_HANSEIHIN_Speck.ks5.keyleng = 8                    ' �L�[��
                                                            ' �L�[�t���O
    ODR_HANSEIHIN_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg
    ODR_HANSEIHIN_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ODR_HANSEIHIN_Speck.ks5.reserve = &H0                   ' �\��ς�
    '--------------------------------------------------- �L�[�Q ��
    
    
    sts = BTRV(BtOpCreate, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_Speck, Len(ODR_HANSEIHIN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�����i�Ǘ��f�[�^")
        Exit Function
    End If
    
    ODR_HANSEIHIN_Create = False

End Function

Public Function ODR_HANSEIHIN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �����i�Ǘ��f�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ODR_HANSEIHIN_Open = True
                                            '�����i�Ǘ��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", ODR_HANSEIHIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_HANSEIHIN]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ODR_HANSEIHIN_Create()    '�����i�Ǘ��f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�����i�Ǘ��f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�����i�Ǘ��f�[�^")
                Exit Function
        End Select
    Loop
    
    ODR_HANSEIHIN_Open = False

End Function
