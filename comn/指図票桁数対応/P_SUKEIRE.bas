Attribute VB_Name = "P_SUKEIRE"
Option Explicit

'********************************************************************
'*
'*              ���i���w�}��������f�[�^  �t�@�C����`
'*
'*          CREATE 2005.12.14
'********************************************************************
'�t�@�C���h�c
Public Const P_SUKEIRE_ID$ = "P_SUKEIRE"

'�y�[�W�T�C�Y
Private Const P_SUKEIRE_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SUKEIRE_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************

Private Type GENKA_TBL_Tag          '��������ð���
    NIN(0 To 2)             As Byte         '�l��
    TIMES(0 To 5)           As Byte         '����
End Type




'���R�[�h��`
Public Type P_SUKEIRE_REC_Tag
    
    SHIJI_NO(0 To 4)       As Byte         '�w�}�[��  ���g�p�Ƃ��� 2007.11.28
    SEQNO(0 To 2)           As Byte         '�ǔ�
    SHIMUKE_CODE(0 To 1)    As Byte         '�d�����溰��
    UKEIRE_DT(0 To 7)       As Byte         '�����
    UKEIRE_QTY(0 To 10)     As Byte         '�������(9(8)V999)
                                            '��������
    GENKA_TBL(0 To 9)       As GENKA_TBL_Tag
    
    JISEKI_NAME(0 To 19)    As Byte         '���ӗv����
    JISEKI_NIN(0 To 2)      As Byte         '����  �l
    JISEKI_TIMES(0 To 5)    As Byte         '����  ��
    TASEKI_NAME(0 To 19)    As Byte         '���ӗv����
    TASEKI_NIN(0 To 2)      As Byte         '����  �l
    TASEKI_TIMES(0 To 5)    As Byte         '����  ��
    
    LAST_F(0 To 0)          As Byte         '�ŏI����׸� 0:�p�� 1:�ŏI
    TORI_CODE(0 To 4)       As Byte         '�����
    
    'SHIJI_NO(0 To 7)        As Byte         '�w�}�[��   2007.11.28
    FILLER(0 To 94)         As Byte         'Filler     2007.11.28
    
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_SUKEIRE_REC        As P_SUKEIRE_REC_Tag

'�L�[��`

Type KEY0_P_SUKEIRE                         '�j�d�x�O
'    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��  2007.11.28
    SHIJI_NO(0 To 7)        As Byte         '�w�}�[��   2007.11.28
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
Type KEY1_P_SUKEIRE                         '�j�d�x�P
    SHIMUKE_CODE(0 To 1)    As Byte         '�d�����溰��
    UKEIRE_DT(0 To 7)       As Byte         '�����
End Type
    
Type KEY2_P_SUKEIRE                         '�j�d�x�Q
    TORI_CODE(0 To 4)       As Byte         '�����
    UKEIRE_DT(0 To 7)       As Byte         '�����
End Type
    
'�L�[�E�f�[�^
Public K0_P_SUKEIRE         As KEY0_P_SUKEIRE
Public K1_P_SUKEIRE         As KEY1_P_SUKEIRE
Public K2_P_SUKEIRE         As KEY2_P_SUKEIRE

Type P_SUKEIRE_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_SUKEIRE_Speck    As P_SUKEIRE_FSpeck
Private Function P_SUKEIRE_Create() As Integer
'********************************************************************
'*
'*              ���i���w�}��������ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SUKEIRE_Create = True
                                            '���i���w�}��������ް��t���p�X�捞��
    sts = GetIni("FILE", P_SUKEIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SUKEIRE]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SUKEIRE_Speck.fs.recoleng = Len(P_SUKEIRE_REC)    ' ���R�[�h��
    P_SUKEIRE_Speck.fs.PageSize = P_SUKEIRE_PG_SIZ      ' �y�[�W�T�C�Y
    P_SUKEIRE_Speck.fs.idexnumb = 3                     ' �C���f�b�N�X��
    P_SUKEIRE_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    P_SUKEIRE_Speck.fs.reserve = &H0                    ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
'2007.11.28    P_SUKEIRE_Speck.ks0.keypos = 1                      ' �L�[�|�W�V����
'2007.11.28    P_SUKEIRE_Speck.ks0.keyleng = 5                     ' �L�[��
    
    P_SUKEIRE_Speck.ks0.keypos = 184                    ' �L�[�|�W�V����    2007.11.28
    P_SUKEIRE_Speck.ks0.keyleng = 8                     ' �L�[��            2007.11.28
    
    
    P_SUKEIRE_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    P_SUKEIRE_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SUKEIRE_Speck.ks0.reserve = &H0                   ' �\��ς�
    
    P_SUKEIRE_Speck.ks1.keypos = 6                      ' �L�[�|�W�V����
    P_SUKEIRE_Speck.ks1.keyleng = 3                     ' �L�[��
    P_SUKEIRE_Speck.ks1.keyflag = BtKfExt               ' �L�[�t���O
    P_SUKEIRE_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SUKEIRE_Speck.ks1.reserve = &H0                   ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    P_SUKEIRE_Speck.ks2.keypos = 9                      ' �L�[�|�W�V����
    P_SUKEIRE_Speck.ks2.keyleng = 2                     ' �L�[��
                                                        ' �L�[�t���O
    P_SUKEIRE_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SUKEIRE_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SUKEIRE_Speck.ks2.reserve = &H0                   ' �\��ς�
    
    
    P_SUKEIRE_Speck.ks3.keypos = 11                     ' �L�[�|�W�V����
    P_SUKEIRE_Speck.ks3.keyleng = 8                     ' �L�[��
                                                        ' �L�[�t���O
    P_SUKEIRE_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_SUKEIRE_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SUKEIRE_Speck.ks3.reserve = &H0                   ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�P ��
    
    '--------------------------------------------------- �L�[�Q ��
    P_SUKEIRE_Speck.ks4.keypos = 179                    ' �L�[�|�W�V����
    P_SUKEIRE_Speck.ks4.keyleng = 5                     ' �L�[��
                                                        ' �L�[�t���O
    P_SUKEIRE_Speck.ks4.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SUKEIRE_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SUKEIRE_Speck.ks4.reserve = &H0                   ' �\��ς�
    
    
    P_SUKEIRE_Speck.ks5.keypos = 11                     ' �L�[�|�W�V����
    P_SUKEIRE_Speck.ks5.keyleng = 8                     ' �L�[��
                                                        ' �L�[�t���O
    P_SUKEIRE_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_SUKEIRE_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SUKEIRE_Speck.ks5.reserve = &H0                   ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�O ��
    
    
    sts = BTRV(BtOpCreate, P_SUKEIRE_POS, P_SUKEIRE_Speck, Len(P_SUKEIRE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���i���w�}��������ް�")
        Exit Function
    End If
    
    P_SUKEIRE_Create = False

End Function

Public Function P_SUKEIRE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���i���w�}��������ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SUKEIRE_Open = True
                                            '���i���w�}��������ް��t���p�X�捞��
    sts = GetIni("FILE", P_SUKEIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SUKEIRE]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SUKEIRE_Create()    '���i���w�}��������ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���i���w�}��������ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���i���w�}��������ް�")
                Exit Function
        End Select
    Loop
    
    P_SUKEIRE_Open = False

End Function

