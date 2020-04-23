Attribute VB_Name = "P_SHUKEIRE"
Option Explicit

'********************************************************************
'*
'*              ���ގ�������ް�  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const P_SHUKEIRE_ID$ = "P_SHUKEIRE"

'�y�[�W�T�C�Y
Private Const P_SHUKEIRE_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SHUKEIRE_POS       As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type P_SHUKEIRE_REC_Tag
    
    ORDER_NO(0 To 4)        As Byte         '������
    SEQNO(0 To 2)           As Byte         '�ǔ�
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
    UKEIRE_DT(0 To 7)       As Byte         '�����
    UKEIRE_QTY(0 To 11)     As Byte         '�������(S9(8)V99)
    UKEIRE_TANKA(0 To 10)   As Byte         '����P��(9(8)V99)
    UKEIRE_KINGAKU(0 To 8)  As Byte         '������z(S9(8))
    LAST_F(0 To 0)          As Byte         '�ŏI����׸� 0:�p�� 1:�ŏI
    KEIJYO_YM(0 To 5)       As Byte         '�v��N��(YYYYMM)
    ZEI_KIN(0 To 8)         As Byte         '����Ŋz(S9(8))    2007.04.29
    FILLER(0 To 44)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_SHUKEIRE_REC       As P_SHUKEIRE_REC_Tag

'�L�[��`

Public Type KEY0_P_SHUKEIRE                     '�j�d�x�O
    ORDER_NO(0 To 4)        As Byte         '������
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type

Public Type KEY1_P_SHUKEIRE                     '�j�d�x�P
    KEIJYO_YM(0 To 5)       As Byte         '�v��N��(YYYYMM)
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
    UKEIRE_DT(0 To 7)        As Byte         '������
End Type
    
Public Type KEY2_P_SHUKEIRE                     '�j�d�x�Q
    KEIJYO_YM(0 To 5)       As Byte         '�v��N��(YYYYMM)
    UKEIRE_DT(0 To 7)        As Byte         '������
End Type
    
    
'�L�[�E�f�[�^
Public K0_P_SHUKEIRE        As KEY0_P_SHUKEIRE
Public K1_P_SHUKEIRE        As KEY1_P_SHUKEIRE
Public K2_P_SHUKEIRE        As KEY2_P_SHUKEIRE

Type P_SHUKEIRE_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����
    ks6                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_SHUKEIRE_Speck    As P_SHUKEIRE_FSpeck
Private Function P_SHUKEIRE_Create() As Integer
'********************************************************************
'*
'*              ���ގ�������ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SHUKEIRE_Create = True
                                            '���ގ�������ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHUKEIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHUKEIRE]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SHUKEIRE_Speck.fs.recoleng = Len(P_SHUKEIRE_REC)  ' ���R�[�h��
    P_SHUKEIRE_Speck.fs.PageSize = P_SHUKEIRE_PG_SIZ    ' �y�[�W�T�C�Y
    P_SHUKEIRE_Speck.fs.idexnumb = 3                    ' �C���f�b�N�X��
    P_SHUKEIRE_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    P_SHUKEIRE_Speck.fs.reserve = &H0                   ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_SHUKEIRE_Speck.ks0.keypos = 1                     ' �L�[�|�W�V����
    P_SHUKEIRE_Speck.ks0.keyleng = 5                    ' �L�[��
    P_SHUKEIRE_Speck.ks0.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    P_SHUKEIRE_Speck.ks0.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHUKEIRE_Speck.ks0.reserve = &H0                  ' �\��ς�
    
    P_SHUKEIRE_Speck.ks1.keypos = 6                     ' �L�[�|�W�V����
    P_SHUKEIRE_Speck.ks1.keyleng = 3                    ' �L�[��
    P_SHUKEIRE_Speck.ks1.keyflag = BtKfExt              ' �L�[�t���O
    P_SHUKEIRE_Speck.ks1.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHUKEIRE_Speck.ks1.reserve = &H0                  ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    P_SHUKEIRE_Speck.ks2.keypos = 55                    ' �L�[�|�W�V����
    P_SHUKEIRE_Speck.ks2.keyleng = 6                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHUKEIRE_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SHUKEIRE_Speck.ks2.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHUKEIRE_Speck.ks2.reserve = &H0                  ' �\��ς�
    
    P_SHUKEIRE_Speck.ks3.keypos = 9                     ' �L�[�|�W�V����
    P_SHUKEIRE_Speck.ks3.keyleng = 5                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHUKEIRE_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SHUKEIRE_Speck.ks3.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHUKEIRE_Speck.ks3.reserve = &H0                  ' �\��ς�
    
    P_SHUKEIRE_Speck.ks4.keypos = 14                    ' �L�[�|�W�V����
    P_SHUKEIRE_Speck.ks4.keyleng = 8                    ' �L�[��
    P_SHUKEIRE_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfDup   ' �L�[�t���O
    P_SHUKEIRE_Speck.ks4.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHUKEIRE_Speck.ks4.reserve = &H0                  ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�Q ��
    
    
    '--------------------------------------------------- �L�[�P ��
    P_SHUKEIRE_Speck.ks5.keypos = 55                    ' �L�[�|�W�V����
    P_SHUKEIRE_Speck.ks5.keyleng = 6                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHUKEIRE_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SHUKEIRE_Speck.ks5.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHUKEIRE_Speck.ks5.reserve = &H0                  ' �\��ς�
    
    
    P_SHUKEIRE_Speck.ks6.keypos = 14                    ' �L�[�|�W�V����
    P_SHUKEIRE_Speck.ks6.keyleng = 8                    ' �L�[��
    P_SHUKEIRE_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfDup   ' �L�[�t���O
    P_SHUKEIRE_Speck.ks6.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHUKEIRE_Speck.ks6.reserve = &H0                  ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�Q ��
    
    
    sts = BTRV(BtOpCreate, P_SHUKEIRE_POS, P_SHUKEIRE_Speck, Len(P_SHUKEIRE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ގ���ް�")
        Exit Function
    End If
    
    P_SHUKEIRE_Create = False

End Function

Public Function P_SHUKEIRE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ގ���ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SHUKEIRE_Open = True
                                            '���ގ���ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHUKEIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHUKEIRE]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHUKEIRE_Create()   '���ގ���ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ގ���ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ގ���ް�")
                Exit Function
        End Select
    Loop
    
    P_SHUKEIRE_Open = False

End Function

