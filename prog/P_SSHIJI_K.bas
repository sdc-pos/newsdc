Attribute VB_Name = "P_SSHIJI_K"
Option Explicit

'********************************************************************
'*
'*              ���i���w�}�f�[�^�i�q�j  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c


Public Const P_SSHIJI_K_ID$ = "P_SSHIJI_K"

'�y�[�W�T�C�Y
Private Const P_SSHIJI_K_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SSHIJI_K_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************

'���R�[�h��`
Public Type P_SSHIJI_K_REC_Tag
    
    xSHIJI_NO(0 To 4)        As Byte        '�w�}�[�� ���g�p�Ƃ��� 2007.11.28
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
    KO_SYUBETSU(0 To 1)     As Byte         '�q�@���
    KO_JGYOBU(0 To 0)       As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�q�@�i��
    KO_QTY(0 To 5)          As Byte         '�q�@����(999V99)
    KO_SHIJI_QTY(0 To 10)   As Byte         '�w����(9(8)V99)
    KO_BIKOU(0 To 39)       As Byte         '�q�@���l
'    KO_ID_NO(0 To 7)        As Byte        '�q �h�c�Q�m�n
    KO_ID_NO(0 To 11)       As Byte         '�q �h�c�Q�m�n (8����12��)  2006/05/24
    CALCEL_F(0 To 0)        As Byte         '��ݾ�F
    CANCEL_DATETIME(0 To 13) As Byte        '��ݾٓ���
'    FILLER(0 To 64)         As Byte         'Filler
    
    SHIJI_No(0 To 7)        As Byte         '�w�}�[��   2007.11.28
    
    
    HIKIATE_QTY(0 To 10)    As Byte         '�݌Ɉ����� 2012.03.09
    IDO_SUMI(0 To 0)        As Byte         '�ړ��ς� ��:���@9:�ς� 2012.03.09
    
    ST_TANABAN(0 To 7)      As Byte         '�W���I�� 2012.03.18
        
        
    IDO_SUMI_QTY(0 To 10)   As Byte         '�ړ��ςݐ��� 2012.04.13
        
        
        
    COMPO_TANTO(0 To 4)     As Byte         '�\������   �S����          2012.04.20
    COMPO_YMDHS(0 To 11)    As Byte         '           ����            2012.04.20
    COMPO_Sumi_Cnt(0 To 1)  As Byte         '           �����ςݐ�      2012.04.20
    COMPO_ALL_Cnt(0 To 1)   As Byte         '           �\����          2012.04.20
        
    FILLER(0 To 0)          As Byte         'Filler 2007.11.28  2012.04.20 �����ύX
    
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_SSHIJI_K_REC       As P_SSHIJI_K_REC_Tag

'�L�[��`

Type KEY0_P_SSHIJI_K                        '�j�d�x�O
'    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��  2007.11.28
    SHIJI_No(0 To 7)        As Byte         '�w�}�[��   '2007.11.28
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
    
Type KEY1_P_SSHIJI_K                        '�j�d�x�P
    KO_JGYOBU(0 To 0)       As Byte         '�q�@���ƕ�
'    KO_ID_NO(0 To 7)        As Byte         '�q �h�c�Q�m�n
    KO_ID_NO(0 To 11)       As Byte         '�q �h�c�Q�m�n (8����12��)  2006/05/24
End Type
    
Type KEY2_P_SSHIJI_K                        '�j�d�x�Q   2012.03.09
    KO_JGYOBU(0 To 0)       As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�q�@�i��
    IDO_SUMI(0 To 0)        As Byte         '�ړ��ς� ��:���@9:�ς�
End Type
    
Type KEY3_P_SSHIJI_K                        '�j�d�x�R   2012.03.18
    SHIJI_No(0 To 7)        As Byte         '�w�}�[��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    ST_TANABAN(0 To 7)      As Byte         '�W���I��
End Type
    
    
    
    
'�L�[�E�f�[�^
Public K0_P_SSHIJI_K        As KEY0_P_SSHIJI_K
Public K1_P_SSHIJI_K        As KEY1_P_SSHIJI_K
Public K2_P_SSHIJI_K        As KEY2_P_SSHIJI_K  '2012.03.09
Public K3_P_SSHIJI_K        As KEY3_P_SSHIJI_K  '2012.03.18

Type P_SSHIJI_K_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����

    ks5                     As BtKeySpeck   ' �� ��߯��\����    2012.03.09
    ks6                     As BtKeySpeck   ' �� ��߯��\����    2012.03.09
    ks7                     As BtKeySpeck   ' �� ��߯��\����    2012.03.09
    ks8                     As BtKeySpeck   ' �� ��߯��\����    2012.03.09

    ks9                     As BtKeySpeck   ' �� ��߯��\����    2012.03.18
    ks10                    As BtKeySpeck   ' �� ��߯��\����    2012.03.18
    ks11                    As BtKeySpeck   ' �� ��߯��\����    2012.03.18


End Type

Private P_SSHIJI_K_Speck    As P_SSHIJI_K_FSpeck
Private Function P_SSHIJI_K_Create() As Integer
'********************************************************************
'*
'*              ���i���w�}�f�[�^�i�q�j  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SSHIJI_K_Create = True
                                            '��z�w�}�f�[�^�i�q�j�t���p�X�捞��
    sts = GetIni("FILE", P_SSHIJI_K_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_K]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SSHIJI_K_Speck.fs.recoleng = Len(P_SSHIJI_K_REC)  ' ���R�[�h��
    P_SSHIJI_K_Speck.fs.PageSize = P_SSHIJI_K_PG_SIZ    ' �y�[�W�T�C�Y
    P_SSHIJI_K_Speck.fs.idexnumb = 4                    ' �C���f�b�N�X��
    P_SSHIJI_K_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    P_SSHIJI_K_Speck.fs.reserve = &H0                   ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
'2008.11.28    P_SSHIJI_K_Speck.ks0.keypos = 1                     ' �L�[�|�W�V����
'2008.11.28    P_SSHIJI_K_Speck.ks0.keyleng = 5                    ' �L�[��
    
    P_SSHIJI_K_Speck.ks0.keypos = 118                   ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks0.keyleng = 8                    ' �L�[��
    
    
    P_SSHIJI_K_Speck.ks0.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    P_SSHIJI_K_Speck.ks0.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks0.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_K_Speck.ks1.keypos = 6                     ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks1.keyleng = 1                    ' �L�[��
    P_SSHIJI_K_Speck.ks1.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    P_SSHIJI_K_Speck.ks1.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks1.reserve = &H0                  ' �\��ς�
    
    
    P_SSHIJI_K_Speck.ks2.keypos = 7                     ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks2.keyleng = 3                    ' �L�[��
    P_SSHIJI_K_Speck.ks2.keyflag = BtKfExt              ' �L�[�t���O
    P_SSHIJI_K_Speck.ks2.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks2.reserve = &H0                  ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    P_SSHIJI_K_Speck.ks3.keypos = 12                    ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks3.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_K_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfSeg
    P_SSHIJI_K_Speck.ks3.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks3.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_K_Speck.ks4.keypos = 91                    ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks4.keyleng = 12                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_K_Speck.ks4.keyflag = BtKfExt + BtKfDup
    P_SSHIJI_K_Speck.ks4.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks4.reserve = &H0                  ' �\��ς�
    '--------------------------------------------------- �L�[�P ��
    
    
    '--------------------------------------------------- �L�[�Q ��  2012.03.09
    P_SSHIJI_K_Speck.ks5.keypos = 12                    ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks5.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_K_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SSHIJI_K_Speck.ks5.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks5.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_K_Speck.ks6.keypos = 13                    ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks6.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_K_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SSHIJI_K_Speck.ks6.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks6.reserve = &H0                  ' �\��ς�
    
    
    P_SSHIJI_K_Speck.ks7.keypos = 14                    ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks7.keyleng = 20                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_K_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SSHIJI_K_Speck.ks7.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks7.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_K_Speck.ks8.keypos = 137                   ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks8.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_K_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_SSHIJI_K_Speck.ks8.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks8.reserve = &H0                  ' �\��ς�
    
    '--------------------------------------------------- �L�[�Q ��
    
    '--------------------------------------------------- �L�[�R  ��  2012.03.09
    P_SSHIJI_K_Speck.ks9.keypos = 118                   ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks9.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_K_Speck.ks9.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SSHIJI_K_Speck.ks9.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks9.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_K_Speck.ks10.keypos = 6                    ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks10.keyleng = 1                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_K_Speck.ks10.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SSHIJI_K_Speck.ks10.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks10.reserve = &H0                 ' �\��ς�
    
    
    P_SSHIJI_K_Speck.ks11.keypos = 138                  ' �L�[�|�W�V����
    P_SSHIJI_K_Speck.ks11.keyleng = 8                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_K_Speck.ks11.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_SSHIJI_K_Speck.ks11.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_K_Speck.ks11.reserve = &H0                 ' �\��ς�
    '--------------------------------------------------- �L�[�R ��
    
    
    sts = BTRV(BtOpCreate, P_SSHIJI_K_POS, P_SSHIJI_K_Speck, Len(P_SSHIJI_K_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "��z�w�}�f�[�^�i�q�j")
        Exit Function
    End If
    
    P_SSHIJI_K_Create = False

End Function

Public Function P_SSHIJI_K_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���i���w�}�f�[�^�i�q�j  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SSHIJI_K_Open = True
                                            '��z�w�}�f�[�^�i�q�j�t���p�X�捞��
    sts = GetIni("FILE", P_SSHIJI_K_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_K]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SSHIJI_K_Create()   '��z�w�}�f�[�^�i�q�j�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "��z�w�}�f�[�^�i�q�j�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "��z�w�}�f�[�^�i�q�j�}�X�^")
                Exit Function
        End Select
    Loop
    
    P_SSHIJI_K_Open = False

End Function