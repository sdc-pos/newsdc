Attribute VB_Name = "P_SSHIJI_O"
Option Explicit

'********************************************************************
'*
'*              ���i���w�}�f�[�^�i�e�j  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const P_SSHIJI_O_ID$ = "P_SSHIJI_O"

'�y�[�W�T�C�Y
Private Const P_SSHIJI_O_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public P_SSHIJI_O_POS As POSBLK
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
Public Type P_SSHIJI_O_REC_Tag
    
    xSHIJI_NO(0 To 4)       As Byte         '�w�}�[��   ���g�p�Ƃ��� 2007.11.28
    HAKKO_DT(0 To 7)        As Byte         '���s��
    Print_datetime(0 To 13) As Byte         '���s����
    TANTO_CODE(0 To 4)      As Byte         '�S���Һ���
    SHONIN_CODE(0 To 4)     As Byte         '���F�Һ���
    SHIMUKE_CODE(0 To 1)    As Byte         '�d�����溰��
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    SHIJI_QTY(0 To 10)      As Byte         '�w����(9(8)V99)
    UKEHARAI_CODE(0 To 4)   As Byte         '��z�溰��
    S_CLASS_CODE(0 To 19)   As Byte         '���i���׽
    F_CLASS_CODE(0 To 19)   As Byte         '�t���׽
    N_CLASS_CODE(0 To 19)   As Byte         '���E�׽
    S_TANTO(0 To 1)         As Byte         '���P�^�S���҃R�[�h
    SAMPLE_F(0 To 0)        As Byte         '���{�쐬
    SHIJI_F(0 To 0)         As Byte         '�w���`�� 0:�ʏ�@1:��߯ā@2�F���i���� 3:�č���(2007.11.09)
    TORI_KBN(0 To 0)        As Byte         '�����R�[�h
    
    PRI_SHIJI(0 To 0)       As Byte         '�o�͑Ώ� �w�}�[
    PRI_PARTS(0 To 0)       As Byte         '�o�͑Ώ� �߰�����
    PRI_GAISOU(0 To 0)      As Byte         '�o�͑Ώ� �O������
    PRI_KISHU(0 To 0)       As Byte         '�o�͑Ώ� �@������
    
    BIKOU(0 To 119)         As Byte         '���l
    
    
    KAN_F(0 To 0)           As Byte         '����F
    KAN_DT(0 To 7)          As Byte         '������
    BUNNOU_CNT(0 To 1)      As Byte         '���[��
    UKEIRE_QTY(0 To 10)     As Byte         '������i���v�j
                                            '��������
    GENKA_TBL(0 To 9)       As GENKA_TBL_Tag
    
    JISEKI_NAME(0 To 19)    As Byte         '���ӗv����
    JISEKI_NIN(0 To 2)      As Byte         '����  �l
    JISEKI_TIMES(0 To 5)    As Byte         '����  ��
    TASEKI_NAME(0 To 19)    As Byte         '���ӗv����
    TASEKI_NIN(0 To 2)      As Byte         '����  �l
    TASEKI_TIMES(0 To 5)    As Byte         '����  ��
    
    
    CANCEL_F(0 To 0)        As Byte         '��ݾ�F
    CANCEL_DATETIME(0 To 13) As Byte        '��ݾٓ���
    
    ORDER_DT(0 To 7)        As Byte         '�󒍓�(������) 2007.02.20
    
    
    SHIJI_No(0 To 7)        As Byte         '�w�}�[��   ���g�p�Ƃ��� 2007.11.28
    
    
    HIN_CHECK_TANTO(0 To 4) As Byte         '�i�������S���Һ��� 2010.09.03
    HIN_CHECK_DATETIME(0 To 13) _
                            As Byte         '�i���������� 2010.09.03
    HIN_CHECK_LABEL_CNT(0 To 2) _
                            As Byte         '�i���������ٌ���   2010.09.03
    HIN_CHECK_GENPIN_CNT(0 To 2) _
                            As Byte         '�i���������i�[����   2010.09.03
            
    ORDER_DT_SEQ(0 To 1)    As Byte         '�󒍓�(������)�}�� 2012.03.27
            
    COMPO_END_F(0 To 0)     As Byte         '�\����������F(���PC) 9:���� 2012.04.13
    
'    FILLER(0 To 2)          As Byte         'Filler 2011.04.13         --> 2015.11.07
    
    HIN_CHECK_GAISOU_CNT(0 To 2) _
                            As Byte         '�i�������O���i�Ԍ���   2015.11.07
    
    
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_SSHIJI_O_REC       As P_SSHIJI_O_REC_Tag

'�L�[��`

Type KEY0_P_SSHIJI_O                        '�j�d�x�O
'    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��
    SHIJI_No(0 To 7)        As Byte         '�w�}�[��   2007.11.28
End Type

Type KEY1_P_SSHIJI_O                        '�j�d�x�P
    KAN_F(0 To 0)           As Byte         '����F
    SHIMUKE_CODE(0 To 1)    As Byte         '�d�����溰��
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    KAN_DT(0 To 7)          As Byte         '������
'    SHIJI_NO(0 To 4)       As Byte         '�w�}�[��
    SHIJI_No(0 To 7)        As Byte         '�w�}�[��   2007.11.28
End Type
    
Type KEY2_P_SSHIJI_O                        '�j�d�x�Q
    ORDER_DT(0 To 7)        As Byte         '�󒍓� 2007.02.20
End Type
    
Type KEY3_P_SSHIJI_O                        '�j�d�x�R   2007.11.14
    HAKKO_DT(0 To 7)        As Byte         '���s��
    TORI_KBN(0 To 0)        As Byte         '�����敪
    UKEHARAI_CODE(0 To 4)   As Byte         '��z�溰��
End Type
    
Type KEY4_P_SSHIJI_O                        '�j�d�x�S   2011.11.11
    SHIMUKE_CODE(0 To 1)    As Byte         '�d�����溰��
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    Print_datetime(0 To 13) As Byte         '���s����
End Type
    
    
Type KEY5_P_SSHIJI_O                        '�j�d�x�T   2012.03.08
    SHIMUKE_CODE(0 To 1)    As Byte         '�d�����溰��
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    KAN_F(0 To 0)           As Byte         '����F
End Type
    
Type KEY6_P_SSHIJI_O                        '�j�d�x�U   2012.03.27
    ORDER_DT(0 To 7)        As Byte         '�󒍓�(������)     2007.02.20
    ORDER_DT_SEQ(0 To 1)    As Byte         '�󒍓�(������)�}�� 2012.03.27
End Type
    
    
    
    
    
'�L�[�E�f�[�^
Public K0_P_SSHIJI_O        As KEY0_P_SSHIJI_O
Public K1_P_SSHIJI_O        As KEY1_P_SSHIJI_O
Public K2_P_SSHIJI_O        As KEY2_P_SSHIJI_O
Public K3_P_SSHIJI_O        As KEY3_P_SSHIJI_O      '2007.11.14

Public K4_P_SSHIJI_O        As KEY4_P_SSHIJI_O      '2011.11.11

Public K5_P_SSHIJI_O        As KEY5_P_SSHIJI_O      '2012.03.08

Public K6_P_SSHIJI_O        As KEY6_P_SSHIJI_O      '2012.03.27



Type P_SSHIJI_O_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����
    ks6                     As BtKeySpeck   ' �� ��߯��\����
    ks7                     As BtKeySpeck   ' �� ��߯��\����
    ks8                     As BtKeySpeck   ' �� ��߯��\����
    ks9                     As BtKeySpeck   ' �� ��߯��\����
    ks10                    As BtKeySpeck   ' �� ��߯��\����
    ks11                    As BtKeySpeck   ' �� ��߯��\����

    '2011.11.11
    ks12                    As BtKeySpeck   ' �� ��߯��\����
    ks13                    As BtKeySpeck   ' �� ��߯��\����
    ks14                    As BtKeySpeck   ' �� ��߯��\����
    ks15                    As BtKeySpeck   ' �� ��߯��\����
    ks16                    As BtKeySpeck   ' �� ��߯��\����
    '2011.11.11

    '2012.03.08
    ks17                    As BtKeySpeck   ' �� ��߯��\����
    ks18                    As BtKeySpeck   ' �� ��߯��\����
    ks19                    As BtKeySpeck   ' �� ��߯��\����
    ks20                    As BtKeySpeck   ' �� ��߯��\����
    ks21                    As BtKeySpeck   ' �� ��߯��\����
    '2012.03.08


    '2012.03.27
    ks22                    As BtKeySpeck   ' �� ��߯��\����
    ks23                    As BtKeySpeck   ' �� ��߯��\����
    '2012.03.27

End Type

Private P_SSHIJI_O_Speck    As P_SSHIJI_O_FSpeck
Private Function P_SSHIJI_O_Create() As Integer
'********************************************************************
'*
'*              ���i���w�}(�e)�ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*      2007.11.14  :KEY3(���s��+�����敪+��z��R�[�h)�@�ǉ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SSHIJI_O_Create = True
                                            '�R�[�h�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_O]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SSHIJI_O_Speck.fs.recoleng = Len(P_SSHIJI_O_REC)  ' ���R�[�h��
    P_SSHIJI_O_Speck.fs.PageSize = P_SSHIJI_O_PG_SIZ    ' �y�[�W�T�C�Y
    P_SSHIJI_O_Speck.fs.idexnumb = 7                    ' �C���f�b�N�X��
    P_SSHIJI_O_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    P_SSHIJI_O_Speck.fs.reserve = &H0                   ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
'2007.11.28    P_SSHIJI_O_Speck.ks0.keypos = 1              ' �L�[�|�W�V����
'2007.11.28    P_SSHIJI_O_Speck.ks0.keyleng = 5             ' �L�[��
    
    
    P_SSHIJI_O_Speck.ks0.keypos = 460                   ' �L�[�|�W�V����    2007.11.28
    P_SSHIJI_O_Speck.ks0.keyleng = 8                    ' �L�[��            2007.11.28
    
    P_SSHIJI_O_Speck.ks0.keyflag = BtKfExt              ' �L�[�t���O
    P_SSHIJI_O_Speck.ks0.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks0.reserve = &H0                  ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    
    '--------------------------------------------------- �L�[�P ��
    P_SSHIJI_O_Speck.ks1.keypos = 267                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks1.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks1.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks1.reserve = &H0                  ' �\��ς�
    
    
    P_SSHIJI_O_Speck.ks2.keypos = 38                    ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks2.keyleng = 2                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks2.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks2.reserve = &H0                  ' �\��ς�
    
    
    P_SSHIJI_O_Speck.ks3.keypos = 40                    ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks3.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks3.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks3.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_O_Speck.ks4.keypos = 41                    ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks4.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks4.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks4.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_O_Speck.ks5.keypos = 42                    ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks5.keyleng = 20                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks5.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks5.reserve = &H0                  ' �\��ς�
    
    
    P_SSHIJI_O_Speck.ks6.keypos = 268                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks6.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks6.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks6.reserve = &H0                  ' �\��ς�
    
    
    
'2007.11.28    P_SSHIJI_O_Speck.ks7.keypos = 1                     ' �L�[�|�W�V����
'2007.11.28    P_SSHIJI_O_Speck.ks7.keyleng = 5                    ' �L�[��
                                                        
    P_SSHIJI_O_Speck.ks7.keypos = 460                   ' �L�[�|�W�V����    2007.11.28
    P_SSHIJI_O_Speck.ks7.keyleng = 8                    ' �L�[��            2007.11.28
                                                        
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks7.keyflag = BtKfExt + BtKfChg
    P_SSHIJI_O_Speck.ks7.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks7.reserve = &H0                  ' �\��ς�
    
    '--------------------------------------------------- �L�[�P ��
    
    '--------------------------------------------------- �L�[�Q  ��
    P_SSHIJI_O_Speck.ks8.keypos = 452                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks8.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SSHIJI_O_Speck.ks8.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks8.reserve = &H0                  ' �\��ς�
    
    '--------------------------------------------------- �L�[�Q ��
    
    
    
    '--------------------------------------------------- �L�[�R ��
    P_SSHIJI_O_Speck.ks9.keypos = 6                     ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks9.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks9.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks9.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks9.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_O_Speck.ks10.keypos = 142                  ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks10.keyleng = 1                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks10.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks10.reserve = &H0                 ' �\��ς�
    
    
    P_SSHIJI_O_Speck.ks11.keypos = 73                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks11.keyleng = 5                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SSHIJI_O_Speck.ks11.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks11.reserve = &H0                 ' �\��ς�
    '--------------------------------------------------- �L�[�R ��
    
    
    
    '--------------------------------------------------- �L�[�S ��  2011.11.11
    P_SSHIJI_O_Speck.ks12.keypos = 38                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks12.keyleng = 2                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks12.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks12.reserve = &H0                 ' �\��ς�
    
    P_SSHIJI_O_Speck.ks13.keypos = 40                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks13.keyleng = 1                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks13.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks13.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks13.reserve = &H0                 ' �\��ς�
    
    
    P_SSHIJI_O_Speck.ks14.keypos = 41                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks14.keyleng = 1                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks14.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks14.reserve = &H0                 ' �\��ς�
    
    
    P_SSHIJI_O_Speck.ks15.keypos = 42                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks15.keyleng = 20                  ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks15.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks15.reserve = &H0                 ' �\��ς�
    
    P_SSHIJI_O_Speck.ks16.keypos = 14                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks16.keyleng = 14                  ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SSHIJI_O_Speck.ks16.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks16.reserve = &H0                 ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�S �� 2011.11.11
    
    '--------------------------------------------------- �L�[�T  �� 2012.03.08
    P_SSHIJI_O_Speck.ks17.keypos = 38                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks17.keyleng = 2                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks17.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks17.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks17.reserve = &H0                 ' �\��ς�
    
    P_SSHIJI_O_Speck.ks18.keypos = 40                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks18.keyleng = 1                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks18.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks18.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_O_Speck.ks19.keypos = 41                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks19.keyleng = 1                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks19.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks19.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks19.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_O_Speck.ks20.keypos = 42                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks20.keyleng = 1                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks20.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks20.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks20.reserve = &H0                  ' �\��ς�
    
    P_SSHIJI_O_Speck.ks21.keypos = 267                   ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks21.keyleng = 1                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks21.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SSHIJI_O_Speck.ks21.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks21.reserve = &H0                  ' �\��ς�
    
    '--------------------------------------------------- �L�[�T �� 2012.03.08
    
    
    '--------------------------------------------------- �L�[�U  �� 2012.03.27
    
    P_SSHIJI_O_Speck.ks22.keypos = 452                  ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks22.keyleng = 8                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks22.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks22.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks22.reserve = &H0                 ' �\��ς�
    
    P_SSHIJI_O_Speck.ks23.keypos = 493                  ' �L�[�|�W�V����
    P_SSHIJI_O_Speck.ks23.keyleng = 2                   ' �L�[��
                                                        ' �L�[�t���O
    P_SSHIJI_O_Speck.ks23.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SSHIJI_O_Speck.ks23.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    P_SSHIJI_O_Speck.ks23.reserve = &H0                 ' �\��ς�
    
    '--------------------------------------------------- �L�[�U �� 2012.03.27
    
    
    
    
    
    
    sts = BTRV(BtOpCreate, P_SSHIJI_O_POS, P_SSHIJI_O_Speck, Len(P_SSHIJI_O_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���i���w�}(�e)�ް�")
        Exit Function
    End If
    
    P_SSHIJI_O_Create = False

End Function

Public Function P_SSHIJI_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���i���w�}(�e)�ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SSHIJI_O_Open = True
                                            '���i���w�}(�e)�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_O]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SSHIJI_O_Create()   '���i���w�}(�e)�ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���i���w�}(�e)�ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���i���w�}(�e)�ް�")
                Exit Function
        End Select
    Loop
    
    P_SSHIJI_O_Open = False

End Function

