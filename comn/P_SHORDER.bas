Attribute VB_Name = "P_SHORDER"
Option Explicit

'********************************************************************
'*
'*              ���ޒ����ް�  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const P_SHORDER_ID$ = "P_SHORDER"

'�y�[�W�T�C�Y
Private Const P_SHORDER_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public P_SHORDER_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type P_SHORDER_REC_Tag
    
    ORDER_NO(0 To 4)        As Byte         '������
    ORDER_DT(0 To 7)        As Byte         '������
    Print_datetime(0 To 13) As Byte         '���s����
    TANTO_CODE(0 To 4)      As Byte         '�S���Һ���
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
    DELI_CODE(0 To 4)       As Byte         '�[���溰��
    ORDER_QTY(0 To 10)      As Byte         '������(9(8)V99)
    Y_NOUKI_DT(0 To 7)      As Byte         '�\��[��
    TANKA(0 To 10)          As Byte         '�����P��(9(8)V99)
    LOT(0 To 7)             As Byte         '����ۯ�
    KAN_F(0 To 0)           As Byte         '����F
    KAN_DT(0 To 7)          As Byte         '������
    BUNNOU_CNT(0 To 1)      As Byte         '���[��
    UKEIRE_QTY(0 To 10)     As Byte         '������i���v�j(9(8)V99)
    
    CANCEL_F(0 To 0)        As Byte         '��ݾ�F
    CANCEL_DATETIME(0 To 13) As Byte        '��ݾٓ���
    PRINT_F(0 To 0)         As Byte         '����������׸�
    WS_NO(0 To 9)           As Byte         '���͒[��
    G_SHIIRE_KBN(0 To 1)    As Byte         '�d���敪
    G_SYUSHI(0 To 2)        As Byte         '���x�P��
    TORI_KBN(0 To 0)        As Byte         '�����敪
    
    ANS_NOUKI_DT(0 To 7)    As Byte         '�[���񓚓�   2008.01.10
    USE_YM(0 To 5)          As Byte         '�g�p�N��     2008.01.10
    
    
    UPD_FLG(0 To 0)         As Byte         '�W�J�X�V�ς�   2012.01.17
    
    FILLER(0 To 70)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_SHORDER_REC       As P_SHORDER_REC_Tag

'�L�[��`

Public Type KEY0_P_SHORDER                         '�j�d�x�O
    ORDER_NO(0 To 4)        As Byte         '������
End Type
    
Public Type KEY1_P_SHORDER                         '�j�d�x�P
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    ORDER_DT(0 To 7)        As Byte         '������
    ORDER_NO(0 To 4)        As Byte         '������
End Type
    
Public Type KEY2_P_SHORDER                         '�j�d�x�Q
    WS_NO(0 To 9)           As Byte         '���͒[��
    PRINT_F(0 To 0)         As Byte         '����������׸�
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
    ORDER_NO(0 To 4)        As Byte         '������
End Type
    
Public Type KEY3_P_SHORDER                         '�j�d�x�R
    KAN_F(0 To 0)           As Byte         '����F
    ORDER_DT(0 To 7)        As Byte         '������
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
End Type
    
    
Public Type KEY4_P_SHORDER                         '�j�d�x�S
    KAN_F(0 To 0)           As Byte         '����F
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
    ORDER_DT(0 To 7)        As Byte         '������
End Type
    
Public Type KEY5_P_SHORDER                         '�j�d�x�T    2007.12.05
    KAN_F(0 To 0)           As Byte         '����F
    Y_NOUKI_DT(0 To 7)      As Byte         '�\��[��
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
End Type
    
    
Public Type KEY6_P_SHORDER                         '�j�d�x�U    2008.03.23
    ANS_NOUKI_DT(0 To 7)    As Byte         '�[���񓚓�
End Type
    
    
Public Type KEY7_P_SHORDER                         '�j�d�x�V    2012.03.06
    USE_YM(0 To 5)          As Byte         '�g�p�N��     2008.01.10
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    CANCEL_F(0 To 0)        As Byte         '��ݾ�F
End Type
    
Public Type KEY8_P_SHORDER                         '�j�d�x�W    2019.03.15
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��

    KAN_F(0 To 0)           As Byte         '����F

    CANCEL_F(0 To 0)        As Byte         '��ݾ�F

    Print_datetime(0 To 13) As Byte         '���s����

End Type
    
    
    
    
    
'�L�[�E�f�[�^
Public K0_P_SHORDER         As KEY0_P_SHORDER
Public K1_P_SHORDER         As KEY1_P_SHORDER
Public K2_P_SHORDER         As KEY2_P_SHORDER
Public K3_P_SHORDER         As KEY3_P_SHORDER
Public K4_P_SHORDER         As KEY4_P_SHORDER
Public K5_P_SHORDER         As KEY5_P_SHORDER   '2007.12.05

Public K6_P_SHORDER         As KEY6_P_SHORDER   '2008.03.23

Public K7_P_SHORDER         As KEY7_P_SHORDER   '2012.03.06

Public K8_P_SHORDER         As KEY8_P_SHORDER   '2019.03.15



Type P_SHORDER_FSpeck
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
    ks12                    As BtKeySpeck   ' �� ��߯��\����

    ks13                    As BtKeySpeck   ' �� ��߯��\����
    ks14                    As BtKeySpeck   ' �� ��߯��\����
    ks15                    As BtKeySpeck   ' �� ��߯��\����

    ks16                    As BtKeySpeck   ' �� ��߯��\����    2007.12.05
    ks17                    As BtKeySpeck   ' �� ��߯��\����    2007.12.05
    ks18                    As BtKeySpeck   ' �� ��߯��\����    2007.12.05

    ks19                    As BtKeySpeck   ' �� ��߯��\����    2008.03.23

    ks20                    As BtKeySpeck   ' �� ��߯��\����    2012.03.06
    ks21                    As BtKeySpeck   ' �� ��߯��\����    2012.03.06
    ks22                    As BtKeySpeck   ' �� ��߯��\����    2012.03.06
    ks23                    As BtKeySpeck   ' �� ��߯��\����    2012.03.06
    ks24                    As BtKeySpeck   ' �� ��߯��\����    2012.03.06

    ks25                    As BtKeySpeck   ' �� ��߯��\����    2019.03.15
    ks26                    As BtKeySpeck   ' �� ��߯��\����    2019.03.15
    ks27                    As BtKeySpeck   ' �� ��߯��\����    2019.03.15
    ks28                    As BtKeySpeck   ' �� ��߯��\����    2019.03.15
    ks29                    As BtKeySpeck   ' �� ��߯��\����    2019.03.15
    ks30                    As BtKeySpeck   ' �� ��߯��\����    2019.03.15

End Type

Private P_SHORDER_Speck    As P_SHORDER_FSpeck
Private Function P_SHORDER_Create() As Integer
'********************************************************************
'*
'*              ���ޒ����ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SHORDER_Create = True
                                            '���ޒ����ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHORDER_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SHORDER]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SHORDER_Speck.fs.recoleng = Len(P_SHORDER_REC)    ' ���R�[�h��
    P_SHORDER_Speck.fs.PageSize = P_SHORDER_PG_SIZ      ' �y�[�W�T�C�Y
    P_SHORDER_Speck.fs.idexnumb = 9                     ' �C���f�b�N�X��
    P_SHORDER_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    P_SHORDER_Speck.fs.reserve = &H0                    ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_SHORDER_Speck.ks0.keypos = 1                      ' �L�[�|�W�V����
    P_SHORDER_Speck.ks0.keyleng = 5                     ' �L�[��
    P_SHORDER_Speck.ks0.keyflag = BtKfExt               ' �L�[�t���O
    P_SHORDER_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHORDER_Speck.ks0.reserve = &H0                   ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    P_SHORDER_Speck.ks1.keypos = 33                     ' �L�[�|�W�V����
    P_SHORDER_Speck.ks1.keyleng = 1                     ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHORDER_Speck.ks1.reserve = &H0                   ' �\��ς�
    
    P_SHORDER_Speck.ks2.keypos = 34                     ' �L�[�|�W�V����
    P_SHORDER_Speck.ks2.keyleng = 1                     ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHORDER_Speck.ks2.reserve = &H0                   ' �\��ς�
    
    P_SHORDER_Speck.ks3.keypos = 35                     ' �L�[�|�W�V����
    P_SHORDER_Speck.ks3.keyleng = 20                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHORDER_Speck.ks3.reserve = &H0                   ' �\��ς�
    
    P_SHORDER_Speck.ks4.keypos = 6                      ' �L�[�|�W�V����
    P_SHORDER_Speck.ks4.keyleng = 8                     ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHORDER_Speck.ks4.reserve = &H0                   ' �\��ς�
    
    P_SHORDER_Speck.ks5.keypos = 1                      ' �L�[�|�W�V����
    P_SHORDER_Speck.ks5.keyleng = 5                     ' �L�[��
    P_SHORDER_Speck.ks5.keyflag = BtKfExt + BtKfChg     ' �L�[�t���O
    P_SHORDER_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHORDER_Speck.ks5.reserve = &H0                   ' �\��ς�
    
    '--------------------------------------------------- �L�[�P ��
    
    
    
    '--------------------------------------------------- �L�[�Q ��
    P_SHORDER_Speck.ks6.keypos = 141                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks6.keyleng = 10                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks6.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHORDER_Speck.ks6.reserve = &H0                   ' �\��ς�
    
    P_SHORDER_Speck.ks7.keypos = 140                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks7.keyleng = 1                     ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks7.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHORDER_Speck.ks7.reserve = &H0                   ' �\��ς�
    
    P_SHORDER_Speck.ks8.keypos = 55                     ' �L�[�|�W�V����
    P_SHORDER_Speck.ks8.keyleng = 5                     ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks8.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHORDER_Speck.ks8.reserve = &H0                   ' �\��ς�
    
    P_SHORDER_Speck.ks9.keypos = 1                      ' �L�[�|�W�V����
    P_SHORDER_Speck.ks9.keyleng = 5                     ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks9.keyflag = BtKfExt + BtKfChg
    P_SHORDER_Speck.ks9.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    P_SHORDER_Speck.ks9.reserve = &H0                   ' �\��ς�
    
    '--------------------------------------------------- �L�[�Q ��
    
    '--------------------------------------------------- �L�[�R ��
    
    P_SHORDER_Speck.ks10.keypos = 103                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks10.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks10.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks10.reserve = &H0                  ' �\��ς�
    
    P_SHORDER_Speck.ks11.keypos = 6                     ' �L�[�|�W�V����
    P_SHORDER_Speck.ks11.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks11.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks11.reserve = &H0                  ' �\��ς�
    
    
    P_SHORDER_Speck.ks12.keypos = 55                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks12.keyleng = 5                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks12.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks12.reserve = &H0                  ' �\��ς�

    '--------------------------------------------------- �L�[�R ��
    
    
    '--------------------------------------------------- �L�[�S ��
    
    P_SHORDER_Speck.ks13.keypos = 103                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks13.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks13.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks13.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks13.reserve = &H0                  ' �\��ς�
    
    P_SHORDER_Speck.ks14.keypos = 55                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks14.keyleng = 5                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks14.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks14.reserve = &H0                  ' �\��ς�
    
    
    P_SHORDER_Speck.ks15.keypos = 6                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks15.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks15.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks15.reserve = &H0                  ' �\��ς�

    '--------------------------------------------------- �L�[�S ��
    
    
    
    
    
    '--------------------------------------------------- �L�[�T 2007.12.05 ��
    
    P_SHORDER_Speck.ks16.keypos = 103                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks16.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks16.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks16.reserve = &H0                  ' �\��ς�
    
    P_SHORDER_Speck.ks17.keypos = 76                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks17.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks17.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks17.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks17.reserve = &H0                  ' �\��ς�
    
    P_SHORDER_Speck.ks18.keypos = 55                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks18.keyleng = 5                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks18.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks18.reserve = &H0                  ' �\��ς�

    '--------------------------------------------------- �L�[�T 2007.12.05 ��
    
    
    '--------------------------------------------------- �L�[�U 2008.03.23 ��


    P_SHORDER_Speck.ks19.keypos = 157                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks19.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks19.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks19.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks19.reserve = &H0                  ' �\��ς�



    '--------------------------------------------------- �L�[�U 2008.03.20 ��
    
    
    '--------------------------------------------------- �L�[�V 2012.03.06 ��
    P_SHORDER_Speck.ks20.keypos = 165                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks20.keyleng = 6                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks20.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks20.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks20.reserve = &H0                  ' �\��ς�

    P_SHORDER_Speck.ks21.keypos = 33                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks21.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks21.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks21.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks21.reserve = &H0                  ' �\��ς�

    P_SHORDER_Speck.ks22.keypos = 34                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks22.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks22.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks22.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks22.reserve = &H0                  ' �\��ς�

    P_SHORDER_Speck.ks23.keypos = 35                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks23.keyleng = 20                   ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks23.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks23.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks23.reserve = &H0                  ' �\��ς�

    P_SHORDER_Speck.ks24.keypos = 125                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks24.keyleng = 1                   ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks24.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks24.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks24.reserve = &H0                  ' �\��ς�
    '--------------------------------------------------- �L�[�V 2012.03.06 ��
    
    
    
    '--------------------------------------------------- �L�[�W 2019.03.15 ��


    P_SHORDER_Speck.ks25.keypos = 33                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks25.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks25.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks25.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks25.reserve = &H0                  ' �\��ς�

    P_SHORDER_Speck.ks26.keypos = 34                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks26.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks26.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks26.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks26.reserve = &H0                  ' �\��ς�


    P_SHORDER_Speck.ks27.keypos = 35                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks27.keyleng = 20                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks27.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks27.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks27.reserve = &H0                  ' �\��ς�


    P_SHORDER_Speck.ks28.keypos = 103                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks28.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks28.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks28.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks28.reserve = &H0                  ' �\��ς�


    P_SHORDER_Speck.ks29.keypos = 125                   ' �L�[�|�W�V����
    P_SHORDER_Speck.ks29.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks29.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks29.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks29.reserve = &H0                  ' �\��ς�


    P_SHORDER_Speck.ks30.keypos = 14                    ' �L�[�|�W�V����
    P_SHORDER_Speck.ks30.keyleng = 14                   ' �L�[��
                                                        ' �L�[�t���O
    P_SHORDER_Speck.ks30.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks30.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHORDER_Speck.ks30.reserve = &H0                  ' �\��ς�



    '--------------------------------------------------- �L�[�W 2019.03.15 ��
    
    
    
    
    sts = BTRV(BtOpCreate, P_SHORDER_POS, P_SHORDER_Speck, Len(P_SHORDER_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ޒ����ް�")
        Exit Function
    End If
    
    P_SHORDER_Create = False

End Function

Public Function P_SHORDER_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޒ����ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SHORDER_Open = True
                                            '���ޒ����ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHORDER_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SHORDER]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHORDER_Create()   '���ޒ����ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ޒ����ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޒ����ް�")
                Exit Function
        End Select
    Loop
    
    P_SHORDER_Open = False

End Function

