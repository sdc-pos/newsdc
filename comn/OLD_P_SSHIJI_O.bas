Attribute VB_Name = "OLD_P_SSHIJI_O"
Option Explicit

'********************************************************************
'*
'*              ���i���w�}�f�[�^�i�e�j  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const OLD_P_SSHIJI_O_ID$ = "OLD_P_SSHIJI_O"

'�y�[�W�T�C�Y
Private Const OLD_P_SSHIJI_O_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public OLD_P_SSHIJI_O_POS As POSBLK
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
Public Type OLD_P_SSHIJI_O_REC_Tag
    
    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��
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
    TORI_KBN(0 To 0)        As Byte
    
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
    
    ORDER_DT(0 To 7)        As Byte         '�󒍓� 2007.02.20
    
    FILLER(0 To 38)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public OLD_P_SSHIJI_O_REC       As OLD_P_SSHIJI_O_REC_Tag

'�L�[��`

Type KEY0_OLD_P_SSHIJI_O                        '�j�d�x�O
    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��
End Type

Type KEY1_OLD_P_SSHIJI_O                        '�j�d�x�P
    KAN_F(0 To 0)           As Byte         '����F
    SHIMUKE_CODE(0 To 1)    As Byte         '�d�����溰��
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    KAN_DT(0 To 7)          As Byte         '������
    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��
End Type
    
Type KEY2_OLD_P_SSHIJI_O                        '�j�d�x�Q
    ORDER_DT(0 To 7)        As Byte         '�󒍓� 2007.02.20
End Type
    
    
    
    
    
    
'�L�[�E�f�[�^
Public K0_OLD_P_SSHIJI_O        As KEY0_OLD_P_SSHIJI_O
Public K1_OLD_P_SSHIJI_O        As KEY1_OLD_P_SSHIJI_O
Public K2_OLD_P_SSHIJI_O        As KEY2_OLD_P_SSHIJI_O


Public Function OLD_P_SSHIJI_O_Open(Mode As Integer) As Integer
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

    OLD_P_SSHIJI_O_Open = True
                                            '���i���w�}(�e)�ް��t���p�X�捞��
    sts = GetIni("FILE", OLD_P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [OLD_P_SSHIJI_O]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, OLD_P_SSHIJI_O_POS, OLD_P_SSHIJI_O_REC, Len(OLD_P_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "���i���w�}(�e)�ް�")
                Exit Function
        End Select
    Loop
    
    OLD_P_SSHIJI_O_Open = False

End Function

