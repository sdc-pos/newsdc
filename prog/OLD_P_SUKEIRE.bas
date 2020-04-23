Attribute VB_Name = "OLD_P_SUKEIRE"
Option Explicit

'********************************************************************
'*
'*              ���i���w�}��������f�[�^  �t�@�C����`
'*
'*          CREATE 2005.12.14
'********************************************************************
'�t�@�C���h�c
Public Const OLD_P_SUKEIRE_ID$ = "OLD_P_SUKEIRE"

'�y�[�W�T�C�Y
Private Const OLD_P_SUKEIRE_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public OLD_P_SUKEIRE_POS As POSBLK
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
Public Type OLD_P_SUKEIRE_REC_Tag
    
    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��
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
    FILLER(0 To 94)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public OLD_P_SUKEIRE_REC        As OLD_P_SUKEIRE_REC_Tag

'�L�[��`

Type KEY0_OLD_P_SUKEIRE                         '�j�d�x�O
    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
Type KEY1_OLD_P_SUKEIRE                         '�j�d�x�P
    SHIMUKE_CODE(0 To 1)    As Byte         '�d�����溰��
    UKEIRE_DT(0 To 7)       As Byte         '�����
End Type
    
Type KEY2_OLD_P_SUKEIRE                         '�j�d�x�Q
    TORI_CODE(0 To 4)       As Byte         '�����
    UKEIRE_DT(0 To 7)       As Byte         '�����
End Type
    
'�L�[�E�f�[�^
Public K0_OLD_P_SUKEIRE         As KEY0_OLD_P_SUKEIRE
Public K1_OLD_P_SUKEIRE         As KEY1_OLD_P_SUKEIRE
Public K2_OLD_P_SUKEIRE         As KEY2_OLD_P_SUKEIRE


Public Function OLD_P_SUKEIRE_Open(Mode As Integer) As Integer
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

    OLD_P_SUKEIRE_Open = True
                                            '���i���w�}��������ް��t���p�X�捞��
    sts = GetIni("FILE", OLD_P_SUKEIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [OLD_P_SUKEIRE]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, OLD_P_SUKEIRE_POS, OLD_P_SUKEIRE_REC, Len(OLD_P_SUKEIRE_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "���i���w�}��������ް�")
                Exit Function
        End Select
    Loop
    
    OLD_P_SUKEIRE_Open = False

End Function

