Attribute VB_Name = "OLD_P_SSHIJI_K"
Option Explicit

'********************************************************************
'*
'*              ���i���w�}�f�[�^�i�q�j  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c


Public Const OLD_P_SSHIJI_K_ID$ = "OLD_P_SSHIJI_K"

'�y�[�W�T�C�Y
Private Const OLD_P_SSHIJI_K_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public OLD_P_SSHIJI_K_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************

'���R�[�h��`
Public Type OLD_P_SSHIJI_K_REC_Tag
    
    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
    KO_SYUBETSU(0 To 1)     As Byte         '�q�@���
    KO_JGYOBU(0 To 0)       As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�q�@�i��
    KO_QTY(0 To 5)          As Byte         '�q�@����(999V99)
    KO_SHIJI_QTY(0 To 10)   As Byte         '�w����(9(8)V99)
    KO_BIKOU(0 To 39)       As Byte         '�q�@���l
'    KO_ID_NO(0 To 7)        As Byte         '�q �h�c�Q�m�n
    KO_ID_NO(0 To 11)       As Byte         '�q �h�c�Q�m�n (8����12��)  2006/05/24
    CALCEL_F(0 To 0)        As Byte         '��ݾ�F
    CANCEL_DATETIME(0 To 13) As Byte        '��ݾٓ���
'    FILLER(0 To 64)         As Byte         'Filler
    FILLER(0 To 60)         As Byte         'Filler                    2006/05/24
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public OLD_P_SSHIJI_K_REC       As OLD_P_SSHIJI_K_REC_Tag

'�L�[��`

Type KEY0_OLD_P_SSHIJI_K                        '�j�d�x�O
    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
    
Type KEY1_OLD_P_SSHIJI_K                        '�j�d�x�P
    KO_JGYOBU(0 To 0)       As Byte         '�q�@���ƕ�
'    KO_ID_NO(0 To 7)        As Byte         '�q �h�c�Q�m�n
    KO_ID_NO(0 To 11)       As Byte         '�q �h�c�Q�m�n (8����12��)  2006/05/24
End Type
    
    
'�L�[�E�f�[�^
Public K0_OLD_P_SSHIJI_K        As KEY0_OLD_P_SSHIJI_K
Public K1_OLD_P_SSHIJI_K        As KEY1_OLD_P_SSHIJI_K


Public Function OLD_P_SSHIJI_K_Open(Mode As Integer) As Integer
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

    OLD_P_SSHIJI_K_Open = True
                                            '��z�w�}�f�[�^�i�q�j�t���p�X�捞��
    sts = GetIni("FILE", OLD_P_SSHIJI_K_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [OLD_P_SSHIJI_K]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, OLD_P_SSHIJI_K_POS, OLD_P_SSHIJI_K_REC, Len(OLD_P_SSHIJI_K_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "��z�w�}�f�[�^�i�q�j�}�X�^")
                Exit Function
        End Select
    Loop
    
    OLD_P_SSHIJI_K_Open = False

End Function
