Attribute VB_Name = "O_P_SSHIJI_K"
Option Explicit

'********************************************************************
'*
'*              ���i���w�}�f�[�^�i�q�j  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c

Public Const O_P_SSHIJI_K_ID$ = "O_P_SSHIJI_K"

'�y�[�W�T�C�Y
Private Const O_P_SSHIJI_K_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public O_P_SSHIJI_K_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************

'���R�[�h��`
Public Type O_P_SSHIJI_K_REC_Tag
    
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
    KO_ID_NO(0 To 7)        As Byte         '�q �h�c�Q�m�n
    CALCEL_F(0 To 0)        As Byte         '��ݾ�F
    CANCEL_DATETIME(0 To 13) As Byte        '��ݾٓ���
    FILLER(0 To 64)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public O_P_SSHIJI_K_REC       As O_P_SSHIJI_K_REC_Tag

'�L�[��`

Type KEY0_O_P_SSHIJI_K                        '�j�d�x�O
    SHIJI_NO(0 To 4)        As Byte         '�w�}�[��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
    
Type KEY1_O_P_SSHIJI_K                        '�j�d�x�O
    KO_JGYOBU(0 To 0)       As Byte         '�q�@���ƕ�
    KO_ID_NO(0 To 7)        As Byte         '�q �h�c�Q�m�n
End Type
    
    
'�L�[�E�f�[�^
Public K0_O_P_SSHIJI_K        As KEY0_O_P_SSHIJI_K

Type O_P_SSHIJI_K_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private O_P_SSHIJI_K_Speck    As O_P_SSHIJI_K_FSpeck
Private Function O_P_SSHIJI_K_Create() As Integer
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

    O_P_SSHIJI_K_Create = True
                                            '��z�w�}�f�[�^�i�q�j�t���p�X�捞��
    sts = GetIni("FILE", O_P_SSHIJI_K_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_P_SSHIJI_K]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    O_P_SSHIJI_K_Speck.fs.recoleng = Len(O_P_SSHIJI_K_REC)  ' ���R�[�h��
    O_P_SSHIJI_K_Speck.fs.PageSize = O_P_SSHIJI_K_PG_SIZ    ' �y�[�W�T�C�Y
    O_P_SSHIJI_K_Speck.fs.idexnumb = 2                    ' �C���f�b�N�X��
    O_P_SSHIJI_K_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    O_P_SSHIJI_K_Speck.fs.reserve = &H0                   ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    O_P_SSHIJI_K_Speck.ks0.keypos = 1                     ' �L�[�|�W�V����
    O_P_SSHIJI_K_Speck.ks0.keyleng = 5                    ' �L�[��
    O_P_SSHIJI_K_Speck.ks0.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    O_P_SSHIJI_K_Speck.ks0.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_P_SSHIJI_K_Speck.ks0.reserve = &H0                  ' �\��ς�
    
    O_P_SSHIJI_K_Speck.ks1.keypos = 6                     ' �L�[�|�W�V����
    O_P_SSHIJI_K_Speck.ks1.keyleng = 1                    ' �L�[��
    O_P_SSHIJI_K_Speck.ks1.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    O_P_SSHIJI_K_Speck.ks1.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_P_SSHIJI_K_Speck.ks1.reserve = &H0                  ' �\��ς�
    
    
    O_P_SSHIJI_K_Speck.ks2.keypos = 7                     ' �L�[�|�W�V����
    O_P_SSHIJI_K_Speck.ks2.keyleng = 3                    ' �L�[��
    O_P_SSHIJI_K_Speck.ks2.keyflag = BtKfExt              ' �L�[�t���O
    O_P_SSHIJI_K_Speck.ks2.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_P_SSHIJI_K_Speck.ks2.reserve = &H0                  ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    O_P_SSHIJI_K_Speck.ks3.keypos = 1                     ' �L�[�|�W�V����
    O_P_SSHIJI_K_Speck.ks3.keyleng = 5                    ' �L�[��
                                                        ' �L�[�t���O
    O_P_SSHIJI_K_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_P_SSHIJI_K_Speck.ks3.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_P_SSHIJI_K_Speck.ks3.reserve = &H0                  ' �\��ς�
    
    O_P_SSHIJI_K_Speck.ks4.keypos = 6                     ' �L�[�|�W�V����
    O_P_SSHIJI_K_Speck.ks4.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    O_P_SSHIJI_K_Speck.ks4.keyflag = BtKfExt + BtKfDup
    O_P_SSHIJI_K_Speck.ks4.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    O_P_SSHIJI_K_Speck.ks4.reserve = &H0                  ' �\��ς�
    '--------------------------------------------------- �L�[�P ��
    
    
    sts = BTRV(BtOpCreate, O_P_SSHIJI_K_POS, O_P_SSHIJI_K_Speck, Len(O_P_SSHIJI_K_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "��z�w�}�f�[�^�i�q�j")
        Exit Function
    End If
    
    O_P_SSHIJI_K_Create = False

End Function

Public Function O_P_SSHIJI_K_Open(Mode As Integer) As Integer
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

    O_P_SSHIJI_K_Open = True
                                            '��z�w�}�f�[�^�i�q�j�t���p�X�捞��
    sts = GetIni("FILE", O_P_SSHIJI_K_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_P_SSHIJI_K]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, O_P_SSHIJI_K_POS, O_P_SSHIJI_K_REC, Len(O_P_SSHIJI_K_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_P_SSHIJI_K_Create()   '��z�w�}�f�[�^�i�q�j�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_P_SSHIJI_K_POS, O_P_SSHIJI_K_REC, Len(O_P_SSHIJI_K_REC), ByVal FullPath, Len(FullPath), Mode)
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
    
    O_P_SSHIJI_K_Open = False

End Function

