Attribute VB_Name = "P_UKEHARAI"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �󕥐�}�X�^  �t�@�C����`                          *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�t�@�C���h�c
Public Const P_UKEHARAI_ID$ = "P_UKEHARAI"

'�y�[�W�T�C�Y
Private Const P_UKEHARAI_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public P_UKEHARAI_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type P_UKEHARAIREC_Tag
    
    
    
    UKEHARAI_CODE(0 To 4)   As Byte         '�󕥐溰��
    SYUSHI_CODE(0 To 2)     As Byte         '���x����
    UKEHARAI_NAME(0 To 49)  As Byte         '�󕥐於��
    UKEHARAI_RNAME(0 To 29) As Byte         '�󕥐旪��
    BUSHO_NAME(0 To 39)     As Byte         '�������^�c�Ə���
    TEL_NO(0 To 14)         As Byte         '�d�b�ԍ�
    FAX_NO(0 To 14)         As Byte         'FAX�ԍ�
    YUBIN_NO(0 To 7)        As Byte         '�X�֔ԍ�
    ADDR1(0 To 39)          As Byte         '�Z��1
    ADDR2(0 To 39)          As Byte         '�Z��2
    TORI_KBN(0 To 0)        As Byte         '�����敪
    FILLER(0 To 117)        As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_UKEHARAIREC        As P_UKEHARAIREC_Tag

'�L�[��`

Type KEY0_P_UKEHARAI                        '�j�d�x�O
    UKEHARAI_CODE(0 To 4)   As Byte         '�󕥐溰��
End Type
    
Type KEY1_P_UKEHARAI                        '�j�d�x�P
    TORI_KBN(0 To 0)        As Byte         '�����敪
    UKEHARAI_CODE(0 To 4)   As Byte         '�󕥐溰��
End Type
    
'�L�[�E�f�[�^
Public K0_P_UKEHARAI        As KEY0_P_UKEHARAI
Public K1_P_UKEHARAI        As KEY1_P_UKEHARAI

Type P_UKEHARAI_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_UKEHARAI_Speck    As P_UKEHARAI_FSpeck
Private Function P_UKEHARAI_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �󕥐�}�X�^  �b�q�d�`�s�d                          *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_UKEHARAI_Create = True
                                            '�󕥐�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_UKEHARAI_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_UKEHARAI]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_UKEHARAI_Speck.fs.recoleng = Len(P_UKEHARAIREC)   ' ���R�[�h��
    P_UKEHARAI_Speck.fs.PageSize = P_UKEHARAI_PG_SIZ     ' �y�[�W�T�C�Y
    P_UKEHARAI_Speck.fs.idexnumb = 2                    ' �C���f�b�N�X��
    P_UKEHARAI_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    P_UKEHARAI_Speck.fs.reserve = &H0                   ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_UKEHARAI_Speck.ks0.keypos = 1                     ' �L�[�|�W�V����
    P_UKEHARAI_Speck.ks0.keyleng = 5                    ' �L�[��
    P_UKEHARAI_Speck.ks0.keyflag = BtKfExt              ' �L�[�t���O
    P_UKEHARAI_Speck.ks0.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_UKEHARAI_Speck.ks0.reserve = &H0                  ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    '--------------------------------------------------- �L�[�O ��
    P_UKEHARAI_Speck.ks1.keypos = 247                   ' �L�[�|�W�V����
    P_UKEHARAI_Speck.ks1.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    P_UKEHARAI_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_UKEHARAI_Speck.ks1.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_UKEHARAI_Speck.ks1.reserve = &H0                  ' �\��ς�
    
    
    P_UKEHARAI_Speck.ks2.keypos = 1                     ' �L�[�|�W�V����
    P_UKEHARAI_Speck.ks2.keyleng = 5                    ' �L�[��
    P_UKEHARAI_Speck.ks2.keyflag = BtKfExt + BtKfChg    ' �L�[�t���O
    P_UKEHARAI_Speck.ks2.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_UKEHARAI_Speck.ks2.reserve = &H0                  ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    
    sts = BTRV(BtOpCreate, P_UKEHARAI_POS, P_UKEHARAI_Speck, Len(P_UKEHARAI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�󕥐�}�X�^")
        Exit Function
    End If
    
    P_UKEHARAI_Create = False

End Function

Public Function P_UKEHARAI_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �󕥐�}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_UKEHARAI_Open = True
                                            '�󕥐�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_UKEHARAI_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_UKEHARAI]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_UKEHARAI_Create()   '�󕥐�}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�󕥐�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�󕥐�}�X�^")
                Exit Function
        End Select
    Loop
    
    P_UKEHARAI_Open = False

End Function
