Attribute VB_Name = "MTS"
Option Explicit
'********************************************************************
'*
'*              ������Ǘ��}�X�^  �t�@�C����`
'*
'*          CREATE 2004.02.19
'********************************************************************
'�t�@�C���h�c
Public Const MTS_ID$ = "MTS"

'�y�[�W�T�C�Y
Public Const MTS_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public MTS_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type MTSREC_Tag
    NAIGAI(0 To 0)          As Byte         '�����O
    DATA_KBN(0 To 0)        As Byte         '�f�[�^�敪�i���g�p�j
    MUKE_CODE(0 To 7)       As Byte         '���Ӑ�R�[�h
    SS_CODE(0 To 7)         As Byte         '�q�Ɂ^�r�r�R�[�h
    MUKE_NAME(0 To 39)      As Byte         '���Ӑ於��
    SS_NAME(0 To 39)        As Byte         '�r�r����
    MUKE_DNAME(0 To 9)      As Byte         '�\������
    DISPLAY_RANKING(0 To 2) As Byte         '�\������
    
    SYUKA_KBN(0 To 1)       As Byte         '�o�׋敪�R�[�h 2008.03.12
    FILLER(0 To 14)         As Byte         'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public MTSREC               As MTSREC_Tag

'�L�[��`

Type KEY0_MTS                 '�j�d�x�O
    MUKE_CODE(0 To 7)       As Byte         '���Ӑ�R�[�h
    SS_CODE(0 To 7)         As Byte         '�q�Ɂ^�r�r�R�[�h
End Type

Type KEY1_MTS                 '�j�d�x�P
    DISPLAY_RANKING(0 To 2) As Byte         '�\������
    MUKE_CODE(0 To 7)       As Byte         '���Ӑ�R�[�h
    SS_CODE(0 To 7)         As Byte         '�q�Ɂ^�r�r�R�[�h
End Type

Type KEY2_MTS                 '�j�d�x�Q
    MUKE_CODE(0 To 7)       As Byte         '���Ӑ�R�[�h
End Type

Type KEY3_MTS                 '�j�d�x�R
    SS_CODE(0 To 7)         As Byte         '�q�Ɂ^�r�r�R�[�h
End Type

'�L�[�E�f�[�^
Public K0_MTS               As KEY0_MTS
Public K1_MTS               As KEY1_MTS
Public K2_MTS               As KEY2_MTS
Public K3_MTS               As KEY3_MTS

Type MTS_FSpeck
    fs  As BtFileSpeck                      '̧�� ��߯��\����
    ks0 As BtKeySpeck                       '�� ��߯��\����
    ks1 As BtKeySpeck
    ks2 As BtKeySpeck
    ks3 As BtKeySpeck
    ks4 As BtKeySpeck
    ks5 As BtKeySpeck
    ks6 As BtKeySpeck
End Type

Private MTS_Speck As MTS_FSpeck
Private Function MTS_Create() As Integer
'********************************************************************
'*
'*              ������Ǘ��}�X�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2004.02.19
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    MTS_Create = True
                                            '������Ǘ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", MTS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [MTS]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    MTS_Speck.fs.recoleng = Len(MTSREC)         ' ���R�[�h��
    MTS_Speck.fs.PageSize = MTS_PG_SIZ          ' �y�[�W�T�C�Y
    MTS_Speck.fs.idexnumb = 4                   ' �C���f�b�N�X��
    MTS_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    MTS_Speck.fs.reserve = &H0                  ' �\��ς�
'------------------------------------------------
                                                ' �L�[�O
    MTS_Speck.ks0.keypos = 3                    ' �L�[�|�W�V����
    MTS_Speck.ks0.keyleng = 8                   ' �L�[��
    MTS_Speck.ks0.keyflag = BtKfExt + BtKfSeg   ' �L�[�t���O
    MTS_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    MTS_Speck.ks0.reserve = &H0                 ' �\��ς�
    
    MTS_Speck.ks1.keypos = 11                   ' �L�[�|�W�V����
    MTS_Speck.ks1.keyleng = 8                   ' �L�[��
    MTS_Speck.ks1.keyflag = BtKfExt             ' �L�[�t���O
    MTS_Speck.ks1.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    MTS_Speck.ks1.reserve = &H0                 ' �\��ς�
'------------------------------------------------
                                                ' �L�[�P
    MTS_Speck.ks2.keypos = 109                  ' �L�[�|�W�V����
    MTS_Speck.ks2.keyleng = 3                   ' �L�[��
                                                ' �L�[�t���O
    MTS_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfChg
    MTS_Speck.ks2.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    MTS_Speck.ks2.reserve = &H0                 ' �\��ς�
    
    MTS_Speck.ks3.keypos = 3                    ' �L�[�|�W�V����
    MTS_Speck.ks3.keyleng = 8                   ' �L�[��
                                                ' �L�[�t���O
    MTS_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg
    MTS_Speck.ks3.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    MTS_Speck.ks3.reserve = &H0                 ' �\��ς�
    
    MTS_Speck.ks4.keypos = 11                   ' �L�[�|�W�V����
    MTS_Speck.ks4.keyleng = 8                   ' �L�[��
    MTS_Speck.ks4.keyflag = BtKfExt + BtKfChg   ' �L�[�t���O
    MTS_Speck.ks4.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    MTS_Speck.ks4.reserve = &H0                 ' �\��ς�
'------------------------------------------------
                                                ' �L�[�Q
    MTS_Speck.ks5.keypos = 3                    ' �L�[�|�W�V����
    MTS_Speck.ks5.keyleng = 8                   ' �L�[��
    MTS_Speck.ks5.keyflag = BtKfExt + BtKfDup   ' �L�[�t���O
    MTS_Speck.ks5.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    MTS_Speck.ks5.reserve = &H0                 ' �\��ς�
'------------------------------------------------
                                                ' �L�[�R
    MTS_Speck.ks6.keypos = 11                   ' �L�[�|�W�V����
    MTS_Speck.ks6.keyleng = 8                   ' �L�[��
    MTS_Speck.ks6.keyflag = BtKfExt + BtKfDup   ' �L�[�t���O
    MTS_Speck.ks6.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    MTS_Speck.ks6.reserve = &H0                 ' �\��ς�


    sts = BTRV(BtOpCreate, MTS_POS, MTS_Speck, Len(MTS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "������Ǘ��}�X�^")
        Exit Function
    End If

    MTS_Create = False

End Function

Public Function MTS_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ������Ǘ��}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    MTS_Open = True
                                            '������Ǘ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", MTS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [MTS]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, MTS_POS, MTSREC, Len(MTSREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = MTS_Create()        '������Ǘ��}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, MTS_POS, MTSREC, Len(MTSREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "������Ǘ��}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "������Ǘ��}�X�^")
                Exit Function
        End Select
    Loop
    
    MTS_Open = False

End Function
