Attribute VB_Name = "YOIN"
Option Explicit
'********************************************************************
'*
'*              �v���}�X�^  �t�@�C����`
'*
'*          CREATE 2001.02.14
'********************************************************************
'�t�@�C���h�c
Public Const YOIN_ID$ = "YOIN"

'�y�[�W�T�C�Y
Public Const YOIN_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public YOIN_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type YOINREC_Tag
    CODE_TYPE(0 To 0)       As Byte     '��o�[�R�[�h�^�C�v
    YOIN_CODE(0 To 0)       As Byte     '�v��
    YOIN_DNAME(0 To 9)      As Byte     '��ƕ\������
    SUM_KBN(0 To 0)         As Byte     '�W�v�敪
    SYSTEM_F(0 To 0)        As Byte     '�V�X�e���\���׸�2004.02
    REGI_F(0 To 0)          As Byte     '�o�^���׸�
    PARAM_F(0 To 0)         As Byte     '�t�����Ұ�(0:�Ȃ� 1:������ 2:�q��)
    Soko_No(0 To 1)         As Byte     '�q�ɇ��i���z�j
    DSP_No(0 To 1)          As Byte     '�\�����@2007.12.10
    FILLER(0 To 3)          As Byte
End Type

'�f�[�^�E�o�b�t�@
Public YOINREC As YOINREC_Tag

'�L�[��`

Type KEY0_YOIN                 '�j�d�x�O
    CODE_TYPE(0 To 0)       As Byte     '��o�[�R�[�h�^�C�v
    YOIN_CODE(0 To 0)       As Byte     '�v��
End Type

Type KEY1_YOIN                 '�j�d�x�O    2007.12.10
    DSP_No(0 To 1)          As Byte     '�\�����@2007.12.10
    CODE_TYPE(0 To 0)       As Byte     '��o�[�R�[�h�^�C�v
End Type

'�L�[�E�f�[�^
Public K0_YOIN As KEY0_YOIN
Public K1_YOIN As KEY1_YOIN             '2007.12.10

Type YOIN_FSpeck
    fs          As BtFileSpeck          ' ̧�� ��߯��\����
    ks0         As BtKeySpeck           ' �� ��߯��\����
    ks1         As BtKeySpeck           ' �� ��߯��\����
    ks2         As BtKeySpeck           ' �� ��߯��\����    2007.12.10
    ks3         As BtKeySpeck           ' �� ��߯��\����    2007.12.10
End Type

Private YOIN_Speck As YOIN_FSpeck

Private Function YOIN_Create() As Integer
'********************************************************************
'*
'*              �v���}�X�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2001.02.14
'*          UPDATE 2007.12.10   KEY1�ǉ�
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    YOIN_Create = True
                                            '�v���}�X�^�t���p�X�捞��
    sts = GetIni("FILE", YOIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    YOIN_Speck.fs.recoleng = Len(YOINREC)           ' ���R�[�h��
    YOIN_Speck.fs.PageSize = YOIN_PG_SIZ            ' �y�[�W�T�C�Y
    YOIN_Speck.fs.idexnumb = 2                      ' �C���f�b�N�X��    2007.12.10
    YOIN_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    YOIN_Speck.fs.reserve = &H0                     ' �\��ς�
'----------------------------------------------------
                                                    ' �L�[�O
    YOIN_Speck.ks0.keypos = 1                       ' �L�[�|�W�V����
    YOIN_Speck.ks0.keyleng = 1                      ' �L�[��
    YOIN_Speck.ks0.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    YOIN_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    YOIN_Speck.ks0.reserve = &H0                    ' �\��ς�
                                                    ' �L�[�O
    YOIN_Speck.ks1.keypos = 2                       ' �L�[�|�W�V����
    YOIN_Speck.ks1.keyleng = 1                      ' �L�[��
    YOIN_Speck.ks1.keyflag = BtKfExt                ' �L�[�t���O
    YOIN_Speck.ks1.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    YOIN_Speck.ks1.reserve = &H0                    ' �\��ς�
    
'----------------------------------------------------
    
'----------------------------------------------------   2007.12.10
                                                    ' �L�[�P
    YOIN_Speck.ks2.keypos = 19                      ' �L�[�|�W�V����
    YOIN_Speck.ks2.keyleng = 2                      ' �L�[��
                                                    ' �L�[�t���O
    YOIN_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    YOIN_Speck.ks2.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    YOIN_Speck.ks2.reserve = &H0                    ' �\��ς�
                                                    ' �L�[�P
    YOIN_Speck.ks3.keypos = 1                       ' �L�[�|�W�V����
    YOIN_Speck.ks3.keyleng = 1                      ' �L�[��
    YOIN_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    YOIN_Speck.ks3.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    YOIN_Speck.ks3.reserve = &H0                    ' �\��ς�
    
'----------------------------------------------------   2007.12.10
    
    
    
    
    sts = BTRV(BtOpCreate, YOIN_POS, YOIN_Speck, Len(YOIN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�v���}�X�^")
        Exit Function
    End If
    YOIN_Create = False
End Function

Function YOIN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �v���}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2001.02.14
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    YOIN_Open = True
                                            '�v���}�X�^�t���p�X�捞��
    sts = GetIni("FILE", YOIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, YOIN_POS, YOINREC, Len(YOINREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = YOIN_Create()        '�v���}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, YOIN_POS, YOINREC, Len(YOINREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�v���}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�v���}�X�^")
                Exit Function
        End Select
    Loop
    
    YOIN_Open = False

End Function



