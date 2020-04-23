Attribute VB_Name = "PCB_U_ITEM"
Option Explicit
'********************************************************************
'*
'*              PCB.U�ݕ�  �t�@�C����`
'*
'*          CREATE 2014.06.18
'********************************************************************
'�t�@�C���h�c
Public Const PCB_U_ITEM_ID$ = "PCB_U_ITEM"

'�y�[�W�T�C�Y
Public Const PCB_U_ITEM_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public PCB_U_ITEM_POS               As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type PCB_U_ITEM_REC_Tag
    JGYOBU(0 To 0)                  As Byte     '���ƕ��敪
    NAIGAI(0 To 0)                  As Byte     '�����O
    HIN_GAI(0 To 19)                As Byte     '�i�ԁi�O���j
    
    NAI_BUHIN(0 To 0)               As Byte     '�����敪
    KENSA_JIGU(0 To 0)              As Byte     '��������
    
    HITUYO_SU(0 To 4)               As Byte     '�K�v���@��
    HITUYO_TUKI(0 To 1)             As Byte     '�K�v���@��
    
    SETUHEN_LAST_DATE(0 To 7)       As Byte     '�݌v�ύX�ŏI��
    SENDO_LAST_DATE(0 To 7)         As Byte     '�N�x�Ǘ��ŏI��
    
    MODULE_KBN(0 To 0)              As Byte     '���W���[���Ώۋ敪
    SETUHEN_KBN(0 To 0)             As Byte     '�݌v�ύX�Ώۋ敪
    
    
    KANRI_NO(0 To 1)                As Byte     '�Ǘ���
    EX_DATE(0 To 7)                 As Byte     '���t
    SETUHEN_NO(0 To 4)              As Byte     '�ݕϊǗ���
    
    BEF_HIN_GAI(0 To 19)            As Byte     '�ύX�O�@���޽�i��
    BEF_HIN_NAI(0 To 19)            As Byte     '�ύX�O�@�H��i��
    AFT_HIN_GAI(0 To 19)            As Byte     '�ύX�O�@���޽�i��
    AFT_HIN_NAI(0 To 19)            As Byte     '�ύX�O�@�H��i��
    
    HEN_BUHIN(0 To 39)              As Byte     '�ύX���i
    HEN_NAIYO(0 To 49)              As Byte     '�ύX���e
    HEN_BASHO(0 To 19)              As Byte     '�����ꏊ
    
    SETUHEN_HOKAN(0 To 19)          As Byte     '�ݕό����ۊ�
        
    BIKOU(0 To 199)                 As Byte     '���l
    
    
    FILLER(0 To 245)                As Byte         'FILLER
    INS_TANTO(0 To 9)               As Byte         '�ǉ��@�S����
    Ins_DateTime(0 To 13)           As Byte         '�ǉ��@����
    UPD_TANTO(0 To 9)               As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)           As Byte         '�X�V�@����



End Type
'�f�[�^�E�o�b�t�@
Public PCB_U_ITEM_REC               As PCB_U_ITEM_REC_Tag

'�L�[��`
Type KEY0_PCB_U_ITEM                '�j�d�x�O
    JGYOBU(0 To 0)                  As Byte     '���ƕ��敪
    NAIGAI(0 To 0)                  As Byte     '�����O
    HIN_GAI(0 To 19)                As Byte     '�i�ԁi�O���j
End Type








'�L�[�E�f�[�^
Public K0_PCB_U_ITEM                As KEY0_PCB_U_ITEM


Private Type PCB_U_ITEM_FSpeck
    fs      As BtFileSpeck              ' ̧�� ��߯��\����
    ks0     As BtKeySpeck               ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck

End Type

Private PCB_U_ITEM_Speck            As PCB_U_ITEM_FSpeck

Private Function PCB_U_ITEM_Create() As Integer
'********************************************************************
'*
'*              PCB.U�ݕ�  �t�@�C����`
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PCB_U_ITEM_Create = True
                                            'PCB.U�ݕρ@�t���p�X�捞��
    sts = GetIni("FILE", PCB_U_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PCB_U_ITEM]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    PCB_U_ITEM_Speck.fs.recoleng = Len(PCB_U_ITEM_REC)      ' ���R�[�h��
    PCB_U_ITEM_Speck.fs.PageSize = PCB_U_ITEM_PG_SIZ        ' �y�[�W�T�C�Y
    PCB_U_ITEM_Speck.fs.idexnumb = 1                        ' �C���f�b�N�X��
    PCB_U_ITEM_Speck.fs.fileflag = 0                        ' �t�@�C���t���O
    PCB_U_ITEM_Speck.fs.reserve = &H0                       ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    PCB_U_ITEM_Speck.ks0.keypos = 1                         ' �L�[�|�W�V����
    PCB_U_ITEM_Speck.ks0.keyleng = 1                        ' �L�[��
    PCB_U_ITEM_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    PCB_U_ITEM_Speck.ks0.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    PCB_U_ITEM_Speck.ks0.reserve = &H0                      ' �\��ς�

    PCB_U_ITEM_Speck.ks1.keypos = 2                         ' �L�[�|�W�V����
    PCB_U_ITEM_Speck.ks1.keyleng = 1                        ' �L�[��
                                                            
    PCB_U_ITEM_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    PCB_U_ITEM_Speck.ks1.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    PCB_U_ITEM_Speck.ks1.reserve = &H0                      ' �\��ς�

    PCB_U_ITEM_Speck.ks2.keypos = 3                         ' �L�[�|�W�V����
    PCB_U_ITEM_Speck.ks2.keyleng = 20                       ' �L�[��
                                                            
    PCB_U_ITEM_Speck.ks2.keyflag = BtKfExt                  ' �L�[�t���O
    PCB_U_ITEM_Speck.ks2.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    PCB_U_ITEM_Speck.ks2.reserve = &H0                      ' �\��ς�




'-----------------------------------------------

    sts = BTRV(BtOpCreate, PCB_U_ITEM_POS, PCB_U_ITEM_Speck, Len(PCB_U_ITEM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "PCB.U�ݕ�")
        Exit Function
    End If

    PCB_U_ITEM_Create = False

End Function

Public Function PCB_U_ITEM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              PCB.U�ݕ�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    PCB_U_ITEM_Open = True
                                            'PCB.U�ݕ� �t���p�X�捞��
    sts = GetIni("FILE", PCB_U_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PCB_U_ITEM]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, PCB_U_ITEM_POS, PCB_U_ITEM_REC, Len(PCB_U_ITEM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PCB_U_ITEM_Create()        'PCB.U�ݕύ쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PCB_U_ITEM_POS, PCB_U_ITEM_REC, Len(PCB_U_ITEM_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�@PCB.U�ݕ�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "PCB.U�ݕ�")
                Exit Function
        End Select
    Loop

    PCB_U_ITEM_Open = False

End Function

