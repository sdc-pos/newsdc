Attribute VB_Name = "M_ITEM"
Option Explicit
'********************************************************************
'*
'*              �i�ڃ}�X�^�i���W���[���j  �t�@�C����`
'*
'*          CREATE 2014.06.18
'********************************************************************
'�t�@�C���h�c
Public Const M_ITEM_ID$ = "M_ITEM"

'�y�[�W�T�C�Y
Public Const M_ITEM_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public M_ITEM_POS               As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type M_ITEM_REC_Tag
    JGYOBU(0 To 0)                  As Byte     '���ƕ��敪
    NAIGAI(0 To 0)                  As Byte     '�����O
    HIN_GAI(0 To 19)                As Byte     '�i�ԁi�O���j
    
    MODULE_KBN(0 To 0)              As Byte     '���W���[���Ώۋ敪
    MODULE_UNIT_KBN(0 To 0)         As Byte     '���W���[�����j�b�g�敪
    
    KENSA_JIGU(0 To 0)              As Byte     '��������
    SETUHEN_KBN(0 To 0)             As Byte     '�݌v�ύX�Ώۋ敪
    
    SETUHEN_LAST_DATE(0 To 7)       As Byte     '�݌v�ύX�ŏI��
    SENDO_LAST_DATE(0 To 7)         As Byte     '�N�x�Ǘ��ŏI��
    
    HITUYO_SU(0 To 4)               As Byte     '�K�v���@��
    HITUYO_TUKI(0 To 3)             As Byte     '�K�v���@��
    
    
    KANRI_NO(0 To 2)                As Byte     '�Ǘ���         '2017.03.27
    FILLER(0 To 133)                As Byte     'FILLER         '2017.03.27 136-->133
    
    INS_TANTO(0 To 9)               As Byte         '�ǉ��@�S����
    Ins_DateTime(0 To 13)           As Byte         '�ǉ��@����
    INS_PROG_ID(0 To 9)             As Byte         '�ǉ��@�v���O����ID
    
    UPD_TANTO(0 To 9)               As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)           As Byte         '�X�V�@����
    UPD_PROG_ID(0 To 9)             As Byte         '�X�V�@�v���O����ID



End Type
'�f�[�^�E�o�b�t�@
Public M_ITEM_REC               As M_ITEM_REC_Tag

'�L�[��`
Type KEY0_M_ITEM                '�j�d�x�O
    JGYOBU(0 To 0)                  As Byte     '���ƕ��敪
    NAIGAI(0 To 0)                  As Byte     '�����O
    HIN_GAI(0 To 19)                As Byte     '�i�ԁi�O���j
End Type








'�L�[�E�f�[�^
Public K0_M_ITEM                As KEY0_M_ITEM


Private Type M_ITEM_FSpeck
    fs      As BtFileSpeck              ' ̧�� ��߯��\����
    ks0     As BtKeySpeck               ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck

End Type

Private M_ITEM_Speck            As M_ITEM_FSpeck

Private Function M_ITEM_Create() As Integer
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

    M_ITEM_Create = True
                                            'PCB.U�ݕρ@�t���p�X�捞��
    sts = GetIni("FILE", M_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [M_ITEM]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    M_ITEM_Speck.fs.recoleng = Len(M_ITEM_REC)      ' ���R�[�h��
    M_ITEM_Speck.fs.PageSize = M_ITEM_PG_SIZ        ' �y�[�W�T�C�Y
    M_ITEM_Speck.fs.idexnumb = 1                        ' �C���f�b�N�X��
    M_ITEM_Speck.fs.fileflag = 0                        ' �t�@�C���t���O
    M_ITEM_Speck.fs.reserve = &H0                       ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    M_ITEM_Speck.ks0.keypos = 1                         ' �L�[�|�W�V����
    M_ITEM_Speck.ks0.keyleng = 1                        ' �L�[��
    M_ITEM_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    M_ITEM_Speck.ks0.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    M_ITEM_Speck.ks0.reserve = &H0                      ' �\��ς�

    M_ITEM_Speck.ks1.keypos = 2                         ' �L�[�|�W�V����
    M_ITEM_Speck.ks1.keyleng = 1                        ' �L�[��
                                                            
    M_ITEM_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    M_ITEM_Speck.ks1.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    M_ITEM_Speck.ks1.reserve = &H0                      ' �\��ς�

    M_ITEM_Speck.ks2.keypos = 3                         ' �L�[�|�W�V����
    M_ITEM_Speck.ks2.keyleng = 20                       ' �L�[��
                                                            
    M_ITEM_Speck.ks2.keyflag = BtKfExt                  ' �L�[�t���O
    M_ITEM_Speck.ks2.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    M_ITEM_Speck.ks2.reserve = &H0                      ' �\��ς�




'-----------------------------------------------

    sts = BTRV(BtOpCreate, M_ITEM_POS, M_ITEM_Speck, Len(M_ITEM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "PCB.U�ݕ�")
        Exit Function
    End If

    M_ITEM_Create = False

End Function

Public Function M_ITEM_Open(Mode As Integer) As Integer
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

    M_ITEM_Open = True
                                            'PCB.U�ݕ� �t���p�X�捞��
    sts = GetIni("FILE", M_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [M_ITEM]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = M_ITEM_Create()        'PCB.U�ݕύ쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), ByVal FullPath, Len(FullPath), Mode)
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

    M_ITEM_Open = False

End Function

