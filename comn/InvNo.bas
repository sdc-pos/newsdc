Attribute VB_Name = "InvNo"
Option Explicit
'********************************************************************
'*
'*              ���g���󇂃f�[�^�@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const INVNO_ID$ = "INVNO"

'�y�[�W�T�C�Y
Public Const INVNO_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public INVNO_POS  As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type INVNOREC_Tag
    INVNO(0 To 19)      As Byte             '����
    Model(0 To 19)      As Byte             '�i��
    LotNo(0 To 19)      As Byte             '�����ԍ�
    OQty(0 To 5)        As Byte             '�o�א�
    ODt(0 To 7)         As Byte             '�o�ד�
    FILLER(0 To 133)    As Byte             '
    EntID(0 To 9)       As Byte             '�o�^ID
    EntDtm(0 To 13)     As Byte             '�o�^����yyyymmddhhmmss
    UpdID(0 To 9)       As Byte             '�X�VID
    UpdDtm(0 To 13)     As Byte             '�X�V���� yyyymmddhhmmss
End Type

'�f�[�^�E�o�b�t�@
Public INVNOREC         As INVNOREC_Tag

'�L�[��`
Type KEY0_INVNO     '�j�d�x�O
    Model(0 To 19)      As Byte             '�i��
    LotNo(0 To 19)      As Byte             '�����ԍ�
End Type

Type KEY1_INVNO     '�j�d�x�P
    INVNO(0 To 19)      As Byte             '����
End Type


'�L�[�E�f�[�^
Public K0_INVNO         As KEY0_INVNO

Type INVNO_FSpeck
    fs      As BtFileSpeck                  '̧�� ��߯��\����
    ks0     As BtKeySpeck                   '�� ��߯��\����
    ks1     As BtKeySpeck                   '�� ��߯��\����
    ks2     As BtKeySpeck                   '�� ��߯��\����
End Type

Private INVNO_Speck     As INVNO_FSpeck

Private Function INVNO_Create() As Integer
'********************************************************************
'*
'*              ���g���󇂃f�[�^�@Create
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    INVNO_Create = True
                                            '���g�Ǘ��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", INVNO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [INVNO]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    INVNO_Speck.fs.recoleng = Len(INVNOREC)             ' ���R�[�h��
    INVNO_Speck.fs.PageSize = INVNO_PG_SIZ              ' �y�[�W�T�C�Y
    INVNO_Speck.fs.idexnumb = 2                         ' �C���f�b�N�X��
    INVNO_Speck.fs.fileflag = 0                         ' �t�@�C���t���O
    INVNO_Speck.fs.reserve = &H0                        ' �\��ς�
                                                    
'---------------------------------------------------
                                                            ' �L�[�O
    INVNO_Speck.ks0.keypos = 21                             ' �L�[�|�W�V����
    INVNO_Speck.ks0.keyleng = 20                            ' �L�[��
    INVNO_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' �L�[�t���O
    INVNO_Speck.ks0.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    INVNO_Speck.ks0.reserve = &H0                           ' �\��ς�

    INVNO_Speck.ks1.keypos = 41                             ' �L�[�|�W�V����
    INVNO_Speck.ks1.keyleng = 20                            ' �L�[��
    INVNO_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfChg               ' �L�[�t���O
    INVNO_Speck.ks1.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    INVNO_Speck.ks1.reserve = &H0                           ' �\��ς�

'---------------------------------------------------


'---------------------------------------------------
                                                            ' �L�[�P
    INVNO_Speck.ks2.keypos = 1                              ' �L�[�|�W�V����
    INVNO_Speck.ks2.keyleng = 20                            ' �L�[��
    INVNO_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg   ' �L�[�t���O
    INVNO_Speck.ks2.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    INVNO_Speck.ks2.reserve = &H0                           ' �\��ς�


    sts = BTRV(BtOpCreate, INVNO_POS, INVNO_Speck, Len(INVNO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���g���󇂃f�[�^")
        Exit Function
    End If

    INVNO_Create = False

End Function

Public Function INVNO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���g���󇂃f�[�^�@Open
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    INVNO_Open = True
                                            '���g���󇂃f�[�^ �t���p�X�捞��
    sts = GetIni("FILE", INVNO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [INVNO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, INVNO_POS, INVNOREC, Len(INVNOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = INVNO_Create()        '���g���󇂃f�[�^ �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, INVNO_POS, INVNOREC, Len(INVNOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���g���󇂃f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���g���󇂃f�[�^")
                Exit Function
        End Select
    Loop

    INVNO_Open = False

End Function
