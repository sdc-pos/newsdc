Attribute VB_Name = "LotNo"
Option Explicit
'********************************************************************
'*
'*              ���g�Ǘ��f�[�^�@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const LOTNO_ID$ = "LOTNO"

'�y�[�W�T�C�Y
Public Const LOTNO_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public LOTNO_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type LOTNOREC_Tag
    Model(0 To 19)      As Byte             '�i��
    PLotNo(0 To 19)     As Byte             '�����ԍ�
    IQty(0 To 5)        As Byte             '���א�
    OQty(0 To 5)        As Byte             '�o�א�
    SQty(0 To 5)        As Byte             '�݌ɐ�
    EDt(0 To 7)         As Byte             '�A�o��(�؍�)
    IDt(0 To 7)         As Byte             '���ד�
    ODt(0 To 7)         As Byte             '�o�ד�
    MemoNo(0 To 19)     As Byte             'No(����)
    EntFN(0 To 39)      As Byte             '�o�^̧�ٖ�
    ITantoCode(0 To 4)  As Byte             '���גS����ID
    OTantoCode(0 To 4)  As Byte             '�o�גS����ID
    FILLER(0 To 69)     As Byte             '
    EntID(0 To 9)       As Byte             '�o�^ID
    EntDtm(0 To 13)     As Byte             '�o�^����yyyymmddhhmmss
    UpdID(0 To 9)       As Byte             '�X�VID
    UpdDtm(0 To 13)     As Byte             '�X�V���� yyyymmddhhmmss
End Type

'�f�[�^�E�o�b�t�@
Public LOTNOREC         As LOTNOREC_Tag

'�L�[��`
Type KEY0_LOTNO         '�j�d�x�O
    Model(0 To 19)      As Byte             '�i��
    PLotNo(0 To 19)     As Byte             '�����ԍ�
End Type

'�L�[�E�f�[�^
Public K0_LOTNO         As KEY0_LOTNO

Type LOTNO_FSpeck
    fs      As BtFileSpeck                  '̧�� ��߯��\����
    ks0     As BtKeySpeck                   '�� ��߯��\����
    ks1     As BtKeySpeck                   '�� ��߯��\����
End Type

Private LOTNO_Speck     As LOTNO_FSpeck

Private Function LOTNO_Create() As Integer
'********************************************************************
'*
'*              ���g�Ǘ��f�[�^�@Create
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    LOTNO_Create = True
                                            '���g�Ǘ��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", LOTNO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [LOTNO]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    LOTNO_Speck.fs.recoleng = Len(LOTNOREC)             ' ���R�[�h��
    LOTNO_Speck.fs.PageSize = LOTNO_PG_SIZ              ' �y�[�W�T�C�Y
    LOTNO_Speck.fs.idexnumb = 1                         ' �C���f�b�N�X��
    LOTNO_Speck.fs.fileflag = 0                         ' �t�@�C���t���O
    LOTNO_Speck.fs.reserve = &H0                        ' �\��ς�
                                                    
'---------------------------------------------------
                                                            ' �L�[�O
    LOTNO_Speck.ks0.keypos = 1                              ' �L�[�|�W�V����
    LOTNO_Speck.ks0.keyleng = 20                            ' �L�[��
    LOTNO_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' �L�[�t���O
    LOTNO_Speck.ks0.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    LOTNO_Speck.ks0.reserve = &H0                           ' �\��ς�

    LOTNO_Speck.ks1.keypos = 21                             ' �L�[�|�W�V����
    LOTNO_Speck.ks1.keyleng = 20                            ' �L�[��
    LOTNO_Speck.ks1.keyflag = BtKfExt + BtKfChg             ' �L�[�t���O
    LOTNO_Speck.ks1.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    LOTNO_Speck.ks1.reserve = &H0                           ' �\��ς�

'---------------------------------------------------

    sts = BTRV(BtOpCreate, LOTNO_POS, LOTNO_Speck, Len(LOTNO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���g�Ǘ��f�[�^")
        Exit Function
    End If

    LOTNO_Create = False

End Function

Public Function LOTNO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���g�Ǘ��f�[�^�@Open
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    LOTNO_Open = True
                                            '���g�Ǘ��f�[�^ �t���p�X�捞��
    sts = GetIni("FILE", LOTNO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [LOTNO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, LOTNO_POS, LOTNOREC, Len(LOTNOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = LOTNO_Create()        '���g�Ǘ��f�[�^ �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, LOTNO_POS, LOTNOREC, Len(LOTNOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���g�Ǘ��f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���g�Ǘ��f�[�^")
                Exit Function
        End Select
    Loop

    LOTNO_Open = False

End Function
