Attribute VB_Name = "KEPPIN"
Option Explicit
'********************************************************************
'*
'*              ���i�f�[�^  �t�@�C����`
'*
'*          CREATE 2013.08.23
'********************************************************************
'�t�@�C���h�c
Public Const KEPPIN_ID$ = "KEPPIN"

'�y�[�W�T�C�Y
Public Const KEPPIN_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public KEPPIN_POS               As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type KEPPINREC_Tag
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
    KEPPIN_CNT(0 To 7)          As Byte     '���i�@����
    KEPPIN_QTY(0 To 7)          As Byte     '���i�@��
End Type
'�f�[�^�E�o�b�t�@
Public KEPPINREC                As KEPPINREC_Tag

'�L�[��`

Type KEY0_KEPPIN                '�j�d�x�O
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
End Type
'�L�[�E�f�[�^
Public K0_KEPPIN                As KEY0_KEPPIN

Type KEPPIN_FSpeck
    fs      As BtFileSpeck                  ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                   ' �� ��߯��\����
End Type

Private KEPPIN_Speck            As KEPPIN_FSpeck

Private Function KEPPIN_Create() As Integer
'********************************************************************
'*
'*              ���i�f�[�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    KEPPIN_Create = True
                                            '���i�f�[�^ �t���p�X�捞��
    sts = GetIni("FILE", KEPPIN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [KEPPIN]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    KEPPIN_Speck.fs.recoleng = Len(KEPPINREC)       ' ���R�[�h��
    KEPPIN_Speck.fs.PageSize = KEPPIN_PG_SIZ        ' �y�[�W�T�C�Y
    KEPPIN_Speck.fs.idexnumb = 1                    ' �C���f�b�N�X��
    KEPPIN_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    KEPPIN_Speck.fs.reserve = &H0                   ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    KEPPIN_Speck.ks0.keypos = 1                     ' �L�[�|�W�V����
    KEPPIN_Speck.ks0.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    KEPPIN_Speck.ks0.keyflag = BtKfExt + BtKfChg
    KEPPIN_Speck.ks0.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    KEPPIN_Speck.ks0.reserve = &H0                  ' �\��ς�
'-----------------------------------------------

    sts = BTRV(BtOpCreate, KEPPIN_POS, KEPPIN_Speck, Len(KEPPIN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���i�f�[�^")
        Exit Function
    End If

    KEPPIN_Create = False

End Function

Public Function KEPPIN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���i�f�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    KEPPIN_Open = True
                                            '���i�f�[�^ �t���p�X�捞��
    sts = GetIni("FILE", KEPPIN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [KEPPIN]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, KEPPIN_POS, KEPPINREC, Len(KEPPINREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = KEPPIN_Create()        '���i�f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, KEPPIN_POS, KEPPINREC, Len(KEPPINREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���i�f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���i�f�[�^")
                Exit Function
        End Select
    Loop

    KEPPIN_Open = False

End Function

