Attribute VB_Name = "TANATBL"
Option Explicit
'********************************************************************
'*
'*              �I�ԓǑւ��e�[�u��  �t�@�C����`
'*
'*          CREATE 2001.06.13
'********************************************************************
'�t�@�C���h�c
Global Const TANATBL_ID = "TANATBL"

'�y�[�W�T�C�Y
Global Const TANATBL_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Global TANATBL_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type TANATBLREC_Tag
    HOST_TANA(0 To 7)   As Byte
    FILLER(0 To 0)      As Byte
    POS_TANA(0 To 7)    As Byte
End Type

'�f�[�^�E�o�b�t�@
Global TANATBLREC As TANATBLREC_Tag

'�L�[��`

Type KEY0_TANATBL                 '�j�d�x�O
    HOST_TANA(0 To 7)   As Byte
End Type

'�L�[�E�f�[�^
Global K0_TANATBL As KEY0_TANATBL

Type TANATBL_FSpeck
    fs As BtFileSpeck           ' ̧�� ��߯��\����
    ks0 As BtKeySpeck           ' �� ��߯��\����
End Type

Global TANATBL_Speck As TANATBL_FSpeck

Private Function TANATBL_Create() As Integer
'********************************************************************
'*
'*              �I�ԓǑւ��e�[�u��  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2001.02.14
'********************************************************************
Dim sts As Integer
Dim FullPath As String
Dim c As String * 128

    TANATBL_Create = False
                                            '�I�ԓǑւ��e�[�u���t���p�X�捞��
    sts = GetIni("FILE", TANATBL_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        TANATBL_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    TANATBL_Speck.fs.recoleng = Len(TANATBLREC)         ' ���R�[�h��
    TANATBL_Speck.fs.PageSize = TANATBL_PG_SIZ          ' �y�[�W�T�C�Y
    TANATBL_Speck.fs.idexnumb = 1                   ' �C���f�b�N�X��
    TANATBL_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    TANATBL_Speck.fs.reserve = &H0                  ' �\��ς�
                                                ' �L�[�O
    TANATBL_Speck.ks0.keypos = 1                    ' �L�[�|�W�V����
    TANATBL_Speck.ks0.keyleng = 8                   ' �L�[��
    TANATBL_Speck.ks0.keyflag = BtKfExt             ' �L�[�t���O
    TANATBL_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    TANATBL_Speck.ks0.reserve = &H0                 ' �\��ς�
    sts = BTRV(BtOpCreate, TANATBL_POS, TANATBL_Speck, Len(TANATBL_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�I�ԓǑւ��e�[�u��")
        TANATBL_Create = True
    End If
End Function

Function TANATBL_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �I�ԓǑւ��e�[�u��  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2001.06.13
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    TANATBL_Open = False
                                            '�I�ԓǑւ��e�[�u���t���p�X�捞��
    sts = GetIni("FILE", TANATBL_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        TANATBL_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, TANATBL_POS, TANATBLREC, Len(TANATBLREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = TANATBL_Create()        '�I�ԓǑւ��e�[�u���쐬
                If sts <> False Then
                    TANATBL_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, TANATBL_POS, TANATBLREC, Len(TANATBLREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�I�ԓǑւ��e�[�u��")
                    TANATBL_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�I�ԓǑւ��e�[�u��")
                TANATBL_Open = True
                Exit Function
        End Select
    Loop
End Function
