Attribute VB_Name = "SOKO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �q�Ƀ}�X�^  �t�@�C����`                            *
'*                                                                  *
'*          CREATE 2004.02.16                                       *
'********************************************************************
'�t�@�C���h�c
Public Const SOKO_ID$ = "SOKO"

'�y�[�W�T�C�Y
Public Const SOKO_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public SOKO_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SOKOREC_Tag
    JGYOBU(0 To 0)      As Byte         '���ƕ��敪
    Soko_No(0 To 1)     As Byte         '�q�ɇ�
    SOKO_NAME(0 To 15)  As Byte         '�q�ɖ���
    SOKO_BUN(0 To 0)    As Byte         '�q�ɕ���
    SOKO_KBN(0 To 0)    As Byte         '�q�ɋ敪
    NAIGAI(0 To 0)      As Byte         '�����O
    KAHI_KBN(0 To 0)    As Byte         '�g�p��
    KONS_KBN(0 To 0)    As Byte         '���ډ�
    RETU_START(0 To 1)  As Byte         '�I�Ԕ͈́@��@�J�n
    RETU_END(0 To 1)    As Byte         '�I�Ԕ͈́@��@�I��
    REN_START(0 To 1)   As Byte         '�I�Ԕ͈́@�A�@�J�n
    REN_END(0 To 1)     As Byte         '�I�Ԕ͈́@�A�@�I��
    DAN_START(0 To 1)   As Byte         '�I�Ԕ͈́@�i�@�J�n
    DAN_END(0 To 1)     As Byte         '�I�Ԕ͈́@�i�@�I��
    
    ORDER_POINT(0 To 2) As Byte         '�����_ 2004.02
    GOODS_ON_F(0 To 0)  As Byte         '���i���q�Ƀt���O 2004.02
    
    
    IO_TANKA_No(0 To 1) As Byte         '���o�ɒP���ݒ躰�� 2008.02.14
    
    FILLER(0 To 13)     As Byte         'FILLER
End Type
'�f�[�^�E�o�b�t�@
Public SOKOREC As SOKOREC_Tag

'�L�[��`

Type KEY0_SOKO            '�j�d�x�O
    Soko_No(0 To 1)     As Byte         '�q�ɇ�
End Type
    
'�L�[�E�f�[�^
Public K0_SOKO          As KEY0_SOKO

Type SOKO_FSpeck
    fs                  As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                 As BtKeySpeck   ' �� ��߯��\����
End Type

Private SOKO_Speck       As SOKO_FSpeck
Private Function SOKO_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �q�Ƀ}�X�^  �b�q�d�`�s�d                            *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SOKO_Create = True
                                            '�q�Ƀ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", SOKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SOKO]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    SOKO_Speck.fs.recoleng = Len(SOKOREC)       ' ���R�[�h��
    SOKO_Speck.fs.PageSize = SOKO_PG_SIZ        ' �y�[�W�T�C�Y
    SOKO_Speck.fs.idexnumb = 1                  ' �C���f�b�N�X��
    SOKO_Speck.fs.fileflag = 0                  ' �t�@�C���t���O
    SOKO_Speck.fs.reserve = &H0                 ' �\��ς�
                                                ' �L�[�O
    SOKO_Speck.ks0.keypos = 2                   ' �L�[�|�W�V����
    SOKO_Speck.ks0.keyleng = 2                  ' �L�[��
    SOKO_Speck.ks0.keyflag = BtKfExt            ' �L�[�t���O
    SOKO_Speck.ks0.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    SOKO_Speck.ks0.reserve = &H0                ' �\��ς�

    sts = BTRV(BtOpCreate, SOKO_POS, SOKO_Speck, Len(SOKO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�q�Ƀ}�X�^")
        Exit Function
    End If
    SOKO_Create = False
End Function

Function SOKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �q�Ƀ}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    SOKO_Open = True
                                            '�q�Ƀ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", SOKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, SOKO_POS, SOKOREC, Len(SOKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SOKO_Create()        '�q�Ƀ}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SOKO_POS, SOKOREC, Len(SOKOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�q�Ƀ}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�q�Ƀ}�X�^")
                Exit Function
        End Select
    Loop
    SOKO_Open = False

End Function
