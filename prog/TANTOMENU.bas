Attribute VB_Name = "TANTOMENU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �S���ҕʃ��j���[  �t�@�C����`                      *
'*                                                                  *
'*          CREATE 2004.02.20                                       *
'********************************************************************
'�t�@�C���h�c
Public Const TMENU_ID$ = "TANTOMENU"

'�y�[�W�T�C�Y
Public Const TMENU_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public TMENU_POS            As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type TMENUREC_Tag
    
    TANTO_CODE(0 To 4)      As Byte         '�S���҃R�[�h
    MENU_GRP_NO(0 To 1)     As Byte         '���j���[�O���[�v
    FILLER(0 To 16)         As Byte         'FILLER

End Type

'�f�[�^�E�o�b�t�@
Public TMENUREC             As TMENUREC_Tag

'�L�[��`

Type KEY0_TMENU                         '�j�d�x�O
    TANTO_CODE(0 To 4)      As Byte         '�S���҃R�[�h
End Type

'�L�[�E�f�[�^
Public K0_TMENU             As KEY0_TMENU

Type TMENU_FSpeck
    fs  As BtFileSpeck          ' ̧�� ��߯��\����
    ks0 As BtKeySpeck           ' �� ��߯��\����
End Type

Private TMENU_Speck          As TMENU_FSpeck
 
Private Function TMENU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �S���ҕʃ��j���[  �b�q�d�`�s�d                      *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.02.14                                       *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    TMENU_Create = False
                                            '�S���ҕʃ��j���[�t���p�X�捞��
    sts = GetIni("FILE", TMENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        TMENU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    TMENU_Speck.fs.recoleng = Len(TMENUREC)             ' ���R�[�h��
    TMENU_Speck.fs.PageSize = TMENU_PG_SIZ%             ' �y�[�W�T�C�Y
    TMENU_Speck.fs.idexnumb = 1                         ' �C���f�b�N�X��
    TMENU_Speck.fs.fileflag = 0                         ' �t�@�C���t���O
    TMENU_Speck.fs.reserve = &H0                        ' �\��ς�
'--------------------------------------------------------
                                                        ' �L�[�O
    TMENU_Speck.ks0.keypos = 1                          ' �L�[�|�W�V����
    TMENU_Speck.ks0.keyleng = 5                         ' �L�[��
    TMENU_Speck.ks0.keyflag = BtKfExt                   ' �L�[�t���O
    TMENU_Speck.ks0.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    TMENU_Speck.ks0.reserve = &H0                       ' �\��ς�
    
'--------------------------------------------------------

    sts = BTRV(BtOpCreate, TMENU_POS, TMENU_Speck, Len(TMENU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�S���ҕʃ��j���[")
        TMENU_Create = True
    End If

End Function

Function TMENU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �S���ҕʃ��j���[  �n�o�d�m                          *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.02.14                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    TMENU_Open = False
                                            '�S���ҕʃ��j���[�t���p�X�捞��
    sts = GetIni("FILE", TMENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        TMENU_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, TMENU_POS, TMENUREC, Len(TMENUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = TMENU_Create()        '�S���ҕʃ��j���[�쐬
                If sts <> False Then
                    TMENU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, TMENU_POS, TMENUREC, Len(TMENUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���j���[�Ǘ��}�X�^")
                    TMENU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "���j���[�Ǘ��}�X�^")
                TMENU_Open = True
                Exit Function
        End Select
    Loop
End Function
