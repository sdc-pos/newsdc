Attribute VB_Name = "P_TANTOMENU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �V�S���ҕʃ��j���[  �t�@�C����`                    *
'*                                                                  *
'*          CREATE 2004.02.20                                       *
'********************************************************************
'�t�@�C���h�c
Public Const P_TMENU_ID$ = "P_TANTOMENU"

'�y�[�W�T�C�Y
Public Const P_TMENU_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public P_TMENU_POS            As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`

Private Type MENU_NO_Tag
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    MENU_NO(0 To 1)         As Byte         '�ƭ���

End Type

Type P_TMENUREC_Tag
    TANTO_CODE(0 To 4)      As Byte         '�S���҃R�[�h
    
    MENU_T(0 To 179)        As MENU_NO_Tag  '�ƭ���     29-->179 2006.10.11
    
    FILLER(0 To 298)        As Byte         'FILLER

End Type

'�f�[�^�E�o�b�t�@
Public P_TMENUREC           As P_TMENUREC_Tag

'�L�[��`

Type KEY0_P_TMENU                           '�j�d�x�O
    TANTO_CODE(0 To 4)      As Byte         '�S���҃R�[�h
End Type

'�L�[�E�f�[�^
Public K0_P_TMENU           As KEY0_P_TMENU

Type P_TMENU_FSpeck
    fs  As BtFileSpeck          ' ̧�� ��߯��\����
    ks0 As BtKeySpeck           ' �� ��߯��\����
End Type

Private P_TMENU_Speck       As P_TMENU_FSpeck
 
Private Function P_TMENU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �V�S���ҕʃ��j���[  �b�q�d�`�s�d                    *
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

    P_TMENU_Create = False
                                            '�S���ҕʃ��j���[�t���p�X�捞��
    sts = GetIni("FILE", P_TMENU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_TMENU_ID]�ǂݍ��݃G���[")
        P_TMENU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    P_TMENU_Speck.fs.recoleng = Len(P_TMENUREC)         ' ���R�[�h��
    P_TMENU_Speck.fs.PageSize = P_TMENU_PG_SIZ%         ' �y�[�W�T�C�Y
    P_TMENU_Speck.fs.idexnumb = 1                       ' �C���f�b�N�X��
    P_TMENU_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    P_TMENU_Speck.fs.reserve = &H0                      ' �\��ς�
'--------------------------------------------------------
                                                        ' �L�[�O
    P_TMENU_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    P_TMENU_Speck.ks0.keyleng = 5                       ' �L�[��
    P_TMENU_Speck.ks0.keyflag = BtKfExt                 ' �L�[�t���O
    P_TMENU_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_TMENU_Speck.ks0.reserve = &H0                     ' �\��ς�
    
'--------------------------------------------------------

    sts = BTRV(BtOpCreate, P_TMENU_POS, P_TMENU_Speck, Len(P_TMENU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�V�S���ҕʃ��j���[")
        P_TMENU_Create = True
    End If

End Function

Function P_TMENU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �V�S���ҕʃ��j���[  �n�o�d�m                        *
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
    
    P_TMENU_Open = False
                                            '�S���ҕʃ��j���[�t���p�X�捞��
    sts = GetIni("FILE", P_TMENU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_TMENU_ID]�ǂݍ��݃G���[")
        P_TMENU_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_TMENU_Create()        '�S���ҕʃ��j���[�쐬
                If sts <> False Then
                    P_TMENU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�V���j���[�Ǘ��}�X�^")
                    P_TMENU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�V���j���[�Ǘ��}�X�^")
                P_TMENU_Open = True
                Exit Function
        End Select
    Loop
End Function
