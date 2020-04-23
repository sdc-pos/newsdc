Attribute VB_Name = "P_MENU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �V���j���[�Ǘ��}�X�^    �t�@�C����`                *
'*                                                                  *
'*          CREATE 2004.02.20                                       *
'********************************************************************
'�t�@�C���h�c
Public Const P_MENU_ID$ = "P_MENU"

'�y�[�W�T�C�Y
Public Const P_MENU_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_MENU_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`

Private Type SAGYO_Tag
    YOIN(0 To 1)            As Byte         '�v��
    PARAM(0 To 15)          As Byte         '���Ұ�(������)
    Disp(0 To 19)           As Byte         '���Ұ�(������)
    LOG_OUT(0 To 0)         As Byte         '۸ޏo�� 0:�o�͂Ȃ� 1:����
End Type


Type P_MENUREC_Tag
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    MENU_NO(0 To 1)         As Byte         '���j���[�O���[�v��
    MENU_DSP(0 To 19)       As Byte         '�\�����e
    SAGYO(0 To 19)          As SAGYO_Tag    '��Ɠ��e
    FILLER(0 To 175)        As Byte         '��Ɠ��e
End Type

'�f�[�^�E�o�b�t�@
Public P_MENUREC            As P_MENUREC_Tag

'�L�[��`

Type KEY0_P_MENU                '�j�d�x�O
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    MENU_NO(0 To 1)         As Byte         '���j���[�O���[�v��
End Type
'�L�[�E�f�[�^
Public K0_P_MENU            As KEY0_P_MENU

Type P_MENU_FSpeck
    fs  As BtFileSpeck              ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck               ' �� ��߯��\����
    ks2 As BtKeySpeck               ' �� ��߯��\����
End Type

Public P_MENU_Speck         As P_MENU_FSpeck
 
Private Function P_MENU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �V���j���[�Ǘ��}�X�^  �b�q�d�`�s�d                    *
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

    P_MENU_Create = False
                                            '���j���[�Ǘ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_MENU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        P_MENU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    P_MENU_Speck.fs.recoleng = Len(P_MENUREC)           ' ���R�[�h��
    P_MENU_Speck.fs.PageSize = P_MENU_PG_SIZ%           ' �y�[�W�T�C�Y
    P_MENU_Speck.fs.idexnumb = 1                        ' �C���f�b�N�X��
    P_MENU_Speck.fs.fileflag = 0                        ' �t�@�C���t���O
    P_MENU_Speck.fs.reserve = &H0                       ' �\��ς�
'-------------------------------------------------------
                                                        ' �L�[�O
    P_MENU_Speck.ks0.keypos = 1                         ' �L�[�|�W�V����
    P_MENU_Speck.ks0.keyleng = 1                        ' �L�[��
    P_MENU_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    P_MENU_Speck.ks0.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    P_MENU_Speck.ks0.reserve = &H0                      ' �\��ς�
    
    P_MENU_Speck.ks1.keypos = 2                         ' �L�[�|�W�V����
    P_MENU_Speck.ks1.keyleng = 1                        ' �L�[��
    P_MENU_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    P_MENU_Speck.ks1.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    P_MENU_Speck.ks1.reserve = &H0                      ' �\��ς�
    
    P_MENU_Speck.ks2.keypos = 3                         ' �L�[�|�W�V����
    P_MENU_Speck.ks2.keyleng = 2                        ' �L�[��
    P_MENU_Speck.ks2.keyflag = BtKfExt                  ' �L�[�t���O
    P_MENU_Speck.ks2.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    P_MENU_Speck.ks2.reserve = &H0                      ' �\��ς�
    
'-------------------------------------------------------

    sts = BTRV(BtOpCreate, P_MENU_POS, P_MENU_Speck, Len(P_MENU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���j���[�Ǘ��}�X�^")
        P_MENU_Create = True
    End If

End Function

Function P_MENU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �V���j���[�Ǘ��}�X�^  �n�o�d�m                      *
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
    
    P_MENU_Open = False
                                            '���j���[�Ǘ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_MENU_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        P_MENU_Open = True
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, P_MENU_POS, P_MENUREC, Len(P_MENUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_MENU_Create()         '���j���[�Ǘ��}�X�^�쐬
                If sts <> False Then
                    P_MENU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_MENU_POS, P_MENUREC, Len(P_MENUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�V���j���[�Ǘ��}�X�^")
                    P_MENU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�V���j���[�Ǘ��}�X�^")
                P_MENU_Open = True
                Exit Function
        End Select
    Loop
End Function
