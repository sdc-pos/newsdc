Attribute VB_Name = "tmpMENU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���j���[�Ǘ��}�X�^�i�ꎞ�t�@�C���j    �t�@�C����`  *
'*                                                                  *
'*          CREATE 2004.02.26                                       *
'********************************************************************
'�t�@�C���h�c
Public Const tmpMENU_ID$ = "tmpMENU"

'�y�[�W�T�C�Y
Public Const tmpMENU_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public tmpMENU_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type tmpMENUREC_Tag
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    MENU_LV1(0 To 2)        As Byte         '���j���[���x���P
    MENU_LV2(0 To 2)        As Byte         '���j���[���x���Q
    MENU_LV3(0 To 2)        As Byte         '���j���[���x���R
    DEL_FLG(0 To 0)         As Byte         '�폜�t���O
    MENU_KBN(0 To 0)        As Byte         '���j���\�敪
    DISPLAY_ITEM(0 To 19)   As Byte         '�\������
    CODE_TYPE(0 To 0)       As Byte         '��o�[�R�[�h�^�C�v
    YOIN_CODE(0 To 0)       As Byte         '�v��
    PARAM_F(0 To 0)         As Byte         '�t�����Ұ�(0:�Ȃ� 1:������)
    PARAM(0 To 15)          As Byte         '�p�����[�^

End Type

'�f�[�^�E�o�b�t�@
Public tmpMENUREC As tmpMENUREC_Tag

'�L�[��`

Type KEY0_tmpMENU               '�j�d�x�O
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    MENU_LV1(0 To 2)        As Byte         '���j���[���x���P
    MENU_LV2(0 To 2)        As Byte         '���j���[���x���Q
    MENU_LV3(0 To 2)        As Byte         '���j���[���x���R
End Type

'�L�[�E�f�[�^
Public K0_tmpMENU           As KEY0_tmpMENU

Type tmpMENU_FSpeck
    fs  As BtFileSpeck          ' ̧�� ��߯��\����
    ks0 As BtKeySpeck           ' �� ��߯��\����
    ks1 As BtKeySpeck           ' �� ��߯��\����
    ks2 As BtKeySpeck           ' �� ��߯��\����
    ks3 As BtKeySpeck           ' �� ��߯��\����
    ks4 As BtKeySpeck           ' �� ��߯��\����
End Type

Public tmpMENU_Speck        As tmpMENU_FSpeck
 
Private Function tmpMENU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���j���[�Ǘ��}�X�^  �b�q�d�`�s�d                    *
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

    tmpMENU_Create = False
                                            '���j���[�Ǘ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", tmpMENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        tmpMENU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    tmpMENU_Speck.fs.recoleng = Len(tmpMENUREC)         ' ���R�[�h��
    tmpMENU_Speck.fs.PageSize = tmpMENU_PG_SIZ%         ' �y�[�W�T�C�Y
    tmpMENU_Speck.fs.idexnumb = 1                       ' �C���f�b�N�X��
    tmpMENU_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    tmpMENU_Speck.fs.reserve = &H0                      ' �\��ς�
'-------------------------------------------------------
                                                        ' �L�[�O
    tmpMENU_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    tmpMENU_Speck.ks0.keyleng = 1                       ' �L�[��
    tmpMENU_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    tmpMENU_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    tmpMENU_Speck.ks0.reserve = &H0                     ' �\��ς�
    
    tmpMENU_Speck.ks1.keypos = 2                        ' �L�[�|�W�V����
    tmpMENU_Speck.ks1.keyleng = 1                       ' �L�[��
    tmpMENU_Speck.ks1.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    tmpMENU_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    tmpMENU_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    tmpMENU_Speck.ks2.keypos = 3                        ' �L�[�|�W�V����
    tmpMENU_Speck.ks2.keyleng = 3                       ' �L�[��
    tmpMENU_Speck.ks2.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    tmpMENU_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    tmpMENU_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    tmpMENU_Speck.ks3.keypos = 6                        ' �L�[�|�W�V����
    tmpMENU_Speck.ks3.keyleng = 3                       ' �L�[��
    tmpMENU_Speck.ks3.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    tmpMENU_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    tmpMENU_Speck.ks3.reserve = &H0                     ' �\��ς�
    
    tmpMENU_Speck.ks4.keypos = 9                        ' �L�[�|�W�V����
    tmpMENU_Speck.ks4.keyleng = 3                       ' �L�[��
    tmpMENU_Speck.ks4.keyflag = BtKfExt                 ' �L�[�t���O
    tmpMENU_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    tmpMENU_Speck.ks4.reserve = &H0                     ' �\��ς�
'-------------------------------------------------------


    sts = BTRV(BtOpCreate, tmpMENU_POS, tmpMENU_Speck, Len(tmpMENU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���j���[�Ǘ��}�X�^(�ꎞ�t�@�C��)")
        tmpMENU_Create = True
    End If

End Function

Function tmpMENU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ���j���[�Ǘ��}�X�^  �n�o�d�m                        *
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
    
    tmpMENU_Open = False
                                            '���j���[�Ǘ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", tmpMENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        tmpMENU_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpMENU_Create()      '���j���[�Ǘ��}�X�^�쐬
                If sts <> False Then
                    tmpMENU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���j���[�Ǘ��}�X�^�i�ꎞ�t�@�C���j")
                    tmpMENU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "���j���[�Ǘ��}�X�^�i�ꎞ�t�@�C���j")
                tmpMENU_Open = True
                Exit Function
        End Select
    Loop
End Function
