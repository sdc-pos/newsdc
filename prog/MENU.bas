Attribute VB_Name = "MENU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���j���[�Ǘ��}�X�^    �t�@�C����`                  *
'*                                                                  *
'*          CREATE 2004.02.20                                       *
'********************************************************************
'�t�@�C���h�c
Public Const MENU_ID$ = "MENU"

'�y�[�W�T�C�Y
Public Const MENU_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public MENU_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type MENUREC_Tag
    MENU_GRP_NO(0 To 1)     As Byte         '���j���[�O���[�v��
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    MENU_LV1(0 To 2)        As Byte         '���j���[���x���P
    MENU_LV2(0 To 2)        As Byte         '���j���[���x���Q
    MENU_LV3(0 To 2)        As Byte         '���j���[���x���R
    MENU_GRP(0 To 19)       As Byte         '���j���[�O���[�v
    MENU_KBN(0 To 0)        As Byte         '���j���\�敪
    DISPLAY_ITEM(0 To 19)   As Byte         '�\������
    CODE_TYPE(0 To 0)       As Byte         '��o�[�R�[�h�^�C�v
    YOIN_CODE(0 To 0)       As Byte         '�v��
    PARAM(0 To 15)          As Byte         '�p�����[�^
    FILLER(0 To 23)         As Byte         'FILLER

End Type

'�f�[�^�E�o�b�t�@
Public MENUREC As MENUREC_Tag

'�L�[��`

Type KEY0_MENU                  '�j�d�x�O
    MENU_GRP_NO(0 To 1)     As Byte         '���j���[�O���[�v��
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    MENU_LV1(0 To 2)        As Byte         '���j���[���x���P
    MENU_LV2(0 To 2)        As Byte         '���j���[���x���Q
    MENU_LV3(0 To 2)        As Byte         '���j���[���x���R
End Type

Type KEY1_MENU                  '�j�d�x�P
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NAIGAI(0 To 0)          As Byte         '�����O
    MENU_GRP_NO(0 To 1)     As Byte         '���j���[�O���[�v��
    MENU_LV1(0 To 2)        As Byte         '���j���[���x���P
    MENU_LV2(0 To 2)        As Byte         '���j���[���x���Q
    MENU_LV3(0 To 2)        As Byte         '���j���[���x���R
End Type

'�L�[�E�f�[�^
Public K0_MENU              As KEY0_MENU
Public K1_MENU              As KEY1_MENU

Type MENU_FSpeck
    fs  As BtFileSpeck              ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck               ' �� ��߯��\����
    ks2 As BtKeySpeck               ' �� ��߯��\����
    ks3 As BtKeySpeck               ' �� ��߯��\����
    ks4 As BtKeySpeck               ' �� ��߯��\����
    ks5 As BtKeySpeck               ' �� ��߯��\����
    ks6 As BtKeySpeck               ' �� ��߯��\����
    ks7 As BtKeySpeck               ' �� ��߯��\����
    ks8 As BtKeySpeck               ' �� ��߯��\����
    ks9 As BtKeySpeck               ' �� ��߯��\����
    ks10 As BtKeySpeck              ' �� ��߯��\����
    ks11 As BtKeySpeck              ' �� ��߯��\����
End Type

Public MENU_Speck           As MENU_FSpeck
 
Private Function MENU_Create() As Integer
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

    MENU_Create = False
                                            '���j���[�Ǘ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", MENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        MENU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    MENU_Speck.fs.recoleng = Len(MENUREC)               ' ���R�[�h��
    MENU_Speck.fs.PageSize = MENU_PG_SIZ%               ' �y�[�W�T�C�Y
    MENU_Speck.fs.idexnumb = 2                          ' �C���f�b�N�X��
    MENU_Speck.fs.fileflag = 0                          ' �t�@�C���t���O
    MENU_Speck.fs.reserve = &H0                         ' �\��ς�
'-------------------------------------------------------
                                                        ' �L�[�O
    MENU_Speck.ks0.keypos = 1                           ' �L�[�|�W�V����
    MENU_Speck.ks0.keyleng = 2                          ' �L�[��
    MENU_Speck.ks0.keyflag = BtKfExt + BtKfSeg          ' �L�[�t���O
    MENU_Speck.ks0.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks0.reserve = &H0                        ' �\��ς�
    
    MENU_Speck.ks1.keypos = 3                           ' �L�[�|�W�V����
    MENU_Speck.ks1.keyleng = 1                          ' �L�[��
    MENU_Speck.ks1.keyflag = BtKfExt + BtKfSeg          ' �L�[�t���O
    MENU_Speck.ks1.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks1.reserve = &H0                        ' �\��ς�
    
    MENU_Speck.ks2.keypos = 4                           ' �L�[�|�W�V����
    MENU_Speck.ks2.keyleng = 1                          ' �L�[��
    MENU_Speck.ks2.keyflag = BtKfExt + BtKfSeg          ' �L�[�t���O
    MENU_Speck.ks2.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks2.reserve = &H0                        ' �\��ς�
    
    MENU_Speck.ks3.keypos = 5                           ' �L�[�|�W�V����
    MENU_Speck.ks3.keyleng = 3                          ' �L�[��
    MENU_Speck.ks3.keyflag = BtKfExt + BtKfSeg          ' �L�[�t���O
    MENU_Speck.ks3.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks3.reserve = &H0                        ' �\��ς�
    
    MENU_Speck.ks4.keypos = 8                           ' �L�[�|�W�V����
    MENU_Speck.ks4.keyleng = 3                          ' �L�[��
    MENU_Speck.ks4.keyflag = BtKfExt + BtKfSeg          ' �L�[�t���O
    MENU_Speck.ks4.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks4.reserve = &H0                        ' �\��ς�
    
    MENU_Speck.ks5.keypos = 11                          ' �L�[�|�W�V����
    MENU_Speck.ks5.keyleng = 3                          ' �L�[��
    MENU_Speck.ks5.keyflag = BtKfExt                    ' �L�[�t���O
    MENU_Speck.ks5.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks5.reserve = &H0                        ' �\��ς�
'-------------------------------------------------------
                                                        ' �L�[�O
    MENU_Speck.ks6.keypos = 3                           ' �L�[�|�W�V����
    MENU_Speck.ks6.keyleng = 1                          ' �L�[��
    MENU_Speck.ks6.keyflag = BtKfExt + BtKfSeg          ' �L�[�t���O
    MENU_Speck.ks6.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks6.reserve = &H0                        ' �\��ς�
    
    MENU_Speck.ks7.keypos = 4                           ' �L�[�|�W�V����
    MENU_Speck.ks7.keyleng = 1                          ' �L�[��
    MENU_Speck.ks7.keyflag = BtKfExt + BtKfSeg          ' �L�[�t���O
    MENU_Speck.ks7.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks7.reserve = &H0                        ' �\��ς�
    
    MENU_Speck.ks8.keypos = 1                           ' �L�[�|�W�V����
    MENU_Speck.ks8.keyleng = 2                          ' �L�[��
    MENU_Speck.ks8.keyflag = BtKfExt + BtKfSeg          ' �L�[�t���O
    MENU_Speck.ks8.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks8.reserve = &H0                        ' �\��ς�
    
    MENU_Speck.ks9.keypos = 5                           ' �L�[�|�W�V����
    MENU_Speck.ks9.keyleng = 3                          ' �L�[��
    MENU_Speck.ks9.keyflag = BtKfExt + BtKfSeg          ' �L�[�t���O
    MENU_Speck.ks9.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks9.reserve = &H0                        ' �\��ς�
    
    MENU_Speck.ks10.keypos = 8                           ' �L�[�|�W�V����
    MENU_Speck.ks10.keyleng = 3                          ' �L�[��
    MENU_Speck.ks10.keyflag = BtKfExt + BtKfSeg          ' �L�[�t���O
    MENU_Speck.ks10.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks10.reserve = &H0                        ' �\��ς�
    
    MENU_Speck.ks11.keypos = 11                          ' �L�[�|�W�V����
    MENU_Speck.ks11.keyleng = 3                          ' �L�[��
    MENU_Speck.ks11.keyflag = BtKfExt                    ' �L�[�t���O
    MENU_Speck.ks11.keytype = Chr(BtKtString)            ' �L�[�^�C�v
    MENU_Speck.ks11.reserve = &H0                        ' �\��ς�
'-------------------------------------------------------

    sts = BTRV(BtOpCreate, MENU_POS, MENU_Speck, Len(MENU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���j���[�Ǘ��}�X�^")
        MENU_Create = True
    End If

End Function

Function MENU_Open(Mode As Integer) As Integer
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
    
    MENU_Open = False
                                            '���j���[�Ǘ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", MENU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        MENU_Open = True
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, MENU_POS, MENUREC, Len(MENUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = MENU_Create()         '���j���[�Ǘ��}�X�^�쐬
                If sts <> False Then
                    MENU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, MENU_POS, MENUREC, Len(MENUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���j���[�Ǘ��}�X�^")
                    MENU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "���j���[�Ǘ��}�X�^")
                MENU_Open = True
                Exit Function
        End Select
    Loop
End Function
