Attribute VB_Name = "TANTO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �S���҃}�X�^  �t�@�C����`                          *
'*                                                                  *
'*          CREATE 2004.02.14                                       *
'********************************************************************
'�t�@�C���h�c
Public Const TANTO_ID$ = "TANTO"

'�y�[�W�T�C�Y
Public Const TANTO_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public TANTO_POS            As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type TANTOREC_Tag
    TANTO_CODE(0 To 4)      As Byte         '�S���҃R�[�h
    TANTO_NAME(0 To 19)     As Byte         '�S���Җ���
    POST_CODE(0 To 1)       As Byte         '����
    KUBUN(0 To 1)           As Byte         '�敪 �󔒁F�ΏۊO 2011.09.30
    FILLER(0 To 18)         As Byte         'FILLER 20-->19-->18 2011.09.30
End Type

'�f�[�^�E�o�b�t�@
Public TANTOREC As TANTOREC_Tag

'�L�[��`

Type KEY0_TANTO                 '�j�d�x�O
    TANTO_CODE(0 To 4)      As Byte         '�S���҃R�[�h
End Type

'�L�[�E�f�[�^
Public K0_TANTO             As KEY0_TANTO

Type TANTO_FSpeck
    fs  As BtFileSpeck          ' ̧�� ��߯��\����
    ks0 As BtKeySpeck           ' �� ��߯��\����
End Type

Public TANTO_Speck As TANTO_FSpeck
 
Private Function TANTO_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �S���҃}�X�^  �b�q�d�`�s�d                      �@  *
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

    TANTO_Create = True
                                            '�S���҃}�X�^�t���p�X�捞��
    sts = GetIni("FILE", TANTO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    TANTO_Speck.fs.recoleng = Len(TANTOREC)             ' ���R�[�h��
    TANTO_Speck.fs.PageSize = TANTO_PG_SIZ%             ' �y�[�W�T�C�Y
    TANTO_Speck.fs.idexnumb = 1                         ' �C���f�b�N�X��
    TANTO_Speck.fs.fileflag = 0                         ' �t�@�C���t���O
    TANTO_Speck.fs.reserve = &H0                        ' �\��ς�
                                                        ' �L�[�O
    TANTO_Speck.ks0.keypos = 1                          ' �L�[�|�W�V����
    TANTO_Speck.ks0.keyleng = 5                         ' �L�[��
    TANTO_Speck.ks0.keyflag = BtKfExt                   ' �L�[�t���O
    TANTO_Speck.ks0.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    TANTO_Speck.ks0.reserve = &H0                       ' �\��ς�

    sts = BTRV(BtOpCreate, TANTO_POS, TANTO_Speck, Len(TANTO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�S���҃}�X�^")
    End If

    TANTO_Create = False

End Function

Function TANTO_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �S���҃}�X�^  �n�o�d�m                          �@  *
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
    
    TANTO_Open = True
                                            '�S���҃}�X�^�t���p�X�捞��
    sts = GetIni("FILE", TANTO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, TANTO_POS, TANTOREC, Len(TANTOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = TANTO_Create()        '�S���҃}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, TANTO_POS, TANTOREC, Len(TANTOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�S���҃}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�S���҃}�X�^")
                Exit Function
        End Select
    Loop

    TANTO_Open = False
    
End Function
