Attribute VB_Name = "O_HATUBN"
Option Explicit
'********************************************************************
'*
'*              ���ԃ}�X�^�@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const O_HATUBAN_ID$ = "O_HATUBAN"

'�y�[�W�T�C�Y
Public Const O_HATUBAN_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public O_HATUBAN_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type O_HATUBANREC_Tag
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NYK_KBN(0 To 0)         As Byte         '���ד`�[���敪
    NYK_DEN_NO(0 To 4)      As Byte         '�����ד`�[��
    SYK_KBN(0 To 0)         As Byte         '�o�ד`�[���敪
    SYK_DEN_NO(0 To 4)      As Byte         '���o�ד`�[��
    NYK_ID_KBN(0 To 0)      As Byte         '����ID���敪
    NYK_ID_NO(0 To 7)       As Byte         '������ID��
    SYK_ID_KBN(0 To 0)      As Byte         '�o��ID���敪
    SYK_ID_NO(0 To 6)       As Byte         '���o��ID��
    FILLER(0 To 11)         As Byte         'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public O_HATUBANREC           As O_HATUBANREC_Tag

'�L�[��`
Type KEY0_O_HATUBAN            '�j�d�x�O
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
End Type

'�L�[�E�f�[�^
Public K0_O_HATUBAN           As KEY0_O_HATUBAN

Type O_HATUBAN_FSpeck
    fs      As BtFileSpeck                  '̧�� ��߯��\����
    ks0     As BtKeySpeck                   '�� ��߯��\����
End Type

Private O_HATUBAN_Speck As O_HATUBAN_FSpeck

Private Function O_HATUBAN_Create() As Integer
'********************************************************************
'*
'*              ���ԃ}�X�^�@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    O_HATUBAN_Create = True
                                            '���ԃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", O_HATUBAN_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_HATUBAN]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    O_HATUBAN_Speck.fs.recoleng = Len(O_HATUBANREC)     ' ���R�[�h��
    O_HATUBAN_Speck.fs.PageSize = O_HATUBAN_PG_SIZ      ' �y�[�W�T�C�Y
    O_HATUBAN_Speck.fs.idexnumb = 1                   ' �C���f�b�N�X��
    O_HATUBAN_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    O_HATUBAN_Speck.fs.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�O
    O_HATUBAN_Speck.ks0.keypos = 1                    ' �L�[�|�W�V����
    O_HATUBAN_Speck.ks0.keyleng = 1                   ' �L�[��
    O_HATUBAN_Speck.ks0.keyflag = BtKfExt             ' �L�[�t���O
    O_HATUBAN_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    O_HATUBAN_Speck.ks0.reserve = &H0                 ' �\��ς�

    sts = BTRV(BtOpCreate, O_HATUBAN_POS, O_HATUBAN_Speck, Len(O_HATUBAN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ԃ}�X�^")
        Exit Function
    End If

    O_HATUBAN_Create = False

End Function

Public Function O_HATUBAN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ԃ}�X�^�@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    O_HATUBAN_Open = True
                                            '���ԃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", O_HATUBAN_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_HATUBAN]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, O_HATUBAN_POS, O_HATUBANREC, Len(O_HATUBANREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_HATUBAN_Create()        '���ԃ}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_HATUBAN_POS, O_HATUBANREC, Len(O_HATUBANREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ԃ}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ԃ}�X�^")
                Exit Function
        End Select
    Loop

    O_HATUBAN_Open = False

End Function
