Attribute VB_Name = "HTDelvNo"
Option Explicit
'********************************************************************
'*
'*              ������ް��@�t�@�C����`
'*              Create 2016.10.14
'********************************************************************
'�t�@�C���h�c
Public Const HTDelvNo_ID$ = "HTDelvNo"

'�y�[�W�T�C�Y
Public Const HTDelvNo_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public HTDelvNo_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type HTDelvNoREC_Tag
    CampName(0 To 19)   As Byte         '�^����Ж�(���}�g�^�A)
    DelvNo(0 To 19)     As Byte         '�����
    ChoCode(0 To 8)     As Byte         '�����溰��
    ChoName(0 To 19)    As Byte         '�����於
    EntID(0 To 11)      As Byte         '�o�^ID
    EntTm(0 To 13)      As Byte         '�o�^����
    UpdID(0 To 11)      As Byte         '�X�VID
    UpdTm(0 To 13)      As Byte         '�X�V����
End Type

'�f�[�^�E�o�b�t�@
Public HTDelvNoREC      As HTDelvNoREC_Tag

'�L�[��`
Type KEY0_HTDelvNo          '�j�d�x�O
    DelvNo(0 To 19)     As Byte         '�����
End Type

'�L�[�E�f�[�^
Public K0_HTDelvNo    As KEY0_HTDelvNo

Type HTDelvNo_FSpeck
    fs                  As BtFileSpeck  '̧�� ��߯��\����
    ks0                 As BtKeySpeck   '�� ��߯��\����
End Type

Private HTDelvNo_Speck  As HTDelvNo_FSpeck

Private Function HTDelvNo_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ������ް��@�b�q�d�`�s�d                          *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    HTDelvNo_Create = True
                                            '������ް�   �t���p�X�捞��
    sts = GetIni("FILE", HTDelvNo_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTDelvNo]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    HTDelvNo_Speck.fs.recoleng = Len(HTDelvNoREC)       ' ���R�[�h��
    HTDelvNo_Speck.fs.PageSize = HTDelvNo_PG_SIZ        ' �y�[�W�T�C�Y
    HTDelvNo_Speck.fs.idexnumb = 1                      ' �C���f�b�N�X��
    HTDelvNo_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    HTDelvNo_Speck.fs.reserve = &H0                     ' �\��ς�
'------------------------------------------------
                                                ' �L�[�O
    HTDelvNo_Speck.ks0.keypos = 21                      ' �L�[�|�W�V����
    HTDelvNo_Speck.ks0.keyleng = 20                     ' �L�[��
    HTDelvNo_Speck.ks0.keyflag = BtKfExt                ' �L�[�t���O
    HTDelvNo_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    HTDelvNo_Speck.ks0.reserve = &H0                    ' �\��ς�
'------------------------------------------------

    sts = BTRV(BtOpCreate, HTDelvNo_POS, HTDelvNo_Speck, Len(HTDelvNo_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "������ް�")
        Exit Function
    End If
    
    HTDelvNo_Create = False

End Function
Public Function HTDelvNo_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ������ް��@�n�o�d�m                              *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    HTDelvNo_Open = True
                                        '������ް��t���p�X�捞��
    sts = GetIni("FILE", HTDelvNo_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTDelvNo]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, HTDelvNo_POS, HTDelvNoREC, Len(HTDelvNoREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = HTDelvNo_Create()        '������ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, HTDelvNo_POS, HTDelvNoREC, Len(HTDelvNoREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "������ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "������ް�")
                Exit Function
        End Select
    Loop

    HTDelvNo_Open = False

End Function


