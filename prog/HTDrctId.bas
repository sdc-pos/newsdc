Attribute VB_Name = "HTDrctId"
Option Explicit
'********************************************************************
'*
'*              ������ID�ް��@�t�@�C����`
'*              Create 2016.10.14
'********************************************************************
'�t�@�C���h�c
Public Const HTDrctId_ID$ = "HTDrctId"

'�y�[�W�T�C�Y
Public Const HTDrctId_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public HTDrctId_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type HTDrctIdREC_Tag
    IDNO(0 To 11)       As Byte         '�`�[ID
    ChoCode(0 To 8)     As Byte         '�����溰��
    
    '>>>>>>>>>> 2016.10.27 �ǉ�
    ChoName(0 To 39)            As Byte '�����於
    ChoZip(0 To 6)              As Byte '������X�֔ԍ�
    ChoTel(0 To 15)             As Byte '������d�b�ԍ�
    ChoAddress(0 To 79)         As Byte '������Z��
    ChoMemo(0 To 39)            As Byte '�����惁��
    TMark(0 To 0)               As Byte '�s�}�[�N�敪
    '>>>>>>>>>> 2016.10.27 �ǉ�
    
    EntID(0 To 11)      As Byte         '�o�^ID
    EntTm(0 To 13)      As Byte         '�o�^����
    UpdID(0 To 11)      As Byte         '�X�VID
    UpdTm(0 To 13)      As Byte         '�X�V����
End Type

'�f�[�^�E�o�b�t�@
Public HTDrctIdREC      As HTDrctIdREC_Tag

'�L�[��`
Type KEY0_HTDrctId          '�j�d�x�O
    IDNO(0 To 19)       As Byte         '�`�[ID
End Type

'�L�[�E�f�[�^
Public K0_HTDrctId    As KEY0_HTDrctId

Type HTDrctId_FSpeck
    fs                  As BtFileSpeck  '̧�� ��߯��\����
    ks0                 As BtKeySpeck   '�� ��߯��\����
End Type

Private HTDrctId_Speck  As HTDrctId_FSpeck

Private Function HTDrctId_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ������ID�ް��@�b�q�d�`�s�d                        �@*
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    HTDrctId_Create = True
                                            '������ID�ް�   �t���p�X�捞��
    sts = GetIni("FILE", HTDrctId_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTDrctId]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    HTDrctId_Speck.fs.recoleng = Len(HTDrctIdREC)       ' ���R�[�h��
    HTDrctId_Speck.fs.PageSize = HTDrctId_PG_SIZ        ' �y�[�W�T�C�Y
    HTDrctId_Speck.fs.idexnumb = 1                      ' �C���f�b�N�X��
    HTDrctId_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    HTDrctId_Speck.fs.reserve = &H0                     ' �\��ς�
'------------------------------------------------
                                                ' �L�[�O
    HTDrctId_Speck.ks0.keypos = 1                       ' �L�[�|�W�V����
    HTDrctId_Speck.ks0.keyleng = 12                     ' �L�[��
    HTDrctId_Speck.ks0.keyflag = BtKfExt                ' �L�[�t���O
    HTDrctId_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    HTDrctId_Speck.ks0.reserve = &H0                    ' �\��ς�
'------------------------------------------------

    sts = BTRV(BtOpCreate, HTDrctId_POS, HTDrctId_Speck, Len(HTDrctId_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "������ID�ް�")
        Exit Function
    End If
    
    HTDrctId_Create = False

End Function
Public Function HTDrctId_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ������ID�ް��@�n�o�d�m                            *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    HTDrctId_Open = True
                                        '������ID�ް��t���p�X�捞��
    sts = GetIni("FILE", HTDrctId_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTDrctId]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, HTDrctId_POS, HTDrctIdREC, Len(HTDrctIdREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = HTDrctId_Create()        '������ID�ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, HTDrctId_POS, HTDrctIdREC, Len(HTDrctIdREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "������ID�ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "������ID�ް�")
                Exit Function
        End Select
    Loop

    HTDrctId_Open = False

End Function


