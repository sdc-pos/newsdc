Attribute VB_Name = "HTIdDelv"
Option Explicit
'********************************************************************
'*
'*              Id������ް��@�t�@�C����`
'*              Create 2016.10.14
'********************************************************************
'�t�@�C���h�c
Public Const HTIdDelv_ID$ = "HTIdDelv"

'�y�[�W�T�C�Y
Public Const HTIdDelv_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public HTIdDelv_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type HTIdDelvREC_Tag
    IDNO(0 To 11)       As Byte         '�`�[ID
    DelvNo(0 To 19)     As Byte         '�����
    EntID(0 To 11)      As Byte         '�o�^ID
    EntTm(0 To 13)      As Byte         '�o�^����
    UpdID(0 To 11)      As Byte         '�X�VID
    UpdTm(0 To 13)      As Byte         '�X�V����
End Type

'�f�[�^�E�o�b�t�@
Public HTIdDelvREC      As HTIdDelvREC_Tag

'�L�[��`
Type KEY0_HTIdDelv          '�j�d�x�O
    IDNO(0 To 11)       As Byte         '�`�[ID
    DelvNo(0 To 19)     As Byte         '�����
End Type

'�L�[�E�f�[�^
Public K0_HTIdDelv    As KEY0_HTIdDelv

Type HTIdDelv_FSpeck
    fs                  As BtFileSpeck  '̧�� ��߯��\����
    ks0                 As BtKeySpeck   '�� ��߯��\����
    ks1                 As BtKeySpeck   '�� ��߯��\����
End Type

Private HTIdDelv_Speck  As HTIdDelv_FSpeck

Private Function HTIdDelv_Create() As Integer
'********************************************************************
'*                                                                  *
'*              Id������ް��@�b�q�d�`�s�d                        *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    HTIdDelv_Create = True
                                            'Id������ް�   �t���p�X�捞��
    sts = GetIni("FILE", HTIdDelv_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTIdDelv]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    HTIdDelv_Speck.fs.recoleng = Len(HTIdDelvREC)       ' ���R�[�h��
    HTIdDelv_Speck.fs.PageSize = HTIdDelv_PG_SIZ        ' �y�[�W�T�C�Y
    HTIdDelv_Speck.fs.idexnumb = 1                      ' �C���f�b�N�X��
    HTIdDelv_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    HTIdDelv_Speck.fs.reserve = &H0                     ' �\��ς�
'------------------------------------------------
                                                ' �L�[�O
    HTIdDelv_Speck.ks0.keypos = 1                       ' �L�[�|�W�V����
    HTIdDelv_Speck.ks0.keyleng = 12                     ' �L�[��
                                                        ' �L�[�t���O
    HTIdDelv_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    HTIdDelv_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    HTIdDelv_Speck.ks0.reserve = &H0                    ' �\��ς�

                                                ' �L�[�O
    HTIdDelv_Speck.ks1.keypos = 13                      ' �L�[�|�W�V����
    HTIdDelv_Speck.ks1.keyleng = 20                     ' �L�[��
                                                        ' �L�[�t���O
    HTIdDelv_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfDup
    HTIdDelv_Speck.ks1.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    HTIdDelv_Speck.ks1.reserve = &H0                    ' �\��ς�


'------------------------------------------------

    sts = BTRV(BtOpCreate, HTIdDelv_POS, HTIdDelv_Speck, Len(HTIdDelv_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "Id������ް�")
        Exit Function
    End If
    
    HTIdDelv_Create = False

End Function
Public Function HTIdDelv_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              Id������ް��@�n�o�d�m                            *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    HTIdDelv_Open = True
                                        'Id������ް��t���p�X�捞��
    sts = GetIni("FILE", HTIdDelv_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HTIdDelv]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, HTIdDelv_POS, HTIdDelvREC, Len(HTIdDelvREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = HTIdDelv_Create()        'Id������ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, HTIdDelv_POS, HTIdDelvREC, Len(HTIdDelvREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "Id������ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "Id������ް�")
                Exit Function
        End Select
    Loop

    HTIdDelv_Open = False

End Function


