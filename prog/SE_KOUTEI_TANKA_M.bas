Attribute VB_Name = "SE_KOUTEI_TANKA_M"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �i�ڕʍ�ƍH���P���ݒ�}�X�^  �t�@�C����`          *
'*                                                                  *
'*          CREATE 2008.02.05                                       *
'********************************************************************
'�t�@�C���h�c
Public Const SE_KOUTEI_TANKA_M_ID$ = "SE_KOUTEI_TANKA_M"

'�y�[�W�T�C�Y
Public Const SE_KOUTEI_TANKA_M_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public SE_KOUTEI_TANKA_M_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************

'�O�H���̍\����
Private Type MAE_KOUTEI_tag
    KOUSU(0 To 6)               As Byte     '�H�� 9(3)V999
    SYUKEI_KBN(0 To 0)          As Byte     '�W�v�敪
    SEIKYU_SAKI(0 To 0)         As Byte     '������
End Type

'��ƍH���̍\����
Private Type SAGYO_KOUTEI_tag
    KOUTEI_NAME(0 To 39)        As Byte     '�H����
    KOUSU(0 To 6)               As Byte     '�H�� 9(3)V999
    SYUKEI_KBN(0 To 0)          As Byte     '�W�v�敪
    SEIKYU_SAKI(0 To 0)         As Byte     '������
End Type

'��H���̍\����
Private Type ATO_KOUTEI_tag
    KOUSU(0 To 6)               As Byte     '�H�� 9(3)V999
    SYUKEI_KBN(0 To 0)          As Byte     '�W�v�敪
    SEIKYU_SAKI(0 To 0)         As Byte     '������
End Type




'���R�[�h��`
Type SE_KOUTEI_TANKA_M_REC_Tag
    
    SE_HIN_GAI(0 To 19)         As Byte     '�i�ڃR�[�h
    
                                            '�O�H��
    SE_MAE_KOUTEI(0 To 9)       As MAE_KOUTEI_tag
                                            '��ƍH��
    SE_SAGYO_KOUTEI(0 To 19)    As SAGYO_KOUTEI_tag
                                            '��H��
    SE_ATO_KOUTEI(0 To 9)       As ATO_KOUTEI_tag
    
    FILLER(0 To 288)            As Byte
    
End Type
'�f�[�^�E�o�b�t�@
Public SE_KOUTEI_TANKA_M_REC    As SE_KOUTEI_TANKA_M_REC_Tag

'�L�[��`

Type KEY0_SE_KOUTEI_TANKA_M                 '�j�d�x�O
    SE_HIN_GAI(0 To 19)         As Byte     '�i�ڃR�[�h
End Type
    
'�L�[�E�f�[�^
Public K0_SE_KOUTEI_TANKA_M     As KEY0_SE_KOUTEI_TANKA_M

Type SE_KOUTEI_TANKA_M_FSpeck
    fs                  As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                 As BtKeySpeck   ' �� ��߯��\����
End Type

Private SE_KOUTEI_TANKA_M_Speck As SE_KOUTEI_TANKA_M_FSpeck
Private Function SE_KOUTEI_TANKA_M_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �i�ڕʍ�ƍH���P���ݒ�}�X�^  �b�q�d�`�s�d          *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SE_KOUTEI_TANKA_M_Create = True
                                            '�i�ڕʍ�ƍH���P���ݒ�}�X�^   �t���p�X�捞��
    sts = GetIni("FILE", SE_KOUTEI_TANKA_M_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_KOUTEI_TANKA_M]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    SE_KOUTEI_TANKA_M_Speck.fs.recoleng = Len(SE_KOUTEI_TANKA_M_REC)    ' ���R�[�h��
    SE_KOUTEI_TANKA_M_Speck.fs.PageSize = SE_KOUTEI_TANKA_M_PG_SIZ      ' �y�[�W�T�C�Y
    SE_KOUTEI_TANKA_M_Speck.fs.idexnumb = 1                             ' �C���f�b�N�X��
    SE_KOUTEI_TANKA_M_Speck.fs.fileflag = 0                             ' �t�@�C���t���O
    SE_KOUTEI_TANKA_M_Speck.fs.reserve = &H0                            ' �\��ς�
    
    
    '-------------------------------------------'   �j�d�x�O
    SE_KOUTEI_TANKA_M_Speck.ks0.keypos = 1                  ' �L�[�|�W�V����
    SE_KOUTEI_TANKA_M_Speck.ks0.keyleng = 20                ' �L�[��
    SE_KOUTEI_TANKA_M_Speck.ks0.keyflag = BtKfExt           ' �L�[�t���O
    SE_KOUTEI_TANKA_M_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    SE_KOUTEI_TANKA_M_Speck.ks0.reserve = &H0               ' �\��ς�
    '-------------------------------------------'   �j�d�x�O

    sts = BTRV(BtOpCreate, SE_KOUTEI_TANKA_M_POS, SE_KOUTEI_TANKA_M_Speck, Len(SE_KOUTEI_TANKA_M_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�i�ڕʍ�ƍH���P���ݒ�}�X�^")
        Exit Function
    End If
    
    SE_KOUTEI_TANKA_M_Create = False

End Function

Function SE_KOUTEI_TANKA_M_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �i�ڕʍ�ƍH���P���ݒ�}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    SE_KOUTEI_TANKA_M_Open = True
                                                '�i�ڕʍ�ƍH���P���ݒ�}�X�^   �t���p�X�捞��
    sts = GetIni("FILE", SE_KOUTEI_TANKA_M_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_KOUTEI_TANKA_M]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, SE_KOUTEI_TANKA_M_POS, SE_KOUTEI_TANKA_M_REC, Len(SE_KOUTEI_TANKA_M_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SE_KOUTEI_TANKA_M_Create()    '�i�ڕʍ�ƍH���P���ݒ�}�X�^ �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SE_KOUTEI_TANKA_M_POS, SE_KOUTEI_TANKA_M_REC, Len(SE_KOUTEI_TANKA_M_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�i�ڕʍ�ƍH���P���ݒ�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�i�ڕʍ�ƍH���P���ݒ�}�X�^")
                Exit Function
        End Select
    Loop
    SE_KOUTEI_TANKA_M_Open = False

End Function
