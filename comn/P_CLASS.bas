Attribute VB_Name = "P_CLASS"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �N���X�}�X�^  �t�@�C����`                          *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�t�@�C���h�c
Public Const P_CLASS_ID$ = "P_CLASS"

'�y�[�W�T�C�Y
Private Const P_CLASS_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public P_CLASS_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type P_CLASSREC_Tag
    
    SHIMUKE_CODE(0 To 1)    As Byte         '�d������
    CLASS_CODE(0 To 19)     As Byte         '�N���X�i�i�ԁj
    CLASS_NAME(0 To 49)     As Byte         '�Ăі�
    TANKA(0 To 10)          As Byte         '���i�����i 9(8)V99
    KOUSU(0 To 6)           As Byte         '�H�� 999V999
    KOURYOU(0 To 10)        As Byte         '�H�� 9(8)V99
    ETC(0 To 10)            As Byte         '���̑�
'''2007.01.11    FILLER(0 To 252)        As Byte         'Filler
    URI_KOURYOU(0 To 10)    As Byte         '�H�� 9(8)V99   2007.01.11
    FILLER(0 To 241)        As Byte         'Filler         2007.01.11
    
    
    
    
    
    
    UPD_TANTO(0 To 4)       As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_CLASSREC           As P_CLASSREC_Tag

'�L�[��`

Type KEY0_P_CLASS                           '�j�d�x�O
    SHIMUKE_CODE(0 To 1)    As Byte         '�d������
    CLASS_CODE(0 To 19)     As Byte         '�N���X�i�i�ԁj
End Type
    
'�L�[�E�f�[�^
Public K0_P_CLASS           As KEY0_P_CLASS

Type P_CLASS_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_CLASS_Speck       As P_CLASS_FSpeck
Private Function P_CLASS_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �N���X�}�X�^  �b�q�d�`�s�d                          *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_CLASS_Create = True
                                            '�N���X�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_CLASS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_CLASS]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_CLASS_Speck.fs.recoleng = Len(P_CLASSREC)         ' ���R�[�h��
    P_CLASS_Speck.fs.PageSize = P_CLASS_PG_SIZ          ' �y�[�W�T�C�Y
    P_CLASS_Speck.fs.idexnumb = 1                       ' �C���f�b�N�X��
    P_CLASS_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    P_CLASS_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_CLASS_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    P_CLASS_Speck.ks0.keyleng = 2                       ' �L�[��
    P_CLASS_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    P_CLASS_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_CLASS_Speck.ks0.reserve = &H0                     ' �\��ς�
    
    P_CLASS_Speck.ks1.keypos = 3                        ' �L�[�|�W�V����
    P_CLASS_Speck.ks1.keyleng = 20                      ' �L�[��
    P_CLASS_Speck.ks1.keyflag = BtKfExt                 ' �L�[�t���O
    P_CLASS_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_CLASS_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    sts = BTRV(BtOpCreate, P_CLASS_POS, P_CLASS_Speck, Len(P_CLASS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�N���X�}�X�^")
        Exit Function
    End If
    
    P_CLASS_Create = False

End Function

Public Function P_Class_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �N���X�}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_Class_Open = True
                                            '�N���X�}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_CLASS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_CLASS]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_CLASS_Create()      '�N���X�}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�N���X�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�N���X�}�X�^")
                Exit Function
        End Select
    Loop
    
    P_Class_Open = False

End Function
