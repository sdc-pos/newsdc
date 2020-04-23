Attribute VB_Name = "KEPPINLOG"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �݌ɏW�v�f�[�^�@�t�@�C����`                        *
'*                                                                  *
'*          CREATE 2004.05.08                                       *
'********************************************************************
'�t�@�C���h�c
Public Const KEPPINLOG_ID$ = "KEPPINLOG"

'�y�[�W�T�C�Y
Public Const KEPPINLOG_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public KEPPINLOG_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                              *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type KEPPINLOGREC_Tag
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    CREATE_DT(0 To 7)       As Byte     '�쐬���t
    FILLER(0 To 17)         As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public KEPPINLOGREC         As KEPPINLOGREC_Tag

'�L�[��`
Private Type KEY0_KEPPINLOG         '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type


'�L�[�E�f�[�^
Public K0_KEPPINLOG As KEY0_KEPPINLOG

Private Type KEPPINLOG_FSpeck
    fs As BtFileSpeck               ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck               ' �� ��߯��\����
    ks2 As BtKeySpeck               ' �� ��߯��\����
End Type

Private KEPPINLOG_Speck As KEPPINLOG_FSpeck

Private Function KEPPINLOG_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���i�h�~�x�����O�@�b�q�d�`�s�d                      *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.05.08                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    KEPPINLOG_Create = True
                                            '���i�h�~�x�����O�t���p�X�捞��
    sts = GetIni("FILE", KEPPINLOG_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[KEPPINLOG] �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    KEPPINLOG_Speck.fs.recoleng = Len(KEPPINLOGREC)         ' ���R�[�h��
    KEPPINLOG_Speck.fs.PageSize = KEPPINLOG_PG_SIZ          ' �y�[�W�T�C�Y
    KEPPINLOG_Speck.fs.idexnumb = 1                         ' �C���f�b�N�X��
    KEPPINLOG_Speck.fs.fileflag = 0                         ' �t�@�C���t���O
    KEPPINLOG_Speck.fs.reserve = &H0                        ' �\��ς�
'-----------------------------------------------' �L�[�O
    KEPPINLOG_Speck.ks0.keypos = 1                          ' �L�[�|�W�V����
    KEPPINLOG_Speck.ks0.keyleng = 1                         ' �L�[��
    KEPPINLOG_Speck.ks0.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    KEPPINLOG_Speck.ks0.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    KEPPINLOG_Speck.ks0.reserve = &H0                       ' �\��ς�

    KEPPINLOG_Speck.ks1.keypos = 2                          ' �L�[�|�W�V����
    KEPPINLOG_Speck.ks1.keyleng = 1                         ' �L�[��
    KEPPINLOG_Speck.ks1.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    KEPPINLOG_Speck.ks1.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    KEPPINLOG_Speck.ks1.reserve = &H0                       ' �\��ς�

    KEPPINLOG_Speck.ks2.keypos = 3                          ' �L�[�|�W�V����
    KEPPINLOG_Speck.ks2.keyleng = 20                        ' �L�[��
    KEPPINLOG_Speck.ks2.keyflag = BtKfExt                   ' �L�[�t���O
    KEPPINLOG_Speck.ks2.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    KEPPINLOG_Speck.ks2.reserve = &H0                       ' �\��ς�

    sts = BTRV(BtOpCreate, KEPPINLOG_POS, KEPPINLOG_Speck, Len(KEPPINLOG_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���i�h�~�x�����O")
        Exit Function
    End If
    
    KEPPINLOG_Create = False

End Function

Function KEPPINLOG_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ���i�h�~�x�����O�@�n�o�d�m                          *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.05.08                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    KEPPINLOG_Open = True
                                            '���i�h�~�x�����O�t���p�X�捞��
    sts = GetIni("FILE", KEPPINLOG_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[KEPPINLOG] �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = KEPPINLOG_Create()    '���i�h�~�x�����O�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���i�h�~�x�����O")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���i�h�~�x�����O")
                Exit Function
        End Select
    Loop

    KEPPINLOG_Open = False

End Function


