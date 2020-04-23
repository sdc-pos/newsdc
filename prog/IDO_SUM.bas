Attribute VB_Name = "IDO_SUM"
Option Explicit
'********************************************************************
'*
'*              �݌Ɉړ����W�v�t�@�C��  �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const IDO_SUM_ID$ = "IDO_SUM"

'�y�[�W�T�C�Y
Public Const IDO_SUM_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public IDO_SUM_POS    As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type IDO_SUMREC_Tag
    JGYOBU(0 To 0)              As Byte     '���ƕ�
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԊO��
    ZAIKO_QTY(0 To 7)           As Byte     '�݌ɐ�
    LAST_DATE(0 To 7)           As Byte     '�ŏI���ѓ��t
    LAST_TIME(0 To 5)           As Byte     '�ŏI���ю���
    
    J_PLUS_CNT(0 To 7)          As Byte     '�݌�+
    J_MAINA_CNT(0 To 7)         As Byte     '�݌�-
    J_SYUKA_CNT(0 To 7)         As Byte     '�o��
    J_IDO_CNT(0 To 7)           As Byte     '�ړ�
    FILLER(0 To 51)             As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public IDO_SUMREC               As IDO_SUMREC_Tag

'�L�[��`
Type KEY0_IDO_SUM               '�j�d�x�O
    JGYOBU(0 To 0)              As Byte     '���ƕ�
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԊO��
End Type


'�L�[�E�f�[�^
Public K0_IDO_SUM               As KEY0_IDO_SUM

Type IDO_SUM_FSpeck
    fs      As BtFileSpeck                  ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                   ' �� ��߯��\����
    ks1     As BtKeySpeck                   ' �� ��߯��\����
    ks2     As BtKeySpeck                   ' �� ��߯��\����
End Type

Private IDO_SUM_Speck As IDO_SUM_FSpeck

Private Function IDO_SUM_Create() As Integer
'********************************************************************
'*
'*              �݌Ɉړ����W�v�t�@�C��  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    IDO_SUM_Create = True
                                            '�݌Ɉړ����W�v�t�@�C���t���p�X�捞��
    sts = GetIni("FILE", IDO_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [IDO_SUM]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    IDO_SUM_Speck.fs.recoleng = Len(Y_SYUREC)       ' ���R�[�h��
    IDO_SUM_Speck.fs.PageSize = IDO_SUM_PG_SIZ      ' �y�[�W�T�C�Y
    IDO_SUM_Speck.fs.idexnumb = 1                   ' �C���f�b�N�X��
    IDO_SUM_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    IDO_SUM_Speck.fs.reserve = &H0                  ' �\��ς�
'---------------------------------------------------' �L�[�O
    IDO_SUM_Speck.ks0.keypos = 1                    ' �L�[�|�W�V����
    IDO_SUM_Speck.ks0.keyleng = 1                   ' �L�[��
                                                    ' �L�[�t���O
    IDO_SUM_Speck.ks0.keyflag = BtKfExt + BtKfSeg
    IDO_SUM_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_SUM_Speck.ks0.reserve = &H0                 ' �\��ς�
    
    IDO_SUM_Speck.ks1.keypos = 2                    ' �L�[�|�W�V����
    IDO_SUM_Speck.ks1.keyleng = 1                   ' �L�[��
    IDO_SUM_Speck.ks1.keyflag = BtKfExt + BtKfSeg   ' �L�[�t���O
    IDO_SUM_Speck.ks1.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_SUM_Speck.ks1.reserve = &H0                 ' �\��ς�
    
    IDO_SUM_Speck.ks2.keypos = 3                    ' �L�[�|�W�V����
    IDO_SUM_Speck.ks2.keyleng = 20                  ' �L�[��
    IDO_SUM_Speck.ks2.keyflag = BtKfExt             ' �L�[�t���O
    IDO_SUM_Speck.ks2.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    IDO_SUM_Speck.ks2.reserve = &H0                 ' �\��ς�

'---------------------------------------------------' �L�[�O
    
    sts = BTRV(BtOpCreate, IDO_SUM_POS, IDO_SUM_Speck, Len(IDO_SUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�݌Ɉړ����W�v�t�@�C��")
        Exit Function
    End If

    IDO_SUM_Create = False

End Function

Function IDO_SUM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �݌Ɉړ����W�v�t�@�C��  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    IDO_SUM_Open = True
                                            '�݌Ɉړ����W�v�t�@�C���t���p�X�捞��
    sts = GetIni("FILE", IDO_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [IDO_SUM]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, IDO_SUM_POS, IDO_SUMREC, Len(IDO_SUMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = IDO_SUM_Create()      '�݌Ɉړ����W�v�t�@�C���쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�݌Ɉړ����W�v�t�@�C��")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌Ɉړ����W�v�t�@�C��")
                Exit Function
        End Select
    Loop
    Y_SYU_Open = False
End Function
