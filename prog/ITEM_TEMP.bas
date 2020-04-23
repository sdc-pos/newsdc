Attribute VB_Name = "ITEM_TEMP"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �i�ڈꎞ�t�@�C��  �t�@�C����`                      *
'*                                                                  *
'*          CREATE 2008.02.03                                       *
'********************************************************************
'�t�@�C���h�c
Public Const ITEM_TEMP_ID$ = "ITEM_TEMP"

'�y�[�W�T�C�Y
Public Const ITEM_TEMP_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public ITEM_TEMP_POS        As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type ITEM_TEMP_REC_Tag
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i�ԁi�O���j
    
    KO_HIN_GAI(0 To 16)     As Byte         '������
    CLASS(0 To 3)           As Byte         '�׽

    ST_SOKO(0 To 1)         As Byte         '�W�����ɑq�� �q��
    ST_RETU(0 To 1)         As Byte         '             ��
    ST_REN(0 To 1)          As Byte         '             �A
    ST_DAN(0 To 1)          As Byte         '             �i
    

End Type

'�f�[�^�E�o�b�t�@
Public ITEM_TEMP_REC        As ITEM_TEMP_REC_Tag

'�L�[��`

Type KEY0_ITEM_TEMP             '�j�d�x�O
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i�ԁi�O���j
End Type

Type KEY1_ITEM_TEMP             '�j�d�x�P
    KO_HIN_GAI(0 To 16)     As Byte         '������
    
    ST_SOKO(0 To 1)         As Byte         '�W�����ɑq�� �q��
    ST_RETU(0 To 1)         As Byte         '             ��
    ST_REN(0 To 1)          As Byte         '             �A
    ST_DAN(0 To 1)          As Byte         '             �i
    
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i�ԁi�O���j

End Type

Type KEY2_ITEM_TEMP             '�j�d�x�Q

    CLASS(0 To 3)           As Byte         '�׽
    
    ST_SOKO(0 To 1)         As Byte         '�W�����ɑq�� �q��
    ST_RETU(0 To 1)         As Byte         '             ��
    ST_REN(0 To 1)          As Byte         '             �A
    ST_DAN(0 To 1)          As Byte         '             �i
    
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i�ԁi�O���j

End Type


'�L�[�E�f�[�^
Public K0_ITEM_TEMP         As KEY0_ITEM_TEMP
Public K1_ITEM_TEMP         As KEY1_ITEM_TEMP
Public K2_ITEM_TEMP         As KEY2_ITEM_TEMP

Type ITEM_TEMP_FSpeck
    fs  As BtFileSpeck          ' ̧�� ��߯��\����
    ks0 As BtKeySpeck           ' �� ��߯��\����
    ks1 As BtKeySpeck           ' �� ��߯��\����
    ks2 As BtKeySpeck           ' �� ��߯��\����
    ks3 As BtKeySpeck           ' �� ��߯��\����
    ks4 As BtKeySpeck           ' �� ��߯��\����
    ks5 As BtKeySpeck           ' �� ��߯��\����
    ks6 As BtKeySpeck           ' �� ��߯��\����
End Type

Public ITEM_TEMP_Speck      As ITEM_TEMP_FSpeck
 
Private Function ITEM_TEMP_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �i�ڈꎞ�t�@�C��  �b�q�d�`�s�d                      *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2008.02.3                                       *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ITEM_TEMP_Create = True
                                            '�S���҃}�X�^�t���p�X�捞��
    sts = GetIni("FILE", ITEM_TEMP_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    ITEM_TEMP_Speck.fs.recoleng = Len(ITEM_TEMP_REC)    ' ���R�[�h��
    ITEM_TEMP_Speck.fs.PageSize = ITEM_TEMP_PG_SIZ%     ' �y�[�W�T�C�Y
    ITEM_TEMP_Speck.fs.idexnumb = 3                     ' �C���f�b�N�X��
    ITEM_TEMP_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    ITEM_TEMP_Speck.fs.reserve = &H0                    ' �\��ς�
                                                        
    
    '---------------------------------------------------' �L�[�O
    ITEM_TEMP_Speck.ks0.keypos = 1                          ' �L�[�|�W�V����
    ITEM_TEMP_Speck.ks0.keyleng = 22                        ' �L�[��
    ITEM_TEMP_Speck.ks0.keyflag = BtKfExt                   ' �L�[�t���O
    ITEM_TEMP_Speck.ks0.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ITEM_TEMP_Speck.ks0.reserve = &H0                       ' �\��ς�
    '---------------------------------------------------' �L�[�O

    '---------------------------------------------------' �L�[�P
    ITEM_TEMP_Speck.ks1.keypos = 23                         ' �L�[�|�W�V����
    ITEM_TEMP_Speck.ks1.keyleng = 20                        ' �L�[��
    ITEM_TEMP_Speck.ks1.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    ITEM_TEMP_Speck.ks1.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ITEM_TEMP_Speck.ks1.reserve = &H0                       ' �\��ς�
    
    ITEM_TEMP_Speck.ks2.keypos = 47                         ' �L�[�|�W�V����
    ITEM_TEMP_Speck.ks2.keyleng = 8                         ' �L�[��
    ITEM_TEMP_Speck.ks2.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    ITEM_TEMP_Speck.ks2.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ITEM_TEMP_Speck.ks2.reserve = &H0                       ' �\��ς�
    
    ITEM_TEMP_Speck.ks3.keypos = 47                         ' �L�[�|�W�V����
    ITEM_TEMP_Speck.ks3.keyleng = 8                         ' �L�[��
    ITEM_TEMP_Speck.ks3.keyflag = BtKfExt                   ' �L�[�t���O
    ITEM_TEMP_Speck.ks3.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ITEM_TEMP_Speck.ks3.reserve = &H0                       ' �\��ς�
    '---------------------------------------------------' �L�[�P

    '---------------------------------------------------' �L�[�Q
    ITEM_TEMP_Speck.ks1.keypos = 43                         ' �L�[�|�W�V����
    ITEM_TEMP_Speck.ks1.keyleng = 4                         ' �L�[��
    ITEM_TEMP_Speck.ks1.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    ITEM_TEMP_Speck.ks1.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ITEM_TEMP_Speck.ks1.reserve = &H0                       ' �\��ς�
    
    ITEM_TEMP_Speck.ks2.keypos = 47                         ' �L�[�|�W�V����
    ITEM_TEMP_Speck.ks2.keyleng = 8                         ' �L�[��
    ITEM_TEMP_Speck.ks2.keyflag = BtKfExt + BtKfSeg         ' �L�[�t���O
    ITEM_TEMP_Speck.ks2.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ITEM_TEMP_Speck.ks2.reserve = &H0                       ' �\��ς�
    
    ITEM_TEMP_Speck.ks3.keypos = 1                          ' �L�[�|�W�V����
    ITEM_TEMP_Speck.ks3.keyleng = 20                        ' �L�[��
    ITEM_TEMP_Speck.ks3.keyflag = BtKfExt                   ' �L�[�t���O
    ITEM_TEMP_Speck.ks3.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ITEM_TEMP_Speck.ks3.reserve = &H0                       ' �\��ς�
    '---------------------------------------------------' �L�[�Q





    sts = BTRV(BtOpCreate, ITEM_TEMP_POS, ITEM_TEMP_Speck, Len(ITEM_TEMP_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�i�ڈꎞ�f�[�^")
    End If

    ITEM_TEMP_Create = False

End Function

Function ITEM_TEMP_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �i�ڈꎞ�f�[�^  �n�o�d�m                          �@*
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
    
    ITEM_TEMP_Open = True
                                            '�i�ڈꎞ�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", ITEM_TEMP_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, ITEM_TEMP_POS, ITEM_TEMP_REC, Len(ITEM_TEMP_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_TEMP_Create()    '�i�ڈꎞ�f�[�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_TEMP_POS, ITEM_TEMP_REC, Len(ITEM_TEMP_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�i�ڈꎞ�f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�i�ڈꎞ�f�[�^")
                Exit Function
        End Select
    Loop

    ITEM_TEMP_Open = False
    
End Function
