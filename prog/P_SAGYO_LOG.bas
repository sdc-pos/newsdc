Attribute VB_Name = "P_SAGYO_LOG"
Option Explicit
'********************************************************************
'*
'*              ��Ǝ���۸�  �t�@�C����`
'*
'*          CREATE 2006.01.30
'********************************************************************
'�t�@�C���h�c
Public Const P_SAGYO_LOG_ID$ = "P_SAGYO_LOG"

'�y�[�W�T�C�Y
Public Const P_SAGYO_LOG_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SAGYO_LOG_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type P_SAGYO_LOG_REC_Tag

    JITU_DT(0 To 7)                     As Byte     '���ѓ��t
    JITU_TM(0 To 5)                     As Byte     '���ю���
    TANTO_CODE(0 To 4)                  As Byte     '�S���҃R�[�h
    WEL_ID(0 To 2)                      As Byte     '�Ώے[����
    JGYOBU(0 To 0)                      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)                      As Byte     '�����O
    MENU_NO(0 To 1)                     As Byte     '���j���[�O���[�v��
    RIRK_ID(0 To 1)                     As Byte     '�������
'    ID_NO(0 To 7)                       As Byte     'ID-NO
    ID_NO(0 To 11)                      As Byte     'ID-NO (8����12��)      2006/05/24
    HIN_GAI(0 To 19)                    As Byte     '�i�ԁi�O���j
    SUMI_JITU_QTY(0 To 7)               As Byte     '���ѐ���(���i���ς�)
    MI_JITU_QTY(0 To 7)                 As Byte     '���ѐ���(�����i)
    MUKE_CODE(0 To 7)                   As Byte     '���Ӑ�R�[�h
    SS_CODE(0 To 7)                     As Byte     '������R�[�h
    FROM_SOKO(0 To 1)                   As Byte     'From �q�ɇ�
    FROM_RETU(0 To 1)                   As Byte     '   �@��
    FROM_REN(0 To 1)                    As Byte     '   �@�A
    FROM_DAN(0 To 1)                    As Byte     '   �@�i
    TO_SOKO(0 To 1)                     As Byte     '�s�n �q�ɇ�
    TO_RETU(0 To 1)                     As Byte     '   �@��
    TO_REN(0 To 1)                      As Byte     '   �@�A
    TO_DAN(0 To 1)                      As Byte     '   �@�i
    PRG_ID(0 To 9)                      As Byte     '�o�͌��v���O����
    WORK_TM(0 To 5)                     As Byte     '��Ǝ���(�b)�E�E�E�ǉ��F2008.08.19
    
    
    SHIJI_No(0 To 7)                    As Byte     '�w�}�[��   ���g�p�Ƃ��� 2010.09.03
    
    
    HIN_CHECK_LABEL_CNT(0 To 2)         As Byte     '�i���������ٌ���   2010.09.03
    HIN_CHECK_GENPIN_CNT(0 To 2)        As Byte     '�i���������i�[����   2010.09.03
    
    JAN_CODE(0 To 19)                   As Byte     'JAN����            2011.08.18
    
    MEMO(0 To 19)                       As Byte     '����               2014.07.01
    
    
    HIN_CHECK_GAISOU_CNT(0 To 2)        As Byte     '�i�������O������   2015.11.07
    
    
    
    
    FILLER(0 To 74)                     As Byte     '                   2011.08.18 117-->97 2014.07.01 97-->77 2015.11.07 77-->74






End Type

'�f�[�^�E�o�b�t�@
Public P_SAGYO_LOG_REC      As P_SAGYO_LOG_REC_Tag

'�L�[��`

Type KEY0_P_SAGYO_LOG           '�j�d�x�O
    JITU_DT(0 To 7)                     As Byte     '���ѓ��t
    JITU_TM(0 To 5)                     As Byte     '���ю���
End Type

Type KEY1_P_SAGYO_LOG           '�j�d�x�P
    TANTO_CODE(0 To 4)                  As Byte     '�S���҃R�[�h
    JITU_DT(0 To 7)                     As Byte     '���ѓ��t
    JITU_TM(0 To 5)                     As Byte     '���ю���
End Type

Type KEY2_P_SAGYO_LOG           '�j�d�x�Q
    TANTO_CODE(0 To 4)                  As Byte     '�S���҃R�[�h
    MENU_NO(0 To 1)                     As Byte     '���j���[�O���[�v��
    JITU_DT(0 To 7)                     As Byte     '���ѓ��t
    JITU_TM(0 To 5)                     As Byte     '���ю���
End Type



'�L�[�E�f�[�^
Public K0_P_SAGYO_LOG       As KEY0_P_SAGYO_LOG
Public K1_P_SAGYO_LOG       As KEY1_P_SAGYO_LOG
Public K2_P_SAGYO_LOG       As KEY2_P_SAGYO_LOG

Type P_SAGYO_LOG_FSpeck
    fs  As BtFileSpeck                      '̧�� ��߯��\����
    ks0 As BtKeySpeck                       '�� ��߯��\����
    ks1 As BtKeySpeck
    ks2 As BtKeySpeck
    ks3 As BtKeySpeck
    ks4 As BtKeySpeck
    ks5 As BtKeySpeck
    ks6 As BtKeySpeck
    ks7 As BtKeySpeck
    ks8 As BtKeySpeck
End Type

Private P_SAGYO_LOG_Speck   As P_SAGYO_LOG_FSpeck
Private Function P_SAGYO_LOG_Create() As Integer
'********************************************************************
'*
'*              ��Ǝ���۸�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2006.01.30
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    P_SAGYO_LOG_Create = True
                                            '��Ǝ���۸ރt���p�X�捞��
    sts = GetIni("FILE", P_SAGYO_LOG_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SAGYO_LOG]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    P_SAGYO_LOG_Speck.fs.recoleng = Len(P_SAGYO_LOG_REC)    ' ���R�[�h��
    P_SAGYO_LOG_Speck.fs.PageSize = P_SAGYO_LOG_PG_SIZ      ' �y�[�W�T�C�Y
    P_SAGYO_LOG_Speck.fs.idexnumb = 3                       ' �C���f�b�N�X��
    P_SAGYO_LOG_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    P_SAGYO_LOG_Speck.fs.reserve = &H0                      ' �\��ς�
'------------------------------------------------
                                                            ' �L�[�O
    P_SAGYO_LOG_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    P_SAGYO_LOG_Speck.ks0.keyleng = 8                       ' �L�[��
                                                            ' �L�[�t���O
    P_SAGYO_LOG_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SAGYO_LOG_Speck.ks0.reserve = &H0                     ' �\��ς�

    P_SAGYO_LOG_Speck.ks1.keypos = 9                        ' �L�[�|�W�V����
    P_SAGYO_LOG_Speck.ks1.keyleng = 6                       ' �L�[��
                                                            ' �L�[�t���O
    P_SAGYO_LOG_Speck.ks1.keyflag = BtKfExt + BtKfDup
    
    P_SAGYO_LOG_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SAGYO_LOG_Speck.ks1.reserve = &H0                     ' �\��ς�
'------------------------------------------------


'------------------------------------------------
                                                            ' �L�[�P
    P_SAGYO_LOG_Speck.ks2.keypos = 15                       ' �L�[�|�W�V����
    P_SAGYO_LOG_Speck.ks2.keyleng = 5                       ' �L�[��
                                                            ' �L�[�t���O
    P_SAGYO_LOG_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SAGYO_LOG_Speck.ks2.reserve = &H0                     ' �\��ς�

    P_SAGYO_LOG_Speck.ks3.keypos = 1                        ' �L�[�|�W�V����
    P_SAGYO_LOG_Speck.ks3.keyleng = 8                       ' �L�[��
                                                            ' �L�[�t���O
    P_SAGYO_LOG_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SAGYO_LOG_Speck.ks3.reserve = &H0                     ' �\��ς�


    P_SAGYO_LOG_Speck.ks4.keypos = 9                        ' �L�[�|�W�V����
    P_SAGYO_LOG_Speck.ks4.keyleng = 6                       ' �L�[��
                                                            ' �L�[�t���O
    P_SAGYO_LOG_Speck.ks4.keyflag = BtKfExt + BtKfDup
    P_SAGYO_LOG_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SAGYO_LOG_Speck.ks4.reserve = &H0                     ' �\��ς�

'------------------------------------------------

'------------------------------------------------
                                                            ' �L�[�Q
    P_SAGYO_LOG_Speck.ks5.keypos = 15                       ' �L�[�|�W�V����
    P_SAGYO_LOG_Speck.ks5.keyleng = 5                       ' �L�[��
                                                            ' �L�[�t���O
    P_SAGYO_LOG_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks5.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SAGYO_LOG_Speck.ks5.reserve = &H0                     ' �\��ς�

    P_SAGYO_LOG_Speck.ks6.keypos = 25                       ' �L�[�|�W�V����
    P_SAGYO_LOG_Speck.ks6.keyleng = 2                       ' �L�[��
                                                            ' �L�[�t���O
    P_SAGYO_LOG_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks6.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SAGYO_LOG_Speck.ks6.reserve = &H0                     ' �\��ς�
                                                            
    P_SAGYO_LOG_Speck.ks7.keypos = 1                        ' �L�[�|�W�V����
    P_SAGYO_LOG_Speck.ks7.keyleng = 8                       ' �L�[��
                                                            ' �L�[�t���O
    P_SAGYO_LOG_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks7.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SAGYO_LOG_Speck.ks7.reserve = &H0                     ' �\��ς�

    P_SAGYO_LOG_Speck.ks8.keypos = 9                        ' �L�[�|�W�V����
    P_SAGYO_LOG_Speck.ks8.keyleng = 6                       ' �L�[��
                                                            ' �L�[�t���O
    P_SAGYO_LOG_Speck.ks8.keyflag = BtKfExt + BtKfDup
    
    P_SAGYO_LOG_Speck.ks8.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SAGYO_LOG_Speck.ks8.reserve = &H0                     ' �\��ς�


'------------------------------------------------




    sts = BTRV(BtOpCreate, P_SAGYO_LOG_POS, P_SAGYO_LOG_Speck, Len(P_SAGYO_LOG_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "��Ǝ���۸�")
        Exit Function
    End If

    P_SAGYO_LOG_Create = False

End Function

Public Function P_SAGYO_LOG_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ��Ǝ���۸�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    P_SAGYO_LOG_Open = True
                                            '��Ǝ���۸ރt���p�X�捞��
    sts = GetIni("FILE", P_SAGYO_LOG_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SAGYO_LOG]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SAGYO_LOG_Create()        '��Ǝ���۸ލ쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "��Ǝ���۸�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "��Ǝ���۸�")
                Exit Function
        End Select
    Loop
    
    P_SAGYO_LOG_Open = False

End Function