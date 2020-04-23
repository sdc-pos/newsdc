Attribute VB_Name = "ITEM_LOC"
Option Explicit
'********************************************************************
'*
'*              �i�ځ|�I�}�X�^  �t�@�C����`
'*
'*          CREATE 2012.06.01
'********************************************************************
'�t�@�C���h�c
Public Const ITEM_LOC_ID$ = "ITEM_LOC"

'�y�[�W�T�C�Y
Public Const ITEM_LOC_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public ITEM_LOC_POS       As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type ITEM_LOCREC_Tag
    No(0 To 7)                  As Byte     'No
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
    
    IRI_QTY(0 To 7)             As Byte     '������萔

    BIKOU(0 To 19)              As Byte     '������l

    SOKO(0 To 1)                As Byte     '�q��
    Retu(0 To 1)                As Byte     '��
    Ren(0 To 1)                 As Byte     '�A
    Dan(0 To 1)                 As Byte     '�i
    
    Print_SU(0 To 7)            As Byte     '�������

    FILLER(0 To 53)             As Byte
        
End Type
'�f�[�^�E�o�b�t�@
Public ITEM_LOCREC              As ITEM_LOCREC_Tag

'�L�[��`

Type KEY0_ITEM_LOC                          '�j�d�x�O
    No(0 To 7)                  As Byte     'No
End Type


Type KEY1_ITEM_LOC                          '�j�d�x�P
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
End Type


Type KEY2_ITEM_LOC                          '�j�d�x�Q
    SOKO(0 To 1)                As Byte     '�q��
    Retu(0 To 1)                As Byte     '��
    Ren(0 To 1)                 As Byte     '�A
    Dan(0 To 1)                 As Byte     '�i
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
End Type




'�L�[�E�f�[�^
Public K0_ITEM_LOC              As KEY0_ITEM_LOC
Public K1_ITEM_LOC              As KEY1_ITEM_LOC
Public K2_ITEM_LOC              As KEY2_ITEM_LOC

Type ITEM_LOC_FSpeck
    fs      As BtFileSpeck                 ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                 ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck
    ks8     As BtKeySpeck
    ks9     As BtKeySpeck
    ks10    As BtKeySpeck
End Type

Private ITEM_LOC_Speck  As ITEM_LOC_FSpeck

Private Function ITEM_LOC_Create() As Integer
'********************************************************************
'*
'*              �i�ځ|�I�}�X�^  �t�@�C���쐬
'*
'*          CREATE 2012.06.01
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ITEM_LOC_Create = True
                                            '�i�ځ|�I�}�X�^�t���p�X�捞��
    sts = GetIni(App.EXEName, ITEM_LOC_ID, App.EXEName, c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, App.EXEName & " " & ITEM_LOC_ID & "�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    ITEM_LOC_Speck.fs.recoleng = Len(ITEM_LOCREC)       ' ���R�[�h��
    ITEM_LOC_Speck.fs.PageSize = ITEM_LOC_PG_SIZ        ' �y�[�W�T�C�Y
    ITEM_LOC_Speck.fs.idexnumb = 3                      ' �C���f�b�N�X��
    ITEM_LOC_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    ITEM_LOC_Speck.fs.reserve = &H0                     ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    ITEM_LOC_Speck.ks0.keypos = 1                       ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks0.keyleng = 8                      ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfDup
    ITEM_LOC_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks0.reserve = &H0                    ' �\��ς�
'-----------------------------------------------


'-----------------------------------------------
                                                ' �L�[�P
    ITEM_LOC_Speck.ks1.keypos = 9                       ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks1.keyleng = 1                      ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks1.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks1.reserve = &H0                    ' �\��ς�

    ITEM_LOC_Speck.ks2.keypos = 10                      ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks2.keyleng = 1                      ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks2.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks2.reserve = &H0                    ' �\��ς�

    ITEM_LOC_Speck.ks3.keypos = 11                      ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks3.keyleng = 20                     ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfDup
    ITEM_LOC_Speck.ks3.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks3.reserve = &H0                    ' �\��ς�
'-----------------------------------------------

'-----------------------------------------------
                                                ' �L�[�Q
    ITEM_LOC_Speck.ks4.keypos = 59                      ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks4.keyleng = 2                      ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks4.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks4.reserve = &H0                    ' �\��ς�
    
    ITEM_LOC_Speck.ks5.keypos = 61                      ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks5.keyleng = 2                      ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks5.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks5.reserve = &H0                    ' �\��ς�
    
    ITEM_LOC_Speck.ks6.keypos = 63                      ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks6.keyleng = 2                      ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks6.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks6.reserve = &H0                    ' �\��ς�
    
    ITEM_LOC_Speck.ks7.keypos = 65                      ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks7.keyleng = 2                      ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks7.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks7.reserve = &H0                    ' �\��ς�
    
    
    
    ITEM_LOC_Speck.ks8.keypos = 9                       ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks8.keyleng = 1                      ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks8.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks8.reserve = &H0                    ' �\��ς�

    ITEM_LOC_Speck.ks9.keypos = 10                      ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks9.keyleng = 1                      ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks9.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks9.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks9.reserve = &H0                    ' �\��ς�

    ITEM_LOC_Speck.ks10.keypos = 11                      ' �L�[�|�W�V����
    ITEM_LOC_Speck.ks10.keyleng = 20                     ' �L�[��
                                                        ' �L�[�t���O
    ITEM_LOC_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup
    ITEM_LOC_Speck.ks10.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    ITEM_LOC_Speck.ks10.reserve = &H0                    ' �\��ς�
'-----------------------------------------------



    sts = BTRV(BtOpCreate, ITEM_LOC_POS, ITEM_LOC_Speck, Len(ITEM_LOC_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���Y�}�X�^")
        Exit Function
    End If

    ITEM_LOC_Create = False

End Function

Public Function ITEM_LOC_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �i�ځ|�I�}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_LOC_Open = True
                                            '���Y�}�X�^�t���p�X�捞��
    sts = GetIni(App.EXEName, ITEM_LOC_ID, App.EXEName, c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, App.EXEName & " " & ITEM_LOC_ID & "�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_LOC_Create()        '���Y�}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���Y�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���Y�}�X�^")
                Exit Function
        End Select
    Loop

    ITEM_LOC_Open = False

End Function

