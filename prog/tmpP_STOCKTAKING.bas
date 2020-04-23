Attribute VB_Name = "tmpP_STOCKTAKING"
Option Explicit

'********************************************************************
'*
'*              ���ޒI�����ް�  �t�@�C����`
'*
'*          CREATE 2006.11.22
'********************************************************************
'�t�@�C���h�c
Public Const tmpP_STOCK_ID$ = "tmpP_STOCK"

'�y�[�W�T�C�Y
Private Const tmpP_STOCK_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public tmpP_STOCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type tmpP_STOCK_REC_Tag
    
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    
    
    CODE(0 To 4)            As Byte         '�d���溰��
    TANKA(0 To 10)          As Byte         '�d���P�� 9(8)V99
    
    INPUT_DATE(0 To 7)      As Byte         '�o�^���t
    
    G_SYUSHI(0 To 2)        As Byte         '���x�P��
    ZEN_ZAIKO_QTY(0 To 7)   As Byte         '�O���݌ɐ���
                            
    NYUKO_QTY(0 To 7)       As Byte         '���ɐ�
    SYUKO_QTY(0 To 7)       As Byte         '�o�ɐ�
    ZAIKO_QTY(0 To 7)       As Byte         '�݌ɐ�
    
    
    LAST_SYUKA_DT(0 To 7)   As Byte         '�ŏI�o�ד�
    LAST_SYUKA_QTY(0 To 7)  As Byte         '�ŏI�o�א���
    
    MOTO_ZAIKO_QTY(0 To 7)  As Byte         '�ďW�v�O
    MAEGARI_QTY(0 To 7)     As Byte         '�O�ؐ�

    
    SYUKA_NON_F(0 To 0)     As Byte         '�o�א��v�Z�L���@0:���Ȃ��@1:����


    ZEN_ZAIKO_KIN(0 To 7)   As Byte         '�O���݌ɋ��z

    FILLER(0 To 5)         As Byte          '

End Type
'�f�[�^�E�o�b�t�@
Public tmpP_STOCK_REC       As tmpP_STOCK_REC_Tag

'�L�[��`
    
Public Type KEY0_tmpP_STOCK                 '�j�d�x�O
    
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    CODE(0 To 4)            As Byte         '�d���溰��
    TANKA(0 To 10)          As Byte         '�d���P�� 9(8)V99
    
End Type
    
Public Type KEY1_tmpP_STOCK                 '�j�d�x�P
    
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    
    INPUT_DATE(0 To 7)      As Byte         '�o�^���t 2006.11.22
    
    
    CODE(0 To 4)            As Byte         '�d���溰��
    TANKA(0 To 10)          As Byte         '�d���P�� 9(8)V99
    
End Type
    
    
Public Type KEY2_tmpP_STOCK                 '�j�d�x�P
    
    
    G_SYUSHI(0 To 2)        As Byte         '���x�P��
    
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    
    INPUT_DATE(0 To 7)      As Byte         '�o�^���t 2006.11.22
    
    
    CODE(0 To 4)            As Byte         '�d���溰��
    TANKA(0 To 10)          As Byte         '�d���P�� 9(8)V99
    
End Type
    
    
    
'�L�[�E�f�[�^
Public K0_tmpP_STOCK        As KEY0_tmpP_STOCK
Public K1_tmpP_STOCK        As KEY1_tmpP_STOCK

Public K2_tmpP_STOCK        As KEY2_tmpP_STOCK


Type tmpP_STOCK_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����
    ks6                     As BtKeySpeck   ' �� ��߯��\����
    ks7                     As BtKeySpeck   ' �� ��߯��\����
    ks8                     As BtKeySpeck   ' �� ��߯��\����
    ks9                     As BtKeySpeck   ' �� ��߯��\����
    ks10                    As BtKeySpeck   ' �� ��߯��\����

    ks11                    As BtKeySpeck   ' �� ��߯��\����
    ks12                    As BtKeySpeck   ' �� ��߯��\����
    ks13                    As BtKeySpeck   ' �� ��߯��\����
    ks14                    As BtKeySpeck   ' �� ��߯��\����
    ks15                    As BtKeySpeck   ' �� ��߯��\����
    ks16                    As BtKeySpeck   ' �� ��߯��\����
    ks17                    As BtKeySpeck   ' �� ��߯��\����

End Type

Private tmpP_STOCK_Speck    As tmpP_STOCK_FSpeck
Private Function tmpP_STOCK_Create() As Integer
'********************************************************************
'*
'*              ���ޒI�����ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128





    tmpP_STOCK_Create = True
                                            '���ޒI�����ް��t���p�X�捞��
    sts = GetIni("FILE", tmpP_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpP_STOCK]�ǂݍ��݃G���[")
        Exit Function
    End If



    FullPath = Trim(c)
    tmpP_STOCK_Speck.fs.recoleng = Len(tmpP_STOCK_REC)  ' ���R�[�h��
    tmpP_STOCK_Speck.fs.PageSize = tmpP_STOCK_PG_SIZ    ' �y�[�W�T�C�Y
    tmpP_STOCK_Speck.fs.idexnumb = 3                    ' �C���f�b�N�X��
    tmpP_STOCK_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    tmpP_STOCK_Speck.fs.reserve = &H0                   ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    tmpP_STOCK_Speck.ks0.keypos = 1                     ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks0.keyleng = 1                    ' �L�[��
    tmpP_STOCK_Speck.ks0.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    tmpP_STOCK_Speck.ks0.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks0.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks1.keypos = 2                     ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks1.keyleng = 1                    ' �L�[��
    tmpP_STOCK_Speck.ks1.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    tmpP_STOCK_Speck.ks1.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks1.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks2.keypos = 3                     ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks2.keyleng = 20                   ' �L�[��
    tmpP_STOCK_Speck.ks2.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    tmpP_STOCK_Speck.ks2.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks2.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks3.keypos = 23                    ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks3.keyleng = 5                    ' �L�[��
    tmpP_STOCK_Speck.ks3.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    tmpP_STOCK_Speck.ks3.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks3.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks4.keypos = 28                    ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks4.keyleng = 11                   ' �L�[��
    tmpP_STOCK_Speck.ks4.keyflag = BtKfExt              ' �L�[�t���O
    tmpP_STOCK_Speck.ks4.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks4.reserve = &H0                  ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    tmpP_STOCK_Speck.ks5.keypos = 1                     ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks5.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_STOCK_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    tmpP_STOCK_Speck.ks5.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks5.reserve = &H0                  ' �\��ς�
    
    
    tmpP_STOCK_Speck.ks6.keypos = 2                     ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks6.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_STOCK_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    tmpP_STOCK_Speck.ks6.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks6.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks7.keypos = 3                     ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks7.keyleng = 20                   ' �L�[��
                                                        ' �L�[�t���O
    tmpP_STOCK_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    tmpP_STOCK_Speck.ks7.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks7.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks8.keypos = 39                    ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks8.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_STOCK_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfChg
    tmpP_STOCK_Speck.ks8.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks8.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks9.keypos = 23                    ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks9.keyleng = 5                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_STOCK_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfChg
    tmpP_STOCK_Speck.ks9.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks9.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks10.keypos = 28                   ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks10.keyleng = 11                  ' �L�[��
    tmpP_STOCK_Speck.ks10.keyflag = BtKfExt + BtKfChg   ' �L�[�t���O
    tmpP_STOCK_Speck.ks10.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks10.reserve = &H0                 ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�Q ��
    
    
    tmpP_STOCK_Speck.ks11.keypos = 47                    ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks11.keyleng = 3                    ' �L�[��
    tmpP_STOCK_Speck.ks11.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    tmpP_STOCK_Speck.ks11.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks11.reserve = &H0                  ' �\��ς�
    
    
    
    tmpP_STOCK_Speck.ks12.keypos = 1                     ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks12.keyleng = 1                    ' �L�[��
    tmpP_STOCK_Speck.ks12.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    tmpP_STOCK_Speck.ks12.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks12.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks13.keypos = 2                     ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks13.keyleng = 1                    ' �L�[��
    tmpP_STOCK_Speck.ks13.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    tmpP_STOCK_Speck.ks13.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks13.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks14.keypos = 3                     ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks14.keyleng = 20                   ' �L�[��
    tmpP_STOCK_Speck.ks14.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    tmpP_STOCK_Speck.ks14.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks14.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks15.keypos = 39                    ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks15.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_STOCK_Speck.ks15.keyflag = BtKfExt + BtKfSeg
    tmpP_STOCK_Speck.ks15.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks15.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks16.keypos = 23                    ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks16.keyleng = 5                    ' �L�[��
    tmpP_STOCK_Speck.ks16.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    tmpP_STOCK_Speck.ks16.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks16.reserve = &H0                  ' �\��ς�
    
    tmpP_STOCK_Speck.ks17.keypos = 28                    ' �L�[�|�W�V����
    tmpP_STOCK_Speck.ks17.keyleng = 11                   ' �L�[��
    tmpP_STOCK_Speck.ks17.keyflag = BtKfExt              ' �L�[�t���O
    tmpP_STOCK_Speck.ks17.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_STOCK_Speck.ks17.reserve = &H0                  ' �\��ς�
    '--------------------------------------------------- �L�[�Q ��
    
    
    sts = BTRV(BtOpCreate, tmpP_STOCK_POS, tmpP_STOCK_Speck, Len(tmpP_STOCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "tmp���ޒI�����ް�")
        Exit Function
    End If
    
    tmpP_STOCK_Create = False

End Function

Public Function tmpP_STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޒI�����ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String



    tmpP_STOCK_Open = True
                                            '���ޒI���f�[�^�t���p�X�捞��
    sts = GetIni("FILE", tmpP_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpP_STOCK]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = Trim(c)

    Do
        sts = BTRV(BtOpOpen, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpP_STOCK_Create()   '���ޒI�����ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "tmp���ޒI�����ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "tmp���ޒI�����ް�")
                Exit Function
        End Select
    Loop
    
    tmpP_STOCK_Open = False

End Function

