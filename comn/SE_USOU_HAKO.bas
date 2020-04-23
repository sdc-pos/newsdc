Attribute VB_Name = "SE_USOU_HAKO"
Option Explicit
'********************************************************************
'*
'*              �A�����g�p����  �t�@�C����`
'*
'*          CREATE 2008.02.25
'********************************************************************
'�t�@�C���h�c
Public Const SE_USOU_HAKO_ID$ = "SE_USOU_HAKO"

'�y�[�W�T�C�Y
Public Const SE_USOU_HAKO_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public SE_USOU_HAKO_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SE_USOU_HAKOREC_Tag
    SHIMUKE_CODE(0 To 1)        As Byte         '�d������
    JITU_DATE(0 To 7)           As Byte         '���ѓ��t
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�i��
    MTS_CODE(0 To 7)            As Byte         '�o�א�
    CYU_KBN(0 To 0)             As Byte         '�����敪(���g�p)
    CYOK_KBN(0 To 0)            As Byte         '�����敪(���g�p)
    MAISU(0 To 5)               As Byte         '�g�p����
    UPD_TANTO(0 To 4)           As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)       As Byte         '�X�V�@����
    
    
    SE_USOU_F(0 To 1)           As Byte         '�A�����@�o���׸�
    
    
    FILLER(0 To 186)            As Byte         'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public SE_USOU_HAKOREC      As SE_USOU_HAKOREC_Tag

'�L�[��`

Type KEY0_SE_USOU_HAKO      '�j�d�x�O
    JITU_DATE(0 To 7)           As Byte         '���ѓ��t
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�i��
    MTS_CODE(0 To 7)            As Byte         '�o�א�
End Type


Type KEY1_SE_USOU_HAKO      '�j�d�x�P
    SE_USOU_F(0 To 1)           As Byte         '�A�����@�o���׸�
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�i��
End Type


Type KEY2_SE_USOU_HAKO      '�j�d�x�Q
    JITU_DATE(0 To 7)           As Byte         '���ѓ��t
    MTS_CODE(0 To 7)            As Byte         '�o�א�
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�i��
End Type






'�L�[�E�f�[�^
Public K0_SE_USOU_HAKO          As KEY0_SE_USOU_HAKO
Public K1_SE_USOU_HAKO          As KEY1_SE_USOU_HAKO
Public K2_SE_USOU_HAKO          As KEY2_SE_USOU_HAKO

Type SE_USOU_HAKO_FSpeck
    fs      As BtFileSpeck                  '̧�� ��߯��\����
    ks0     As BtKeySpeck                   '�� ��߯��\����
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
    ks11    As BtKeySpeck
    ks12    As BtKeySpeck
    ks13    As BtKeySpeck

End Type

Private SE_USOU_HAKO_Speck As SE_USOU_HAKO_FSpeck
Private Function SE_USOU_HAKO_Create() As Integer
'********************************************************************
'*
'*              �A�����g�p����  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2004.02.19
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SE_USOU_HAKO_Create = True
                                            '�A�����g�p���уt���p�X�捞��
    sts = GetIni("FILE", SE_USOU_HAKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_USOU_HAKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    SE_USOU_HAKO_Speck.fs.recoleng = Len(SE_USOU_HAKOREC)           ' ���R�[�h��
    SE_USOU_HAKO_Speck.fs.PageSize = SE_USOU_HAKO_PG_SIZ            ' �y�[�W�T�C�Y
    SE_USOU_HAKO_Speck.fs.idexnumb = 3                              ' �C���f�b�N�X��
    SE_USOU_HAKO_Speck.fs.fileflag = 0                              ' �t�@�C���t���O
    SE_USOU_HAKO_Speck.fs.reserve = &H0                             ' �\��ς�
'------------------------------------------------
    SE_USOU_HAKO_Speck.ks0.keypos = 3                               ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks0.keyleng = 8                              ' �L�[��
    SE_USOU_HAKO_Speck.ks0.keyflag = BtKfExt + BtKfSeg              ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks0.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks0.reserve = &H0                            ' �\��ς�

    SE_USOU_HAKO_Speck.ks1.keypos = 11                              ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks1.keyleng = 1                              ' �L�[��
    SE_USOU_HAKO_Speck.ks1.keyflag = BtKfExt + BtKfSeg              ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks1.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks1.reserve = &H0                            ' �\��ς�

    SE_USOU_HAKO_Speck.ks2.keypos = 12                              ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks2.keyleng = 1                              ' �L�[��
    SE_USOU_HAKO_Speck.ks2.keyflag = BtKfExt + BtKfSeg              ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks2.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks2.reserve = &H0                            ' �\��ς�

    SE_USOU_HAKO_Speck.ks3.keypos = 13                              ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks3.keyleng = 20                             ' �L�[��
    SE_USOU_HAKO_Speck.ks3.keyflag = BtKfExt + BtKfSeg              ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks3.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks3.reserve = &H0                            ' �\��ς�

    SE_USOU_HAKO_Speck.ks4.keypos = 33                              ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks4.keyleng = 8                              ' �L�[��
    SE_USOU_HAKO_Speck.ks4.keyflag = BtKfExt                        ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks4.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks4.reserve = &H0                            ' �\��ς�
'------------------------------------------------


'------------------------------------------------
    SE_USOU_HAKO_Speck.ks5.keypos = 68                              ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks5.keyleng = 2                              ' �L�[��
    SE_USOU_HAKO_Speck.ks5.keyflag = BtKfExt + _
                                        BtKfSeg + _
                                        BtKfDup + _
                                        BtKfChg                     ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks5.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks5.reserve = &H0                            ' �\��ς�

    SE_USOU_HAKO_Speck.ks6.keypos = 11                              ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks6.keyleng = 1                              ' �L�[��
    SE_USOU_HAKO_Speck.ks6.keyflag = BtKfExt + _
                                        BtKfSeg + _
                                        BtKfDup + _
                                        BtKfChg                     ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks6.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks6.reserve = &H0                            ' �\��ς�

    SE_USOU_HAKO_Speck.ks7.keypos = 12                              ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks7.keyleng = 1                              ' �L�[��
    SE_USOU_HAKO_Speck.ks7.keyflag = BtKfExt + _
                                        BtKfSeg + _
                                        BtKfDup + _
                                        BtKfChg                     ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks7.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks7.reserve = &H0                            ' �\��ς�

    SE_USOU_HAKO_Speck.ks8.keypos = 13                              ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks8.keyleng = 20                             ' �L�[��
    SE_USOU_HAKO_Speck.ks8.keyflag = BtKfExt + _
                                        BtKfDup + _
                                        BtKfChg                     ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks8.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks8.reserve = &H0                            ' �\��ς�

'------------------------------------------------
                                                                ' �L�[�P

                                                                ' �L�[�Q
'------------------------------------------------
    SE_USOU_HAKO_Speck.ks9.keypos = 3                               ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks9.keyleng = 8                              ' �L�[��
    SE_USOU_HAKO_Speck.ks9.keyflag = BtKfExt + BtKfSeg              ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks9.keytype = Chr(BtKtString)                ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks9.reserve = &H0                            ' �\��ς�

    SE_USOU_HAKO_Speck.ks10.keypos = 33                             ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks10.keyleng = 8                             ' �L�[��
    SE_USOU_HAKO_Speck.ks10.keyflag = BtKfExt + BtKfSeg             ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks10.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks10.reserve = &H0                           ' �\��ς�

    SE_USOU_HAKO_Speck.ks11.keypos = 11                             ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks11.keyleng = 1                             ' �L�[��
    SE_USOU_HAKO_Speck.ks11.keyflag = BtKfExt + BtKfSeg             ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks11.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks11.reserve = &H0                           ' �\��ς�

    SE_USOU_HAKO_Speck.ks12.keypos = 12                             ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks12.keyleng = 1                             ' �L�[��
    SE_USOU_HAKO_Speck.ks12.keyflag = BtKfExt + BtKfSeg             ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks12.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks12.reserve = &H0                           ' �\��ς�

    SE_USOU_HAKO_Speck.ks13.keypos = 13                             ' �L�[�|�W�V����
    SE_USOU_HAKO_Speck.ks13.keyleng = 20                            ' �L�[��
    SE_USOU_HAKO_Speck.ks13.keyflag = BtKfExt                       ' �L�[�t���O
    SE_USOU_HAKO_Speck.ks13.keytype = Chr(BtKtString)               ' �L�[�^�C�v
    SE_USOU_HAKO_Speck.ks13.reserve = &H0                           ' �\��ς�
'------------------------------------------------
                                                                ' �L�[�Q
    sts = BTRV(BtOpCreate, SE_USOU_HAKO_POS, SE_USOU_HAKO_Speck, Len(SE_USOU_HAKO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�A��������")
        Exit Function
    End If

    SE_USOU_HAKO_Create = False

End Function

Public Function SE_USOU_HAKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �A�����g�p����  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    SE_USOU_HAKO_Open = True
                                            '�A�����g�p���� �t���p�X�捞��
    sts = GetIni("FILE", SE_USOU_HAKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_USOU_HAKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SE_USOU_HAKO_Create()        '�A�����g�p���э쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�A�����g�p����")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�A�����g�p����")
                Exit Function
        End Select
    Loop
    
    SE_USOU_HAKO_Open = False

End Function
