Attribute VB_Name = "P_KANRI"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �Ǘ��}�X�^  �t�@�C����`                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�t�@�C���h�c
Public Const P_KANRI_ID$ = "P_KANRI"

'�y�[�W�T�C�Y
Private Const P_KANRI_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public P_KANRI_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type P_KANRIREC_Tag
    REC_NO(0 To 1)          As Byte         'ں��އ�
    SHIME_DD(0 To 1)        As Byte         'SDC�������ߓ�
    xSASHIZU_NO(0 To 4)     As Byte         '�w�}�[��(���ݒl+1) ���g�p�Ƃ��� 2007.11.28
    ORDER_NO(0 To 4)        As Byte         '������(���ݒl+1)
    URIAGE_NO(0 To 4)       As Byte         '���ޔ���ں��އ�(���ݒl+1)
    
    ZEI_CHANGE_YMD(0 To 7)  As Byte         '����ŕύX���t
    NOW_ZEI_RITU(0 To 3)    As Byte         '���@����ŗ�
    NOW_MARUME(0 To 0)      As Byte         '    �ۂ�
    NEW_ZEI_RITU(0 To 3)    As Byte         '�V�@����ŗ�
    NEW_MARUME(0 To 0)      As Byte         '    �ۂ�
    
    SHONIN_CODE(0 To 4)     As Byte         '���F�Һ���
    KAISHA_NAME(0 To 29)    As Byte         '��Ж�
    CENTER_NAME(0 To 29)    As Byte         '�Z���^�[��
    TEL_NO(0 To 14)         As Byte         '�d�b�ԍ�
    FAX_NO(0 To 14)         As Byte         'FAX�ԍ�
    
    URI_MARUME(0 To 0)      As Byte         '������z�ۂ�
    SHI_MARUME(0 To 0)      As Byte         '�d�����z�ۂ�
    
    SASHIZU_NO(0 To 7)      As Byte         '������(���ݒl+1)   2007.11.28
    
    
    NYUKO_S_RATE(0 To 6)    As Byte         '���Ɂ@�����[�g     2008.02.13
    NYUKO_R_RATE(0 To 6)    As Byte         '���Ɂ@�]�T��       2008.02.13
    
    SYUKO_S_RATE(0 To 6)    As Byte         '�o�Ɂ@�����[�g     2008.02.13
    SYUKO_R_RATE(0 To 6)    As Byte         '�o�Ɂ@�]�T��       2008.02.13
    
    SYUKA_S_RATE(0 To 6)    As Byte         '�o�Ɂ@�����[�g     2008.02.13
    SYUKA_R_RATE(0 To 6)    As Byte         '�o�Ɂ@�]�T��       2008.02.13
    
    KOUTEI_LOT(0 To 5)      As Byte         '�H���@�O��H���W�����b�g   2008.02.13
    KOUTEI_S_RATE(0 To 6)   As Byte         '�H���@�����[�g             2008.02.13
    KOUTEI_R_RATE(0 To 6)   As Byte         '�H���@�]�T��               2008.02.13
    KOUTEI_SHIZAI(0 To 2)   As Byte         '�H���@�����ފm�F�_��       2008.02.13
    KOUTEI_BUHIN(0 To 2)    As Byte         '�H���@�������i�m�F�_��     2008.02.13
    KOUTEI_LABEL(0 To 2)    As Byte         '�H���@���x���\�t����       2008.02.13
    
    MITSUMORI_NO(0 To 7)    As Byte         '���Ϗ���   2008.02.13
    SEIKYU_NO(0 To 7)       As Byte         '��������   2008.02.13
        
    
    MIN_URIAGE_NO(0 To 7)   As Byte         '�~�j�}�����㇂     2008.02.13
    
    
    FILLER(0 To 18)         As Byte         'FILLER
End Type
'�f�[�^�E�o�b�t�@
Public P_KANRIREC           As P_KANRIREC_Tag




Private Type P_KOTEI_Tag                    '2008.02.13
    KOTEI(0 To 2)       As Byte
End Type

Public Type P_KANRIREC02_Tag                '2008.02.13
    REC_NO(0 To 1)          As Byte         'ں��އ�
        
    BEF_KOTEI(0 To 9)       As P_KOTEI_Tag  '�O�H��
    MAIN_KOTEI(0 To 9)      As P_KOTEI_Tag  '��ƍH��
    AFT_KOTEI(0 To 9)       As P_KOTEI_Tag  '��H��
        
    FUTAI_KOTEI(0 To 4)     As P_KOTEI_Tag  '�t�эH���@(���ݖ��g�p)
    KEIHI(0 To 4)           As P_KOTEI_Tag  '�o��@(���ݖ��g�p)
    
    FILLER(0 To 133)        As Byte         'FILLER
End Type
'�f�[�^�E�o�b�t�@
Public P_KANRIREC02         As P_KANRIREC02_Tag



'�L�[��`

Type KEY0_P_KANRI           '�j�d�x�O
    REC_NO(0 To 1)          As Byte         'ں��އ�
End Type
    
'�L�[�E�f�[�^
Public K0_P_KANRI           As KEY0_P_KANRI

Type P_KANRI_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_KANRI_Speck       As P_KANRI_FSpeck
Private Function P_KANRI_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �Ǘ��}�X�^  �b�q�d�`�s�d                            *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_KANRI_Create = True
                                            '�Ǘ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_KANRI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_KANRI]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_KANRI_Speck.fs.recoleng = Len(P_KANRIREC)            ' ���R�[�h��
    P_KANRI_Speck.fs.PageSize = P_KANRI_PG_SIZ          ' �y�[�W�T�C�Y
    P_KANRI_Speck.fs.idexnumb = 1                       ' �C���f�b�N�X��
    P_KANRI_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    P_KANRI_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_KANRI_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    P_KANRI_Speck.ks0.keyleng = 2                       ' �L�[��
    P_KANRI_Speck.ks0.keyflag = BtKfExt                 ' �L�[�t���O
    P_KANRI_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_KANRI_Speck.ks0.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    sts = BTRV(BtOpCreate, P_KANRI_POS, P_KANRI_Speck, Len(P_KANRI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�Ǘ��}�X�^")
        Exit Function
    End If
    
    P_KANRI_Create = False

End Function

Public Function P_KANRI_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �Ǘ��}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_KANRI_Open = True
                                            '�Ǘ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", P_KANRI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_KANRI]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_KANRI_Create()      '�Ǘ��}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�Ǘ��}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�Ǘ��}�X�^")
                Exit Function
        End Select
    Loop
    
    P_KANRI_Open = False

End Function
Public Function P_KANRI_MAKE_Proc() As Integer
'----------------------------------------------------------------------------
'                   �Ǘ��}�X�^�̎����쐬
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

    P_KANRI_MAKE_Proc = True

    Call UniCode_Conv(P_KANRIREC.REC_NO, P_ST_KANRI_No)     'ں��އ�
    Call UniCode_Conv(P_KANRIREC.SHIME_DD, "31")            '�������ߓ�
    Call UniCode_Conv(P_KANRIREC.SASHIZU_NO, "00000")       '�w�}�[��
    Call UniCode_Conv(P_KANRIREC.ORDER_NO, "00000")         '������
    Call UniCode_Conv(P_KANRIREC.URIAGE_NO, "00000")        '���ޔ���ں��އ�

    Call UniCode_Conv(P_KANRIREC.ZEI_CHANGE_YMD, "")        '����ŕύX���t
    Call UniCode_Conv(P_KANRIREC.NOW_ZEI_RITU, "00.0")      '���@����ŗ�
    Call UniCode_Conv(P_KANRIREC.NOW_MARUME, "0")           '���@�ۂ�
    Call UniCode_Conv(P_KANRIREC.NEW_ZEI_RITU, "00.0")      '�V�@����ŗ�
    Call UniCode_Conv(P_KANRIREC.NEW_MARUME, "0")           '�V�@�ۂ�

    Call UniCode_Conv(P_KANRIREC.SHONIN_CODE, "")           '���F�Һ���
    Call UniCode_Conv(P_KANRIREC.KAISHA_NAME, "")           '��Ж���
    Call UniCode_Conv(P_KANRIREC.TEL_NO, "")                '�d�b�ԍ�
    Call UniCode_Conv(P_KANRIREC.FAX_NO, "")                'FAX�ԍ�
    
    Call UniCode_Conv(P_KANRIREC.FILLER, "")

    Do
        
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        
        sts = BTRV(BtOpInsert, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "�Ǘ��}�X�^")
                Exit Function
        End Select
    Loop
    
    
    P_KANRI_MAKE_Proc = False



End Function

