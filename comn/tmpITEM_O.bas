Attribute VB_Name = "tmpITEM_O"
Option Explicit
'********************************************************************
'*
'*              ��㎖�@���ϗp�i�ڃ}�X�^  �t�@�C����`
'*
'*          CREATE 2016.05.24
'********************************************************************
'�t�@�C���h�c
Public Const tmpITEM_O_ID$ = "tmpITEM_O"

'�y�[�W�T�C�Y
Public Const tmpITEM_O_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public tmpITEM_O_POS               As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************



'�f�[�^�E�o�b�t�@
Public tmpITEM_O_REC                As ITEM_O_REC_Tag

'�L�[��`
'�L�[�E�f�[�^
Public K0_tmpITEM_O                 As KEY0_ITEM_O
Public K1_tmpITEM_O                 As KEY1_ITEM_O




Private tmpITEM_O_Speck            As ITEM_O_FSpeck

Private Function tmpITEM_O_Create() As Integer
'********************************************************************
'*
'*              ��㎖�@���ϗp�i�ڃ}�X�^  CREATE
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************

Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    tmpITEM_O_Create = True
                                        '��㎖�@���ϗp�i�ڃ}�X�^ �t���p�X�捞��
    sts = GetIni("FILE", tmpITEM_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpITEM_O]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    tmpITEM_O_Speck.fs.recoleng = Len(tmpITEM_O_REC)          ' ���R�[�h��
    tmpITEM_O_Speck.fs.PageSize = tmpITEM_O_PG_SIZ            ' �y�[�W�T�C�Y
    tmpITEM_O_Speck.fs.idexnumb = 2                        ' �C���f�b�N�X��
    tmpITEM_O_Speck.fs.fileflag = 0                        ' �t�@�C���t���O
    tmpITEM_O_Speck.fs.reserve = &H0                       ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    tmpITEM_O_Speck.ks0.keypos = 1                         ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks0.keyleng = 1                        ' �L�[��
    tmpITEM_O_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    tmpITEM_O_Speck.ks0.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks0.reserve = &H0                      ' �\��ς�

    tmpITEM_O_Speck.ks1.keypos = 2                         ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks1.keyleng = 1                        ' �L�[��
    tmpITEM_O_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    tmpITEM_O_Speck.ks1.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks1.reserve = &H0                      ' �\��ς�

    tmpITEM_O_Speck.ks2.keypos = 3                         ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks2.keyleng = 20                       ' �L�[��
    tmpITEM_O_Speck.ks2.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    tmpITEM_O_Speck.ks2.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks2.reserve = &H0                      ' �\��ς�


    tmpITEM_O_Speck.ks3.keypos = 23                        ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks3.keyleng = 3                        ' �L�[��
    tmpITEM_O_Speck.ks3.keyflag = BtKfExt                  ' �L�[�t���O
    tmpITEM_O_Speck.ks3.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks3.reserve = &H0                      ' �\��ς�


'-----------------------------------------------
                                                ' �L�[�P
    tmpITEM_O_Speck.ks4.keypos = 1                         ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks4.keyleng = 1                        ' �L�[��
    tmpITEM_O_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' �L�[�t���O
    tmpITEM_O_Speck.ks4.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks4.reserve = &H0                      ' �\��ς�

    tmpITEM_O_Speck.ks5.keypos = 2                         ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks5.keyleng = 1                        ' �L�[��
    tmpITEM_O_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' �L�[�t���O
    tmpITEM_O_Speck.ks5.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks5.reserve = &H0                      ' �\��ς�

    tmpITEM_O_Speck.ks6.keypos = 3                         ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks6.keyleng = 20                       ' �L�[��
    tmpITEM_O_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' �L�[�t���O
    tmpITEM_O_Speck.ks6.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks6.reserve = &H0                      ' �\��ς�

    tmpITEM_O_Speck.ks7.keypos = 26                         ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks7.keyleng = 1                        ' �L�[��
    tmpITEM_O_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' �L�[�t���O
    tmpITEM_O_Speck.ks7.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks7.reserve = &H0                      ' �\��ς�

    tmpITEM_O_Speck.ks8.keypos = 27                         ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks8.keyleng = 1                        ' �L�[��
    tmpITEM_O_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' �L�[�t���O
    tmpITEM_O_Speck.ks8.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks8.reserve = &H0                      ' �\��ς�

    tmpITEM_O_Speck.ks9.keypos = 28                         ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks9.keyleng = 20                       ' �L�[��
    tmpITEM_O_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' �L�[�t���O
    tmpITEM_O_Speck.ks9.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks9.reserve = &H0                      ' �\��ς�


    tmpITEM_O_Speck.ks10.keypos = 23                        ' �L�[�|�W�V����
    tmpITEM_O_Speck.ks10.keyleng = 3                        ' �L�[��
    tmpITEM_O_Speck.ks10.keyflag = BtKfExt + BtKfDup + BtKfChg              ' �L�[�t���O
    tmpITEM_O_Speck.ks10.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpITEM_O_Speck.ks10.reserve = &H0                      ' �\��ς�

'-----------------------------------------------
    sts = BTRV(BtOpCreate, tmpITEM_O_POS, tmpITEM_O_Speck, Len(tmpITEM_O_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "��㎖�@���ϗp�i�ڃ}�X�^")
        Exit Function
    End If

    tmpITEM_O_Create = False

End Function

Public Function tmpITEM_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ��㎖�@���ϗp�i�ڃ}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    tmpITEM_O_Open = True
                                            '��㎖�@���ϗp�i�ڃ}�X�^ �t���p�X�捞��
    sts = GetIni("FILE", tmpITEM_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [tmpITEM_O]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpITEM_O_Create()    '��㎖�@���ϗp�i�ڃ}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "��㎖�@���ϗp�i�ڃ}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "��㎖�@���ϗp�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop

    tmpITEM_O_Open = False

End Function

Public Sub Rclr_tmpITEM_O_REC()

'********************************************************************
'*
'*              ��㎖�@���ϗp�i�ڃ}�X�^  ���R�[�h������
'*
'********************************************************************

    Call UniCode_Conv(tmpITEM_O_REC.JGYOBU, "")            '���ƕ��敪
    Call UniCode_Conv(tmpITEM_O_REC.NAIGAI, "")            '�����O
    Call UniCode_Conv(tmpITEM_O_REC.HIN_GAI, "")           '�i�ԁi�O���j


    Call UniCode_Conv(tmpITEM_O_REC.KO_JGYOBU, "")         '���ƕ��敪     2017.09.28
    Call UniCode_Conv(tmpITEM_O_REC.KO_NAIGAI, "")         '�����O         2017.09.28
    Call UniCode_Conv(tmpITEM_O_REC.KO_HIN_GAI, "")        '�i�ԁi�O���j   2017.09.28


    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_TANI, "")    '�����H���@�P��
    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_KIN, "")     '�����H���@���z

    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_TANI, "")       '���i���H���@�P��
    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_KIN, "")        '���i���H���@���z

    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_TANI, "")     'PF���H�@�P��
    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_KIN, "")      'PF���H�@���z

    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_TANI, "")     'PE���H�@�P��
    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_KIN, "")      'PE���H�@���z

    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_TANI, "")    'PF���ށ@�P��
    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_KIN, "")     'PF���ށ@���z

    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_TANI, "") '�i�ԕ\�����ف@�P��
    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_KIN, "") '�i�ԕ\�����ف@���z

    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_TANI, "")  '�ݒu�H���������@�P��
    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_KIN, "")  '�ݒu�H���������@���z

    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_TANI, "")       '����ށ@�P��
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_KIN, "")        '����ށ@���z

    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_TANI, "")  '�����ށ@�P��
    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_KIN, "")   '�����ށ@���z

    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_TANI, "")  '����ASSY�@�P��
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_KIN, "")   '����ASSY�@���z

    Call UniCode_Conv(tmpITEM_O_REC.KANRI_TANI, "")        '�Ǘ���@�P��
    Call UniCode_Conv(tmpITEM_O_REC.KANRI_KIN, "")         '�Ǘ���@���z
    
    Call UniCode_Conv(tmpITEM_O_REC.GOUKEI_KIN, "")        '���v�@���z
    Call UniCode_Conv(tmpITEM_O_REC.INPUT_TANTO_CODE, "")  '���͒S���Һ���
    
    
 '>>>>>>>>>>>>>>>>>>>>> 2017.09.27
    Call UniCode_Conv(tmpITEM_O_REC.BUZAI_TANTO_NAME, "")  '���ޒS���Җ�
    Call UniCode_Conv(tmpITEM_O_REC.T_HIN_NAME, "")        '��o�i��
    Call UniCode_Conv(tmpITEM_O_REC.TANI, "")              '�P��
    Call UniCode_Conv(tmpITEM_O_REC.T_TANKA, "")           '��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.T_KINGAKU, "")         '��o���z
    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_T_KIN, "")   '�����H���@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_T_KIN, "")      '���i���H���@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_T_KIN, "")    'PF���H�@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_T_KIN, "")    'PE���H�@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_T_KIN, "")   'PF���ށ@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_T_KIN, "") '�i�ԕ\�����ف@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_T_KIN, "") '�ݒu�H���������@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_T_KIN, "")      '����ށ@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_T_KIN, "") '�����ށ@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_T_KIN, "") '����ASSY�@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.KANRI_T_KIN, "")       '�Ǘ���@��o���z
    Call UniCode_Conv(tmpITEM_O_REC.GOUKEI_T_KIN, "")      '��o���v���z

    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_F, "")       '�����H�� ���Ϗ��\���׸�
    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_F, "")          '���i���H�� ���Ϗ��\���׸�
    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_F, "")        'PF���H ���Ϗ��\���׸�
    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_F, "")        'PE���H ���Ϗ��\���׸�
    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_F, "")       'PF���� ���Ϗ��\���׸�
    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_F, "")    '�i�ԕ\������ ���Ϗ��\���׸�
    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_F, "")     '�ݒu�H�������� ���Ϗ��\���׸�
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_F, "")          '����� ���Ϗ��\���׸�
    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_F, "")     '������ ���Ϗ��\���׸�
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_F, "")     '����ASSY ���Ϗ��\���׸�
    Call UniCode_Conv(tmpITEM_O_REC.KANRI_F, "")           '�Ǘ��� ���Ϗ��\���׸�
 '>>>>>>>>>>>>>>>>>>>>> 2017.09.27
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    Call UniCode_Conv(tmpITEM_O_REC.KO_QTY, "")            '����
    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_QTY, "")     '�����H���@����
    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_QTY, "")        '���i���H���@����
    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_QTY, "")      'PF���H�@����
    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_QTY, "")      'PE���H�@����
    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_QTY, "")      'PF���ށ@����
    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_QTY, "")  '�i�ԕ\�����ف@����
    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_QTY, "")   '�ݒu�H���������@����
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_QTY, "")        '����ށ@����
    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_QTY, "")   '�����ށ@����
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_QTY, "")   '����ASSY�@����
    Call UniCode_Conv(tmpITEM_O_REC.KANRI_QTY, "")         '�Ǘ���@����
    
    
    Call UniCode_Conv(tmpITEM_O_REC.NAKANISHI_T_TAN, "")   '�����H���@��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.SHOHIN_T_TAN, "")      '���i���H���@��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.PF_KAKOU_T_TAN, "")    'PF���H�@��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.PE_KAKOU_T_TAN, "")    'PE���H�@��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.PE_SHIZAI_T_TAN, "")   'PF���ށ@��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.HINBAN_LABEL_T_TAN, "") '�i�ԕ\�����ف@��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.KOUJI_SETSU_T_TAN, "") '�ݒu�H����+�����@��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_T_TAN, "")      '����ށ@��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.FUKU_SHIZAI_T_TAN, "") '�����ށ@��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.KONPOU_ASSY_T_TAN, "") '����ASSY�@��o�P��
    Call UniCode_Conv(tmpITEM_O_REC.KANRI_T_TAN, "")       '�Ǘ���@��o�P��
    
    
    
    Call UniCode_Conv(tmpITEM_O_REC.FILLER, "")
    
    Call UniCode_Conv(tmpITEM_O_REC.INS_TANTO, "")         '�ǉ��@�S����
    Call UniCode_Conv(tmpITEM_O_REC.Ins_DateTime, "")      '�ǉ��@����
    Call UniCode_Conv(tmpITEM_O_REC.UPD_TANTO, "")         '�X�V�@�S����
    Call UniCode_Conv(tmpITEM_O_REC.UPD_DATETIME, "")      '�X�V�@����



End Sub

