Attribute VB_Name = "ODR_ORDER"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �e�i�ԁ@�����e �t�@�C����`                           *
'*                                                                  *
'*          CREATE 2008.02.19                                       *
'
'           2012.04.13          PRT_FLG �ǉ�
'
'********************************************************************
'�t�@�C���h�c
Public Const ODR_ORDER_ID$ = "ODR_ORDER"

'�y�[�W�T�C�Y
Private Const ODR_ORDER_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ODR_ORDER_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type ODR_ORDER_REC_Tag
    SHIMUKE(0 To 1)             As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��
    USE_YM(0 To 5)              As Byte         '�g�p���iYYYYMM)
    BUN_KB(0 To 2)              As Byte         '���[�@�L���敪
    REQ_KB(0 To 0)              As Byte         '�W�J�敪
    ODR_QTY(0 To 4)             As Byte         '������
    CYUMON_DT(0 To 7)           As Byte         '���ރZ���^�[�����[���iYYYYMMDD�j
    KAITO_DT(0 To 7)            As Byte         '�񓚔[��
    FIN_DT(0 To 7)              As Byte         '�������t
    KUMI_OK_DT(0 To 7)          As Byte         '�g���\���t
    ODR_BMN(0 To 4)             As Byte         '��������
    DEN_NO(0 To 9)              As Byte         '�`�[��
    UPD_TANTO(0 To 4)           As Byte         '�X�V�@�S����
    INS_DT(0 To 7)              As Byte         '�ǉ��@���t
    INS_TM(0 To 5)              As Byte         '�ǉ��@����
    USE_YM_MOTO(0 To 5)         As Byte         '�g�p���iYYYYMM�j�v���O�����N�����̓��e
'    FILLER(0 To 21)             As Byte         'Filler
    UPD_DT(0 To 7)              As Byte         '�X�V�@���t
    UPD_TM(0 To 5)              As Byte         '�X�V�@����
    UPD_PG(0 To 6)              As Byte         '�X�V�@�v���O����
    PRT_FLG(0 To 0)             As Byte         '�w�}�\���         F:�ς݁A��:�����   2012.04.13

End Type
'�f�[�^�E�o�b�t�@
Public ODR_ORDER_REC            As ODR_ORDER_REC_Tag



'�L�[��`

Type KEY0_ODR_ORDER                           '�j�d�x�O
    SHIMUKE(0 To 1)             As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��
End Type

Type KEY1_ODR_ORDER                           '�j�d�x�P
    SHIMUKE(0 To 1)             As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    USE_YM(0 To 5)              As Byte         '�g�p���iYYYYMM)
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��
End Type

Type KEY2_ODR_ORDER                           '�j�d�x�Q
    ODR_QTY(0 To 4)             As Byte         '������
End Type

Type KEY3_ODR_ORDER                           '�j�d�x�R
    KAITO_DT(0 To 7)            As Byte         '�񓚔[��
End Type

Type KEY4_ODR_ORDER                           '�j�d�x�S
    FIN_DT(0 To 7)              As Byte         '�������t
End Type

Type KEY5_ODR_ORDER                           '�j�d�x�T         '2009/03/12�ǉ�
    SHIMUKE(0 To 1)             As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    USE_YM(0 To 5)              As Byte         '�g�p���iYYYYMM)
    INS_DT(0 To 7)              As Byte         '�ǉ��@���t
    INS_TM(0 To 5)              As Byte         '�ǉ��@����
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��
End Type

Type KEY6_ODR_ORDER                             '�j�d�x�U         '20012/04/13�ǉ�
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
End Type


'�L�[�E�f�[�^
Public K0_ODR_ORDER           As KEY0_ODR_ORDER
Public K1_ODR_ORDER           As KEY1_ODR_ORDER
Public K2_ODR_ORDER           As KEY2_ODR_ORDER

Public K3_ODR_ORDER           As KEY3_ODR_ORDER
Public K4_ODR_ORDER           As KEY4_ODR_ORDER
Public K5_ODR_ORDER           As KEY5_ODR_ORDER

Public K6_ODR_ORDER           As KEY6_ODR_ORDER


Type ODR_ORDER_FSpeck
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

End Type

Private ODR_ORDER_Speck       As ODR_ORDER_FSpeck
Private Function ODR_ORDER_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �e�Q�����e  �b�q�d�`�s�d                            *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ODR_ORDER_Create = True
                                            '�e�Q�����e�t���p�X�捞��
    sts = GetIni("FILE", ODR_ORDER_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ODR_ORDER]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    ODR_ORDER_Speck.fs.recoleng = Len(ODR_ORDER_REC)      ' ���R�[�h��
    ODR_ORDER_Speck.fs.PageSize = ODR_ORDER_PG_SIZ        ' �y�[�W�T�C�Y
    ODR_ORDER_Speck.fs.idexnumb = 7                       ' �C���f�b�N�X��
    ODR_ORDER_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    ODR_ORDER_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    ODR_ORDER_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks0.keyleng = 41                      ' �L�[��
    ODR_ORDER_Speck.ks0.keyflag = BtKfChg + BtKfExt       ' �L�[�t���O
    ODR_ORDER_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_ORDER_Speck.ks0.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�P ��
    ODR_ORDER_Speck.ks1.keypos = 1                        ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks1.keyleng = 4                       ' �L�[��
    ODR_ORDER_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' �L�[�t���O
    ODR_ORDER_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_ORDER_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    ODR_ORDER_Speck.ks2.keypos = 42                       ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks2.keyleng = 6                       ' �L�[��
    ODR_ORDER_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' �L�[�t���O
    ODR_ORDER_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_ORDER_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    ODR_ORDER_Speck.ks3.keypos = 5                        ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks3.keyleng = 37                      ' �L�[��
    ODR_ORDER_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg      ' �L�[�t���O
    ODR_ORDER_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_ORDER_Speck.ks3.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�P ��
    
    '--------------------------------------------------- �L�[�Q ��
    ODR_ORDER_Speck.ks4.keypos = 52                       ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks4.keyleng = 5                       ' �L�[��
    ODR_ORDER_Speck.ks4.keyflag = BtKfDup + BtKfChg + BtKfExt       ' �L�[�t���O
    ODR_ORDER_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_ORDER_Speck.ks4.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�Q ��
    
    
    '--------------------------------------------------- �L�[�R ��
    
    ODR_ORDER_Speck.ks5.keypos = 65                         ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks5.keyleng = 8                         ' �L�[��
                                                            ' �L�[�t���O
    ODR_ORDER_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg
    ODR_ORDER_Speck.ks5.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ODR_ORDER_Speck.ks5.reserve = &H0                       ' �\��ς�
    '--------------------------------------------------- �L�[�R ��
    
    '--------------------------------------------------- �L�[�S ��
    
    ODR_ORDER_Speck.ks6.keypos = 73                         ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks6.keyleng = 8                         ' �L�[��
                                                            ' �L�[�t���O
    ODR_ORDER_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg
    ODR_ORDER_Speck.ks6.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    ODR_ORDER_Speck.ks6.reserve = &H0                       ' �\��ς�
    '--------------------------------------------------- �L�[�S ��
    
    '--------------------------------------------------- �L�[�T ��
    ODR_ORDER_Speck.ks7.keypos = 1                        ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks7.keyleng = 4                       ' �L�[��
    ODR_ORDER_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' �L�[�t���O
    ODR_ORDER_Speck.ks7.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_ORDER_Speck.ks7.reserve = &H0                     ' �\��ς�
    
    ODR_ORDER_Speck.ks8.keypos = 42                       ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks8.keyleng = 6                       ' �L�[��
    ODR_ORDER_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' �L�[�t���O
    ODR_ORDER_Speck.ks8.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_ORDER_Speck.ks8.reserve = &H0                     ' �\��ς�
    
    ODR_ORDER_Speck.ks9.keypos = 109                      ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks9.keyleng = 14                      ' �L�[��
    ODR_ORDER_Speck.ks9.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' �L�[�t���O
    ODR_ORDER_Speck.ks9.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_ORDER_Speck.ks9.reserve = &H0                     ' �\��ς�
    
    
    ODR_ORDER_Speck.ks10.keypos = 5                        ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks10.keyleng = 37                      ' �L�[��
    ODR_ORDER_Speck.ks10.keyflag = BtKfExt + BtKfDup + BtKfChg      ' �L�[�t���O
    ODR_ORDER_Speck.ks10.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_ORDER_Speck.ks10.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�T ��
    '--------------------------------------------------- �L�[�U ��
    ODR_ORDER_Speck.ks11.keypos = 25                        ' �L�[�|�W�V����
    ODR_ORDER_Speck.ks11.keyleng = 10                       ' �L�[��
                                                            ' �L�[�t���O
    ODR_ORDER_Speck.ks11.keyflag = BtKfChg + BtKfDup + BtKfExt
    ODR_ORDER_Speck.ks11.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    ODR_ORDER_Speck.ks11.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�U ��
    
    sts = BTRV(BtOpCreate, ODR_ORDER_POS, ODR_ORDER_Speck, Len(ODR_ORDER_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�e�Q�����e")
        Exit Function
    End If
    
    ODR_ORDER_Create = False

End Function

Public Function ODR_ORDER_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �e�Q�����e  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ODR_ORDER_Open = True
                                            '�e�Q�����e�t���p�X�捞��
    sts = GetIni("FILE", ODR_ORDER_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ODR_ORDER]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ODR_ORDER_Create()      '�e�Q�����e�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�e �����e")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�e �����e")
                Exit Function
        End Select
    Loop
    
    ODR_ORDER_Open = False
    
End Function
Public Function ODR_ORDER_GET(SM As String, JB As String, NG As String, HG As String, _
                               i_NO As String, OD As String, BN As String, Locked As Integer) As Integer
'           ����
'   JB      ���ƕ�
'   NG      ���O
'   HG      �i��
'   OD      ������
'   I_No    Key�@��
'   BN      ���[��
'   Locked  �f�����k������


Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    ODR_ORDER_GET = True
    Call UniCode_Conv(K0_ODR_ORDER.SHIMUKE, SM)
    Call UniCode_Conv(K0_ODR_ORDER.JGYOBU, JB)
    Call UniCode_Conv(K0_ODR_ORDER.NAIGAI, NG)
    Call UniCode_Conv(K0_ODR_ORDER.HIN_GAI, HG)
    Call UniCode_Conv(K0_ODR_ORDER.INS_NO, i_NO)
    Call UniCode_Conv(K0_ODR_ORDER.ORDER_NO, OD)
    Call UniCode_Conv(K0_ODR_ORDER.BUN_NO, BN)
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       '���R�[�h����
                'Beep
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                yn = MsgBox("���Ŏg�p���ł��I<�����e>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_ORDER")
                Exit Function
        End Select
    Loop
    
    ODR_ORDER_GET = False

End Function


Public Sub ODR_ORDER_CLR()
    
    Call UniCode_Conv(ODR_ORDER_REC.SHIMUKE, "")
    Call UniCode_Conv(ODR_ORDER_REC.JGYOBU, "")
    Call UniCode_Conv(ODR_ORDER_REC.NAIGAI, "")
    Call UniCode_Conv(ODR_ORDER_REC.USE_YM, "")
    Call UniCode_Conv(ODR_ORDER_REC.INS_NO, String(UBound(ODR_ORDER_REC.INS_NO) + 1, "0"))
    Call UniCode_Conv(ODR_ORDER_REC.ORDER_NO, "")
    Call UniCode_Conv(ODR_ORDER_REC.BUN_NO, "")
    Call UniCode_Conv(ODR_ORDER_REC.HIN_GAI, "")
    Call UniCode_Conv(ODR_ORDER_REC.BUN_KB, String(UBound(ODR_ORDER_REC.BUN_KB) + 1, "0"))
    Call UniCode_Conv(ODR_ORDER_REC.REQ_KB, String(UBound(ODR_ORDER_REC.REQ_KB) + 1, "0"))
    Call UniCode_Conv(ODR_ORDER_REC.ODR_QTY, String(UBound(ODR_ORDER_REC.ODR_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_ORDER_REC.CYUMON_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.KAITO_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.FIN_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.KUMI_OK_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.ODR_BMN, "")
    Call UniCode_Conv(ODR_ORDER_REC.DEN_NO, "")
    Call UniCode_Conv(ODR_ORDER_REC.UPD_TANTO, "")
    Call UniCode_Conv(ODR_ORDER_REC.INS_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.INS_TM, "")
    Call UniCode_Conv(ODR_ORDER_REC.USE_YM_MOTO, "")
    'Call UniCode_Conv(ODR_ORDER_REC.FILLER, "")
    Call UniCode_Conv(ODR_ORDER_REC.UPD_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.UPD_TM, "")
    Call UniCode_Conv(ODR_ORDER_REC.UPD_PG, "")
    Call UniCode_Conv(ODR_ORDER_REC.PRT_FLG, "")

End Sub

