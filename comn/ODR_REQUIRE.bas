Attribute VB_Name = "ODR_REQUIRE"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���v�ʂe �t�@�C����`                           �@�@�@*
'*                                                                  *
'*          CREATE 2008.02.19                                       *
'********************************************************************
'�t�@�C���h�c
Public Const ODR_REQUIRE_ID$ = "ODR_REQUIRE"

'�y�[�W�T�C�Y
Private Const ODR_REQUIRE_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ODR_REQ_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type ODR_REQ_R_Tag

    SHIMUKE(0 To 1)        As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    KO_SYUBETSU(0 To 1)         As Byte         '�q�@���
    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    USE_YM(0 To 5)              As Byte         '�g�p���iYYYYMM)
    CYUMON_DT(0 To 7)           As Byte         '���ރZ���^�[�����[���iYYYYMMDD�j
    REQ_QTY(0 To 7)             As Byte         '�W�J��     9(5)v9(2)
    ODR_QTY(0 To 7)             As Byte         '���v��     9(5)v9(2)
    FUSOKU_QTY(0 To 7)          As Byte         '�s����     9(5)v9(2)
    UPD_TANTO(0 To 4)           As Byte         '�X�V�@�S����
    UPD_DT(0 To 7)              As Byte         '�X�V�@��
    UPD_TM(0 To 5)              As Byte         '�X�V�@����
    OK_DT(0 To 7)               As Byte         '
    FILLER(0 To 19)             As Byte         'Filler

End Type
'�f�[�^�E�o�b�t�@
Public ODR_REQ_R            As ODR_REQ_R_Tag



'�L�[��`

Type KEY0_ODR_REQUIRE                           '�j�d�x�O

    SHIMUKE(0 To 1)        As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
     
End Type

Type KEY1_ODR_REQUIRE                           '�j�d�x�P

    SHIMUKE(0 To 1)        As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��

End Type

Type KEY2_ODR_REQUIRE                           '�j�d�x�Q

    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    SHIMUKE(0 To 1)        As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��

End Type

Type KEY3_ODR_REQUIRE                           '�j�d�x�R

    USE_YM(0 To 5)              As Byte         '�g�p���iYYYYMM)
    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    SHIMUKE(0 To 1)        As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��

End Type


'�L�[�E�f�[�^
Public K0_ODR_REQ           As KEY0_ODR_REQUIRE
Public K1_ODR_REQ           As KEY1_ODR_REQUIRE
Public K2_ODR_REQ           As KEY2_ODR_REQUIRE
Public K3_ODR_REQ           As KEY3_ODR_REQUIRE

Type ODR_REQUIRE_FSpeck
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

End Type

Private ODR_REQUIRE_Speck       As ODR_REQUIRE_FSpeck
Private Function ODR_REQUIRE_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���v�ʂe  �b�q�d�`�s�d                               *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ODR_REQUIRE_Create = True
                                            '���v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_REQUIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_REQUIRE]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    ODR_REQUIRE_Speck.fs.recoleng = Len(ODR_REQ_R)      ' ���R�[�h��
    ODR_REQUIRE_Speck.fs.PageSize = ODR_REQUIRE_PG_SIZ          ' �y�[�W�T�C�Y
    ODR_REQUIRE_Speck.fs.idexnumb = 4                       ' �C���f�b�N�X��
    ODR_REQUIRE_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    ODR_REQUIRE_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    ODR_REQUIRE_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks0.keyleng = 61                      ' �L�[��
    ODR_REQUIRE_Speck.ks0.keyflag = BtKfChg + BtKfExt       ' �L�[�t���O
    ODR_REQUIRE_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks0.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    '--------------------------------------------------- �L�[�P ��
    ODR_REQUIRE_Speck.ks1.keypos = 1                        ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks1.keyleng = 4                       ' �L�[��
    ODR_REQUIRE_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' �L�[�t���O
    ODR_REQUIRE_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    ODR_REQUIRE_Speck.ks2.keypos = 42                       ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks2.keyleng = 20                      ' �L�[��
    ODR_REQUIRE_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' �L�[�t���O
    ODR_REQUIRE_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    ODR_REQUIRE_Speck.ks3.keypos = 15                        ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks3.keyleng = 37                      ' �L�[��
    ODR_REQUIRE_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg      ' �L�[�t���O
    ODR_REQUIRE_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks3.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�P ��
    
    '--------------------------------------------------- �L�[�Q ��
    ODR_REQUIRE_Speck.ks4.keypos = 64                        ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks4.keyleng = 2                       ' �L�[��
    ODR_REQUIRE_Speck.ks4.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' �L�[�t���O
    ODR_REQUIRE_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks4.reserve = &H0                     ' �\��ς�
    
    ODR_REQUIRE_Speck.ks5.keypos = 42                       ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks5.keyleng = 20                      ' �L�[��
    ODR_REQUIRE_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' �L�[�t���O
    ODR_REQUIRE_Speck.ks5.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks5.reserve = &H0                     ' �\��ς�
    
    ODR_REQUIRE_Speck.ks6.keypos = 1                        ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks6.keyleng = 41                      ' �L�[��
    ODR_REQUIRE_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg      ' �L�[�t���O
    ODR_REQUIRE_Speck.ks6.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks6.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�Q ��
    '--------------------------------------------------- �L�[�R ��
    ODR_REQUIRE_Speck.ks7.keypos = 66                        ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks7.keyleng = 6                       ' �L�[��
    ODR_REQUIRE_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' �L�[�t���O
    ODR_REQUIRE_Speck.ks7.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks7.reserve = &H0                     ' �\��ς�
    
    ODR_REQUIRE_Speck.ks8.keypos = 64                        ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks8.keyleng = 2                       ' �L�[��
    ODR_REQUIRE_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' �L�[�t���O
    ODR_REQUIRE_Speck.ks8.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks8.reserve = &H0                     ' �\��ς�
    
    ODR_REQUIRE_Speck.ks9.keypos = 42                       ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks9.keyleng = 20                      ' �L�[��
    ODR_REQUIRE_Speck.ks9.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' �L�[�t���O
    ODR_REQUIRE_Speck.ks9.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks9.reserve = &H0                     ' �\��ς�
    
    ODR_REQUIRE_Speck.ks10.keypos = 1                        ' �L�[�|�W�V����
    ODR_REQUIRE_Speck.ks10.keyleng = 41                      ' �L�[��
    ODR_REQUIRE_Speck.ks10.keyflag = BtKfExt + BtKfDup + BtKfChg      ' �L�[�t���O
    ODR_REQUIRE_Speck.ks10.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_REQUIRE_Speck.ks10.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�R ��
    
    sts = BTRV(BtOpCreate, ODR_REQ_POS, ODR_REQUIRE_Speck, Len(ODR_REQUIRE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���v�ʂe")
        Exit Function
    End If
    
    ODR_REQUIRE_Create = False

End Function

Public Function ODR_REQUIRE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���v�ʂe  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim yn          As Integer
Dim c           As String * 128
Dim FullPath    As String
Dim W_STR       As String
Dim W_PC        As String
Dim X_i         As Long

    ODR_REQUIRE_Open = True
                                            '���v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_REQUIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_REQUIRE]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)


    Do
        sts = BTRV(BtOpOpen, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                yn = MsgBox("���Ŏg�p���ł��I<���v�ʂe>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function

            Case BtErrFileNotFound
                sts = ODR_REQUIRE_Create()      '���v�ʂe �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���v�ʂe")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���v�ʂe")
                Exit Function
        End Select
    Loop
    
    ODR_REQUIRE_Open = False
    
End Function

Public Function ODR_REQUIRE_GET(SM As String, JB As String, NG As String, _
                    YM As String, HG As String, OD As String, BN As String, Locked As Integer) As Integer
'           ����
'   SM      �d����
'   JB      ���ƕ�
'   NG      ���O
'   YM      �g�p��
'   HG      �e�i��
'   OD      ������
'   BN      ���[��
'   Locked  �f�����k������


Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    ODR_REQUIRE_GET = True
    
    Call UniCode_Conv(K0_ODR_REQ.SHIMUKE, SM)
    Call UniCode_Conv(K0_ODR_REQ.JGYOBU, JB)
    Call UniCode_Conv(K0_ODR_REQ.NAIGAI, NG)
    Call UniCode_Conv(K0_ODR_REQ.HIN_GAI, HG)
    Call UniCode_Conv(K0_ODR_REQ.ORDER_NO, OD)
    Call UniCode_Conv(K0_ODR_REQ.BUN_NO, BN)
    
'2019.01.08    com = BtOpGetEqual + Locked
    com = BtOpGetEqual
    Do
        sts = BTRV(com, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       '���R�[�h����
                'Beep
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                yn = MsgBox("���Ŏg�p���ł��I<���v�ʂe>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_REQUIRE")
                Exit Function
        End Select
    Loop
    
    ODR_REQUIRE_GET = False

End Function


Public Sub ODR_REQUIRE_CLR()
    
    Call UniCode_Conv(ODR_REQ_R.SHIMUKE, "")
    Call UniCode_Conv(ODR_REQ_R.JGYOBU, "")
    Call UniCode_Conv(ODR_REQ_R.NAIGAI, "")
    Call UniCode_Conv(ODR_REQ_R.USE_YM, "")
    Call UniCode_Conv(ODR_REQ_R.ORDER_NO, "")
    Call UniCode_Conv(ODR_REQ_R.INS_NO, "")
    Call UniCode_Conv(ODR_REQ_R.BUN_NO, "")
    Call UniCode_Conv(ODR_REQ_R.HIN_GAI, "")
    Call UniCode_Conv(ODR_REQ_R.REQ_QTY, String(UBound(ODR_REQ_R.REQ_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_REQ_R.ODR_QTY, String(UBound(ODR_REQ_R.ODR_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_REQ_R.FUSOKU_QTY, String(UBound(ODR_REQ_R.FUSOKU_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_REQ_R.KO_HIN_GAI, "")
    Call UniCode_Conv(ODR_REQ_R.KO_SYUBETSU, "")
    Call UniCode_Conv(ODR_REQ_R.KO_JGYOBU, "")
    Call UniCode_Conv(ODR_REQ_R.KO_NAIGAI, "")
    Call UniCode_Conv(ODR_REQ_R.CYUMON_DT, "")
    Call UniCode_Conv(ODR_REQ_R.UPD_TANTO, "")
    Call UniCode_Conv(ODR_REQ_R.UPD_DT, "")
    Call UniCode_Conv(ODR_REQ_R.UPD_TM, "")
    Call UniCode_Conv(ODR_REQ_R.OK_DT, "")
    Call UniCode_Conv(ODR_REQ_R.FILLER, "")
End Sub

