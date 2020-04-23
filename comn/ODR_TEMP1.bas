Attribute VB_Name = "ODR_TEMP1"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���ԁ@���v�ʂe�iWORK) �t�@�C����`              �@�@�@*
'*                                                                  *
'*          CREATE 2008.02.19                                       *
'********************************************************************
'�t�@�C���h�c
Public Const ODR_TEMP1_ID$ = "ODR_TEMP1"

'�y�[�W�T�C�Y
Private Const ODR_TEMP1_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ODR_TP1_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type ODR_TP1_R_Tag

    KAITO_DT(0 To 7)            As Byte         '�e�����̉񓚔[��
    CYUMON_DT(0 To 7)           As Byte         '���ރZ���^�[�����[���iYYYYMMDD�j
    USE_YM(0 To 5)              As Byte         '�g�p��
    SHIMUKE(0 To 1)             As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��
    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    KO_SYUBETSU(0 To 1)         As Byte         '�q�@���
    KO_QTY(0 To 5)              As Byte         '�q�@����(999V99)
    OK_DT(0 To 7)               As Byte         '�o�ɉ\�� YYYYMMDD
    KAN_KB(0 To 0)              As Byte         '�e�����̊����敪
    ALL_QTY(0 To 8)             As Byte         '�W�J��     9(5)v9(2)
    USE_QTY(0 To 8)             As Byte         '�g�p��     9(5)v9(2)
    NED_QTY(0 To 8)             As Byte         '�K�v��     9(5)v9(2)
    REQ_QTY(0 To 8)             As Byte         '���v��     9(5)v9(2)
    FUSOKU_QTY(0 To 8)          As Byte         '�s����     9(5)v9(2)
    UPDT_DT(0 To 5)             As Byte         '�X�V��     YYMMDD
    UPDT_TM(0 To 3)             As Byte         '�X�V����   hhmm
    FILLER(0 To 17)             As Byte         'Filler

End Type
'�f�[�^�E�o�b�t�@
Public ODR_TP1_R            As ODR_TP1_R_Tag



'�L�[��`

Type KEY0_ODR_TEMP1                           '�j�d�x�O

    SHIMUKE(0 To 1)             As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��
    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    
End Type

Type KEY1_ODR_TEMP1                           '�j�d�x�P

    SHIMUKE(0 To 1)             As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��
    
    OK_DT(0 To 7)               As Byte         '�o�ɉ\�� YYYYMMDD

End Type

Type KEY2_ODR_TEMP1                           '�j�d�x�Q

    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    
    SHIMUKE(0 To 1)             As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��

End Type

Type KEY3_ODR_TEMP1                           '�j�d�x�R

    KAN_KB(0 To 0)              As Byte         '�e�����̊����敪
    
    KAITO_DT(0 To 7)            As Byte         '�e�����̉񓚔[��
    CYUMON_DT(0 To 7)           As Byte         '���ރZ���^�[�����[���iYYYYMMDD�j
    USE_YM(0 To 5)              As Byte         '�g�p��
    SHIMUKE(0 To 1)             As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��
    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��

End Type

Type KEY4_ODR_TEMP1                           '�j�d�x�S�i2010/05/07�ǉ��j
    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    
    KAN_KB(0 To 0)              As Byte         '�e�����̊����敪
    
    KAITO_DT(0 To 7)            As Byte         '�e�����̉񓚔[��
    CYUMON_DT(0 To 7)           As Byte         '���ރZ���^�[�����[���iYYYYMMDD�j
    USE_YM(0 To 5)              As Byte         '�g�p��
    SHIMUKE(0 To 1)             As Byte         '�d������
    JGYOBU(0 To 0)              As Byte         '���ƕ�
    NAIGAI(0 To 0)              As Byte         '�����O
    HIN_GAI(0 To 19)            As Byte         '�e�i��
    ORDER_NO(0 To 9)            As Byte         '�e�i�ԁ@������
    INS_NO(0 To 3)              As Byte         '�o�^��
    BUN_NO(0 To 2)              As Byte         '���[��

End Type

'�L�[�E�f�[�^
Public K0_ODR_TEMP1           As KEY0_ODR_TEMP1
Public K1_ODR_TEMP1           As KEY1_ODR_TEMP1
Public K2_ODR_TEMP1           As KEY2_ODR_TEMP1
Public K3_ODR_TEMP1           As KEY3_ODR_TEMP1
Public K4_ODR_TEMP1           As KEY4_ODR_TEMP1

Type ODR_TEMP1_FSpeck
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

End Type

Private ODR_TEMP1_Speck       As ODR_TEMP1_FSpeck
Private Function ODR_TEMP1_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���ԏ��v�ʂe  �b�q�d�`�s�d                            *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128
Dim W_STR       As String
Dim W_PC        As String
Dim X_i         As Long

    ODR_TEMP1_Create = True
                                            '���ԏ��v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_TEMP1_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP1]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    c = Space(255)
    If GetComputerNameA(c, 255) <> 0 Then
        W_PC = Left(c, InStr(c, vbNullChar) - 1)
    Else
        W_PC = "000"
    End If
    
    X_i = InStr(1, FullPath, "*") - 1
    If X_i <= 0 Then
        X_i = Len(Trim(FullPath)) - 4
    End If
    W_STR = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)
    
    FullPath = W_STR


    ODR_TEMP1_Speck.fs.recoleng = Len(ODR_TP1_R)      ' ���R�[�h��
    ODR_TEMP1_Speck.fs.PageSize = ODR_TEMP1_PG_SIZ          ' �y�[�W�T�C�Y
    ODR_TEMP1_Speck.fs.idexnumb = 5                       ' �C���f�b�N�X��
    ODR_TEMP1_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    ODR_TEMP1_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    ODR_TEMP1_Speck.ks0.keypos = 23                       ' �L�[�|�W�V����
    ODR_TEMP1_Speck.ks0.keyleng = 63                      ' �L�[��
    ODR_TEMP1_Speck.ks0.keyflag = BtKfChg + BtKfDup + BtKfExt      ' �L�[�t���O
    ODR_TEMP1_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP1_Speck.ks0.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    ODR_TEMP1_Speck.ks1.keypos = 23                        ' �L�[�|�W�V����
    ODR_TEMP1_Speck.ks1.keyleng = 41                      ' �L�[��
    ODR_TEMP1_Speck.ks1.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt    ' �L�[�t���O
    ODR_TEMP1_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP1_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    ODR_TEMP1_Speck.ks2.keypos = 89                       ' �L�[�|�W�V����
    ODR_TEMP1_Speck.ks2.keyleng = 8                       ' �L�[��
    ODR_TEMP1_Speck.ks2.keyflag = BtKfChg + BtKfDup + BtKfExt      ' �L�[�t���O
    ODR_TEMP1_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP1_Speck.ks2.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�P ��
    
    '--------------------------------------------------- �L�[�Q ��
    ODR_TEMP1_Speck.ks3.keypos = 64                       ' �L�[�|�W�V����
    ODR_TEMP1_Speck.ks3.keyleng = 22                      ' �L�[��
    ODR_TEMP1_Speck.ks3.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt       ' �L�[�t���O
    ODR_TEMP1_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP1_Speck.ks3.reserve = &H0                     ' �\��ς�
    
    ODR_TEMP1_Speck.ks4.keypos = 23                       ' �L�[�|�W�V����
    ODR_TEMP1_Speck.ks4.keyleng = 41                      ' �L�[��
    ODR_TEMP1_Speck.ks4.keyflag = BtKfChg + BtKfDup + BtKfExt               ' �L�[�t���O
    ODR_TEMP1_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP1_Speck.ks4.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�Q ��
    
    '--------------------------------------------------- �L�[�R ��
    ODR_TEMP1_Speck.ks5.keypos = 102                     ' �L�[�|�W�V����
    ODR_TEMP1_Speck.ks5.keyleng = 1                      ' �L�[��
    ODR_TEMP1_Speck.ks5.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt       ' �L�[�t���O
    ODR_TEMP1_Speck.ks5.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP1_Speck.ks5.reserve = &H0                     ' �\��ς�
    
    ODR_TEMP1_Speck.ks6.keypos = 1                        ' �L�[�|�W�V����
    ODR_TEMP1_Speck.ks6.keyleng = 85                       ' �L�[��
    ODR_TEMP1_Speck.ks6.keyflag = BtKfChg + BtKfDup + BtKfExt      ' �L�[�t���O
    ODR_TEMP1_Speck.ks6.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP1_Speck.ks6.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�R ��
    
    
    '--------------------------------------------------- �L�[�S ��      '2010/05/07�ǉ�
    ODR_TEMP1_Speck.ks7.keypos = 64                       ' �L�[�|�W�V����
    ODR_TEMP1_Speck.ks7.keyleng = 22                      ' �L�[��
    ODR_TEMP1_Speck.ks7.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt       ' �L�[�t���O
    ODR_TEMP1_Speck.ks7.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP1_Speck.ks7.reserve = &H0                     ' �\��ς�
    
    ODR_TEMP1_Speck.ks8.keypos = 102                     ' �L�[�|�W�V����
    ODR_TEMP1_Speck.ks8.keyleng = 1                      ' �L�[��
    ODR_TEMP1_Speck.ks8.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt       ' �L�[�t���O
    ODR_TEMP1_Speck.ks8.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP1_Speck.ks8.reserve = &H0                     ' �\��ς�
    
    ODR_TEMP1_Speck.ks9.keypos = 1                        ' �L�[�|�W�V����
    ODR_TEMP1_Speck.ks9.keyleng = 63                       ' �L�[��
    ODR_TEMP1_Speck.ks9.keyflag = BtKfChg + BtKfDup + BtKfExt      ' �L�[�t���O
    ODR_TEMP1_Speck.ks9.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP1_Speck.ks9.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�S ��
    

    sts = BTRV(BtOpCreate, ODR_TP1_POS, ODR_TEMP1_Speck, Len(ODR_TEMP1_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ODR_TEMP1")
        Exit Function
    End If
    
    ODR_TEMP1_Create = False

End Function

Public Function ODR_TEMP1_Open(Mode As Integer) As Integer
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

    ODR_TEMP1_Open = True
                                            '���v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_TEMP1_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP1]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)


    
    c = Space(255)
    If GetComputerNameA(c, 255) <> 0 Then
        W_PC = Left(c, InStr(c, vbNullChar) - 1)
    Else
        W_PC = "000"
    End If
    
    X_i = InStr(1, FullPath, "*") - 1
    If X_i <= 0 Then
        X_i = Len(Trim(FullPath)) - 4
    End If
    W_STR = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)
    
    FullPath = W_STR



    Do
        sts = BTRV(BtOpOpen, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                yn = MsgBox("���Ŏg�p���ł��I<���ԏ��v�ʂe>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function

            Case BtErrFileNotFound
                sts = ODR_TEMP1_Create()      '���v�ʂe �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ODR_TEMP1")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ODR_TEMP1")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP1_Open = False
    
End Function

Public Function ODR_TEMP1_KILL() As Integer
'********************************************************************
'*
'*              ���v�ʂe  �폜���č쐬�i�n�������j
'*
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
Dim W_STR       As String
Dim W_PC        As String
Dim X_i         As Long
Dim X_j         As Long

    ODR_TEMP1_KILL = True
                                            '���v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_TEMP1_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP1]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    c = Space(255)
    If GetComputerNameA(c, 255) <> 0 Then
        W_PC = Left(c, InStr(c, vbNullChar) - 1)
    Else
        W_PC = "000"
    End If
    
    X_i = InStr(1, FullPath, "*") - 1
    If X_i <= 0 Then
        X_i = Len(Trim(FullPath)) - 4
    End If

    W_STR = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)

    FullPath = W_STR
    
    Kill FullPath
    
    ODR_TEMP1_KILL = False
    
End Function

Public Function ODR_TEMP1_GET(SM As String, JB As String, NG As String, _
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

    ODR_TEMP1_GET = True
    
    Call UniCode_Conv(K0_ODR_TEMP1.SHIMUKE, SM)
    Call UniCode_Conv(K0_ODR_TEMP1.JGYOBU, JB)
    Call UniCode_Conv(K0_ODR_TEMP1.NAIGAI, NG)
    Call UniCode_Conv(K0_ODR_TEMP1.HIN_GAI, HG)
    Call UniCode_Conv(K0_ODR_TEMP1.ORDER_NO, OD)
    Call UniCode_Conv(K0_ODR_TEMP1.BUN_NO, BN)
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       '���R�[�h����
                'Beep
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                yn = MsgBox("���Ŏg�p���ł��I<���ԏ��v�ʂe>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_TEMP1")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP1_GET = False

End Function

Public Sub ODR_TEMP1_CLR()
    
    Call UniCode_Conv(ODR_TP1_R.KAITO_DT, "")
    Call UniCode_Conv(ODR_TP1_R.CYUMON_DT, "")
    Call UniCode_Conv(ODR_TP1_R.USE_YM, "")
    Call UniCode_Conv(ODR_TP1_R.SHIMUKE, "")
    Call UniCode_Conv(ODR_TP1_R.JGYOBU, "")
    Call UniCode_Conv(ODR_TP1_R.NAIGAI, "")
    Call UniCode_Conv(ODR_TP1_R.HIN_GAI, "")
    Call UniCode_Conv(ODR_TP1_R.ORDER_NO, "")
    Call UniCode_Conv(ODR_TP1_R.INS_NO, "")
    Call UniCode_Conv(ODR_TP1_R.BUN_NO, "")
    
    Call UniCode_Conv(ODR_TP1_R.KO_JGYOBU, "")
    Call UniCode_Conv(ODR_TP1_R.KO_NAIGAI, "")
    Call UniCode_Conv(ODR_TP1_R.KO_HIN_GAI, "")
    Call UniCode_Conv(ODR_TP1_R.KO_SYUBETSU, "")
    Call UniCode_Conv(ODR_TP1_R.KO_QTY, String(UBound(ODR_TP1_R.KO_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_TP1_R.OK_DT, "")
    Call UniCode_Conv(ODR_TP1_R.KAN_KB, "1")
    Call UniCode_Conv(ODR_TP1_R.REQ_QTY, String(UBound(ODR_TP1_R.REQ_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_TP1_R.USE_QTY, String(UBound(ODR_TP1_R.USE_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_TP1_R.NED_QTY, String(UBound(ODR_TP1_R.NED_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_TP1_R.FUSOKU_QTY, String(UBound(ODR_TP1_R.FUSOKU_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_TP1_R.UPDT_DT, "")
    
    Call UniCode_Conv(ODR_TP1_R.UPDT_TM, "")
    Call UniCode_Conv(ODR_TP1_R.FILLER, "")
    
End Sub

