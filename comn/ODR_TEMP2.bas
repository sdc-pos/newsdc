Attribute VB_Name = "ODR_TEMP2"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���ԁ@���v�ʂe�iWORK) �t�@�C����`              �@�@�@*
'*                                                                  *
'*          CREATE 2008.03.06                                       *
'********************************************************************
'�t�@�C���h�c
Public Const ODR_TEMP2_ID$ = "ODR_TEMP2"

'�y�[�W�T�C�Y
Private Const ODR_TEMP2_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ODR_TP2_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type ODR_TP2_R_Tag


    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    IO_KB(0 To 0)               As Byte         'io�敪
    USE_YM(0 To 5)              As Byte         '�g�p��
    ANS_NOUKI_DT(0 To 7)        As Byte         '�Ώۓ��t   YYYYMMDD    (�񓚔[���j
    ORDER_NO(0 To 4)            As Byte         '������
    ZAI_QTY(0 To 8)             As Byte         '�݌ɐ��^������         9(5)v9(2)
    MOTO_QTY(0 To 8)            As Byte         '���X�̍݌ɐ��^������ 9(5)v9(2)
    UPDT_DT(0 To 5)             As Byte         '�X�V��     YYMMDD
    UPDT_TM(0 To 3)             As Byte         '�X�V����   hhmm
    FILLER(0 To 7)              As Byte         'Filler

End Type
'�f�[�^�E�o�b�t�@
Public ODR_TP2_R            As ODR_TP2_R_Tag



'�L�[��`

Type KEY0_ODR_TEMP2                           '�j�d�x�O

    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    IO_KB(0 To 0)               As Byte         'io�敪
    USE_YM(0 To 5)              As Byte         '�g�p��
    ANS_NOUKI_DT(0 To 7)        As Byte         '�Ώۓ��t   YYYYMMDD    (�񓚔[���j
    ORDER_NO(0 To 4)            As Byte         '������
    
End Type

Type KEY1_ODR_TEMP2                           '�j�d�x�P

    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    USE_YM(0 To 5)              As Byte         '�g�p��
    IO_KB(0 To 0)               As Byte         'io�敪
    
End Type
'�L�[�E�f�[�^
Public K0_ODR_TEMP2           As KEY0_ODR_TEMP2
Public K1_ODR_TEMP2           As KEY1_ODR_TEMP2

Type ODR_TEMP2_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private ODR_TEMP2_Speck       As ODR_TEMP2_FSpeck
Private Function ODR_TEMP2_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ODR_TEMP2  �b�q�d�`�s�d                            *
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

    ODR_TEMP2_Create = True
                                            'ODR_TEMP2 �t���p�X�捞��
    sts = GetIni("FILE", ODR_TEMP2_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP2]�ǂݍ��݃G���[")
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



    ODR_TEMP2_Speck.fs.recoleng = Len(ODR_TP2_R)      ' ���R�[�h��
    ODR_TEMP2_Speck.fs.PageSize = ODR_TEMP2_PG_SIZ          ' �y�[�W�T�C�Y
    ODR_TEMP2_Speck.fs.idexnumb = 2                       ' �C���f�b�N�X��
    ODR_TEMP2_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    ODR_TEMP2_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    ODR_TEMP2_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    ODR_TEMP2_Speck.ks0.keyleng = 42                      ' �L�[��
    ODR_TEMP2_Speck.ks0.keyflag = BtKfChg + BtKfDup + BtKfExt      ' �L�[�t���O
    ODR_TEMP2_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP2_Speck.ks0.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    '--------------------------------------------------- �L�[�P ��
    ODR_TEMP2_Speck.ks1.keypos = 1                        ' �L�[�|�W�V����
    ODR_TEMP2_Speck.ks1.keyleng = 22                      ' �L�[��
    ODR_TEMP2_Speck.ks1.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt     ' �L�[�t���O
    ODR_TEMP2_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP2_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    ODR_TEMP2_Speck.ks2.keypos = 24                       ' �L�[�|�W�V����
    ODR_TEMP2_Speck.ks2.keyleng = 6                       ' �L�[��
    ODR_TEMP2_Speck.ks2.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt     ' �L�[�t���O
    ODR_TEMP2_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP2_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    ODR_TEMP2_Speck.ks3.keypos = 23                       ' �L�[�|�W�V����
    ODR_TEMP2_Speck.ks3.keyleng = 1                       ' �L�[��
    ODR_TEMP2_Speck.ks3.keyflag = BtKfChg + BtKfDup + BtKfExt      ' �L�[�t���O
    ODR_TEMP2_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP2_Speck.ks3.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�P ��
    
    

    sts = BTRV(BtOpCreate, ODR_TP2_POS, ODR_TEMP2_Speck, Len(ODR_TEMP2_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ODR_TEMP2")
        Exit Function
    End If
    
    ODR_TEMP2_Create = False

End Function

Public Function ODR_TEMP2_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ODR_TEMP2  �n�o�d�m
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

    ODR_TEMP2_Open = True
                                            '���v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_TEMP2_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP2]�ǂݍ��݃G���[")
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
        sts = BTRV(BtOpOpen, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                yn = MsgBox("���Ŏg�p���ł��I<ODR_TEMP2>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function

            Case BtErrFileNotFound
                sts = ODR_TEMP2_Create()      'ODR_TEMP2 �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ODR_TEMP2")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ODR_TEMP2")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP2_Open = False
    
End Function

Public Function ODR_TEMP2_KILL() As Integer
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

    ODR_TEMP2_KILL = True
                                            '���v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_TEMP2_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP2]�ǂݍ��݃G���[")
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
    
    ODR_TEMP2_KILL = False
    
End Function

Public Function ODR_TEMP2_GET(JB As String, NG As String, HG As String, _
                    Kb As String, DT As String, OD As String, Locked As Integer) As Integer
'           ����

'   JB      ���ƕ�
'   NG      ���O
'   HG      �q�i��
'   KB      io�敪
'   DT      �[��
'   OD      ������

'   Locked  �f�����k������
    
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    ODR_TEMP2_GET = True
    
    Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, JB)       '�q�@���ƕ�
    Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, NG)       '�q�@�����O
    Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, HG)      '�q�i��
    Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, Kb)           'io�敪
    Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, DT)       '������
    Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, OD)        '������
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       '���R�[�h����
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                yn = MsgBox("���Ŏg�p���ł��I<ODR_TEMP2>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_TEMP2")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP2_GET = False

End Function
    

Public Sub ODR_TEMP2_CLR()
    '�q�@���ƕ�
    Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, "")
    '�q�@�����O
    Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, "")
    '�q�i��
    Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, "")
    'io�敪
    Call UniCode_Conv(ODR_TP2_R.IO_KB, "")
    '�g�p��
    Call UniCode_Conv(ODR_TP2_R.USE_YM, "")
    '�[��
    Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, "")
    '������
    Call UniCode_Conv(ODR_TP2_R.ORDER_NO, "")
    '�݌ɐ�     9(5)v9(2)
    Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, String(UBound(ODR_TP2_R.ZAI_QTY) + 1, "0"))
    '�݌ɐ�     9(5)v9(2)
    Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, String(UBound(ODR_TP2_R.MOTO_QTY) + 1, "0"))
    '�X�V��     yymmdd
    Call UniCode_Conv(ODR_TP2_R.UPDT_DT, "")
    '�X�V����   hhmm
    Call UniCode_Conv(ODR_TP2_R.UPDT_TM, "")
    
    Call UniCode_Conv(ODR_TP2_R.FILLER, "")
End Sub

