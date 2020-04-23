Attribute VB_Name = "ODR_TEMP3"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���ԁ@���v�ʂe�iWORK) �t�@�C����`              �@�@�@*
'*                                                                  *
'*          CREATE 2008.03.06                                       *
'********************************************************************
'�t�@�C���h�c
Public Const ODR_TEMP3_ID$ = "ODR_TEMP3"

'�y�[�W�T�C�Y
Private Const ODR_TEMP3_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ODR_TP3_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type ODR_TP3_R_Tag
    USE_YM(0 To 5)              As Byte         '�g�p��
    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    USE_QTY(0 To 10)            As Byte         '�g�p��         9(8)v9(2)
    REQ_QTY(0 To 10)            As Byte         '�K�v��         9(8)v9(2)
    ZAI_QTY(0 To 10)            As Byte         '�����݌ɐ�     9(8)v9(2)
    MAI_QTY(0 To 10)            As Byte         '�s����         9(8)v9(2)
    ODR_QTY(0 To 10)            As Byte         '������         9(8)v9(2)
    SHI_QTY(0 To 10)            As Byte         '�d���c��       9(8)v9(2)
    HANSEIHIN_QTY(0 To 10)      As Byte         '�����i��       9(8)v9(2)
    
    HANSEIHIN_USE_QTY(0 To 10)  As Byte         '�����i��       9(8)v9(2)
    
    
    
    UKE_Z_QTY(0 To 10)          As Byte         '����ς݁i�O���ȑO�j   9(8)v9(2)
    UKE_T_QTY(0 To 10)          As Byte         '����ς݁i�����j       9(8)v9(2)
    
    
    
    
    
    
    LOT_QTY(0 To 10)            As Byte         '���b�g��       9(8)v9(2)
    SECT(0 To 4)                As Byte         '�d����
    TANKA(0 To 10)              As Byte         '�����P��       9(8)V9(2)
    NOUKI(0 To 7)               As Byte         '��]�[��
    KAITO(0 To 7)               As Byte         '�񓚔[��
    ITEM_NM(0 To 39)            As Byte         '�i��
    FILLER(0 To 1)             As Byte         'Filler

End Type
'�f�[�^�E�o�b�t�@
Public ODR_TP3_R            As ODR_TP3_R_Tag



'�L�[��`

Type KEY0_ODR_TEMP3                           '�j�d�x�O

    USE_YM(0 To 5)              As Byte         '�g�p��
    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    
End Type
Type KEY1_ODR_TEMP3                           '�j�d�x�P
    
    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    
End Type

'�L�[�E�f�[�^
Public K0_ODR_TEMP3           As KEY0_ODR_TEMP3
Public K1_ODR_TEMP3           As KEY1_ODR_TEMP3

Type ODR_TEMP3_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private ODR_TEMP3_Speck       As ODR_TEMP3_FSpeck
Private Function ODR_TEMP3_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ODR_TEMP3  �b�q�d�`�s�d                            *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128
Dim W_Str       As String
Dim W_PC        As String
Dim X_i         As Long
    
    
    ODR_TEMP3_Create = True
                                            'ODR_TEMP3 �t���p�X�捞��
    sts = GetIni("FILE", ODR_TEMP3_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP3]�ǂݍ��݃G���[")
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
    W_Str = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)
    
    FullPath = W_Str
    
    ODR_TEMP3_Speck.fs.recoleng = Len(ODR_TP3_R)      ' ���R�[�h��
    ODR_TEMP3_Speck.fs.PageSize = ODR_TEMP3_PG_SIZ          ' �y�[�W�T�C�Y
    ODR_TEMP3_Speck.fs.idexnumb = 2                       ' �C���f�b�N�X��
    ODR_TEMP3_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    ODR_TEMP3_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    ODR_TEMP3_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    ODR_TEMP3_Speck.ks0.keyleng = 28                      ' �L�[��
    ODR_TEMP3_Speck.ks0.keyflag = BtKfChg + BtKfDup + BtKfExt      ' �L�[�t���O
    ODR_TEMP3_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP3_Speck.ks0.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    '--------------------------------------------------- �L�[�P ��
    ODR_TEMP3_Speck.ks1.keypos = 7                        ' �L�[�|�W�V����
    ODR_TEMP3_Speck.ks1.keyleng = 22                      ' �L�[��
    ODR_TEMP3_Speck.ks1.keyflag = BtKfChg + BtKfDup + BtKfExt      ' �L�[�t���O
    ODR_TEMP3_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_TEMP3_Speck.ks1.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�P ��
    

    sts = BTRV(BtOpCreate, ODR_TP3_POS, ODR_TEMP3_Speck, Len(ODR_TEMP3_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ODR_TEMP3")
        Exit Function
    End If
    
    ODR_TEMP3_Create = False

End Function

Public Function ODR_TEMP3_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ODR_TEMP3  �n�o�d�m
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
Dim W_Str       As String
Dim W_PC        As String
Dim X_i         As Long

    ODR_TEMP3_Open = True
                                            '���v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_TEMP3_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP3]�ǂݍ��݃G���[")
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
    W_Str = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)
    
    FullPath = W_Str



    Do
        sts = BTRV(BtOpOpen, ODR_TP3_POS, ODR_TP3_R, Len(ODR_TP3_R), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                yn = MsgBox("���Ŏg�p���ł��I<ODR_TEMP3>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function

            Case BtErrFileNotFound
                sts = ODR_TEMP3_Create()      'ODR_TEMP3 �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_TP3_POS, ODR_TP3_R, Len(ODR_TP3_R), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ODR_TEMP3")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ODR_TEMP3")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP3_Open = False
    
End Function

Public Function ODR_TEMP3_KILL() As Integer
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
Dim W_Str       As String
Dim W_PC        As String
Dim X_i         As Long
Dim X_j         As Long

    ODR_TEMP3_KILL = True
                                            '���v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_TEMP3_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP3]�ǂݍ��݃G���[")
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

    W_Str = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)

    FullPath = W_Str
    
    On Error Resume Next
    Kill FullPath
    On Error GoTo 0
    
    ODR_TEMP3_KILL = False
    
End Function

Public Function ODR_TEMP3_GET(JB As String, NG As String, HG As String, Locked As Integer) As Integer
'           ����

'   JB      ���ƕ�
'   NG      ���O
'   HG      �q�i��

'   Locked  �f�����k������
    
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    ODR_TEMP3_GET = True
    
    Call UniCode_Conv(K0_ODR_TEMP3.KO_JGYOBU, JB)       '�q�@���ƕ�
    Call UniCode_Conv(K0_ODR_TEMP3.KO_NAIGAI, NG)       '�q�@�����O
    Call UniCode_Conv(K0_ODR_TEMP3.KO_HIN_GAI, HG)      '�q�i��
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, ODR_TP3_POS, ODR_TP3_R, Len(ODR_TP3_R), K0_ODR_TEMP3, Len(K0_ODR_TEMP3), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       '���R�[�h����
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                yn = MsgBox("���Ŏg�p���ł��I<ODR_TEMP3>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_TEMP3")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP3_GET = False

End Function
    

Public Sub ODR_TEMP3_CLR()
    '�g�p��
    Call UniCode_Conv(ODR_TP3_R.USE_YM, "")
    '�q�@���ƕ�
    Call UniCode_Conv(ODR_TP3_R.KO_JGYOBU, "")
    '�q�@�����O
    Call UniCode_Conv(ODR_TP3_R.KO_NAIGAI, "")
    '�q�i��
    Call UniCode_Conv(ODR_TP3_R.KO_HIN_GAI, "")
    
    '�g�p��         9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.USE_QTY, String(UBound(ODR_TP3_R.USE_QTY) + 1, "0"))
    '�K�v��         9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.REQ_QTY, String(UBound(ODR_TP3_R.REQ_QTY) + 1, "0"))
    '�����݌ɐ�     9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.ZAI_QTY, String(UBound(ODR_TP3_R.ZAI_QTY) + 1, "0"))
    '�s����         9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.MAI_QTY, String(UBound(ODR_TP3_R.MAI_QTY) + 1, "0"))
    '������         9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.ODR_QTY, String(UBound(ODR_TP3_R.ODR_QTY) + 1, "0"))
    
    
    '����ς�(�O���܂ŕ�)    9(8)v9(2)  2008.05.21
    Call UniCode_Conv(ODR_TP3_R.UKE_Z_QTY, String(UBound(ODR_TP3_R.UKE_Z_QTY) + 1, "0"))
    '����ς�(������)    9(8)v9(2)  2008.05.21
    Call UniCode_Conv(ODR_TP3_R.UKE_T_QTY, String(UBound(ODR_TP3_R.UKE_T_QTY) + 1, "0"))
    
    
    
    '�d���c��       9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.SHI_QTY, String(UBound(ODR_TP3_R.SHI_QTY) + 1, "0"))
    '�����i��       9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.HANSEIHIN_QTY, String(UBound(ODR_TP3_R.HANSEIHIN_QTY) + 1, "0"))
    
    
    '�����i��       9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.HANSEIHIN_USE_QTY, String(UBound(ODR_TP3_R.HANSEIHIN_USE_QTY) + 1, "0"))
    
    
    '���b�g��       9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.LOT_QTY, String(UBound(ODR_TP3_R.LOT_QTY) + 1, "0"))
    
    '�d����
    Call UniCode_Conv(ODR_TP3_R.SECT, "")
    '�����P��       9(8)V9(2)
    Call UniCode_Conv(ODR_TP3_R.TANKA, String(UBound(ODR_TP3_R.TANKA) + 1, "0"))
    '��]�[��
    Call UniCode_Conv(ODR_TP3_R.NOUKI, "")
    '�񓚔[��
    Call UniCode_Conv(ODR_TP3_R.KAITO, "")
    '�i��
    Call UniCode_Conv(ODR_TP3_R.ITEM_NM, "")
    
    Call UniCode_Conv(ODR_TP3_R.FILLER, "")
End Sub

