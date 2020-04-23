Attribute VB_Name = "P_SEISAN_GK"
Option Explicit

'********************************************************************
'*
'*              ���Y���і��׏W�v�ް�  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const P_SEISAN_GK_ID$ = "P_SEISAN_GK"

'�y�[�W�T�C�Y
Private Const P_SEISAN_GK_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SEISAN_GK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
Private Type UCHIWAKE_TBL_Tag
    KIN(0 To 10)            As Byte         '���������z
End Type

'���R�[�h��`
Public Type P_SEISAN_GK_REC_Tag
    
    TORI_KBN(0 To 0)        As Byte         '�����敪
    TORI_CODE(0 To 4)       As Byte         '����溰��
    
    UCHIWAKE_TBL(0 To 9)    As UCHIWAKE_TBL_Tag

    CNT(0 To 10)            As Byte         '����
    QTY(0 To 10)            As Byte         '����
    KAZEI(0 To 10)          As Byte         '�ېőΏۊz

End Type
'�f�[�^�E�o�b�t�@
Public P_SEISAN_GK_REC      As P_SEISAN_GK_REC_Tag

'�L�[��`
Public Type KEY0_P_SEISAN_GK                '�j�d�x�O
    TORI_CODE(0 To 4)       As Byte         '����溰��
End Type
    
Public Type KEY1_P_SEISAN_GK                '�j�d�x�P
    TORI_KBN(0 To 0)        As Byte         '�����敪
    TORI_CODE(0 To 4)       As Byte         '����溰��
End Type
    
'�L�[�E�f�[�^
Public K0_P_SEISAN_GK       As KEY0_P_SEISAN_GK
Public K1_P_SEISAN_GK       As KEY1_P_SEISAN_GK

Type P_SEISAN_GK_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_SEISAN_GK_Speck   As P_SEISAN_GK_FSpeck
Private Function P_SEISAN_GK_Create() As Integer
'********************************************************************
'*
'*              ���Y���і��׏W�v�ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim sBuffer         As String * 255
Dim com             As String


Dim Ret             As Integer


    P_SEISAN_GK_Create = True
                                            '���Y���і��׏W�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SEISAN_GK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_GK]�ǂݍ��݃G���[")
        Exit Function
    End If

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)

    P_SEISAN_GK_Speck.fs.recoleng = Len(P_SEISAN_GK_REC)    ' ���R�[�h��
    P_SEISAN_GK_Speck.fs.PageSize = P_SEISAN_GK_PG_SIZ      ' �y�[�W�T�C�Y
    P_SEISAN_GK_Speck.fs.idexnumb = 2                       ' �C���f�b�N�X��
    P_SEISAN_GK_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    P_SEISAN_GK_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_SEISAN_GK_Speck.ks0.keypos = 2                        ' �L�[�|�W�V����
    P_SEISAN_GK_Speck.ks0.keyleng = 4                       ' �L�[��
    P_SEISAN_GK_Speck.ks0.keyflag = BtKfExt                 ' �L�[�t���O
    P_SEISAN_GK_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SEISAN_GK_Speck.ks0.reserve = &H0                     ' �\��ς�
    
    
    
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    P_SEISAN_GK_Speck.ks1.keypos = 1                        ' �L�[�|�W�V����
    P_SEISAN_GK_Speck.ks1.keyleng = 1                       ' �L�[��
    P_SEISAN_GK_Speck.ks1.keyflag = BtKfExt + BtKfSeg       ' �L�[�t���O
    P_SEISAN_GK_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SEISAN_GK_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    P_SEISAN_GK_Speck.ks2.keypos = 2                        ' �L�[�|�W�V����
    P_SEISAN_GK_Speck.ks2.keyleng = 5                       ' �L�[��
    P_SEISAN_GK_Speck.ks2.keyflag = BtKfExt                 ' �L�[�t���O
    P_SEISAN_GK_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SEISAN_GK_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    
    
    '--------------------------------------------------- �L�[�P ��
    
    
    
    sts = BTRV(BtOpCreate, P_SEISAN_GK_POS, P_SEISAN_GK_Speck, Len(P_SEISAN_GK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���Y���і��׏W�v�ް�")
        Exit Function
    End If
    
    P_SEISAN_GK_Create = False

End Function

Public Function P_SEISAN_GK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���Y���і��׏W�v�ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim sBuffer         As String * 255
Dim com             As String


Dim Ret             As Integer

    P_SEISAN_GK_Open = True
                                            '���Y���і��׏W�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SEISAN_GK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_GK]�ǂݍ��݃G���[")
        Exit Function
    End If
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)

    Do
        sts = BTRV(BtOpOpen, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SEISAN_GK_Create()  '���Y���і��׏W�v�ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���Y���і��׏W�v�ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���Y���і��׏W�v�ް�")
                Exit Function
        End Select
    Loop
    
    P_SEISAN_GK_Open = False

End Function

