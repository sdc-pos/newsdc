Attribute VB_Name = "P_SHURI_SUM"
Option Explicit

'********************************************************************
'*
'*              ���ޔ���W�v�ް�  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const P_SHURI_SUM_ID$ = "P_SHURI_SUM"

'�y�[�W�T�C�Y
Private Const P_SHURI_SUM_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SHURI_SUM_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
Private Type URIAGE_TBL_Tag
    URIAGE(0 To 9)          As Byte
End Type

'���R�[�h��`
Public Type P_SHURI_SUM_REC_Tag
    
    G_SYUSHI(0 To 2)        As Byte         '���x�P��
    
    TORI_KBN(0 To 0)        As Byte         '�����敪
    TOKUI_CODE(0 To 4)      As Byte         '���Ӑ溰��
    URIAGE_TBL(0 To 5)      As URIAGE_TBL_Tag

End Type
'�f�[�^�E�o�b�t�@
Public P_SHURI_SUM_REC      As P_SHURI_SUM_REC_Tag

'�L�[��`
Public Type KEY0_P_SHURI_SUM                '�j�d�x�O
    TORI_KBN(0 To 0)        As Byte         '�����敪
    TOKUI_CODE(0 To 4)      As Byte         '���Ӑ溰��
End Type
    
Public Type KEY1_P_SHURI_SUM                '�j�d�x�P
    G_SYUSHI(0 To 2)        As Byte         '���x�P��
    TOKUI_CODE(0 To 4)      As Byte         '���Ӑ溰��
End Type
    
    
'�L�[�E�f�[�^
Public K0_P_SHURI_SUM       As KEY0_P_SHURI_SUM
Public K1_P_SHURI_SUM       As KEY1_P_SHURI_SUM

Type P_SHURI_SUM_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_SHURI_SUM_Speck  As P_SHURI_SUM_FSpeck
Private Function P_SHURI_SUM_Create() As Integer
'********************************************************************
'*
'*              ���ޔ���W�v�ް�(1)  �b�q�d�`�s�d
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


    P_SHURI_SUM_Create = True
                                            '���ޔ���W�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHURI_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHURI_SUM]�ǂݍ��݃G���[")
        Exit Function
    End If

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


'    Ret = InStr(1, Trim(c), ".") - 1
    
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)



    P_SHURI_SUM_Speck.fs.recoleng = Len(P_SHURI_SUM_REC)  ' ���R�[�h��
    P_SHURI_SUM_Speck.fs.PageSize = P_SHURI_SUM_PG_SIZ    ' �y�[�W�T�C�Y
    P_SHURI_SUM_Speck.fs.idexnumb = 2                      ' �C���f�b�N�X��
    P_SHURI_SUM_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    P_SHURI_SUM_Speck.fs.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_SHURI_SUM_Speck.ks0.keypos = 4                        ' �L�[�|�W�V����
    P_SHURI_SUM_Speck.ks0.keyleng = 1                       ' �L�[��
                                                            ' �L�[�t���O
    P_SHURI_SUM_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfSeg
    P_SHURI_SUM_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SHURI_SUM_Speck.ks0.reserve = &H0                     ' �\��ς�
    
    
    P_SHURI_SUM_Speck.ks1.keypos = 5                        ' �L�[�|�W�V����
    P_SHURI_SUM_Speck.ks1.keyleng = 5                       ' �L�[��
    P_SHURI_SUM_Speck.ks1.keyflag = BtKfExt + BtKfDup       ' �L�[�t���O
    P_SHURI_SUM_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SHURI_SUM_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    P_SHURI_SUM_Speck.ks2.keypos = 1                        ' �L�[�|�W�V����
    P_SHURI_SUM_Speck.ks2.keyleng = 3                       ' �L�[��
                                                            ' �L�[�t���O
    P_SHURI_SUM_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfSeg
    P_SHURI_SUM_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SHURI_SUM_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    
    P_SHURI_SUM_Speck.ks3.keypos = 5                        ' �L�[�|�W�V����
    P_SHURI_SUM_Speck.ks3.keyleng = 5                       ' �L�[��
    P_SHURI_SUM_Speck.ks3.keyflag = BtKfExt + BtKfDup       ' �L�[�t���O
    P_SHURI_SUM_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_SHURI_SUM_Speck.ks3.reserve = &H0                     ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    
    
    sts = BTRV(BtOpCreate, P_SHURI_SUM_POS, P_SHURI_SUM_Speck, Len(P_SHURI_SUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ޔ���W�v�ް�")
        Exit Function
    End If
    
    P_SHURI_SUM_Create = False

End Function

Public Function P_SHURI_SUM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޔ���W�v�ް�  �n�o�d�m
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
    
    
    P_SHURI_SUM_Open = True
                                            '���ޔ���W�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHURI_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHURI_SUM]�ǂݍ��݃G���[")
        Exit Function
    End If
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


'    Ret = InStr(1, Trim(c), ".") - 1
    
    Ret = InStrRev(Trim(c), ".") - 1
   
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)

    Do
        sts = BTRV(BtOpOpen, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHURI_SUM_Create()  '���ޔ���W�v�ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ޔ���W�v�ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޒ����W�v�ް�")
                Exit Function
        End Select
    Loop
    
    P_SHURI_SUM_Open = False

End Function

