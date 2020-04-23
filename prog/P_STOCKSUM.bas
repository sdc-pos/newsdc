Attribute VB_Name = "P_STOCKSUM"
Option Explicit

'********************************************************************
'*
'*              ���ޒI���W�v�ް�  �t�@�C����`
'*
'*          CREATE 2006.02.15
'********************************************************************
'�t�@�C���h�c
Public Const P_STOCKSUM_ID$ = "P_STOCKSUM"

'�y�[�W�T�C�Y
Private Const P_STOCKSUM_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_STOCKSUM_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`


Public Type P_STOCKSUM_REC_Tag
    G_SYUSHI(0 To 2)            As Byte         '���x�P��
    ZEN_ZAIKO_KIN(0 To 10)      As Byte         '�O���݌ɋ��z

    NYUKO_KIN(0 To 10)          As Byte         '�������ɋ��z
    SYUKO_KIN(0 To 10)          As Byte         '�����o�ɋ��z
    ZAIKO_KIN(0 To 10)          As Byte         '���݌ɋ��z
    FILLER(0 To 16)             As Byte         '


End Type
'�f�[�^�E�o�b�t�@
Public P_STOCKSUM_REC          As P_STOCKSUM_REC_Tag

'�L�[��`
    
Public Type KEY0_P_STOCKSUM                    '�j�d�x�O
    G_SYUSHI(0 To 2)            As Byte         '���x�P��
End Type
    
    
'�L�[�E�f�[�^
Public K0_P_STOCKSUM        As KEY0_P_STOCKSUM

Type P_STOCKSUM_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_STOCKSUM_Speck       As P_STOCKSUM_FSpeck
Private Function P_STOCKSUM_Create() As Integer
'********************************************************************
'*
'*              ���ޒI���W�v�ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*      ���x���Ƀt�@�C�����𕪂���  2007.11.13
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim Ret             As Long     '2007.11.13




    P_STOCKSUM_Create = True
                                            '���ޒI���W�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_STOCKSUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_STOCKSUM]�ǂݍ��݃G���[")
        Exit Function
    End If



    '2007.11.13
'    FullPath = Trim(c)
    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - Ret)
    '2007.11.13
   
    
    P_STOCKSUM_Speck.fs.recoleng = Len(P_STOCKSUM_REC)        ' ���R�[�h��
    P_STOCKSUM_Speck.fs.PageSize = P_STOCKSUM_PG_SIZ          ' �y�[�W�T�C�Y
    P_STOCKSUM_Speck.fs.idexnumb = 1                       ' �C���f�b�N�X��
    P_STOCKSUM_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    P_STOCKSUM_Speck.fs.reserve = &H0                      ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    P_STOCKSUM_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    P_STOCKSUM_Speck.ks0.keyleng = 3                       ' �L�[��
    P_STOCKSUM_Speck.ks0.keyflag = BtKfExt                 ' �L�[�t���O
    P_STOCKSUM_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCKSUM_Speck.ks0.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    
    sts = BTRV(BtOpCreate, P_STOCKSUM_POS, P_STOCKSUM_Speck, Len(P_STOCKSUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ޒI���W�v�ް�")
        Exit Function
    End If
    
    P_STOCKSUM_Create = False

End Function

Public Function P_STOCKSUM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޒI���W�v�ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*      ���x���Ƀt�@�C�����𕪂���  2007.11.13
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim Ret             As Long     '2007.11.13


    P_STOCKSUM_Open = True
                                            '���ޒI���W�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_STOCKSUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_STOCKSUM]�ǂݍ��݃G���[")
        Exit Function
    End If
    '2007.11.13
'    FullPath = Trim(c)
    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - Ret)
    '2007.11.13

    Do
        sts = BTRV(BtOpOpen, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_STOCKSUM_Create()   '���ޒI���W�v�ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ޒI���W�v�ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޒI���W�v�ް�")
                Exit Function
        End Select
    Loop
    
    P_STOCKSUM_Open = False

End Function

