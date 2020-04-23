Attribute VB_Name = "OLD_ZAIKO2"
Option Explicit
'********************************************************************
'*
'*              �݌Ƀf�[�^ �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const OLD_ZAIKO2_ID$ = "OLD_ZAIKO2"

'�y�[�W�T�C�Y
Public Const OLD_ZAIKO2_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public OLD_ZAIKO2_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type OLD_ZAIKO2REC_Tag
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 12)    As Byte     '�i�ԁi�O���j
    GOODS_ON(0 To 0)    As Byte     '���i���^�����i��
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
    NYUKO_DT(0 To 7)    As Byte     '���ɓ��t
    HIN_NAI(0 To 12)    As Byte     '�i�ԁi�����j
    YUKO_Z_QTY(0 To 7)  As Byte     '�L���݌ɐ�
    LOCK_F(0 To 0)      As Byte     '�r���t���O
    WEL_ID(0 To 2)      As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)      As Byte     '�g�p���v���O����
    GOODS_YMD(0 To 7)   As Byte     '���i�����t
    FILLER(0 To 46)     As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public OLD_ZAIKO2REC         As OLD_ZAIKO2REC_Tag

'�L�[��`
Type KEY0_OLD_ZAIKO2                    '�j�d�x�O
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 12)    As Byte     '�i�ԁi�O���j
    GOODS_ON(0 To 0)    As Byte     '���i���^�����i��
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
End Type


'�L�[�E�f�[�^
Public K0_OLD_ZAIKO2         As KEY0_OLD_ZAIKO2

Public Function OLD_ZAIKO2_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �݌Ƀf�[�^�@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_ZAIKO2_Open = True
                                            '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", OLD_ZAIKO2_ID, "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI [OLD_ZAIKO2]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_ZAIKO2_POS, OLD_ZAIKO2REC, Len(OLD_ZAIKO2REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌Ƀf�[�^")
                Exit Function
        End Select
    Loop
    OLD_ZAIKO2_Open = False

End Function

