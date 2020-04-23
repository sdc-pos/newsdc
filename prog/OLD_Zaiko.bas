Attribute VB_Name = "OLD_ZAIKO"
Option Explicit
'********************************************************************
'*
'*              �i���j�݌Ƀf�[�^ �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const OLD_ZAIKO_ID$ = "OLD_ZAIKO"

'�y�[�W�T�C�Y
Public Const OLD_ZAIKO_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public OLD_ZAIKO_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type OLD_ZAIKOREC_Tag
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
Public OLD_ZAIKOREC     As OLD_ZAIKOREC_Tag

'�L�[��`
Type KEY0_OLD_ZAIKO                 '�j�d�x�O
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
Public K0_OLD_ZAIKO         As KEY0_OLD_ZAIKO
Public Function OLD_ZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �i���j�݌Ƀf�[�^�@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_ZAIKO_Open = True
                                            '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", OLD_ZAIKO_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI [OLD_ZAIKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_ZAIKO_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "(��)�݌Ƀf�[�^")
                Exit Function
        End Select
    Loop
    OLD_ZAIKO_Open = False

End Function

