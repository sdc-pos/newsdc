Attribute VB_Name = "tmpZAIKO"
Option Explicit
'********************************************************************
'*
'*              �݌Ƀf�[�^�i�ꎞ�f�[�^�j �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const tmpZAIKO_ID$ = "tmpZAIKO"

'�y�[�W�T�C�Y
Public Const tmpZAIKO_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public tmpZAIKO_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type tmpZAIKOREC_Tag
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.12.05 13-->20
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    GOODS_ON(0 To 0)    As Byte     '���i���^�����i��
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
    NYUKO_DT(0 To 7)    As Byte     '���ɓ��t
    '2005.12.05 13-->20
    HIN_NAI(0 To 19)    As Byte     '�i�ԁi�����j
    YUKO_Z_QTY(0 To 7)  As Byte     '�L���݌ɐ�
    LOCK_F(0 To 0)      As Byte     '�r���t���O
    WEL_ID(0 To 2)      As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)      As Byte     '�g�p���v���O����
    GOODS_YMD(0 To 7)   As Byte     '���i�����t
    
    '2005.12.05 ���ڒǉ�
    SHIIRE_CODE(0 To 4) As Byte     '�d���溰��
    SHIIRE_TANKA(0 To 10) As Byte   '�d���P��(9(8)V99)
    KEIJYO_YM(0 To 5)   As Byte     '�v��N��
    '2005.12.05 ���ڒǉ�
    
    FILLER(0 To 74)     As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public tmpZAIKOREC      As tmpZAIKOREC_Tag

'�L�[��`

Type KEY0_tmpZAIKO                    '�j�d�x�O
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    GOODS_ON(0 To 0)    As Byte     '���i���^�����i��
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
End Type

Type KEY1_tmpZAIKO                     '�j�d�x�P
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    GOODS_ON(0 To 0)    As Byte     '���i���^�����i��
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
End Type

'�L�[�E�f�[�^
Public K0_tmpZAIKO      As KEY0_tmpZAIKO
Public K1_tmpZAIKO      As KEY1_tmpZAIKO

Type tmpZAIKO_FSpeck
    fs      As BtFileSpeck          ' ̧�� ��߯��\����
    ks0     As BtKeySpeck           ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck
    ks8     As BtKeySpeck
    ks9     As BtKeySpeck
    ks10    As BtKeySpeck
    ks11    As BtKeySpeck
    ks12    As BtKeySpeck
    ks13    As BtKeySpeck
    ks14    As BtKeySpeck
    ks15    As BtKeySpeck
End Type

Private tmpZAIKO_Speck As tmpZAIKO_FSpeck
Public Function tmpZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �݌Ƀf�[�^�i�ꎞ�f�[�^�j�@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    tmpZAIKO_Open = True
                                            '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", tmpZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [tmpZAIKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                Exit Function
        End Select
    Loop
    tmpZAIKO_Open = False

End Function

