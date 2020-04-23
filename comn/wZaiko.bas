Attribute VB_Name = "wZAIKO"
Option Explicit
'********************************************************************
'*
'*              �݌Ƀf�[�^(ܰ�) �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Global Const wZAIKO_ID = "wZAIKO"

'�y�[�W�T�C�Y
Global Const wZAIKO_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Global wZAIKO_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type wZAIKOREC_Tag
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
Global wZAIKOREC As wZAIKOREC_Tag

'�L�[��`
Type KEY0_wZAIKO                    '�j�d�x�O
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
End Type

Type KEY1_wZAIKO                    '�j�d�x�P
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
End Type

Type KEY2_wZAIKO                    '�j�d�x�Q
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
End Type

Type KEY3_wZAIKO                    '�j�d�x�R
    WEL_ID(0 To 1)      As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)      As Byte     '�g�p���v���O����
End Type

Type KEY4_wZAIKO                     '�j�d�x�S
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
End Type

'�L�[�E�f�[�^
Global K0_wZAIKO As KEY0_wZAIKO
Global K1_wZAIKO As KEY1_wZAIKO
Global K2_wZAIKO As KEY2_wZAIKO
Global K3_wZAIKO As KEY3_wZAIKO
Global K4_wZAIKO As KEY4_wZAIKO

Function wZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �݌Ƀf�[�^�@�n�o�d�m                                *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    wZAIKO_Open = False
                                            '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", wZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        wZAIKO_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌Ƀf�[�^(ܰ�)")
                wZAIKO_Open = True
                Exit Function
        End Select
    Loop
End Function

