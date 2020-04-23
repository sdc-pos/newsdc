Attribute VB_Name = "OLD_SUMZ"
Option Explicit
'********************************************************************
'*
'*              �i���j�݌ɏW�v�f�[�^�@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const OLD_SUMZ_ID$ = "OLD_SUMZ"

'�y�[�W�T�C�Y
Public Const OLD_SUMZ_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public OLD_SUMZ_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OLD_SUMZREC_Tag
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 12)        As Byte     '�i�ԁi�O���j
    ST_SOKO(0 To 1)         As Byte     '�W�����ɑq�� �q��
    ST_RETU(0 To 1)         As Byte     '             ��
    ST_REN(0 To 1)          As Byte     '             �A
    ST_DAN(0 To 1)          As Byte     '             �i
    T_Zai_Qty(0 To 7)       As Byte     '�݌ɑ���(����)
    ZEN_Zai_Qty(0 To 7)     As Byte     '�݌ɑ���(�O��)
    SYK_E_QTY(0 To 7)       As Byte     '�o�ɍςݐ�
    NYUKA_YQTY(0 To 7)      As Byte     '���ח\�萔
    HS_ZAIQTY(0 To 7)       As Byte     'νč݌ɐ�(����)
    ZEN_HS_ZAIQTY(0 To 7)   As Byte     'νč݌ɐ�(�O��)
    SAI_QTY(0 To 7)         As Byte     '���ِ�
    SUM_DT(0 To 7)          As Byte     '�W�v���t
    FILLER(0 To 8)          As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public OLD_SUMZREC As OLD_SUMZREC_Tag

'�L�[��`
Private Type KEY0_OLD_SUMZ          '�j�d�x�O
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    NAIGAI(0 To 0) As Byte          '�����O
    HIN_GAI(0 To 12) As Byte        '�i�ԁi�O���j
End Type


'�L�[�E�f�[�^
Public K0_OLD_SUMZ As KEY0_OLD_SUMZ
Function OLD_SUMZ_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �݌ɏW�v�f�[�^�@�n�o�d�m                            *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_SUMZ_Open = True
                                            '�݌ɏW�v�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", OLD_SUMZ_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI[OLD_SUMZ] �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_SUMZ_POS, OLD_SUMZREC, Len(OLD_SUMZREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_SUMZ_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�i���j�݌ɏW�v�f�[�^")
                Exit Function
        End Select
    Loop

    OLD_SUMZ_Open = False
End Function


