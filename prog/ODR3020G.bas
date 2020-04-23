Attribute VB_Name = "ODR3020G"
Option Explicit
'********************************************************************
'*
'*              �n�c�q�R�O�Q�O�p�@���ʕϐ�
'*
'********************************************************************
Public NAIGAI_CODE()   As String * 1
Public NAIGAI_NAME()   As String

'---------------------------------------------- *���i���w�}�ް��i�e�j�ʃ|�C���^
'�|�W�V���j���O
Public wODR_BUHIN_SUII_POS  As POSBLK
'�f�[�^�E�o�b�t�@
Public wODR_BUHIN_SUII_REC  As ODR_BUHIN_SUII_REC_Tag
'�L�[�E�f�[�^
Public K0_wODR_BUHIN_SUII   As KEY0_ODR_BUHIN_SUII
Public K1_wODR_BUHIN_SUII   As KEY1_ODR_BUHIN_SUII

Public Function wODR_BUHIN_SUII_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���i���ڃf�[�^  �n�o�d�m
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


    wODR_BUHIN_SUII_Open = True
                                            '���i���ڃf�[�^�t���p�X�捞��
    sts = GetIni("FILE", ODR_BUHIN_SUII_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_SUII]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)


    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)


    Do
        sts = BTRV(BtOpOpen, wODR_BUHIN_SUII_POS, wODR_BUHIN_SUII_REC, Len(wODR_BUHIN_SUII_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "���i���ڃf�[�^")
                Exit Function
        End Select
    Loop
    
    wODR_BUHIN_SUII_Open = False

End Function

