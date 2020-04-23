Attribute VB_Name = "OLD_J_NYU"
Option Explicit
'********************************************************************
'*
'*              �i���j���׃`�F�b�N�f�[�^�@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const OLD_J_NYU_ID$ = "OLD_J_NYU"

'�y�[�W�T�C�Y
Public Const OLD_J_NYU_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public OLD_J_NYU_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OLD_J_NYUREC_Tag
    JGYOBU(0 To 0)      As Byte         '���ƕ��敪
    NAIGAI(0 To 0)      As Byte         '�����O
    HIN_GAI(0 To 12)    As Byte         '�i�ԁi�O���j
    JITU_QTY(0 To 7)    As Byte         '���ѐ���
    FILLER(0 To 12)     As Byte         'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public OLD_J_NYUREC     As OLD_J_NYUREC_Tag

'�L�[��`
Type KEY0_OLD_J_NYU            '�j�d�x�O
    JGYOBU(0 To 0)      As Byte         '���ƕ��敪
    NAIGAI(0 To 0)      As Byte         '�����O
    HIN_GAI(0 To 12)    As Byte         '�i�ԁi�O���j
End Type

'�L�[�E�f�[�^
Public K0_OLD_J_NYU     As KEY0_OLD_J_NYU
Public Function OLD_J_NYU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �i���j���׃`�F�b�N�f�[�^�@�n�o�d�m                  *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_J_NYU_Open = True
                                        '���׃`�F�b�N�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", OLD_J_NYU_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI [OLD_J_NYU]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_J_NYU_POS, OLD_J_NYUREC, Len(OLD_J_NYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_J_NYU_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "(��)���׃`�F�b�N�f�[�^")
                Exit Function
        End Select
    Loop

    OLD_J_NYU_Open = False

End Function


