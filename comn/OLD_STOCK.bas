Attribute VB_Name = "OLD_STOCK"
Option Explicit
'********************************************************************
'*
'*              �i���j�I�����f�[�^  �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const OLD_STOCK_ID$ = "OLD_STOCK"

'�y�[�W�T�C�Y
Public Const OLD_STOCK_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public OLD_STOCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OLD_STOCKREC_Tag
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 12)        As Byte     '�i�ԁi�O���j
    ST_LOCATION(0 To 7)     As Byte     '�W�����ɑq��
    HOST_ZAIKO(0 To 7)      As Byte     '�������_�݌�
    POS_ZAIKO(0 To 7)       As Byte     '�o�n�r���݌�
    ST_ZAIKO(0 To 7)        As Byte     '�W���I�ԍ݌�
    
    EE1_LOCATION(0 To 7)    As Byte     '�ʒu���P
    EE1_ZAIKO(0 To 7)       As Byte     '�݌�
    EE2_LOCATION(0 To 7)    As Byte     '�ʒu���Q
    EE2_ZAIKO(0 To 7)       As Byte     '�݌�
    EE3_LOCATION(0 To 7)    As Byte     '�ʒu���R
    EE3_ZAIKO(0 To 7)       As Byte     '�݌�
    
    ETC_ZAIKO(0 To 7)       As Byte     '���̑��݌�
    CHECK_MARK(0 To 0)      As Byte     '�ƍ��}�[�N
    PRINT_YMD(0 To 7)       As Byte     '������t
    INPUT_YMD(0 To 7)       As Byte     '���͓��t
    
    SAI_QTY(0 To 8)         As Byte     '���ِ��@2004.06.29
    
    FILLER(0 To 30)         As Byte
    
End Type
'�f�[�^�E�o�b�t�@
Public OLD_STOCKREC         As OLD_STOCKREC_Tag

'�L�[��`

Type KEY0_OLD_STOCK                     '�j�d�x�O
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 12)        As Byte     '�i�ԁi�O���j
End Type


'�L�[�E�f�[�^
Public K0_OLD_STOCK         As KEY0_OLD_STOCK
Function OLD_STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �I�����f�[�^  �n�o�d�m                              *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_STOCK_Open = True
                                    '�I�����f�[�^�t���p�X�捞��
    sts = GetIni("FILE", OLD_STOCK_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI [OLD_STOCK]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_STOCK_POS, OLD_STOCKREC, Len(OLD_STOCKREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_STOCK_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "(��)�I�����f�[�^")
                Exit Function
        End Select
    Loop
    
    OLD_STOCK_Open = False

End Function
