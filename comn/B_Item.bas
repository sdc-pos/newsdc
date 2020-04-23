Attribute VB_Name = "B_ITEM"
Option Explicit
'********************************************************************
'*
'*              ���I�i�ԊǗ��f�[�^  �t�@�C����`
'*
'*          CREATE 2013.10.17
'********************************************************************
'�t�@�C���h�c
Public Const B_ITEM_ID$ = "B_ITEM"

'�y�[�W�T�C�Y
Public Const B_ITEM_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public B_ITEM_POS         As POSBLK
'=
'====================================================================
'=          ���R�[�h�������v���V�[�W��     ( Rclr_ITEMREC )
'====================================================================
'=
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************

'���R�[�h��`
Type B_ITEMREC_Tag
    JGYOBU(0 To 0)              As Byte     '���ƕ��敪
    NAIGAI(0 To 0)              As Byte     '�����O
    HIN_GAI(0 To 19)            As Byte     '�i�ԁi�O���j
    
    B_HIN_CODE(0 To 69)         As Byte     '���I�i�Ժ���
    
    FILLER(0 To 371)            As Byte     'FILLER
    

    INS_TANTO(0 To 9)           As Byte     '�ǉ��@�S����
    Ins_DateTime(0 To 13)       As Byte     '�ǉ��@����
    UPD_TANTO(0 To 9)           As Byte     '�X�V�@�S����
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public B_ITEMREC                As B_ITEMREC_Tag

'�L�[��`

Type KEY0_B_ITEM                    '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type


Type KEY1_B_ITEM                    '�j�d�x�P
    B_HIN_CODE(0 To 69)         As Byte     '���I�i�Ժ���
End Type


'�L�[�E�f�[�^
Public K0_B_ITEM                As KEY0_B_ITEM
Public K1_B_ITEM                As KEY1_B_ITEM

Type B_ITEM_FSpeck
    fs      As BtFileSpeck          ' ̧�� ��߯��\����
    ks0     As BtKeySpeck           ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
End Type

Private B_ITEM_Speck    As B_ITEM_FSpeck

Private Function B_ITEMreate() As Integer
'********************************************************************
'*
'*              ���I�i�ԊǗ��f�[�^  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    B_ITEMreate = True
                                            '���I�i�ԊǗ��f�[�^ �t���p�X�捞��
    sts = GetIni("FILE", B_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [B_ITEM]�ǂݍ��݃G���[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    B_ITEM_Speck.fs.recoleng = Len(B_ITEMREC)   ' ���R�[�h��
    B_ITEM_Speck.fs.PageSize = B_ITEM_PG_SIZ    ' �y�[�W�T�C�Y
    B_ITEM_Speck.fs.idexnumb = 2                ' �C���f�b�N�X��
    B_ITEM_Speck.fs.fileflag = 0                ' �t�@�C���t���O
    B_ITEM_Speck.fs.reserve = &H0               ' �\��ς�
'-----------------------------------------------
                                                ' �L�[�O
    B_ITEM_Speck.ks0.keypos = 1                             ' �L�[�|�W�V����
    B_ITEM_Speck.ks0.keyleng = 1                            ' �L�[��
    B_ITEM_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg  ' �L�[�t���O
    B_ITEM_Speck.ks0.keytype = Chr(BtKtString)              ' �L�[�^�C�v
    B_ITEM_Speck.ks0.reserve = &H0                          ' �\��ς�

    B_ITEM_Speck.ks1.keypos = 2                             ' �L�[�|�W�V����
    B_ITEM_Speck.ks1.keyleng = 1                            ' �L�[��
    B_ITEM_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg  ' �L�[�t���O
    B_ITEM_Speck.ks1.keytype = Chr(BtKtString)              ' �L�[�^�C�v
    B_ITEM_Speck.ks1.reserve = &H0                          ' �\��ς�

    B_ITEM_Speck.ks2.keypos = 3                             ' �L�[�|�W�V����
    B_ITEM_Speck.ks2.keyleng = 20                           ' �L�[��
    B_ITEM_Speck.ks2.keyflag = BtKfExt + BtKfChg            ' �L�[�t���O
    B_ITEM_Speck.ks2.keytype = Chr(BtKtString)              ' �L�[�^�C�v
    B_ITEM_Speck.ks2.reserve = &H0                          ' �\��ς�
'-----------------------------------------------

'-----------------------------------------------
                                                ' �L�[�P
    B_ITEM_Speck.ks3.keypos = 23                            ' �L�[�|�W�V����
    B_ITEM_Speck.ks3.keyleng = 70                           ' �L�[��
    B_ITEM_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg  ' �L�[�t���O
    B_ITEM_Speck.ks3.keytype = Chr(BtKtString)              ' �L�[�^�C�v
    B_ITEM_Speck.ks3.reserve = &H0                          ' �\��ς�
'-----------------------------------------------



    sts = BTRV(BtOpCreate, B_ITEM_POS, B_ITEM_Speck, Len(B_ITEM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���I�i�ԊǗ��ް�")
        Exit Function
    End If

    B_ITEMreate = False

End Function

Public Function B_ITEM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���I�i�ԊǗ��f�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    B_ITEM_Open = True
                                                '���I�i�ԊǗ��f�[�^ �t���p�X�捞��
    sts = GetIni("FILE", B_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [B_ITEM]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = B_ITEMreate()             '���I�i�ԊǗ��f�[�^    �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���I�i�ԊǗ��f�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���I�i�ԊǗ��f�[�^")
                Exit Function
        End Select
    Loop

    B_ITEM_Open = False

End Function

Public Sub Rclr_B_ITEMREC()
'********************************************************************
'*
'*              ���I�i�ԊǗ��f�[�^  ���R�[�h������
'*
'********************************************************************


    Call UniCode_Conv(B_ITEMREC.JGYOBU, "")             '���ƕ��敪
    Call UniCode_Conv(B_ITEMREC.NAIGAI, "")             '�����O
    Call UniCode_Conv(B_ITEMREC.HIN_GAI, "")            '�i�ԁi�O���j


    Call UniCode_Conv(B_ITEMREC.B_HIN_CODE, "")         '���I�i��
    
    Call UniCode_Conv(B_ITEMREC.FILLER, "")

    Call UniCode_Conv(B_ITEMREC.INS_TANTO, "")          '�ǉ��S��
    Call UniCode_Conv(B_ITEMREC.Ins_DateTime, "")       '�ǉ�����

    Call UniCode_Conv(B_ITEMREC.UPD_TANTO, "")          '�X�V�S��
    Call UniCode_Conv(B_ITEMREC.UPD_DATETIME, "")       '�X�V����

End Sub
