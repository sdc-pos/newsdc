Attribute VB_Name = "tmpP_SHORDER"
Option Explicit

'********************************************************************
'*
'*              ���ޒ����ް�  �t�@�C����`
'*
'*          CREATE 2007.10.31
'********************************************************************
'�t�@�C���h�c
Public Const tmpP_SHORDER_ID$ = "tmpP_SHORDER"

'�y�[�W�T�C�Y
Private Const tmpP_SHORDER_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public tmpP_SHORDER_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type tmpP_SHORDER_REC_Tag
    
    ORDER_NO(0 To 4)        As Byte         '������
    ORDER_DT(0 To 7)        As Byte         '������
    Print_datetime(0 To 13) As Byte         '���s����
    TANTO_CODE(0 To 4)      As Byte         '�S���Һ���
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
    DELI_CODE(0 To 4)       As Byte         '�[���溰��
    ORDER_QTY(0 To 10)      As Byte         '������(9(8)V99)
    Y_NOUKI_DT(0 To 7)      As Byte         '�\��[��
    TANKA(0 To 10)          As Byte         '�����P��(9(8)V99)
    LOT(0 To 7)             As Byte         '����ۯ�
    KAN_F(0 To 0)           As Byte         '����F
    KAN_DT(0 To 7)          As Byte         '������
    BUNNOU_CNT(0 To 1)      As Byte         '���[��
    UKEIRE_QTY(0 To 10)     As Byte         '������i���v�j(9(8)V99)
    
    CANCEL_F(0 To 0)        As Byte         '��ݾ�F
    CANCEL_DATETIME(0 To 13) As Byte        '��ݾٓ���
    PRINT_F(0 To 0)         As Byte         '����������׸�
    WS_NO(0 To 9)           As Byte         '���͒[��
    G_SHIIRE_KBN(0 To 1)    As Byte         '�d���敪
    G_SYUSHI(0 To 2)        As Byte         '���x�P��
    TORI_KBN(0 To 0)        As Byte         '�����敪
    
    ANS_NOUKI_DT(0 To 7)    As Byte         '�[���񓚓�   2007.12.05
    USE_YM(0 To 5)          As Byte         '�g�p�N��     2007.12.05
    
    
    FILLER(0 To 71)         As Byte         'Filler
    
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public tmpP_SHORDER_REC     As tmpP_SHORDER_REC_Tag

'�L�[��`

Public Type KEY0_tmpP_SHORDER                       '�j�d�x�O
    ORDER_NO(0 To 4)        As Byte         '������
End Type
    
Public Type KEY1_tmpP_SHORDER                       '�j�d�x�P
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    ORDER_DT(0 To 7)        As Byte         '������
    ORDER_NO(0 To 4)        As Byte         '������
End Type
    
Public Type KEY2_tmpP_SHORDER                       '�j�d�x�Q
    WS_NO(0 To 9)           As Byte         '���͒[��
    PRINT_F(0 To 0)         As Byte         '����������׸�
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
    ORDER_NO(0 To 4)        As Byte         '������
End Type
    
Public Type KEY3_tmpP_SHORDER                       '�j�d�x�R
    KAN_F(0 To 0)           As Byte         '����F
    ORDER_DT(0 To 7)        As Byte         '������
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
End Type
    
    
Public Type KEY4_tmpP_SHORDER                       '�j�d�x�S
    KAN_F(0 To 0)           As Byte         '����F
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
    ORDER_DT(0 To 7)        As Byte         '������
End Type
    
Public Type KEY5_tmpP_SHORDER                         '�j�d�x�T    2007.12.05
    KAN_F(0 To 0)           As Byte         '����F
    Y_NOUKI_DT(0 To 7)      As Byte         '�\��[��
    ORDER_CODE(0 To 4)      As Byte         '�����溰��
End Type
    
'�L�[�E�f�[�^
Public K0_tmpP_SHORDER      As KEY0_tmpP_SHORDER
Public K1_tmpP_SHORDER      As KEY1_tmpP_SHORDER
Public K2_tmpP_SHORDER      As KEY2_tmpP_SHORDER
Public K3_tmpP_SHORDER      As KEY3_tmpP_SHORDER
Public K4_tmpP_SHORDER      As KEY4_tmpP_SHORDER
Public K5_tmpP_SHORDER      As KEY5_tmpP_SHORDER

Type tmpP_SHORDER_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����
    ks6                     As BtKeySpeck   ' �� ��߯��\����
    ks7                     As BtKeySpeck   ' �� ��߯��\����
    ks8                     As BtKeySpeck   ' �� ��߯��\����
    ks9                     As BtKeySpeck   ' �� ��߯��\����
    ks10                    As BtKeySpeck   ' �� ��߯��\����
    ks11                    As BtKeySpeck   ' �� ��߯��\����
    ks12                    As BtKeySpeck   ' �� ��߯��\����

    ks13                    As BtKeySpeck   ' �� ��߯��\����
    ks14                    As BtKeySpeck   ' �� ��߯��\����
    ks15                    As BtKeySpeck   ' �� ��߯��\����


    ks16                    As BtKeySpeck   ' �� ��߯��\����    2007.12.05
    ks17                    As BtKeySpeck   ' �� ��߯��\����    2007.12.05
    ks18                    As BtKeySpeck   ' �� ��߯��\����    2007.12.05

End Type

Private tmpP_SHORDER_Speck  As tmpP_SHORDER_FSpeck
Private Function tmpP_SHORDER_Create() As Integer
'********************************************************************
'*
'*              ���ޒ����ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim sBuffer         As String * 255
Dim com             As String


Dim Ret             As Integer


    tmpP_SHORDER_Create = True
                                            '���ޒ����ް��t���p�X�捞��
    sts = GetIni("FILE", tmpP_SHORDER_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHORDER]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


'    Ret = InStr(1, Trim(c), ".") - 1
    
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)


    tmpP_SHORDER_Speck.fs.recoleng = Len(P_SHORDER_REC)    ' ���R�[�h��
    tmpP_SHORDER_Speck.fs.PageSize = tmpP_SHORDER_PG_SIZ      ' �y�[�W�T�C�Y
    tmpP_SHORDER_Speck.fs.idexnumb = 6                     ' �C���f�b�N�X��
    tmpP_SHORDER_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    tmpP_SHORDER_Speck.fs.reserve = &H0                    ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    tmpP_SHORDER_Speck.ks0.keypos = 1                      ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks0.keyleng = 5                     ' �L�[��
    tmpP_SHORDER_Speck.ks0.keyflag = BtKfExt               ' �L�[�t���O
    tmpP_SHORDER_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks0.reserve = &H0                   ' �\��ς�
    
    
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    tmpP_SHORDER_Speck.ks1.keypos = 33                     ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks1.keyleng = 1                     ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    tmpP_SHORDER_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks1.reserve = &H0                   ' �\��ς�
    
    tmpP_SHORDER_Speck.ks2.keypos = 34                     ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks2.keyleng = 1                     ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg
    tmpP_SHORDER_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks2.reserve = &H0                   ' �\��ς�
    
    tmpP_SHORDER_Speck.ks3.keypos = 35                     ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks3.keyleng = 20                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg
    tmpP_SHORDER_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks3.reserve = &H0                   ' �\��ς�
    
    tmpP_SHORDER_Speck.ks4.keypos = 6                      ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks4.keyleng = 8                     ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg
    tmpP_SHORDER_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks4.reserve = &H0                   ' �\��ς�
    
    tmpP_SHORDER_Speck.ks5.keypos = 1                      ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks5.keyleng = 5                     ' �L�[��
    tmpP_SHORDER_Speck.ks5.keyflag = BtKfExt + BtKfChg     ' �L�[�t���O
    tmpP_SHORDER_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks5.reserve = &H0                   ' �\��ς�
    
    '--------------------------------------------------- �L�[�P ��
    
    
    
    '--------------------------------------------------- �L�[�Q ��
    tmpP_SHORDER_Speck.ks6.keypos = 141                    ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks6.keyleng = 10                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg
    tmpP_SHORDER_Speck.ks6.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks6.reserve = &H0                   ' �\��ς�
    
    tmpP_SHORDER_Speck.ks7.keypos = 140                    ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks7.keyleng = 1                     ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfSeg
    tmpP_SHORDER_Speck.ks7.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks7.reserve = &H0                   ' �\��ς�
    
    tmpP_SHORDER_Speck.ks8.keypos = 55                     ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks8.keyleng = 5                     ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfSeg
    tmpP_SHORDER_Speck.ks8.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks8.reserve = &H0                   ' �\��ς�
    
    tmpP_SHORDER_Speck.ks9.keypos = 1                      ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks9.keyleng = 5                     ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks9.keyflag = BtKfExt + BtKfChg
    tmpP_SHORDER_Speck.ks9.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks9.reserve = &H0                   ' �\��ς�
    
    '--------------------------------------------------- �L�[�Q ��
    
    '--------------------------------------------------- �L�[�R ��
    
    tmpP_SHORDER_Speck.ks10.keypos = 103                   ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks10.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    tmpP_SHORDER_Speck.ks10.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks10.reserve = &H0                  ' �\��ς�
    
    tmpP_SHORDER_Speck.ks11.keypos = 6                     ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks11.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    tmpP_SHORDER_Speck.ks11.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks11.reserve = &H0                  ' �\��ς�
    
    
    tmpP_SHORDER_Speck.ks12.keypos = 55                    ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks12.keyleng = 5                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfDup
    tmpP_SHORDER_Speck.ks12.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks12.reserve = &H0                  ' �\��ς�

    '--------------------------------------------------- �L�[�R ��
    
    
    '--------------------------------------------------- �L�[�S ��
    
    tmpP_SHORDER_Speck.ks13.keypos = 103                   ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks13.keyleng = 1                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks13.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    tmpP_SHORDER_Speck.ks13.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks13.reserve = &H0                  ' �\��ς�
    
    tmpP_SHORDER_Speck.ks14.keypos = 55                    ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks14.keyleng = 5                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    tmpP_SHORDER_Speck.ks14.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks14.reserve = &H0                  ' �\��ς�
    
    
    tmpP_SHORDER_Speck.ks15.keypos = 6                    ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks15.keyleng = 8                    ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfDup
    tmpP_SHORDER_Speck.ks15.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks15.reserve = &H0                  ' �\��ς�

    '--------------------------------------------------- �L�[�S ��
    
    
    '--------------------------------------------------- �L�[�T 2007.12.05 ��
    
    tmpP_SHORDER_Speck.ks16.keypos = 103                ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks16.keyleng = 1                 ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    tmpP_SHORDER_Speck.ks16.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks16.reserve = &H0               ' �\��ς�
    
    tmpP_SHORDER_Speck.ks17.keypos = 76                 ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks17.keyleng = 8                 ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks17.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    tmpP_SHORDER_Speck.ks17.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks17.reserve = &H0               ' �\��ς�
    
    tmpP_SHORDER_Speck.ks18.keypos = 55                 ' �L�[�|�W�V����
    tmpP_SHORDER_Speck.ks18.keyleng = 5                 ' �L�[��
                                                        ' �L�[�t���O
    tmpP_SHORDER_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfDup
    tmpP_SHORDER_Speck.ks18.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    tmpP_SHORDER_Speck.ks18.reserve = &H0               ' �\��ς�

    '--------------------------------------------------- �L�[�T 2007.12.05 ��
    
    
    
    sts = BTRV(BtOpCreate, tmpP_SHORDER_POS, tmpP_SHORDER_Speck, Len(tmpP_SHORDER_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ޒ����ް�")
        Exit Function
    End If
    
    tmpP_SHORDER_Create = False

End Function

Public Function tmpP_SHORDER_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޒ����ް�  �n�o�d�m
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


    tmpP_SHORDER_Open = True
                                            '���ޒ����ް��t���p�X�捞��
    sts = GetIni("FILE", tmpP_SHORDER_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHORDER]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


'    Ret = InStr(1, Trim(c), ".") - 1
    
    Ret = InStrRev(Trim(c), ".") - 1
    
    FullPath = Left(Trim(c), Ret) & com & Right(Trim(c), Len(Trim(c)) - Ret)



    On Error Resume Next
    Kill (FullPath)
    On Error GoTo 0


    Do
        sts = BTRV(BtOpOpen, tmpP_SHORDER_POS, tmpP_SHORDER_REC, Len(P_SHORDER_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpP_SHORDER_Create()   '���ޒ����ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpP_SHORDER_POS, tmpP_SHORDER_REC, Len(tmpP_SHORDER_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ޒ����ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޒ����ް�")
                Exit Function
        End Select
    Loop
    
    tmpP_SHORDER_Open = False

End Function

