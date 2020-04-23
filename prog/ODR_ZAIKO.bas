Attribute VB_Name = "ODR_ZAIKO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �����݌ɂe�iWORK) �t�@�C����`              �@�@�@*
'*                                                                  *
'*          CREATE 2008.08.09                                       *
'********************************************************************
'�t�@�C���h�c
Public Const ODR_ZAIKO_ID$ = "ODR_ZAIKO"

'�y�[�W�T�C�Y
Private Const ODR_ZAIKO_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public ODR_ZK_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type ODR_Z_QTY_Tag
    Z_QTY(0 To 8)               As Byte         '�����݌�
    O_QTY(0 To 8)               As Byte         '�ǉ�������
    Y_QTY(0 To 8)               As Byte         '�\������
End Type


Public Type ODR_ZK_R_Tag

    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��
    ALL_ZAI(0 To 23)            As ODR_Z_QTY_Tag            '����`�Q�S�P��
    FILLER(0 To 29)             As Byte         'Filler

End Type
'�f�[�^�E�o�b�t�@
Public ODR_ZK_R            As ODR_ZK_R_Tag



'�L�[��`

Type KEY0_ODR_ZAIKO                           '�j�d�x�O

    KO_JGYOBU(0 To 0)           As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)           As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)         As Byte         '�q�i��

End Type

'�L�[�E�f�[�^
Public K0_ODR_ZK            As KEY0_ODR_ZAIKO


Type ODR_ZAIKO_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����

End Type

Private ODR_ZAIKO_Speck       As ODR_ZAIKO_FSpeck
Private Function ODR_ZAIKO_Create() As Integer
'*******************************************************************
'*                                                                 *
'*              ODR_ZAIKO  �b�q�d�`�s�d                             *
'*                                                                 *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                          *
'*                                                                 *
'*******************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ODR_ZAIKO_Create = True
                                            'ODR_ZAIKO �t���p�X�捞��
    sts = GetIni("FILE", ODR_ZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_ZAIKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    ODR_ZAIKO_Speck.fs.recoleng = Len(ODR_ZK_R)      ' ���R�[�h��
    ODR_ZAIKO_Speck.fs.PageSize = ODR_ZAIKO_PG_SIZ          ' �y�[�W�T�C�Y
    ODR_ZAIKO_Speck.fs.idexnumb = 1                       ' �C���f�b�N�X��
    ODR_ZAIKO_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    ODR_ZAIKO_Speck.fs.reserve = &H0                      ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    ODR_ZAIKO_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    ODR_ZAIKO_Speck.ks0.keyleng = 22                      ' �L�[��
    ODR_ZAIKO_Speck.ks0.keyflag = BtKfChg + BtKfDup + BtKfExt      ' �L�[�t���O
    ODR_ZAIKO_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    ODR_ZAIKO_Speck.ks0.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    

    sts = BTRV(BtOpCreate, ODR_ZK_POS, ODR_ZAIKO_Speck, Len(ODR_ZAIKO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ODR_ZAIKO")
        Exit Function
    End If
    
    ODR_ZAIKO_Create = False

End Function

Public Function ODR_ZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ODR_ZAIKO  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim yn          As Integer
Dim c           As String * 128
Dim FullPath    As String
Dim W_STR       As String
Dim W_PC        As String
Dim X_i         As Long

    ODR_ZAIKO_Open = True
                                            '���v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_ZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_ZAIKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                yn = MsgBox("���Ŏg�p���ł��I<ODR_ZAIKO>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function

            Case BtErrFileNotFound
                sts = ODR_ZAIKO_Create()            'ODR_ZAIKO �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ODR_ZAIKO")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ODR_ZAIKO")
                Exit Function
        End Select
    Loop
    
    ODR_ZAIKO_Open = False
    
End Function

Public Function ODR_ZAIKO_KILL() As Integer
'********************************************************************
'*
'*              ���v�ʂe  �폜���č쐬�i�n�������j
'*
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
Dim W_STR       As String
Dim W_PC        As String
Dim X_i         As Long
Dim X_j         As Long

    ODR_ZAIKO_KILL = True
                                            '���v�ʂe �t���p�X�捞��
    sts = GetIni("FILE", ODR_ZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_ZAIKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)


    
    Kill FullPath
    
    ODR_ZAIKO_KILL = False
    
End Function

Public Function ODR_ZAIKO_GET(JB As String, NG As String, HG As String, _
                      Locked As Integer) As Integer
'           ����

'   JB      ���ƕ�
'   NG      ���O
'   HG      �q�i��

'   Locked  �f�����k������
    
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    ODR_ZAIKO_GET = True
    
    Call UniCode_Conv(K0_ODR_ZK.KO_JGYOBU, JB)       '�q�@���ƕ�
    Call UniCode_Conv(K0_ODR_ZK.KO_NAIGAI, NG)       '�q�@�����O
    Call UniCode_Conv(K0_ODR_ZK.KO_HIN_GAI, HG)      '�q�i��
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                yn = MsgBox("���Ŏg�p���ł��I<ODR_ZAIKO>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_ZAIKO")
                Exit Function
        End Select
    Loop
    
    ODR_ZAIKO_GET = False

End Function
    

Public Sub ODR_ZAIKO_CLR()
Dim X_i As Integer

    '�q�@���ƕ�
    Call UniCode_Conv(ODR_ZK_R.KO_JGYOBU, "")
    '�q�@�����O
    Call UniCode_Conv(ODR_ZK_R.KO_NAIGAI, "")
    '�q�i��
    Call UniCode_Conv(ODR_ZK_R.KO_HIN_GAI, "")
    
    For X_i = 0 To UBound(ODR_ZK_R.ALL_ZAI)
                                            '�݌ɐ�     9(5)v9(2)
        Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_i).Z_QTY, String(UBound(ODR_ZK_R.ALL_ZAI(X_i).Z_QTY) + 1, "0"))
                                            '������     9(5)v9(2)
        Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_i).O_QTY, String(UBound(ODR_ZK_R.ALL_ZAI(X_i).O_QTY) + 1, "0"))
                                            '�\��
        Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_i).Y_QTY, String(UBound(ODR_ZK_R.ALL_ZAI(X_i).Y_QTY) + 1, "0"))
    Next X_i
    
    Call UniCode_Conv(ODR_ZK_R.FILLER, "")
End Sub

