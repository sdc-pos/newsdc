Attribute VB_Name = "PI00090com"
Option Explicit


Public pubBikou_1   As String   '���l�P 2007.07.20
Public pubBikou_2   As String   '���l�Q 2007.07.20
Public pubBikou_3   As String   '���l�R 2007.07.20


'Glid�p��
Public SHORDER   As New XArrayDB

Public Const Min_Row% = 1                   '�ŏ��s��
Public Const Min_Col% = 0                   '�ŏ���
Public Const Max_Col% = 16                  '�ő��

    
Public Const colJGYOBU% = 0                 '���ƕ�
Public Const colNAIGAI% = 1                 '�����O
Public Const colHIN_GAI% = 2                '�i��
Public Const colHIN_NAME% = 3               '�i��

Public Const colSO_SUU% = 4                 '���K�v��
Public Const colTANKA% = 5                  '�d���P��

Public Const colST_LOCATION% = 6            '�W���I��

Public Const colZAIKO_QTY% = 7              '�݌ɐ�

Public Const colSHIJI_Z_QTY% = 8            '�����c

Public Const colHIKIATE_Z_QTY% = 9          '�����c

Public Const colFUSOKU_QTY% = 10            '�s����

Public Const colORDER_QTY% = 11             '������

Public Const colLOT% = 12                   'ۯĐ�

Public Const colORDER_CODE% = 13            '�d���溰��
Public Const colORDER_NAME% = 14            '�d���於

Public Const colLT% = 15                    'ذ�����

Public Const colY_NOUKI_DT% = 16            '�\��[��

'�X�e�[�V������
Public WS_NO       As String * 10

'---------------------------------------------- *�����p���ޒ����ް�
'�|�W�V���j���O
Public wP_SHORDER_POS       As POSBLK
'�f�[�^�E�o�b�t�@
Public wP_SHORDER_REC       As P_SHORDER_REC_Tag
'�L�[�E�f�[�^
Public K2_wP_SHORDER        As KEY2_P_SHORDER
Public Function wP_SHORDER_Open(Mode As Integer) As Integer
'****************************************************
'*      �u���ޒ����ް��v    �n�o�d�m����
'*
'*  ���ޒ����ް���ʃ|�C���^�łn�o�d�m����
'*  (�Ăь��ŋN�����ɂP�x�����Ăяo��)
'*  �߂�l: false       :����
'*          true        :�ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    wP_SHORDER_Open = True
                                    '���ޒ����ް��@�t���p�X�捞��
    sts = GetIni("FILE", P_SHORDER_ID, "SYS", c)
    
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHORDER]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    wP_SHORDER_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޒ����ް�")
                Exit Function
        End Select
    Loop

    wP_SHORDER_Open = False

End Function

Public Function wP_SHORDER_CLOSE() As Integer

'****************************************************
'*      �u���ޒ����ް��v    �b�k�n�r�d����
'*
'*  ���ޒ����ް���ʃ|�C���^�łb�k�n�r�d����
'*  (�Ăь��ŏI�����ɂP�x�����Ăяo��)
'*  �߂�l: false       :����
'*          true        :�ُ�
'****************************************************
Dim sts As Integer
    
    wP_SHORDER_CLOSE = True
    
    sts = BTRV(BtOpClose, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), K2_wP_SHORDER, Len(K2_wP_SHORDER), 2)
    
    Select Case sts
        Case BtNoErr, BtErrNoOpen
        Case Else
            Call File_Error(sts, BtOpClose, "���ޒ����ް�")
            Exit Function
    End Select

    wP_SHORDER_CLOSE = False

End Function

