Attribute VB_Name = "ODR3010G"
Option Explicit
'********************************************************************
'*
'*              �n�c�q�R�O�P�O�p�@���ʕϐ�
'*
'********************************************************************
Public ODR30102_Return As Integer         '�m�F��ʏI�����
Public ODR30104_Return As Integer         '�m�F��ʏI�����
Public ODR30105_Return As Integer         '�m�F��ʏI�����


Public KIBOU_DT         As String       '��]�[��


Public DIS_ITEM         As String       '�q���i�R�[�h
Public DIS_ITEM_NM      As String       '�q���i��
Public DIS_USE_QTY      As String       '�g�p����
Public DIS_MRP_QTY      As String       '�K�v��
Public DIS_ZAI_QTY      As String       '�����݌�
Public DIS_FUSOKU       As String       '�s����
Public DIS_ORDR_QTY     As String       '������
Public DIS_ZAN_QTY      As String       '�d���c

Public DIS_HANSEIHIN_QTY _
                        As String       '�����i��
Public DIS_TEI_QTY      As String       '�ݒ��}


Public DIS_LOT_QTY      As String       '���b�g��
Public DIS_SECT_CD      As String       '�d����
Public DIS_SECT_NM      As String       '�d���於
Public DIS_TANKA        As String       '�d���P��
Public DIS_KIBOU_DT     As String       '��]�[��
Public DIS_KAITO_DE     As String       '�񓚔[��
Public DIS_KEY          As String       '�j�d�x����


Public DIS_DELI_CD      As String       '�[����R�[�h
Public DIS_DELI_NM      As String       '�[���於

Public DIS_Item_Zaiko      As String    '�O�����݌�
Public DIS_ZAIKO_ODR      As String     '�݌Ɂ{������
Public DIS_ZAIKO_UKE      As String     '�݌Ɂ{�d���ϐ�

Public Key_SIMUKE       As String       '�d������
Public Key_JIGYOBU      As String       '���ƕ�
Public Key_NAIGAI       As String       '�����O
Public Key_USE_YM       As String       '�g�p���iYYYYMM)
Public Key_INS_NO       As String       '�o�^��
Public Key_HinGai      As String        '�e�i��
Public Key_ORDER_NO     As String       '�e�i�ԁ@������
Public Key_BUN_NO       As String       '���[��

Public pubBikou_1   As String           '���l�P
Public pubBikou_2   As String           '���l�Q
Public pubBikou_3   As String           '���l�R


'�O���b�h�p��`
Public ORDR_GRID   As New XArrayDB

'Public Const Col_No% = 0                '�s��

Public Const Col_ITEM% = 0              '�q���i�R�[�h
Public Const Col_ITEM_NM% = 1           '�q���i��
Public Const Col_USE_QTY% = 2           '�g�p����
Public Const Col_MRP_QTY% = 3           '�K�v��
Public Const Col_ZAI_QTY% = 4           '�����݌�
Public Const Col_FUSOKU% = 5            '�s����
Public Const Col_ORDR_QTY% = 6          '������

Public Const Col_ZAN_QTY% = 7           '�d���c

Public Const Col_HANSEIHIN_QTY% = 8     '�����i��

Public Const Col_TEI_QTY% = 9          '�ݒ��}


Public Const Col_LOT_QTY% = 10          '���b�g��
Public Const Col_KAITO_DT% = 11         '�񓚔[��
Public Const Col_KIBOU_DT% = 12         '��]�[��

Public Const Col_SECT_CD% = 13           '�d����
Public Const Col_SECT_NM% = 14          '�d���於
Public Const Col_TANKA% = 15            '�d���P��
Public Const Col_KEY% = 16              '�g�p��
Public Const Col_JIGYOBU% = 17          '���ƕ�
Public Const Col_NAIGAI% = 18           '�����O

Public Const Col_DELI_CD% = 19          '�[����
Public Const Col_DELI_NM% = 20          '�[���於

Public Const Col_Item_Zaiko% = 21       '�i�ڂl�݌ɐ�
Public Const Col_ZAIKO_ODR% = 22        '�݌Ɂ{�e���O��
Public Const Col_ZAIKO_UKE% = 23        '�݌Ɂ{�d���ϐ�

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

