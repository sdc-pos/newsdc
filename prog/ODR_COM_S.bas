Attribute VB_Name = "ODR_COMN"
Option Explicit
'********************************************************************
'*
'*              �n�c�q�p�@���ʕϐ�&�r����
'*
'********************************************************************

Public ODR_Return       As Integer      '�m�F��ʏI�����

Public GW_SHIMEBI       As String       '�J�z���t

Public GW_TOUGETU       As String       '���ߓ����瓾�������iyyyymm)

Public GW_MAX_YYMM      As String       '�����iyyyymm)����̍ő�g�p��


Public GW_PC_NM As String               '���s�[����


Public GW_SIMUKE        As String       '�d������
Public GW_JIGYOBU       As String       '���ƕ�
Public GW_NAIGAI        As String       '�����O
Public GW_TANTO         As String       '�S����

Public GW_USE_YM        As String       '�g�p�� yyyymm

Public GW_HINGAI        As String       '�Ώەi��

Public GW_HINGAI_KO     As String       '�i�ԁ@�i�q�i�ԁj
Public GW_JIGYOBU_KO    As String       '���ƕ��i�q�i�ԁj
Public GW_NAIGAI_KO     As String       '�����O�i�q�i�ԁj


Public Type SE_JGYOBU_TBL

    SHIMUKE    As String * 2
    JGYOBU          As String * 1
    NAIGAI          As String * 1

End Type


Public SE_JGYOBU_T()       As SE_JGYOBU_TBL

Function SET_JGYOBU_T()
'
'           �d�������ð��قɾ��
'           P_CODE�t�@�C���́A�Ăъ��Open/Close
'

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer
Dim i           As Integer
    
    SET_JGYOBU_T = True
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    i = -1
    Do
        DoEvents
        
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN04_CD Then
                    Exit Do
                End If
            
                i = i + 1
                ReDim Preserve SE_JGYOBU_T(0 To i)
                SE_JGYOBU_T(i).SHIMUKE = Trim(StrConv(P_CODEREC.C_Code, vbUnicode))
                SE_JGYOBU_T(i).JGYOBU = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                SE_JGYOBU_T(i).NAIGAI = Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
                        
                        
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "�R�[�h�}�X�^")
                'Unload Me
                Exit Function
        End Select
    
        com = BtOpGetNext
    
    Loop
    

    SET_JGYOBU_T = False

End Function
