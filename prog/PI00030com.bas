Attribute VB_Name = "PI00030com"
Option Explicit


Public pubBikou_1   As String   '���l�P 2007.07.20
Public pubBikou_2   As String   '���l�Q 2007.07.20
Public pubBikou_3   As String   '���l�R 2007.07.20



'---------------------------------------------- *�����p���ޒ����ް�
'�|�W�V���j���O
Public wP_SHORDER_POS       As POSBLK
'�f�[�^�E�o�b�t�@
Public wP_SHORDER_REC       As P_SHORDER_REC_Tag
'�L�[�E�f�[�^
Public K2_wP_SHORDER        As KEY2_P_SHORDER

Public GLB_SYUSHI_F     As String           '2017.11.17

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
        Call LOG_OUT(LOG_F, "SYS.INI [P_SHORDER]�ǂݍ��݃G���[")
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

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub Main()
    
Dim lngReturnValue      As Long
Dim strMyTitle          As String
Dim lngPrevHwnd         As Long
Dim lngTopHwnd          As Long
Dim lngThreadID1        As Long
Dim lngThreadID2        As Long
    
    
    
    GLB_SYUSHI_F = Trim(Command)


    ' 2�d�N���̏ꍇ�́A��O�Ɏ����Ă��Ď������g�͏I������
    strMyTitle = App.Title
    App.Title = "$" & App.Title
    lngPrevHwnd = FindWindow("ThunderRT6Main", strMyTitle)
    If lngPrevHwnd <> 0 Then
    lngTopHwnd = GetLastActivePopup(lngPrevHwnd)
    If IsIconic(lngTopHwnd) = WIN32API_TRUE Then
    lngReturnValue = ShowWindow(lngTopHwnd, SW_NORMAL)
    End If
    lngThreadID1 = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
    lngThreadID2 = GetCurrentThreadId()
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 1)
    lngReturnValue = SetForegroundWindow(lngTopHwnd)
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 0)
    Exit Sub
    End If
    App.Title = strMyTitle




    PI000301.Show
End Sub

