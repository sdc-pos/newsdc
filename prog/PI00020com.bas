Attribute VB_Name = "PI00020com"
Option Explicit

Public Taget_Key    As String * 8       '�X�V�Ώۂ̎w�}�[��

Public BUNNOU_CNT   As Integer          '���[��

Public Doukon_Tbl_No(0 To 19)   As String * 1

Public POS_UMU          As Boolean

Public PRI_S_TANTO      As Boolean      '���x�^�S���҈�� OFF:����Ȃ� ON:�������
Public PRI_MAIN_BCR     As Boolean      'Ҳ��ް���� OFF:����Ȃ� ON:�������
Public PRI_BIKOU_BCR    As Boolean      '���l���@OFF�F���͒l�@ON:�o��BCR
Public PRI_DOUKON       As Boolean      '���i�������@���� OFF:����Ȃ� ON:�������

Public PRI_NYUKO_IN     As Boolean      '���Ɋ�����@���� OFF:����Ȃ� ON:�������
Public PRI_INPUT_IN     As Boolean      '���͊�����@���� OFF:����Ȃ� ON:�������

Public PRI_SAGYO_DAY    As Boolean      '��Ɠ��^���ʁ^�S�� OFF:����Ȃ� ON:������� 2007.05.22
Public PRI_HINBAN_BIKOU As Boolean      '�����@�i�ԁ^���^���� OFF:����Ȃ� ON:������� 2007.05.22


Public JISEKI_TITLE     As Variant      '���ӂ̖��̃^�C�g��
Public TASEKI_TITLE     As Variant      '���ӂ̖��̃^�C�g��

Public JISSEKI_DSP      As String * 1   '2008.08.19

'---------------------------------------------- *���i���w�}�ް��i�e�j�ʃ|�C���^
'�|�W�V���j���O
Public wP_SSHIJI_O_POS  As POSBLK
'�f�[�^�E�o�b�t�@
Public wP_SSHIJI_O_REC  As P_SSHIJI_O_REC_Tag
'�L�[�E�f�[�^
Public K0_wP_SSHIJI_O   As KEY0_P_SSHIJI_O
Public K1_wP_SSHIJI_O   As KEY1_P_SSHIJI_O
'2016.01.06 �_�~�[
Private Const LAST_UPDATE_DAY$ = "([PI00020] 2016.01.06 15:30) "


' ------------------------------------------------------------------------
'       �w�肵�����x�̐��l�Ɏl�̌ܓ����܂��B
'
' @Param    dValue      �ۂߑΏۂ̔{���x���������_���B
' @Param    iDigits     �߂�l�̗L�������̐��x�B
' @Return               iDigits �ɓ��������x�̐��l�Ɏl�̌ܓ����ꂽ���l�B
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function


Sub Main()
    
Dim lngReturnValue      As Long
Dim strMyTitle          As String
Dim lngPrevHwnd         As Long
Dim lngTopHwnd          As Long
Dim lngThreadID1        As Long
Dim lngThreadID2        As Long
    
    
    


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




    PI000201.Show
End Sub

