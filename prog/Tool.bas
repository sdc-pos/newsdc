Attribute VB_Name = "Tool"
Option Explicit
'****************************************************
'*      �O���[�o����`
'*
'*
'****************************************************
Type JGYOBU_TBL                 '�L�����ƕ��e�[�u��
    CODE As String * 1
    NAME As String * 20
    COLOR As Long
End Type
Public JGYOBU_T() As JGYOBU_TBL

Public Last_JGYOBU As String    '���ݏ��������ƕ��R�[�h

Public LOG_F  As String         '���O�t�@�C������




'[ ���l���e�`�F�b�N�C�ҏW���� ]�p�萔
'                           ���@�����^�C�v�@��
Public Const CHK_EDIT% = 0              '�`�F�b�N���ҏW
Public Const EDIT_ONLY% = 1             '�ҏW�̂�
'                           ���@�����C�s�@��
Public Const NEGA_DIS% = 0              '�s��
Public Const NEGA_ENA% = 1              '��
'                           ���@�[���}���@��
Public Const ZSUP_DIS% = 0              '����
Public Const ZSUP_ENA% = 1              '�L��
'                           ���@�J���}�ҏW�@��
Public Const COMA_ENA% = 0              '�L��
Public Const COMA_DIS% = 1              '����
'[ �J�����g�t�H�[���n�[�h�R�s�[ ]�p�萔
'                           ���@�L�[�R�[�h�萔�@��
Public Const VK_LMENU = &HA4                'Alt�L�[
Public Const VK_SNAPSHOT = &H2C             'PrintScreen�L�[
'                           ���@���ް�޲�����׸ޒ萔�@��
Public Const KEYEVENTF_EXTENDEDKEY = &H1    '�L�[������
Public Const KEYEVENTF_KEYUP = &H2          '�L�[�𗣂�


'2012.12.21 HCOPY
Private Const SRCCOPY = &HCC0020        ' (DWORD) dest = source
Private Const SRCPAINT = &HEE0086       ' (DWORD) dest = source OR dest
Private Const SRCAND = &H8800C6         ' (DWORD) dest = source AND dest
Private Const SRCINVERT = &H660046      ' (DWORD) dest = source XOR dest
Private Const SRCERASE = &H440328       ' (DWORD) dest = source AND (NOT dest )
Private Const NOTSRCCOPY = &H330008     ' (DWORD) dest = (NOT source)
Private Const NOTSRCERASE = &H1100A6    ' (DWORD) dest = (NOT src) AND (NOT dest)
Private Const MERGECOPY = &HC000CA      ' (DWORD) dest = (source AND pattern)
Private Const MERGEPAINT = &HBB0226     ' (DWORD) dest = (NOT source) OR dest
Private Const PATCOPY = &HF00021        ' (DWORD) dest = pattern
Private Const PATPAINT = &HFB0A09       ' (DWORD) dest = DPSnoo
Private Const PATINVERT = &H5A0049      ' (DWORD) dest = pattern XOR dest
Private Const DSTINVERT = &H550009      ' (DWORD) dest = (NOT dest)
Private Const BLACKNESS = &H42          ' (DWORD) dest = BLACK
Private Const WHITENESS = &HFF0062      ' (DWORD) dest = WHITE

' API�̐錾(API�r���[���[���g�p)
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" _
                    (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" _
                    (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long _
                    , ByVal nWidth As Long, ByVal nHeight As Long _
                    , ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long _
                    , ByVal dwROP As Long) As Long




'[ �V�X�e���\��ϗv���R�[�h ]�p����
'Public YOIN_HENPIN          As String * 2       '�u�Ǖi�ԕi�v�̗v��
'2004Public YOIN_MAE_SOUSAI      As String * 2       '�u�O�؂葊�E�v�̗v��
'Public YOIN_SIKYU           As String * 2       '�u�x���v�̗v��
'2004Public YOIN_CHOKUSO         As String * 2       '�u�o��(����)�v�̗v��
'2004Public YOIN_CHOKU_MODOSI    As String * 2       '�u�o��(����)�̖߂��v�̗v��
'2004Public YOIN_HSP             As String * 2       '�u�o�ׁi��^�X�j�v�̗v��
'2004Public YOIN_TUK             As String * 2       '�u�o�ׁi���؁j�v�̗v��
'2004Public YOIN_SPO             As String * 2       '�u�o�ׁi�X�|�b�g�j�v�̗v��
'2004Public YOIN_HJU             As String * 2       '�u�o�ׁi��[�j�v�̗v��
'2004Public YOIN_TOK             As String * 2       '�u�o�ׁi�����j�v�̗v��
'2004Public YOIN_BOU             As String * 2       '�u�o�ׁi�f�Ձj�v�̗v��
'2004Public YOIN_SYU_HSP         As String * 2       '�u�o�ׁi��^�X�j�o�ɕ\�o�Ɂv�̗v��
'2004Public YOIN_SYU_TUK         As String * 2       '�u�o�ׁi���؁j�o�ɕ\�o�Ɂv�̗v��
'2004Public YOIN_SYU_SPO         As String * 2       '�u�o�ׁi�X�|�b�g�j�o�ɕ\�o�Ɂv�̗v��
'2004Public YOIN_SYU_HJU         As String * 2       '�u�o�ׁi��[�j�o�ɕ\�o�Ɂv�̗v��
'2004Public YOIN_SYU_TOK         As String * 2       '�u�o�ׁi�����j�o�ɕ\�o�Ɂv�̗v��
'2004Public YOIN_SYU_BOU         As String * 2       '�u�o�ׁi�f�Ձj�o�ɕ\�o�Ɂv�̗v��
'2004Public YOIN_KIN             As String * 2       '�u�o�ׁi�ً}�j�v�̗v��
'Public YOIN_NYUKA           As String * 2       '�u�ʏ���Ɂi���בq�ɂ��j�v�̗v��


Sub File_Error(sts As Integer, Opretion As Integer, file As String, Optional Mode As Integer = 1, Optional FileName As String = "")
'****************************************************
'*      �t�@�C���G���[����
'*
'*  ��  ��: �t�@�C���X�e�[�^�X
'*          �I�y���[�V�����R�[�h
'*          �t�@�C������
'*          ���[�h 1: �\���L�� 0: �\������
'*
'*  �߂�l: �Ȃ�
'*          CREATE 1997.01.09  M.Yoshizawa                          *
'****************************************************
    Dim Buf As String

    Dim c   As String * 128                     '2016.06.23
    
    Dim Ret As Integer                          '2017.12.09
    
    
    
    Ret = GetIni("FILE", FileName, "SYS", c)    '2016.06.23 --> sts --> ret 2017.12.09
    If Ret <> False Then                        '2016.06.23 --> sts --> ret 2017.12.09
        c = ""                                  '2016.06.23
    End If                                      '2016.06.23




'    Buf = "Op= " + Str$(Opretion) + " " + "sts = " + Str$(sts) + " " + file
    
    Buf = "Op= " + Str$(Opretion) + " " + "sts = " + Str$(sts) + " " + file + " " + Trim(c)
    Call LOG_OUT(LOG_F, Buf)
    
    
    If App.EXEName = "F110010" Then
        Mode = 0
        
'        F1100101.errText(0).Text = Format(Now, "YYYY/MM/DD�@HH:MM:SS") & "�ُ픭��"
'        F1100101.errText(0).Visible = True
'        F1100101.errText(1).Visible = True

'        F1100101.Label1.BackColor = vbRed
    
    Else
        If Mode = 1 Then
'            Call Bt_Error(sts, Opretion, file)
            Call Bt_Error(sts, Opretion, file, FileName)
        End If
    End If
End Sub
Sub Ctrl_Lock(F_Obj As Form)
'*****************************************************
'*�@�@�@�R���g���[���@���b�N
'*
'*�@���@���F�t�H�[���I�u�W�F�N�g
'*
'*�@�߂�l�F�Ȃ�
'*          CREATE 1999.03.16  S.Shibano
'*****************************************************
Dim i As Integer

    For i = 0 To F_Obj.Count - 1
                                    '�uEnabled�v�����µ�޼ު�Ă��H
        If TypeOf F_Obj.Controls(i) Is CommandButton Or _
           TypeOf F_Obj.Controls(i) Is ComboBox Or _
           TypeOf F_Obj.Controls(i) Is CheckBox Or _
           TypeOf F_Obj.Controls(i) Is DirListBox Or _
           TypeOf F_Obj.Controls(i) Is TextBox Or _
           TypeOf F_Obj.Controls(i) Is DriveListBox Or _
           TypeOf F_Obj.Controls(i) Is FileListBox Or _
           TypeOf F_Obj.Controls(i) Is ListBox Or _
           TypeOf F_Obj.Controls(i) Is HScrollBar Or _
           TypeOf F_Obj.Controls(i) Is VScrollBar Then
        
        
        
        
            F_Obj.Controls(i).Tag = F_Obj.Controls(i).Enabled
            F_Obj.Controls(i).Enabled = False
        End If
    
    
    Next i

End Sub

Sub Ctrl_UnLock(F_Obj As Form)
'*****************************************************
'*�@�@�@�R���g���[���@�A�����b�N
'*
'*�@���@���F�t�H�[���I�u�W�F�N�g
'*
'*�@�߂�l�F�Ȃ�
'*          CREATE 1999.03.16  S.Shibano
'*****************************************************
Dim i As Integer

    For i = 0 To F_Obj.Count - 1
                                    '�uEnabled�v�����µ�޼ު�Ă��H
        If TypeOf F_Obj.Controls(i) Is CommandButton Or _
           TypeOf F_Obj.Controls(i) Is ComboBox Or _
           TypeOf F_Obj.Controls(i) Is CheckBox Or _
           TypeOf F_Obj.Controls(i) Is DirListBox Or _
           TypeOf F_Obj.Controls(i) Is TextBox Or _
           TypeOf F_Obj.Controls(i) Is DriveListBox Or _
           TypeOf F_Obj.Controls(i) Is FileListBox Or _
           TypeOf F_Obj.Controls(i) Is ListBox Or _
           TypeOf F_Obj.Controls(i) Is HScrollBar Or _
           TypeOf F_Obj.Controls(i) Is VScrollBar Then
        
           F_Obj.Controls(i).Enabled = F_Obj.Controls(i).Tag
        End If
    Next i


End Sub

Function GetIni(Section As String, ITEM As String, NAME As String, c As String) As Integer
'****************************************************
'*      �h�m�h�t�@�C����荞�ݏ���
'*
'*  ��  ��: �Z�N�V������
'*          �A�C�e����
'*          �h�m�h�t�@�C����
'*          ��荞�ݗ̈�i�n�t�s�o�t�s�j
'*
'*  �߂�l: false ����
'*          true  �ُ�
'*          CREATE 1997.01.09  M.Yoshizawa
'****************************************************
Dim FileName        As String
Dim sts             As Long
'Dim Work(0 To 127)  As Byte
'Dim buf1            As String * 128

Dim Work(0 To 1023)  As Byte
Dim buf1            As String * 1024



Dim buf2            As String
    
    GetIni = False
    FileName = App.Path
    If Right(FileName, 1) <> "\" Then
        FileName = FileName & "\"
    End If
    FileName = FileName & NAME & ".ini"
    c = Space(Len(c))
    sts = GetPrivateProfileString(Section, ITEM, "", buf1, 1024, FileName)
    If sts = False Then
        GetIni = True
    Else
        buf2 = RTrim(buf1)
        Call UniCode_Conv(Work, buf2)
        c = StrConv(LeftB(Work, sts), vbUnicode)
    End If
End Function
Function WriteIni(Section As String, ITEM As String, NAME As String, c As String) As Integer
'****************************************************
'*      �h�m�h�t�@�C���������ݏ���
'*
'*  ��  ��: �Z�N�V������
'*          �A�C�e����
'*          �h�m�h�t�@�C����
'*          �������ݓ��e
'*
'*  �߂�l: false ����
'*          true  �ُ�
'*          CREATE 1997.02.15  M.Yoshizawa
'****************************************************
Dim FileName As String
Dim sts As Long
    
    WriteIni = False
    FileName = App.Path
    If Right(FileName, 1) <> "\" Then
        FileName = FileName & "\"
    End If
    FileName = FileName & NAME & ".ini"
    sts = WritePrivateProfileString(Section, ITEM, c, FileName)
    If sts = False Then
        WriteIni = True
    End If

End Function


Sub LOG_OUT(file As String, MSG As String)
'****************************************************
'*      ���O�t�@�C���o�͏���
'*
'*  ��  ��: ���O�t�@�C����
'*          �o�͓��e
'*
'*  �߂�l: �Ȃ�
'*          CREATE 1997.01.09  M.Yoshizawa
'****************************************************
Dim stream  As Integer                       '�t�@�C���ԍ�
Dim Buf     As String                           '�ǂݍ��݃o�b�t�@
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

    
    stream = FreeFile
    Open file For Append As stream
    prog = StrConv(App.EXEName, vbUpperCase)
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If
    
    Buf = (Date$ & " " & Time$ & " " & com & " " & prog & " " & MSG)
    
    
    Print #stream, Buf
    Close stream
End Sub

Sub UniCode_Conv(Buffer() As Byte, Unicode As String)
'****************************************************
'*      �t�m�h�b�n�c�d�ϊ�
'*
'*  ��  ��: �`�m�r�h�i�n�t�s�o�t�s�j
'*          �t�m�h�b�n�c�d
'*
'*  �߂�l: �Ȃ�
'*          CREATE 1997.01.09  M.Yoshizawa
'****************************************************
Dim TmpBuf() As Byte
Dim TmpStr As String
Dim TmpStrlen As Integer
Dim i As Integer
Dim Swork As String
                            '������
    Swork = Space(UBound(Buffer) + 1)
    TmpBuf = ""
    TmpStr = StrConv(Swork, vbFromUnicode)
    TmpStrlen = LenB(TmpStr) - 1
    TmpBuf = StrConv(Swork, vbFromUnicode)
    For i = 0 To TmpStrlen
        Buffer(i) = TmpBuf(i)
    Next i

                            '�ϊ�
    TmpBuf = ""
    TmpStr = StrConv(Unicode, vbFromUnicode)
    TmpStrlen = LenB(TmpStr) - 1
    TmpBuf = StrConv(Unicode, vbFromUnicode)
    For i = 0 To TmpStrlen
                            '�󂯎�葤�̌����𒴂����ꍇ�͐؂�̂Ă�
        If i > (UBound(Buffer)) Then
           Exit For
        End If
        
        Buffer(i) = TmpBuf(i)
    Next i
End Sub



Function Numeric_Check(Mode As Integer, Keta As Integer, Dec As Integer, NEGA As Integer, ZSUP As Integer, COMA As Integer, Buf As String, RetBuf As String) As Integer
'*****************************************************
'*�@�@�@���l���e�`�F�b�N�C�ҏW����
'*
'*�@���@���F�����^�C�v�i�O�F�`�F�b�N���ҏW
'*�@�@�@�@�@�@�@�@�@�@�@�P�F�ҏW�̂݁j
'*�@�@�@�@�@�����i�����_�C�����C�J���}�܂ށj
'*�@�@�@�@�@��������
'*�@�@�@�@�@�����C�s�i�O�F�s�C�P�F�j
'*�@�@�@�@�@�[���}���@�@�i�O�F�s�C�P�F�j
'*�@�@�@�@�@�J���}�ҏW�@�i�O�F�L��C�P�F�����j
'*�@�@�@�@�@�`�F�b�N�Ώ�
'*�@�@�@�@�@�ҏW���e
'*
'*�@�߂�l�F�����������@����
'*�@�@�@�@�@���������@�@�ُ�
'*          CREATE 1997.01.09  M.Yoshizawa
'*****************************************************
Dim Using_Value As String
Dim Using_wk As String
Dim dNum As Double
Dim iLen As Integer
Dim iSei_Len As Integer
Dim iDec_Len As Integer
Dim iDec_Pos As Integer
Dim iGW_EDIT_pos As Integer
Dim iKeta_cnt As Integer
Dim GW_EDIT_Str As String
    
On Error GoTo Error_Proc
    
    Numeric_Check = True
    RetBuf = Space(Keta)
    Using_wk = Trim(Buf)
    
    '�p�����[�^�`�F�b�N
    If Mode <> CHK_EDIT And Mode <> EDIT_ONLY Then Exit Function
    If Keta < 0 Or Dec < 0 Then Exit Function
    If NEGA <> NEGA_DIS And NEGA <> NEGA_ENA Then Exit Function
    If ZSUP <> ZSUP_DIS And ZSUP <> ZSUP_ENA Then Exit Function
    If COMA <> COMA_ENA And COMA <> COMA_DIS Then Exit Function
    
    If (IsNumeric(Using_wk) = False) Then   '���l�ȊO�G���[
        Exit Function
    End If
    
    dNum = CDbl(Using_wk)
    iDec_Pos = InStr(Using_wk, ".")         '�����_�̈ʒu�i�O�������j
    If iDec_Pos = 0 Then
        iDec_Len = 0
    Else
        iDec_Len = Len(Mid(Using_wk, iDec_Pos + 1)) '�����_�ȉ��̌���
    End If
    
    If Mode = EDIT_ONLY Then GoTo Numeric_EDIT      '*** ->���� ����� ***
    
    If NEGA = NEGA_DIS And (Sgn(dNum) < 0) Then    '�}�C�i�X�s�Ń}�C�i�X�l
        Exit Function
    End If

    If Dec < iDec_Len Then                  '�����_�ȉ��̌����I�[�o�[
        Exit Function
    End If
    
Numeric_EDIT:       '*** �ҏW�t�H�[�}�b�g�쐬 ***
    
                        '** �ҏW��̐����������`�F�b�N **
    If Keta = 0 Then        '�������w��
        Using_Value = "#0"
    Else                    '�����w��
        If Dec = 0 Then             '�����_����
            iSei_Len = Keta
        Else                        '�����_�L��
            iSei_Len = Keta - Dec - 1
        End If
        If iSei_Len <= 0 Then Exit Function     '�����������s���G���[
                    '*** �ҏW������쐬 ***
        If COMA = COMA_ENA Then                  '�J���}�L��
            If ZSUP = ZSUP_DIS Then                  '�[���T�v���X����
                GW_EDIT_Str = "0"
                If NEGA = NEGA_ENA Then
                    iSei_Len = iSei_Len - 1     '�}�C�i�X�Ȃ�1�����炷
                End If
            Else                            '�[���T�v���X
                GW_EDIT_Str = "#"
            End If
            Using_Value = "0"
            iKeta_cnt = 1
            For iGW_EDIT_pos = 1 To iSei_Len - 1
                If (iKeta_cnt Mod 3) = 0 Then
                    iGW_EDIT_pos = iGW_EDIT_pos + 1
                    If iGW_EDIT_pos < iSei_Len Then
                        Using_Value = GW_EDIT_Str & "," & Using_Value
                    End If
                Else
                    Using_Value = GW_EDIT_Str & Using_Value
                End If
                iKeta_cnt = iKeta_cnt + 1
            Next iGW_EDIT_pos
        Else                            '�J���}����
            If ZSUP = ZSUP_DIS Then          '�[���T�v���X����
                If Sgn(dNum) < 0 Then
                    Using_Value = String(iSei_Len - 1, "0") '�l���}�C�i�X�Ȃ�1�����炷
                Else
                    Using_Value = String(iSei_Len, "0")
                End If
            Else                            '�[���T�v���X
                Using_Value = String(iSei_Len - 1, "#") & "0"
            End If
        End If
    End If

    If Dec > 0 Then                 '�����_�ȉ�
        Using_Value = Using_Value & "." & String(Dec, "0")
    End If
    
    iLen = Len(Using_Value)
    If Keta = 0 Then        '�������w��
        RetBuf = Format(dNum, Using_Value)
    Else                    '�����w��
        If ZSUP = ZSUP_DIS Then      '�[���T�v���X�����Ł`
            '�J���}�L�� & �}�C�i�X�� ���H
            '�J���}���� & �}�C�i�X�l �Ȃ�1�����₷
            If (COMA = COMA_ENA And NEGA = NEGA_ENA) Or _
               (COMA = COMA_DIS And Sgn(dNum) < 0) Then
                iLen = iLen + 1
            End If
        End If
        If iLen <> Keta Then Exit Function      '->�ҏW�����s��v
        Using_wk = Format(dNum, Using_Value)
        iLen = Len(Using_wk)
        Select Case iLen            '�ҏW�㌅��
          Case Keta
            RetBuf = Using_wk
          Case Is < Keta
            RetBuf = Space(Keta - iLen) & Using_wk
          Case Else                     '�����I�[�o�[
            Exit Function
        End Select
    End If
    
    Numeric_Check = False
    
Exit Function

Error_Proc:

    Numeric_Check = True

End Function


Function JGYOB_TB_Set(Optional JGYOBU As Integer = 0) As Integer
'****************************************************
'*      ���ƕ��e�[�u���Z�b�g
'*
'*  �߂�l: false ����
'*          true  �ُ�
'*          CREATE 1997.07.05  S.Shibano
'****************************************************
Dim c   As String
Dim i   As Long
Dim j   As Integer

    JGYOB_TB_Set = False

'    For i = 0 To UBound(JGYOBU_T)
'        JGYOBU_T(i).Code = " "
'        JGYOBU_T(i).NAME = "                    "
'    Next i

                                '���ƕ���荞��
    i = 0
    j = 0
    Do
        If GetIni("JIGYOBU", "code" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
            Call LOG_OUT(LOG_F, "[SYS.INI] [JIGYOBU] [CODE] READ ERROR")
            JGYOB_TB_Set = True
            Exit Function
        End If
        If RTrim(c) = "0" Then
            Exit Do
        End If

        If JGYOBU = 1 And _
            RTrim(c) = SHIZAI Then
            '���ނ𖳎�
        Else
            ReDim Preserve JGYOBU_T(j)

            JGYOBU_T(j).CODE = RTrim(c)
            If GetIni("JIGYOBU", "name" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
                Call LOG_OUT(LOG_F, "[SYS.INI] [JIGYOBU] [NAME] READ ERROR")
                JGYOB_TB_Set = True
                Exit Function
            End If
            JGYOBU_T(j).NAME = RTrim(c)

            If GetIni("JIGYOBU", "color" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
                Call LOG_OUT(LOG_F, "[SYS.INI] [JIGYOBU] [COLOR] READ ERROR")
                JGYOB_TB_Set = True
                Exit Function
            End If
            JGYOBU_T(j).COLOR = CLng(RTrim(c))
            j = j + 1
        End If
        i = i + 1
    Loop
                                
                                '�f�t�H���g���ƕ���荞��
    If Trim(Last_JGYOBU) = "" Then
        If GetIni("JIGYOBU", "DEF_NO", "SYS", c) Then
            Call LOG_OUT(LOG_F, "[SYS.INI] [JIGYOBU] [DEF_NO] READ ERROR")
            JGYOB_TB_Set = True
            Exit Function
        End If
        Last_JGYOBU = RTrim(c)
    End If

End Function

Public Sub Data_Select(In_Dat As String, Get_Pos As Integer, Max_Pos As Integer, Out_Dat As String)
'****************************************************
'*      �f�[�^�̐؂�o��
'*�@���@���F�؂�o�����f�[�^(","��؂�̃f�[�^)
'*�@�@�@�@�@�؂�o���|�W�V����
'*�@�@�@�@�@�ő��
'*�@�@�@�@�@�؂�o���ꂽ�f�[�^
'*
'*  �߂�l: �Ȃ�
'*          CREATE 2001.04.10  M.Yoshizawa
'****************************************************

Dim i           As Integer
Dim Start_Pos   As Integer
Dim End_Pos     As Integer

    Out_Dat = ""

    Start_Pos = 1
    For i = 1 To Max_Pos
        End_Pos = InStr(Start_Pos, In_Dat, ",")
        If End_Pos = 0 And i <> Max_Pos Then
            Exit Sub
        End If
    
        If Get_Pos = i Then
            If End_Pos > Start_Pos Then
                Out_Dat = Mid(In_Dat, Start_Pos, End_Pos - Start_Pos)
            Else
                Out_Dat = Mid(In_Dat, Start_Pos)
            End If
            If Out_Dat = "NON" Then
                Out_Dat = ""
            End If
            Exit Sub
        End If
        Start_Pos = End_Pos + 1
    Next i

End Sub

'Public Function SYSTEM_YOIN_Set() As Integer
''****************************************************
''*      �V�X�e���\��ϗv���̎捞��
''*
''*  ���� :  �Ȃ�
''*  �߂�l: false       ����
''*          SYS_ERR     �p���ł��Ȃ��ُ�
''****************************************************
'Dim c As String
'
'    SYSTEM_YOIN_Set = SYS_ERR
'
'
'
'                                        '�u�ʏ���ׁv�̗v��
'    If GetIni("YOIN", "YOIN_TU_NYUKA", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_TU_NYUKA] READ ERROR")
'        Exit Function
'    End If
'    YOIN_TU_NYUKA = Trim(c)
'                                        '�u�O�؂���ׁv�̗v��
'    If GetIni("YOIN", "YOIN_MAEGARI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_MAEGARI] READ ERROR")
'        Exit Function
'    End If
'    YOIN_MAEGARI = Trim(c)
'                                        '�u�Ǖi�ԕi�v�̗v��
''    If GetIni("YOIN", "YOIN_HENPIN", "SYS", c) Then
''        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_HENPIN] READ ERROR")
''        Exit Function
''    End If
''    YOIN_HENPIN = Trim(c)
'                                        '�u�O�؂葊�E�v�̗v��
'    If GetIni("YOIN", "YOIN_MAE_SOUSAI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_MAE_SOUSAI] READ ERROR")
'        Exit Function
'    End If
'    YOIN_MAE_SOUSAI = Trim(c)
'                                        '�u�x���v�̗v��
''    If GetIni("YOIN", "YOIN_SIKYU", "SYS", c) Then
''        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SIKYU] READ ERROR")
''        Exit Function
''    End If
''   YOIN_SIKYU = Trim(c)
'                                        '�u�o��(����)�v�̗v��
'    If GetIni("YOIN", "YOIN_CHOKUSO", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_CHOKUSO] READ ERROR")
'        Exit Function
'    End If
'    YOIN_CHOKUSO = Trim(c)
'                                        '�u�o��(����)�߂��v�̗v��
'    If GetIni("YOIN", "YOIN_CHOKU_MODOSI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_CHOKU_MODOSI] READ ERROR")
'        Exit Function
'    End If
'    YOIN_CHOKU_MODOSI = Trim(c)
'                                        '�u�o�ׁi��^�X�j�v�̗v��
'    If GetIni("YOIN", "YOIN_HSP", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_HSP] READ ERROR")
'        Exit Function
'    End If
'    YOIN_HSP = Trim(c)
'                                        '�u�o�ׁi���؁j�v�̗v��
'    If GetIni("YOIN", "YOIN_TUK", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_TUK] READ ERROR")
'        Exit Function
'    End If
'    YOIN_TUK = Trim(c)
'                                        '�u�o�ׁi�X�|�b�g�j�v�̗v��
'    If GetIni("YOIN", "YOIN_SPO", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SPO] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SPO = Trim(c)
'                                        '�u�o�ׁi��[�j�v�̗v��
'    If GetIni("YOIN", "YOIN_HJU", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_HJU] READ ERROR")
'        Exit Function
'    End If
'    YOIN_HJU = Trim(c)
'                                        '�u�o�ׁi�����j�v�̗v��
'    If GetIni("YOIN", "YOIN_TOK", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_TOK] READ ERROR")
'        Exit Function
'    End If
'    YOIN_TOK = Trim(c)
'                                        '�u�o�ׁi�f�Ձj�v�̗v��
'    If GetIni("YOIN", "YOIN_BOU", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_BOU] READ ERROR")
'        Exit Function
'    End If
'    YOIN_BOU = Trim(c)
'                                        '�u�o�ׁi��^�X�j�o�ɕ\�o�Ɂv�̗v��
'    If GetIni("YOIN", "YOIN_SYU_HSP", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_HSP] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_HSP = Trim(c)
'                                        '�u�o�ׁi���؁j�o�ɕ\�o�Ɂv�̗v��
'    If GetIni("YOIN", "YOIN_SYU_TUK", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_TUK] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_TUK = Trim(c)
'                                        '�u�o�ׁi�X�|�b�g�j�o�ɕ\�o�Ɂv�̗v��
'    If GetIni("YOIN", "YOIN_SYU_SPO", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_SPO] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_SPO = Trim(c)
'                                        '�u�o�ׁi��[�j�o�ɕ\�o�Ɂv�̗v��
'    If GetIni("YOIN", "YOIN_SYU_HJU", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_HJU] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_HJU = Trim(c)
'                                        '�u�o�ׁi�����j�o�ɕ\�o�Ɂv�̗v��
'    If GetIni("YOIN", "YOIN_SYU_TOK", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_TOK] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_TOK = Trim(c)
'                                        '�u�o�ׁi�f�Ձj�o�ɕ\�o�Ɂv�̗v��
'    If GetIni("YOIN", "YOIN_SYU_BOU", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_BOU] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_BOU = Trim(c)
'                                        '�u�o�ׁi�ً}�j�o�ɕ\�o�Ɂv�̗v��
'    If GetIni("YOIN", "YOIN_KIN", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_KIN] READ ERROR")
'        Exit Function
'    End If
'    YOIN_KIN = Trim(c)
'                                        '�u�ʏ���Ɂi���בq�ɂ��j�v�̗v��
''    If GetIni("YOIN", "YOIN_NYUKA", "SYS", c) Then
''        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_NYUKA] READ ERROR")
''        Exit Function
''    End If
''    YOIN_NYUKA = Trim(c)
'                                        '�u�����O�U�ւ��v�̗v��
'    If GetIni("YOIN", "YOIN_FURIKAE", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE] READ ERROR")
'        Exit Function
'    End If
'    YOIN_FURIKAE = Trim(c)
'                                        '�u�����O�U�ւ����̏o�Ɂv�̗v��
'    If GetIni("YOIN", "YOIN_FURIKAE_OUT", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE_OUT] READ ERROR")
'        Exit Function
'    End If
'    YOIN_FURIKAE_OUT = Trim(c)
'                                        '�u�����O�U�ւ����̓��Ɂv�̗v��
'    If GetIni("YOIN", "YOIN_FURIKAE_IN", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE_IN] READ ERROR")
'        Exit Function
'    End If
'    YOIN_FURIKAE_IN = Trim(c)
'
'                                        '�uWEL �I�����v�̗v��
'    If GetIni("YOIN", "YOIN_WEL_TANAOROSI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_WEL_TANAOROSI] READ ERROR")
'        Exit Function
'    End If
'    YOIN_WEL_TANAOROSI = Trim(c)
'                                        '�uWEL �I�ԕ\���v�̗v��
'    If GetIni("YOIN", "YOIN_WEL_TANAHYOJI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_WEL_TANAHYOJI] READ ERROR")
'        Exit Function
'    End If
'    YOIN_WEL_TANAHYOJI = Trim(c)
'                                        '�uWEL �I�ƍ��v�̗v��
'    If GetIni("YOIN", "YOIN_WEL_TANASHOGO", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_WEL_TANASHOGO] READ ERROR")
'        Exit Function
'    End If
'    YOIN_WEL_TANASHOGO = Trim(c)
'
'
'    SYSTEM_YOIN_Set = False
'End Function



Sub Form_HCopy(obj_Pic As Object, pr_Size As Integer, pr_Orient As Integer)
'00/02/12�u�g�l�v�����p
'---------------------------------------------------------------------------
'           �J�����g�t�H�[���̃n�[�h�R�s�[
'
'�m�����nobj_Pic   �F�Ұ�ގ捞�ݗp�߸�����޼ު�āiFORM�̌����Ȃ��ʒu�ɔz�u�j
'�@�@�@�@pr_Size   �F����p���T�C�Y
'�@�@�@�@pr_Orient �F����p������
'
'�s�L�[����ɂ��āt
'�@�@�v�����X�T�^�X�W�ł̓L�[���u�����v�u�����v���܂Ƃ߂čs���邪�A
'�@�@�v�����m�s�ł́u�����v�u�����v��ʁX�ɂ��Ȃ��ƔF�����Ă���Ȃ�
'
'�s�n�[�h�R�s�[�g�p��̒��Ӂt
'�@�T�uCALL���_�Ńt�H�[�J�X������FORM����������B
'�@��U�N���b�v�{�[�h�Ɏ�荞�񂾉摜���A�s�N�`���{�b�N�X�ɓǂݍ���ň������ׁA
'�@�摜�ǂݍ��ݗp�̃s�N�`���{�b�N�X�R���g���[���������Ƃ��ēn���B
'�@�s�N�`���{�b�N�X�́AFORM��̌����Ȃ��ʒu�ɔz�u���邩�AVisible=False�ɂ���B
'
'---------------------------------------------------------------------------
Dim sngPrnRatio As Single
Dim sngPrnHeight As Single
Dim sngPrnWidth As Single
Dim sngPicPosX As Single
Dim sngPicPosY As Single
Dim sngPicRatio As Single
Dim sngPicWidth As Single
Dim sngPicHeight As Single

Dim c As String
Dim USE_Printer As String
Dim Wk_Printer As Printer

Dim Pri_Name As Printer





'�n�[�h�R�s�[�p�v�����^��I���i�V�X�e���v�����^�j
'''    If GetIni("SYSTEM", "PRINTER", "SYS", c) Then
'''        Beep
'''        MsgBox "�V�X�e���v�����^����`����Ă��܂���B", vbCritical
'''        Exit Sub
'''   End If
'''    USE_Printer = RTrim(c)


    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            USE_Printer = Pri_Name.DeviceName
            Exit For
        End If
    Next


    For Each Wk_Printer In Printers
        c = RTrim(Wk_Printer.DeviceName)
        If Wk_Printer.DeviceName = USE_Printer Then
            Set Printer = Wk_Printer
            Exit For
        End If
    Next





'�N���b�v�{�[�h���N���A
    Clipboard.Clear
'Alt�L�[������ 0-->1
    Keybd_Event VK_LMENU, 1, KEYEVENTF_EXTENDEDKEY, 0
'PrintScreen�L�[������
    Keybd_Event VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY, 0
'�L�[��������s�i�d�v�F���ꂪ��������ۼ��ެ�𔲂��閘�L�[���삪�������Ȃ��j
    DoEvents
'PrintScreen�L�[�𗣂�
    Keybd_Event VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
'Alt�L�[�𗣂�
    Keybd_Event VK_LMENU, 1, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0






'�N���b�v�{�[�h����t�H�[���̉摜���擾
    obj_Pic.Picture = Clipboard.GetData()
'�摜�̈���ʒu�ƃT�C�Y���C��
    With obj_Pic.Picture
        sngPicRatio = .Width / .Height
    End With

    With Printer
        '����p���̐ݒ�
        .PaperSize = pr_Size         '�p���T�C�Y
        .Orientation = pr_Orient     '��ɂ��Ĉ������p���̕Ӂi���ӁC�Z�Ӂj

        '����p���̐ݒ�
        sngPrnRatio = .ScaleWidth / .ScaleHeight
        sngPrnWidth = .ScaleX(.ScaleWidth, _
                              .ScaleMode, _
                              vbHimetric)
        sngPrnHeight = .ScaleY(.ScaleHeight, _
                               .ScaleMode, _
                               vbHimetric)
        If sngPicRatio > sngPrnRatio Then
            sngPicHeight = _
                .ScaleY(sngPrnWidth / sngPicRatio, _
                        vbHimetric, _
                        .ScaleMode)
            sngPicWidth = _
                .ScaleX(sngPrnWidth, _
                        vbHimetric, _
                        .ScaleMode)
        Else
            sngPicHeight = _
                .ScaleY(sngPrnHeight, _
                        vbHimetric, _
                        .ScaleMode)
            sngPicWidth = _
                .ScaleX(sngPrnHeight * sngPicRatio, _
                        vbHimetric, _
                        .ScaleMode)
        End If
        sngPicPosX = (.ScaleWidth - sngPicWidth) / 2
        sngPicPosY = (.ScaleHeight - sngPicHeight) / 2

        '�t�H�[���̉摜�����
        .PaintPicture obj_Pic.Picture, _
                      sngPicPosX, _
                      sngPicPosY, _
                      sngPicWidth, _
                      sngPicHeight
        '������I�����A������v�����^�ɓn��
        .EndDoc
    End With

'�N���b�v�{�[�h���N���A
    Clipboard.Clear


End Sub


Sub Tab_Ctrl(Sf As Integer)
'******************************************************
'*�@�@�@�^�u�R���g���[��
'*
'*�@���@���FShift  (Shift�̂�)
'*
'*�@�߂�l�F�Ȃ�
'******************************************************
Dim S_Wk As String
'
    S_Wk = ""
    If Sf = vbShiftMask Then S_Wk = "+"
    S_Wk = S_Wk & "{TAB}"
    SendKeys S_Wk           ', True

End Sub

Public Sub Moji_Cut_Proc(IN_WORD As String, OUT_WORD As String, Keta As Integer)
'-------------------------------------------------------
'
'   �w�w�蕶������؂�o���x
'     2011.05.09
'-------------------------------------------------------
Dim Work    As String
Dim i       As Integer


    For i = 1 To Len(IN_WORD)
        Work = Left(IN_WORD, i)
        Work = StrConv(Work, vbFromUnicode)
    
        If LenB(Work) > Keta Then
            Exit For
        End If
    
    Next i

    OUT_WORD = Left(IN_WORD, i - 1)

End Sub


'**********************************************************************
' @(f)
'
' �@�\�@�@�@: ������̍��[����w�肵���o�C�g�����̕������Ԃ��܂�
'
' �Ԃ�l�@�@: �w��o�C�g�����̕�����
'
' �������@�@: p_Str - IN �؂�o���ΏۂƂȂ镶����
' �@�@�@�@�@: p_Len - IN �؂�o���o�C�g��
'
' �@�\�����@: VB�W����LeftB�֐��̓o�C�g���łȂ��������ŕԂ��̂�
' �@�@�@�@�@: �������o�C�g���ŕԂ��֐����쐬
'
' ���l�@�@�@: 2�o�C�g�����������ɂȂ����ꍇ�A�S�~�����͐؂�̂�
' �@�@�@�@�@:
' �@�@�@�@�@: ���ʗ�j
' �@�@�@�@�@: f_LeftB("abcdefg",5) �� abcde '���ʂ�2�o�C�g�擾
' �@�@�@�@�@: f_LeftB("a������",5) �� a���� '���p�A�S�p����
' �@�@�@�@�@: f_LeftB("��������",5) �� ���� '2�o�C�g�����������ɂȂ�
'
'       2013.06.06
'**********************************************************************
Public Function f_MidB(p_Str As String, p_Start As Integer, p_Len As Integer) As String
    Dim stRtn As String
    
    f_MidB = ""
    
    '�w��o�C�g���Ő؂�o��
    stRtn = StrConv(MidB(StrConv(p_Str, vbFromUnicode), p_Start, p_Len), vbUnicode)
    
    '2�o�C�g�����������ɂȂ����ꍇ�A�S�~�����͐؂�̂�
    'stRtn�̕���������p_Str���ēx�؂�o���Ă݂āA���̃o�C�g���𒲂ׂ܂��B
'    If f_LenB(Mid(p_Str, p_Start, Len(stRtn))) > p_Len Then
'        f_MidB = Mid(p_Str, p_Start, Len(stRtn) - 1)
'    Else
        f_MidB = stRtn
'    End If
    
End Function
    


'**********************************************************************
' @(f)
'
' �@�\�@�@�@: ������̃o�C�g����Ԃ��܂�
'
' �Ԃ�l�@�@: ������̃o�C�g��
'
' �������@�@: p_Str - IN �����ΏۂƂȂ镶����
'
' �@�\�����@: VB�W����LenB�֐��̓o�C�g���łȂ��������ŕԂ��̂�
' �@�@�@�@�@: �������o�C�g���ŕԂ��֐����쐬
'
' ���l�@�@�@: �S�p�����P�����͂Q�o�C�g�ł�
' �@�@�@�@�@:
' �@�@�@�@�@: ���ʗ�j
' �@�@�@�@�@: f_LenB("abcdefg") �� 7
' �@�@�@�@�@: f_LenB("a����") �� 5
'
'       2013.06.06
'**********************************************************************
Public Function f_LenB(p_Str As String) As Long
    f_LenB = LenB(StrConv(p_Str, vbFromUnicode))
End Function







Sub Form_HCopy_API(F_Obj As Form)
'---------------------------------------------------------------------------
'           �J�����g�t�H�[���̃n�[�h�R�s�[
'
'       API�g�p
'   2012.12.21
'---------------------------------------------------------------------------
Dim hDesktopWindow As Long
    Dim hDCScreen As Long
    Dim Ret As Long
    Dim i As Long

    F_Obj.Move 0, 0, Screen.Width, Screen.Height
    F_Obj.AutoRedraw = True

    hDesktopWindow = GetDesktopWindow
    hDCScreen = GetWindowDC(hDesktopWindow)
    Ret = BitBlt(F_Obj.hDC, 0, 0, Screen.Width, Screen.Height, hDCScreen, 0, 0, SRCCOPY)
    Ret = ReleaseDC(hDesktopWindow, hDCScreen)

    F_Obj.PrintForm

End Sub
Sub Form_HCopy_Win7(obj_Pic As Object, pr_Size As Integer, pr_Orient As Integer)
'00/02/12�u�g�l�v�����p
'---------------------------------------------------------------------------
'           �J�����g�t�H�[���̃n�[�h�R�s�[
'
'�m�����nobj_Pic   �F�Ұ�ގ捞�ݗp�߸�����޼ު�āiFORM�̌����Ȃ��ʒu�ɔz�u�j
'�@�@�@�@pr_Size   �F����p���T�C�Y
'�@�@�@�@pr_Orient �F����p������
'
'�s�L�[����ɂ��āt
'�@�@�v�����X�T�^�X�W�ł̓L�[���u�����v�u�����v���܂Ƃ߂čs���邪�A
'�@�@�v�����m�s�ł́u�����v�u�����v��ʁX�ɂ��Ȃ��ƔF�����Ă���Ȃ�
'
'�s�n�[�h�R�s�[�g�p��̒��Ӂt
'�@�T�uCALL���_�Ńt�H�[�J�X������FORM����������B
'�@��U�N���b�v�{�[�h�Ɏ�荞�񂾉摜���A�s�N�`���{�b�N�X�ɓǂݍ���ň������ׁA
'�@�摜�ǂݍ��ݗp�̃s�N�`���{�b�N�X�R���g���[���������Ƃ��ēn���B
'�@�s�N�`���{�b�N�X�́AFORM��̌����Ȃ��ʒu�ɔz�u���邩�AVisible=False�ɂ���B
'
'---------------------------------------------------------------------------
Dim sngPrnRatio As Single
Dim sngPrnHeight As Single
Dim sngPrnWidth As Single
Dim sngPicPosX As Single
Dim sngPicPosY As Single
Dim sngPicRatio As Single
Dim sngPicWidth As Single
Dim sngPicHeight As Single

Dim c As String
Dim USE_Printer As String
Dim Wk_Printer As Printer

Dim Pri_Name As Printer





'�n�[�h�R�s�[�p�v�����^��I���i�V�X�e���v�����^�j
'''    If GetIni("SYSTEM", "PRINTER", "SYS", c) Then
'''        Beep
'''        MsgBox "�V�X�e���v�����^����`����Ă��܂���B", vbCritical
'''        Exit Sub
'''   End If
'''    USE_Printer = RTrim(c)


    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            USE_Printer = Pri_Name.DeviceName
            Exit For
        End If
    Next


    For Each Wk_Printer In Printers
        c = RTrim(Wk_Printer.DeviceName)
        If Wk_Printer.DeviceName = USE_Printer Then
            Set Printer = Wk_Printer
            Exit For
        End If
    Next





'�N���b�v�{�[�h���N���A
    Clipboard.Clear
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents
'Call LOG_OUT(LOG_F, "Clipboard.Clear")
    
'Alt�L�[������ 0-->1
    Keybd_Event VK_LMENU, 1, KEYEVENTF_EXTENDEDKEY, 0
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents
'Call LOG_OUT(LOG_F, "Keybd_Event VK_LMENU, 1, KEYEVENTF_EXTENDEDKEY, 0")

'PrintScreen�L�[������
    Keybd_Event VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY, 0
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents
'Call LOG_OUT(LOG_F, "Keybd_Event VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY, 0")

'�L�[��������s�i�d�v�F���ꂪ��������ۼ��ެ�𔲂��閘�L�[���삪�������Ȃ��j
    DoEvents
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents

'PrintScreen�L�[�𗣂�
    Keybd_Event VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents
'Call LOG_OUT(LOG_F, "Keybd_Event VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0")

'Alt�L�[�𗣂�
    Keybd_Event VK_LMENU, 1, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents

'Call LOG_OUT(LOG_F, "Keybd_Event VK_LMENU, 1, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0")


'�N���b�v�{�[�h����t�H�[���̉摜���擾
    obj_Pic.Picture = Clipboard.GetData()
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents

'Call LOG_OUT(LOG_F, "obj_Pic.Picture = Clipboard.GetData()")

'�摜�̈���ʒu�ƃT�C�Y���C��
    With obj_Pic.Picture
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents


'Call LOG_OUT(LOG_F, "With obj_Pic.Picture")


        sngPicRatio = .Width / .Height
    
    End With
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents


'Call LOG_OUT(LOG_F, "sngPicRatio = .Width / .Height")

    With Printer
        '����p���̐ݒ�
        .PaperSize = pr_Size         '�p���T�C�Y
        .Orientation = pr_Orient     '��ɂ��Ĉ������p���̕Ӂi���ӁC�Z�Ӂj

        '����p���̐ݒ�
        sngPrnRatio = .ScaleWidth / .ScaleHeight
        sngPrnWidth = .ScaleX(.ScaleWidth, _
                              .ScaleMode, _
                              vbHimetric)
        sngPrnHeight = .ScaleY(.ScaleHeight, _
                               .ScaleMode, _
                               vbHimetric)
        If sngPicRatio > sngPrnRatio Then
            sngPicHeight = _
                .ScaleY(sngPrnWidth / sngPicRatio, _
                        vbHimetric, _
                        .ScaleMode)
            sngPicWidth = _
                .ScaleX(sngPrnWidth, _
                        vbHimetric, _
                        .ScaleMode)
        Else
            sngPicHeight = _
                .ScaleY(sngPrnHeight, _
                        vbHimetric, _
                        .ScaleMode)
            sngPicWidth = _
                .ScaleX(sngPrnHeight * sngPicRatio, _
                        vbHimetric, _
                        .ScaleMode)
        End If
        sngPicPosX = (.ScaleWidth - sngPicWidth) / 2
        sngPicPosY = (.ScaleHeight - sngPicHeight) / 2
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents

'Call LOG_OUT(LOG_F, "sngPicPosY = (.ScaleHeight - sngPicHeight) / 2")

        '�t�H�[���̉摜�����
        .PaintPicture obj_Pic.Picture, _
                      sngPicPosX, _
                      sngPicPosY, _
                      sngPicWidth, _
                      sngPicHeight
        '������I�����A������v�����^�ɓn��
        .EndDoc
    End With

'�N���b�v�{�[�h���N���A
    Clipboard.Clear
    
    Sleep (500)                     '2012.12.22 DoEvents�ƃZ�b�g�Œǉ��B    Sleep�̃^�C�}�[�l�͔������K�v�I�H
    DoEvents


End Sub


Sub Form_HCopy_Win7_NEW(obj_Pic As Object, pr_Size As Integer, pr_Orient As Integer)
'00/02/12�u�g�l�v�����p
'---------------------------------------------------------------------------
'           �J�����g�t�H�[���̃n�[�h�R�s�[
'
'�m�����nobj_Pic   �F�Ұ�ގ捞�ݗp�߸�����޼ު�āiFORM�̌����Ȃ��ʒu�ɔz�u�j
'�@�@�@�@pr_Size   �F����p���T�C�Y
'�@�@�@�@pr_Orient �F����p������
'
'�s�L�[����ɂ��āt
'�@�@�v�����X�T�^�X�W�ł̓L�[���u�����v�u�����v���܂Ƃ߂čs���邪�A
'�@�@�v�����m�s�ł́u�����v�u�����v��ʁX�ɂ��Ȃ��ƔF�����Ă���Ȃ�
'
'�s�n�[�h�R�s�[�g�p��̒��Ӂt
'�@�T�uCALL���_�Ńt�H�[�J�X������FORM����������B
'�@��U�N���b�v�{�[�h�Ɏ�荞�񂾉摜���A�s�N�`���{�b�N�X�ɓǂݍ���ň������ׁA
'�@�摜�ǂݍ��ݗp�̃s�N�`���{�b�N�X�R���g���[���������Ƃ��ēn���B
'�@�s�N�`���{�b�N�X�́AFORM��̌����Ȃ��ʒu�ɔz�u���邩�AVisible=False�ɂ���B
'
'   2017.01.07 KG_TOOL ���
'---------------------------------------------------------------------------
Dim sngPrnRatio As Single
Dim sngPrnHeight As Single
Dim sngPrnWidth As Single
Dim sngPicPosX As Single
Dim sngPicPosY As Single
Dim sngPicRatio As Single
Dim sngPicWidth As Single
Dim sngPicHeight As Single

Dim c As String
Dim USE_Printer As String
Dim Wk_Printer As Printer


Dim Pri_Name As Printer


'�n�[�h�R�s�[�p�v�����^��I���i�V�X�e���v�����^�j
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            USE_Printer = Pri_Name.DeviceName
            Exit For
        End If
    Next


    For Each Wk_Printer In Printers
        c = RTrim(Wk_Printer.DeviceName)
        If Wk_Printer.DeviceName = USE_Printer Then
            Set Printer = Wk_Printer
            Exit For
        End If
    Next

'�N���b�v�{�[�h���N���A
    Clipboard.Clear
    Sleep (500)
    DoEvents
    
'Alt�L�[������
    Keybd_Event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY, 0
    Sleep (500)
    DoEvents
'PrintScreen�L�[������
    Keybd_Event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0
    Sleep (500)
'�L�[��������s�i�d�v�F���ꂪ��������ۼ��ެ�𔲂��閘�L�[���삪�������Ȃ��j
    DoEvents
'PrintScreen�L�[�𗣂�
    Keybd_Event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    Sleep (500)
    DoEvents
'Alt�L�[�𗣂�
    Keybd_Event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    Sleep (500)
    DoEvents
'�N���b�v�{�[�h����t�H�[���̉摜���擾
    obj_Pic.Picture = Clipboard.GetData()
    Sleep (500)
    DoEvents
    
'�摜�̈���ʒu�ƃT�C�Y���C��
    With obj_Pic.Picture
    Sleep (500)
    DoEvents
        If .Height = 0 Then
'            MsgBox ".Height�␳"
            '.Width = 21722
            '.Height = 14790
            sngPicRatio = 21722 / 14790
        End If
'        MsgBox "�摜�̈���ʒu�ƃT�C�Y���C�� Width<" & .Width & " Height<" & .Height & ">"
        If .Height <> 0 Then
            sngPicRatio = .Width / .Height
        End If
    End With
    
    
    Sleep (500)
    DoEvents
    With Printer
        '����p���̐ݒ�
        .PaperSize = pr_Size         '�p���T�C�Y
        .Orientation = pr_Orient     '��ɂ��Ĉ������p���̕Ӂi���ӁC�Z�Ӂj

        '����p���̐ݒ�
'        MsgBox "����p���̐ݒ�=<" & .ScaleHeight & ">"
        
        sngPrnRatio = .ScaleWidth / .ScaleHeight
        
        sngPrnWidth = .ScaleX(.ScaleWidth, _
                              .ScaleMode, _
                              vbHimetric)
        sngPrnHeight = .ScaleY(.ScaleHeight, _
                               .ScaleMode, _
                               vbHimetric)
                               
                               
        If sngPicRatio > sngPrnRatio Then
        
'            MsgBox "sngPicRatio <" & sngPicRatio & ">"
            sngPicHeight = _
                .ScaleY(sngPrnWidth / sngPicRatio, vbHimetric, .ScaleMode)
                
                
            sngPicWidth = _
                .ScaleX(sngPrnWidth, _
                        vbHimetric, _
                        .ScaleMode)
        Else
            sngPicHeight = _
                .ScaleY(sngPrnHeight, _
                        vbHimetric, _
                        .ScaleMode)
            sngPicWidth = _
                .ScaleX(sngPrnHeight * sngPicRatio, _
                        vbHimetric, _
                        .ScaleMode)
        End If
        
'        MsgBox "sngPicPosX <" & (.ScaleWidth - sngPicWidth) & ">"
        sngPicPosX = (.ScaleWidth - sngPicWidth) / 2
        
'        MsgBox "sngPicPosY <" & (.ScaleHeight - sngPicHeight) & ">"
        sngPicPosY = (.ScaleHeight - sngPicHeight) / 2
        
    Sleep (500)
    DoEvents
        
        '�t�H�[���̉摜�����
'        MsgBox "�t�H�[���̉摜�����"
        .PaintPicture obj_Pic.Picture, _
                      sngPicPosX, _
                      sngPicPosY, _
                      sngPicWidth, _
                      sngPicHeight
                      
                      
        '������I�����A������v�����^�ɓn��
'        .EndDoc
    End With
    
    Printer.EndDoc
    
    DoEvents
    

End Sub


Public Function f16sinTo10sin(ByVal str16sin As String) As String
'---------------------------------------------------------------------------
'           16�i�|�|�|��10�i�@�ϊ�
'
'   2017.09.15
'---------------------------------------------------------------------------
    
    
    
    Dim i As Long, N As Long, dbl10Sin As Double
    Const Table As String = "0123456789ABCDEF"
    '�O��̋󔒂���菜���啶���ɕϊ�
    str16sin = Trim$(UCase(str16sin))
    '�Ώە�����̃`�F�b�N
    If Len(str16sin) = 0 Or Len(str16sin) > 8 Then Exit Function
    For i = 1 To Len(str16sin)
        '������0�`F�͈͓̔����`�F�b�N
        If Mid$(str16sin, i, 1) < Chr$(48) Or Mid$(str16sin, i, 1) > Chr$(70) Then
            Exit Function
        End If
    Next i
    '�P�����Â�10�i���ɕϊ�
    For i = 1 To Len(str16sin)
        '10�i���̂����ɂȂ邩���ׂ�
        N = (InStr(Table, Mid$(str16sin, i, 1)) - 1)
        '���オ�蕪�̌v�Z�Ə��v�����߂�
        dbl10Sin = dbl10Sin * 16 + N
    Next i
    f16sinTo10sin = CStr(dbl10Sin)
End Function

Public Function f10sinTo16sin(ByVal str10sin As String) As String
'---------------------------------------------------------------------------
'           10�i�|�|�|��16�i�@�ϊ�
'
'   2017.09.15
'---------------------------------------------------------------------------
    
    Dim i        As Long, j           As Long, k As Integer
    Dim RetValue As Variant, ModValue As Variant
    Dim strSum   As String, Keta(8)   As Double
    Const Table As String = "0123456789ABCDEF"
    str10sin = Trim$(str10sin)      '�󔒂���菜��
    
    If Len(str10sin) < 1 Then
        Exit Function
    End If
    
    
    For i = 1 To Len(str10sin)      '0�`9�͈͓̔��ɂ��邩�`�F�b�N
        If Mid$(str10sin, i, 1) < Chr$(48) Or Mid$(str10sin, i, 1) > Chr$(57) Then
            Exit Function
        End If
    Next i
    
        
    
    
    RetValue = CDec(str10sin)
    '16�i���͈͓̔����`�F�b�N
    If RetValue < 0 Or RetValue > 4294967295# Then Exit Function
    Keta(0) = 1: i = 0
    Do
        i = i + 1: k = i    'RetValue ��16�i���̌��������߂�
        Keta(i) = Keta(i - 1) * 16
    Loop Until Keta(i) > RetValue
    For i = 1 To k
        ModValue = Keta(k - i)
        '�������߂Ă��̒l��16�i���̉��ɂȂ邩�����߂�
        strSum = strSum & Mid$(Table, Int(RetValue / ModValue) + 1, 1)
        '�]������߂�16���傫���ꍇ�͍ēx�v�Z
        RetValue = RetValue - Int(RetValue / ModValue) * ModValue
    Next i
    f10sinTo16sin = strSum
End Function

Public Function GetLng(ByVal s) As Long

'2019/12/2 �����Ϗo�א��Z�o�����Ő��Y�v�捀�ڂ��󔒂ł����삷��l�ɒǉ�

    GetLng = 0
    s = StrConv(s, vbUnicode)
    If RTrim(s) <> "" Then
        GetLng = CLng(s)
    End If
End Function

