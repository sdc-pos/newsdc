Attribute VB_Name = "mdlSocket"
Option Explicit

'[2014/02/10 - M.MATSUYAMA �ǉ�(Ver2.0.0)] �\�P�b�g�ʐM�p�ǉ�

'----- �������t�@�C��(�\�P�b�g�ʐM�p) -----
Public Const SEC_SOCKET         As String = "F110010"               '�\�P�b�g�ʐM�p�ݒ�Z�N�V����
Public Const KEY_LOCALPORT      As String = "SocketPort"            '���[�J���|�[�g�ԍ�

Public Const DEF_LOCALPORT      As Long = 2222                      '�f�t�H���g���[�J���|�[�g�ԍ�

Public Const SEC_LOG            As String = "F110010"               '���O�t�@�C���o�͗p�ݒ�Z�N�V����
Public Const KEY_LOGWRITE       As String = "LogWrite"              '���O�t�@�C���o�̓t���O
Public Const KEY_LOGPATH        As String = "LogPath"               '���O�t�@�C���ۑ��t�H���_
Public Const KEY_LOGSAVE        As String = "LogSave"               '���O�t�@�C���ۑ�����

Public Const DEF_LOGWRITE       As Boolean = True                   '�f�t�H���g���O�t�@�C���o�̓t���O
Public Const DEF_LOGSAVE        As Integer = 30                     '�f�t�H���g���O�t�@�C���ۑ�����

Public Const FNC_LOGSAVECHK     As String = "���O�ۑ����ԃ`�F�b�N����"
Public Const FNC_DATEMONITOR    As String = "���t�X�V�Ď�����"

Public Const FNC_PARENTCONN     As String = "�A�N�Z�X�|�C���g�ڑ�����"
Public Const FNC_PARENTDISCONN  As String = "�A�N�Z�X�|�C���g�ؒf����"
Public Const FNC_RECVDATA       As String = "�f�[�^��M����"
Public Const FNC_RESPDATA       As String = "�f�[�^�ԐM����"
Public Const FNC_SENDDATA       As String = "�f�[�^���M����"
Public Const FNC_RECVMESSAGE    As String = "���b�Z�[�W��M����"
Public Const FNC_FILESEND       As String = "�t�@�C�����M����"
Public Const FNC_SOCKCLOSE      As String = "�\�P�b�g�ʐM�ؒf����"
Public Const FNC_SOCKCONNECT    As String = "�\�P�b�g�ʐM�ڑ�����"
Public Const FNC_SOCKCONNREQ    As String = "�\�P�b�g�ʐM�ڑ��v������"
Public Const FNC_SOCKSEND       As String = "�\�P�b�g�f�[�^���M����"
Public Const FNC_SOCKRECEIVE    As String = "�\�P�b�g�f�[�^��M����"
Public Const FNC_SOCKERROR      As String = "�\�P�b�g�ʐM�G���["

Public Const MAX_FSENDDATASIZE  As Integer = 1445                   '�ő�t�@�C�����M�f�[�^�T�C�Y

Public Const P_FILELOAD         As String = "FILELOAD"              '�t�@�C����M�p�R�}���h

Public Const RESP_OK            As String = "1"                     '����I��
Public Const RESP_NG            As String = "9"                     '�G���[

Public Type SOCKET_CONFIG
    m_IsListen      As Boolean          '�\�P�b�g�ʐM�J�n�t���O(True:���X�i�[�������ς�)
    m_LocalPort     As Long             '���[�J���|�[�g�ԍ�
End Type

Public Type LOG_CONFIG
    m_LogWrite      As Boolean          '���O�t�@�C���o�̓t���O
    m_LogPath       As String           '���O�t�@�C���ۑ��t�H���_
    m_LogFName      As String           '���O�t�@�C����
    m_LogSave       As Long             '���O�t�@�C���ۑ�����(�P��[��] 0 �̏ꍇ�͍폜���Ȃ�)
End Type

Public gbl_SockCfg  As SOCKET_CONFIG
Public gbl_LogCfg   As LOG_CONFIG

'----- �A�v���P�[�V�����N������ -----
Public gbl_StartApp As Date
    
Public Enum LogMsgIcon
    icoDownload = 1
    icoUpload = 2
    icoError = 3
    icoMessage = 4
End Enum

'*******************************************************************************
' ���O���b�Z�[�W�\��
' process   :   �w�肳�ꂽ���O���b�Z�[�W�����X�g�ɒǉ�����
' input     :   strMsg      ���b�Z�[�W������
'           :   strFunc     ������
'           :   [intID]     �q�@ID
'           :   [strIP]     IP�A�h���X
'           :   [lngIcon]   �A�C�R�����
' output    :   �Ȃ�
' return    :   �Ȃ�
'*******************************************************************************
Public Sub WriteLogMsg(ByVal strMsg As String, ByVal strFunc As String, Optional ByVal intID As Integer = -1, Optional strIP As String = "", Optional lngIcon As LogMsgIcon = icoMessage)
    Static varIcon  As Variant
    Dim strNow      As String
    Dim intFNo      As Integer
    Dim strLogText  As String
    Dim strLogFile  As String
    Dim strID       As String
    Dim strIPAddr   As String
    Dim flgOpen     As Boolean
    
    If IsArray(varIcon) = False Then
        varIcon = Array("(��)", "(��)", "(�~)", "(��)")
    End If
    
    '----- ���ݓ����̕�����쐬 -----
    strNow = Format$(Now, "yyyy/mm/dd hh:mm:ss")
    
    '----- �q�@ID�̕�����쐬 -----
    strID = IIf(intID < 0, "---", Format$(intID, "000"))
    
    strIPAddr = IIf(Len(strIP) > 0, "(" & strIP & ")", "")
    
    '----- ���O�o�͕�������쐬 -----
    strLogText = "[" & strNow & "]" & vbTab & CStr(varIcon(lngIcon - 1)) & vbTab & "(" & strID & ")" & vbTab & strIPAddr & vbTab & "�y" & strFunc & "�z" & vbTab & strMsg
    
    '----- ���O�o�̓I�v�V���� -----
    If gbl_LogCfg.m_LogWrite = False Then Exit Sub
    
    '------------------------
    '   ���O�t�@�C���ɏo��
    '------------------------
    flgOpen = False
    
    strLogFile = gbl_LogCfg.m_LogFName
    
    On Error GoTo WriteLogMsg_Exit
    intFNo = FreeFile
    Open strLogFile For Append Access Write Lock Read Write As #intFNo
    flgOpen = True
    
    '----- ���O���b�Z�[�W���������� -----
    Print #intFNo, strLogText
    
WriteLogMsg_Exit:
    
    If flgOpen = True Then
        Close #intFNo
    End If

End Sub

'******************************************************************************************
' ���s���G���[���b�Z�[�W�o��
' process   :   �������ɔ����������s���G���[�̃��b�Z�[�W�����O�ɏo�͂���B
' input     :   objErr      �G���[�I�u�W�F�N�g
'           :   strFunc     ���s������
' output    :   �Ȃ�
' return    :   �G���[���b�Z�[�W
'******************************************************************************************
Public Sub WriteLogErr(objErr As ErrObject, strFunc As String)
    Dim strMsg  As String
    Dim lngIcon As Long
    
    '----- ���b�Z�[�W�̍쐬�ƃA�C�R���̌��� -----
    If objErr.Number = 0 Then
        strMsg = "�G���[�͔������Ă��܂���B"
        lngIcon = LogMsgIcon.icoMessage
    Else
        strMsg = strFunc & "���Ɏ��s���G���[���������܂����B[���e] : " & ChStr(objErr.Description, vbCrLf, " ") & " [�ԍ�] : " & CStr(objErr.Number)
        lngIcon = LogMsgIcon.icoError
    End If
    
    '----- ���O�o�� -----
    Call WriteLogMsg(strMsg, strFunc, , , lngIcon)
End Sub

'*******************************************************************************
' ���O�t�@�C���폜
' process   :   �����Ԉȏ�O�̃��O�t�@�C�����폜����
' input     :   strTagName  ���O�t�@�C������
'           :   lngSaveDay  ���O�ێ�����
' output    :   �Ȃ�
' return    :   [strTagName]_yyyymmdd.log �̌`���̃t�@�C�������������܂��B
'*******************************************************************************
Public Sub DeleteLogFile(strTagName As String, lngSaveDay As Long)
    Dim strPath As String
    Dim strFind As String
    Dim strSave As String
    Dim strCheck As String
    Dim strFName As String
    
    Call WriteLogMsg(FNC_LOGSAVECHK & "���J�n���܂��B(�����Ώ�:" & strTagName & "_YYYYMMDD.log / �ۑ�����:" & CStr(lngSaveDay) & " ����)", FNC_LOGSAVECHK, , , icoMessage)
    
    On Error GoTo DeleteLogFile_Exit
    Err.Clear
    
    If lngSaveDay <= 0 Then Exit Sub
    
    '----- �����p������쐬 -----
    strFind = GetFullPath(gbl_LogCfg.m_LogPath, strTagName) & "_????????.log"
    
    '----- �ۑ����Ԃ̓��t��������쐬 -----
    strSave = Format$(DateAdd("d", -(lngSaveDay), gbl_StartApp), "yyyymmdd")
    
    '----- ���O�t�@�C�������J�n -----
    strCheck = Dir$(strFind, vbArchive)
    Do While Len(strCheck) > 0
        If StrComp(Mid$(strCheck, Len(strTagName) + 2, 8), strSave, vbTextCompare) < 0 Then
            '----- �ۑ����Ԃ��Â����O�t�@�C���̏ꍇ -----
            strFName = GetFullPath(gbl_LogCfg.m_LogPath, strCheck)
            Kill strFName
            Call WriteLogMsg("�ۑ����Ԃ��Â����O�t�@�C��(" & strFName & ")���폜���܂����B", FNC_LOGSAVECHK, , , icoMessage)
        End If
        
        '----- ���̃t�@�C�������� -----
        strCheck = Dir$
    Loop
    
DeleteLogFile_Exit:
    If Err.Number <> 0 Then
        Call WriteLogErr(Err, FNC_LOGSAVECHK)
    End If
    
    Call WriteLogMsg(FNC_LOGSAVECHK & "���I�����܂��B", FNC_LOGSAVECHK, , , icoMessage)
End Sub

'*******************************************************************************
' ���M�p�f�[�^���H
' process   :   �w�肳�ꂽ��������n���f�B�ւ̑��M�p�f�[�^�ɉ��H����
' input     :   strMsg      ���M�f�[�^
' output    :   �Ȃ�
' return    :   ���M�p������
'*******************************************************************************
Public Function ConvTextMsg(strMsg As String) As String
    Dim strBuf As String
    Dim intLoop As Integer
    
    '----- ���M����������H -----
    strBuf = ChStr(strMsg, "[CRLF]", vbCrLf)
    strBuf = ChStr(strBuf, "[CR]", vbCr)
    strBuf = ChStr(strBuf, "[LF]", vbLf)
    strBuf = ChStr(strBuf, "[TAB]", vbTab)
    strBuf = ChStr(strBuf, "[ESC]", Chr$(27))
    For intLoop = 1 To 255
        strBuf = ChStr(strBuf, "[0x" & Right$("0" & UCase(Hex(intLoop)), 2) & "]", Chr$(intLoop))
        strBuf = ChStr(strBuf, "[0x" & Right$("0" & LCase(Hex(intLoop)), 2) & "]", Chr$(intLoop))
    Next intLoop
    strBuf = ChStr(strBuf, "[DATE]", Format$(Now, "yyyymmdd"))
    strBuf = ChStr(strBuf, "[TIME]", Format$(Now, "hhmmss"))
    
    ConvTextMsg = strBuf
End Function

'*******************************************************************************
' ��M�f�[�^���H
' process   :   �n���f�B�����M�����f�[�^���e�L�X�g�`���ɉ��H����
' input     :   strMsg      ��M�f�[�^
' output    :   �Ȃ�
' return    :   ��M�e�L�X�g������
'*******************************************************************************
Public Function ConvBinaryMsg(strMsg As String) As String
    Dim strBuf As String
    Dim strWrk As String
    Dim strChr As String
    Dim intLoop As Integer
    
    '----- ��M����������H -----
    strBuf = ChStr(strMsg, vbCrLf, "[CRLF]")
    strBuf = ChStr(strBuf, vbCr, "[CR]")
    strBuf = ChStr(strBuf, vbLf, "[LF]")
    strBuf = ChStr(strBuf, vbTab, "[TAB]")
    strBuf = ChStr(strBuf, Chr$(27), "[ESC]")
    strWrk = ""
    For intLoop = 1 To Len(strBuf)
        strChr = Mid$(strBuf, intLoop, 1)
        If (Asc(strChr) < 0) Or (Asc(strChr) >= &H20 And Asc(strChr) <= &H7E) Or (Asc(strChr) >= &HA0 And Asc(strChr) <= &HDF) Then
            '----- �\���\�����̏ꍇ -----
            strWrk = strWrk & strChr
        Else
            '----- �\���s�\�����̏ꍇ -----
            strWrk = strWrk & "[0x" & Right$("0" & UCase(Hex(Asc(strChr))), 2) & "]"
        End If
    Next intLoop
    
    ConvBinaryMsg = strWrk
End Function

'********************************************************************************
' �t���p�X�t�@�C�����쐬
' process   :   �f�B���N�g�����A�t�@�C���������Ƀt���p�X�̃t�@�C�������쐬����
' intput    :   strPath     �p�X��
'           :   strFile     �t�@�C����
'           :   varToken    �p�X���A�t�@�C�����̋�؂蕶��
' output    :   �Ȃ�
' return    :   �t���p�X�t�@�C����
'********************************************************************************
Public Function GetFullPath(strPath As String, strFile As String, Optional varToken As Variant = "\") As String
    Dim strFPath     As String
    Dim strName1    As String
    Dim strName2    As String
    Dim strMake     As String
    Dim intTokLen   As Integer
    
    If strPath = "" And strFile <> "" Then
        '�t�@�C���������w�肳��Ă���
        strFPath = strFile
        GoTo GetFullPath_Exit
    ElseIf strPath <> "" And strFile = "" Then
        '�p�X�������w�肳��Ă���
        strFPath = strPath
        GoTo GetFullPath_Exit
    ElseIf strPath = "" And strFile = "" Then
        '�ǂ�����w�肳��Ă��Ȃ�
        strFPath = ""
        GoTo GetFullPath_Exit
    End If
    
    intTokLen = Len(CStr(varToken))
    
    If Right$(strPath, intTokLen) = CStr(varToken) Then
        '��؂蕶�����Ȃ�
        strName1 = Left$(strPath, Len(strPath) - intTokLen)
    Else
        strName1 = strPath
    End If
    If Left$(strFile, intTokLen) = CStr(varToken) Then
        '��؂蕶�����Ȃ�
        strName2 = Right$(strFile, Len(strFile) - intTokLen)
    Else
        strName2 = strFile
    End If
    
    '�t���p�X���쐬
    strFPath = strName1 & CStr(varToken) & strName2
    
GetFullPath_Exit:
    GetFullPath = strFPath
End Function

'******************************************************************************************
' �f�B���N�g���쐬
' process   :   �w�肳�ꂽ�p�X���쐬����B
' input     :   strPath     �쐬����p�X
' output    :   ����
' date      :   1999/02/25 - K.HAYASHI �C��
' return    :   True:����I��, False:�ُ�I��
'******************************************************************************************
Public Function MkDirEx(strPath As String) As Boolean
    Dim strDrv   As String
    Dim flgRet   As Variant
    Dim strDir() As String
    Dim strWork  As String
    Dim intCnt   As Integer
    Dim intLoop  As Integer
    
    Err.Clear
    On Error GoTo MkDirEx_Exit
    
    '----- ���^�[���l������ -----
    flgRet = False
    
    '----- �h���C�u�`�F�b�N -----
    If InStr(strPath, ":\") > 0 Then
        '----- �h���C�u���������o�� -----
        strDrv = Left$(strPath, 2)
        
        Call ChDir(strDrv)
    End If
    
    '----- �e�f�B���N�g���𕪊� -----
    intCnt = Explode(Mid$(strPath, 4), strDir, "\")
    If intCnt > 0 Then
        strWork = strDrv
        intLoop = 0
        For intLoop = 0 To intCnt - 1
            If Len(strDir(intLoop)) = 0 Then
                Exit For
            End If
            
            '----- �f�B���N�g�����쐬 -----
            strWork = strWork & "\" & strDir(intLoop)
            
            '----- �f�B���N�g���L���`�F�b�N -----
            If Len(Dir2(strWork, vbDirectory)) = 0 Then
                '----- �f�B���N�g���쐬 -----
                MkDir strWork
            End If
        Next intLoop
        
        flgRet = True
    End If
MkDirEx_Exit:
    MkDirEx = flgRet
End Function

'*******************************************************************************
' ���X�g������擾
' process   :   1�ɂ܂Ƃ߂�ꂽ���X�g�������z��Ɋi�[����
' input     :   strList     ���X�g������
'           :   varBuf      ������i�[�o�b�t�@
'           :   strToken    ���������؂�g�[�N��
' output    :   varBuf      �z��Ɋi�[���ꂽ���X�g������
' return    :   �z��, 0:�f�[�^��
'*******************************************************************************
Public Function Explode(strList As String, varBuf As Variant, strToken As String) As Integer
    Dim Index As Integer
    Dim intSp As Integer
    Dim intEp As Integer
    
    '----- ���^�[���l������ -----
    Index = 0
    
    '----- �����`�F�b�N -----
    If strToken = "" Then
        GoTo Explode_Exit
    End If
    
    ReDim varBuf(0) As String
    intSp = 1
    intEp = InStr(intSp, strList, strToken)
    Do While intEp > 0
        ReDim Preserve varBuf(Index) As String
        varBuf(Index) = Mid$(strList, intSp, intEp - intSp)
        intSp = intEp + Len(strToken)
        intEp = InStr(intSp, strList, strToken)
        Index = Index + 1
    Loop
    ReDim Preserve varBuf(Index) As String
    varBuf(Index) = Mid$(strList, intSp)
    Index = Index + 1
    
Explode_Exit:
    Explode = Index
End Function

'*******************************************************************************
' �t�@�C���L���`�F�b�N
' process   :   �w��t�@�C���������݂��邩�ǂ����`�F�b�N����
' input     :   szFilePath      �m�F����t�@�C���̃p�X
'           :   intAttr           �m�F����t�@�C���̃^�C�v
' output    :   �Ȃ�
' return    :   �L:�t�@�C����, ��:""
'*******************************************************************************
Public Function Dir2(Optional varFilePath As Variant, Optional intAttr As VbFileAttribute = vbNormal) As String
    Static strPath As String
    Dim strBuf As String
    Dim strRet As String
    Dim strChk  As String
    
    On Error Resume Next
    
    If IsMissing(varFilePath) = False Then
        strPath = ""
        Call DivFileName(CStr(varFilePath), strPath, strBuf)
    End If
    
    strRet = Dir$(varFilePath, intAttr)
    If Len(strRet) > 0 Then
        If Len(strPath) > 0 Then
            strChk = GetFullPath(strPath, strRet)
        Else
            strChk = strRet
        End If
        
        If (GetAttr(strChk) And intAttr) <> intAttr Then
            '�t�@�C�������`�F�b�N
            strRet = ""
        End If
    End If
    
    On Error GoTo 0
    
    Dir2 = strRet
End Function

'*******************************************************************************
' �t�@�C��������
' process   :   �w��t�@�C�����̃f�B���N�g���ƃt�@�C���̕����𕪉�����
' input     :   strFilePath     ��������t�@�C���̃p�X
'           :   strPathName     �����㕶����i�[�o�b�t�@�P
'           :   strFileName     �����㕶����i�[�o�b�t�@�Q
'           :   strToken        �t�@�C������؂蕶��
' output    :   strPathName     �����㕶����P
'           :   strFileName     �����㕶����Q
' return    :   True:����, False:�ُ�
'*******************************************************************************
Public Function DivFileName(strFilePath As String, strPathName As String, strFileName As String, Optional strToken As String = "\") As Boolean
    Dim strFBuf As String
    Dim strPath As String
    Dim strFile As String
    Dim intPos  As Integer
    
    If StrComp(Right$(strFilePath, Len(strToken)), strToken) = 0 Then
        '----- ��؂蕶�����Ō�ɂ���ꍇ�͏Ȃ� -----
        strFBuf = Left$(strFilePath, Len(strFilePath) - Len(strToken))
    Else
        strFBuf = strFilePath
    End If
    
    For intPos = Len(strFBuf) To 1 Step -1
        '----------------------------------------
        '   �Ō�̕��������؂蕶������������
        '   �ŏ��ɊY�����������ȍ~���t�@�C����
        '----------------------------------------
        If StrComp(Mid$(strFBuf, intPos, Len(strToken)), strToken) = 0 Then
            strFile = Mid$(strFBuf, intPos + Len(strToken))
            strPath = Left$(strFBuf, intPos - 1)
            If StrComp(strToken, "/") = 0 Then
                If Len(strPath) = 0 Then
                    strPath = "/"
                End If
            ElseIf StrComp(strToken, "\") = 0 Then
                If StrComp(Right$(strPath, 1), ":") = 0 Then
                    '----- ���[�g�p�X(C:\��)�̂Ƃ� -----
                    strPath = strPath & strToken
                End If
            End If
            Exit For
        End If
    Next intPos
    
    If Len(strFile) * Len(strPath) <> 0 Then
        strPathName = strPath
        strFileName = strFile
        DivFileName = True
    Else
        strPathName = ""
        strFileName = ""
        DivFileName = False
    End If
End Function

'******************************************************************************************
' ������u������
' process   :   �w�肳�ꂽ������̒��̎w�蕶����S�Ēu��������
' input     :   strText     �ҏW������
'           :   strToken1   ����������
'           :   strToken2   �u������������
' output    :   ����
' return    :   �ҏW��̕�����
'******************************************************************************************
Public Function ChStr(strText As String, strToken1 As String, strToken2 As String) As String
    Dim strWork As String
    Dim strLeft As String
    Dim strRight As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    strWork = strText
    intPos1 = 1
    Do
        '�u�����������������
        intPos2 = InStr(intPos1, strWork, strToken1)
        If intPos2 = 0 Then
            Exit Do
        Else
            strLeft = Left$(strWork, intPos2 - 1)
            strRight = Mid$(strWork, intPos2 + Len(strToken1))
        End If
        
        '�������u��������
        strWork = strLeft & strToken2 & strRight
        
        intPos1 = Len(strLeft) + Len(strToken2) + 1
    Loop
    ChStr = strWork
End Function

'******************************************************************************************
' ������o�C�g����
' process   :   �w�肳�ꂽ�o�C�g���ɒ���������������쐬����B
' input     :   strText     �������镶����
'           :   intLength   ��������o�C�g��
'           :   lngFormat   �e�L�X�g�z�u�ʒu
' output    :   �Ȃ�
' return    :   �w��o�C�g���ɒ�������������
'******************************************************************************************
Function AlignText(strText As String, intLength As Integer, Optional lngFormat As AlignmentConstants = -1) As String
    Dim strMake     As String       '���H������o�b�t�@
    Dim strBuf      As String       '���[�N�o�b�t�@
    Dim strBody     As String       '���[�N�o�b�t�@
    Dim strLast     As String       '���[�N�o�b�t�@
    Dim intSpcMax   As Integer      '�X�y�[�X�ǉ���
    Dim intSpcL     As Integer      '�����X�y�[�X��
    Dim intSpcR     As Integer      '�E���X�y�[�X��
    
    If Len(strText) = 0 Then
        '----- ����0�̕����̏ꍇ -----
        AlignText = Space$(intLength)
        Exit Function
    End If
    
    '----- �w�肳�ꂽ�o�C�g���ŕ������؂� -----
    strBuf = StrConv(LeftB(StrConv(strText, vbFromUnicode), intLength), vbUnicode)
    
    '----- 1�������O�̕�����ƍŌ��1�������擾 -----
    strBody = Left$(strText, Len(strBuf) - 1)
    strLast = Mid$(strText, Len(strBuf), 1)
    
    '----- �Ō�̕����̃o�C�g�����`�F�b�N -----
    If LenB(StrConv(strLast, vbFromUnicode)) = 2 Then
        '----- 2�o�C�g�����̏ꍇ�͍��v�o�C�g�����`�F�b�N -----
        If LenB(StrConv(strBody, vbFromUnicode)) + LenB(StrConv(strLast, vbFromUnicode)) > intLength Then
            '----- �w��o�C�g�����傫���ꍇ��1�������Ȃ� -----
            strMake = strBody
        Else
            '----- ���̂܂܌������� -----
            strMake = strBody & strLast
        End If
    Else
        '----- �Ō�̕�����1�o�C�g�����Ƃ������Ƃ͐���ȃT�C�Y -----
        strMake = strBody & strLast
    End If
    
    '----- �����ʒu�ݒ� -----
    intSpcMax = intLength - LenB(StrConv(strMake, vbFromUnicode))
    
    Select Case (lngFormat)
    Case vbCenter
        '----- �������� -----
        intSpcL = intSpcMax \ 2
        intSpcR = intSpcMax - intSpcL
        AlignText = Space$(intSpcL) & strMake & Space$(intSpcR)
    Case vbRightJustify
        '----- �E�l�� -----
        AlignText = Space$(intSpcMax) & strMake
    Case vbLeftJustify
        '----- ���l�� -----
        AlignText = strMake & Space$(intSpcMax)
    Case Else
        '----- �ݒ�Ȃ� -----
        AlignText = strMake
    End Select
    
End Function

Public Function GetFileData(ByVal FileName As String) As String
    Dim strRetBuf  As String
    Dim strReadBuf  As String
    Dim intFNo      As Integer
    Dim blnOpen     As Boolean
    
    On Error GoTo GetFileData_Exit
    
    '----- �߂�l�̏����� -----
    strRetBuf = ""
    
    blnOpen = False
    
'    If IsExist(FileName) = False Then
'        '----- �t�@�C�������݂��Ȃ��ꍇ -----
'        GoTo GetFileData_Exit
'    End If
    
    '----------------------
    '   �t�@�C���I�[�v��
    '----------------------
    intFNo = FreeFile()
    Open FileName For Input Access Read Lock Write As #intFNo
    blnOpen = True
    
    Do While Not EOF(intFNo)
        Line Input #intFNo, strReadBuf
        strRetBuf = strRetBuf & strReadBuf & vbCrLf
    Loop
    
GetFileData_Exit:
    If blnOpen = True Then
        '----------------------
        '   �t�@�C���N���[�Y
        '----------------------
        Close #intFNo
    End If
    
    GetFileData = strRetBuf
End Function

