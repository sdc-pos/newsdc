Attribute VB_Name = "Y_SYU_LOG"
Option Explicit
'********************************************************************
'*
'*              �o�ח\��f�[�^���O
'*
'*          CREATE 2001.05.09
'********************************************************************
Public SYUKA_LOGF   As String   '�o�׃��O�t�@�C������
Public SYUKA_LOG_ON As Boolean  '�o�׃��O�o�͂n�m�^�n�e�e

Public Function SYUKA_LOGF_GET_PROC() As Integer
'****************************************************
'*      �o�׃��O�t�@�C�����̂̎捞��
'*
'*  ���� :  �Ȃ�
'*  �߂�l: false       ����
'*          SYS_ERR     �p���ł��Ȃ��ُ�
'****************************************************
Dim c       As String
Dim Ret     As Integer

    SYUKA_LOGF_GET_PROC = SYS_ERR
    
                                        '�o�׃��O�o�͗L����荞��
    If GetIni("SYUKA_LOG", StrConv(App.EXEName, vbUpperCase), "SYS", c) Then
        SYUKA_LOG_ON = False
        SYUKA_LOGF_GET_PROC = False
        Exit Function
    End If
    If Trim(c) = "0" Then
        SYUKA_LOG_ON = False
        SYUKA_LOGF_GET_PROC = False
        Exit Function
    End If

    SYUKA_LOG_ON = True

    If GetIni("FILE", "SYU_LOG", "SYS", c) Then
        Exit Function
    End If

    Ret = InStr(1, Trim(c), ".") - 1
    SYUKA_LOGF = Left(Trim(c), Ret) & Right(Format(Date, "yyyymmdd"), 2) & Right(Trim(c), Len(Trim(c)) - Ret)
    SYUKA_LOGF_GET_PROC = False
End Function


Public Sub SYUKA_LOG_OUT_PROC(YOIN1 As String, YOIN2 As String)
'****************************************************
'*      �o�׃��O�t�@�C���̏o��
'*
'*  ���� :  �o�͗v��
'*          �f�[�^��
'*  �߂�l:  �Ȃ�
'*  �Ăь��ŕێ����Ă���ŏI�o�ח\����e���o�͂���
'*
'****************************************************
Dim stream  As Integer                       '�t�@�C���ԍ�
Dim Buf     As String                           '�ǂݍ��݃o�b�t�@
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

    
    stream = FreeFile
    
    On Error Resume Next
    
    If Format(Date, "yyyymmdd") <> Format(FileDateTime(SYUKA_LOGF), "yyyymmdd") Then
        Open SYUKA_LOGF For Output As stream
    Else
        Open SYUKA_LOGF For Append As stream
    End If
    prog = StrConv(App.EXEName, vbUpperCase)
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If
    
    Buf = (Date$ & " " & Time$ & " " & com & " " & prog)
    Buf = (Buf & " " & YOIN1 & " " & YOIN2 & " ")
    Buf = (Buf & "�`���t�F" & StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) & " ")
    Buf = (Buf & "�`ID�F" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & " ")
    Buf = (Buf & "�`���F" & StrConv(Y_SYUREC.DEN_NO, vbUnicode) & " ")
    Buf = (Buf & "����F" & StrConv(Y_SYUREC.CYU_KBN, vbUnicode) & " ")
    Buf = (Buf & "������F" & StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & " ")
    Buf = (Buf & "�i�ԁF" & StrConv(Y_SYUREC.HIN_NO, vbUnicode) & " ")
    Buf = (Buf & "���F" & Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#0"))
    Print #stream, Buf
    Close stream

End Sub

