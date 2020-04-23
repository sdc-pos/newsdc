Attribute VB_Name = "SDC_FLD_M"
Option Explicit
Public SDC_FLD_Root As String            '�o�͐惋�[�g
Public SDC_FLD_Folder As String          '�t�H���_��
Public SDC_FLD_File As String            '�t�@�C����
Public SDC_FLD_xxx As String             '���ʎq

Public SDC_FLD_Return As Integer         '�m�F��ʏI�����

Function SDC_FLD_GET(Ini As String, Sect As String, Out_Nm As String) As Integer
'------------------------------------------------------------------------------
'                       �o�͐�m�F�@���@�f����
'
'   �o�͐惋�[�g��������΁A�m�F��ɍ쐬�i��ݾَ��͉�ʷ�ݾقƓ��l�j
'   �t�H���_��������΁A�m�F��ɍ쐬�i��ݾَ��͍ē��́j
'
'   Ini    �Fini�t�@�C����
'   Sect   �F����ݖ�
'   Out_Nm �F�o��̧�فi�٥�߽�j
'
'
'   �߂�l False�F����
'   �@�@�@ True �F��ݾق܂���ini�擾�ُ�
'------------------------------------------------------------------------------
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    '�o�͐惋�[�g
    sts = GetIni(Sect, "ROOT", Ini, c)
    If sts <> False Then GoTo SDC_FLD_GET_ERR
    SDC_FLD_Root = Trim(c)

    '�t�H���_��
    sts = GetIni(Sect, "FOLDER", Ini, c)
    If sts <> False Then GoTo SDC_FLD_GET_ERR
    SDC_FLD_Folder = Trim(c)

    '�t�@�C����
    sts = GetIni(Sect, "FILE", Ini, c)
    If sts <> False Then GoTo SDC_FLD_GET_ERR
    SDC_FLD_File = Trim(c)

    '���ʎq
    sts = GetIni(Sect, "xxx", Ini, c)
    If sts <> False Then GoTo SDC_FLD_GET_ERR
    SDC_FLD_xxx = Trim(c)

    SDC_FLD_F.Show vbModal

    If SDC_FLD_Return Then
        Out_Nm = ""
    Else
        Out_Nm = SDC_FLD_Root & "\" & SDC_FLD_Folder & _
                 "\" & SDC_FLD_File & "." & SDC_FLD_xxx
    End If

    SDC_FLD_GET = SDC_FLD_Return
    Exit Function


SDC_FLD_GET_ERR:
    SDC_FLD_GET = True
    Call Log_Out(LOG_F, Ini & " �ǂݍ��݃G���[<" & Sect & ">")
    MsgBox Ini & " �ǂݍ��݃G���[<" & Sect & ">", vbCritical

End Function
