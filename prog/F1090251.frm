VERSION 5.00
Begin VB.Form F1090251 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�݌ɍ��ك`�F�b�N����"
   ClientHeight    =   4710
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   7320
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�z�X�g�f�[�^�W�v��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�f�[�^��������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�݌ɍ��ك`�F�b�N����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   4800
   End
End
Attribute VB_Name = "F1090251"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type YUKO_SOKO_TBL             '�L��νđq�Ɏ�荞�݃e�[�u��
    HS_SOKO As String * 2
    NAIGAI As String * 1
End Type

Private SOKO_T() As YUKO_SOKO_TBL       '�q�ɏ��

Private HS_NaiG As String               '�����O�i������e�j�����@ν��ް����e�ɂ��ݒ�


Private JGYOBA_CODE As String           '���Ə�R�[�h
Private SYUSI_CODE  As String           '���x�R�[�h
Private HIN_GAI     As String           '�i�ԁi�O���j
Private HIN_NAI     As String           '�i�ԁi�����j
Private HIN_MEI     As String           '�i��
Private Dummy       As String           '�_�~�[
Private ZAIKO_QTY   As String           '�݌ɐ�

Private FileNo      As Integer          '�t�@�C���ԍ�
Private HS_ZAI      As String           '���C���t�@�C���l�[��

Private ZENKAI_YMD      As String       '�O�񏈗��N����
Private ZENZENKAI_YMD   As String       '�O�X�񏈗��N����


Private Function SumZ_Init() As Integer
'----------------------------------------------------------------------------
'                   �u�݌ɏW�v�f�[�^�v�z�X�g�݌ɃN���A�[����
'                   �����g�p�@  2006.07.03
'----------------------------------------------------------------------------

Dim sts As Integer
Dim com As Integer
Dim ans As Integer

    SumZ_Init = True
    
    com = BtOpGetGreater


    com = BtOpGetFirst

    Do
                
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    
                    
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        Loop
        
        If sts Then
            Exit Do
        End If
        
        '�O�񁨑O�X��
        Call UniCode_Conv(SUMZREC.ZEN_HS_ZAIQTY, StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
        
        Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
        Call UniCode_Conv(SUMZREC.SAI_QTY, StrConv(SUMZREC.T_Zai_Qty, vbUnicode))
        
        Do
            sts = BTRV(BtOpUpdate, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
        
        DoEvents
    Loop

    SumZ_Init = False

End Function




Private Function New_SumZ_Init() As Integer
'----------------------------------------------------------------------------
'                   �u�݌ɏW�v�f�[�^�v�z�X�g�݌ɃN���A�[����
'----------------------------------------------------------------------------

Dim sts As Integer
Dim com As Integer
Dim ans As Integer

    New_SumZ_Init = True
    
    com = BtOpGetFirst


    Do
                
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    
                    
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        Loop
        
        If sts Then
            Exit Do
        End If
        
        Call UniCode_Conv(SUMZREC.BU_ZAI_QTY, "00000000")
        
        Do
            sts = BTRV(BtOpUpdate, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
        
        DoEvents
    Loop

    New_SumZ_Init = False

End Function


Private Function SumZ_Update(JGYOBU As String) As Integer
'----------------------------------------------------------------------------
'                   �u�݌ɏW�v�f�[�^�v�z�X�g�݌ɍX�V����
'----------------------------------------------------------------------------

Dim i               As Integer
Dim j               As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim T_Zai_Qty       As Long
Dim SAI_QTY         As Long
        
Dim ans             As Integer
        
Dim Input_Buffer    As String
Dim Pos             As Integer
        
Dim Skip_Flg        As Boolean
        
        
Dim Input_Wk        As Variant
        
        
    SumZ_Update = True
    
    Do While Not EOF(FileNo)
        DoEvents
        
        Input #FileNo, Input_Buffer

        Input_Wk = Split(Input_Buffer, vbTab, -1)

        JGYOBA_CODE = ""
        HIN_GAI = ""
        HIN_NAI = ""
        HIN_MEI = ""
        Dummy = ""
        ZAIKO_QTY = "0"

        If UBound(Input_Wk) >= 0 Then
            JGYOBA_CODE = Input_Wk(0)
        End If

        If UBound(Input_Wk) >= 1 Then
            SYUSI_CODE = Input_Wk(1)
        End If

        If UBound(Input_Wk) >= 2 Then
            HIN_GAI = Input_Wk(2)
        End If

        If UBound(Input_Wk) >= 3 Then
            HIN_NAI = Input_Wk(3)
        End If

        If UBound(Input_Wk) >= 4 Then
            HIN_MEI = Input_Wk(4)
        End If

        If UBound(Input_Wk) >= 5 Then
            Dummy = Input_Wk(5)
        End If

        If UBound(Input_Wk) >= 6 Then
            ZAIKO_QTY = Input_Wk(6)
        End If


'''        i = 0
'''        Do
'''
'''
'''            If InStr(Input_Buffer, Chr(9)) <> 0 Then
'''               ' �^�u��؂�L���܂ňʒu���ړ����܂��B
'''               Pos = InStr(Input_Buffer, Chr(9))
'''               ' �t�B�[���h �e�L�X�g�� strField �ϐ��Ɋ��蓖�Ă܂��B
'''
'''                Select Case i
'''                    Case 0
'''                       JGYOBA_CODE = Left(Input_Buffer, Pos - 1)
'''                    Case 1
'''                       SYUSI_CODE = Left(Input_Buffer, Pos - 1)
'''                    Case 2
'''                       HIN_GAI = Left(Input_Buffer, Pos - 1)
'''                    Case 3
'''                       HIN_NAI = Left(Input_Buffer, Pos - 1)
'''                    Case 4
'''                       HIN_MEI = Left(Input_Buffer, Pos - 1)
'''                    Case 5
'''                       Dummy = Left(Input_Buffer, Pos - 1)
'''                End Select
'''
'''                Input_Buffer = Right(Input_Buffer, Len(Input_Buffer) - Pos)
'''                i = i + 1
'''            Else
'''               ' �^�u��؂�L����������Ȃ���΁A
'''               ' �t�B�[���h �e�L�X�g�͍s�̍Ō�̃t�B�[���h�ł��B
'''                ZAIKO_QTY = Input_Buffer
'''                Exit Do
'''            End If
'''
'''
'''        Loop
        
        '�݌ɐ����O�A�����i�ԁ��󔒂͏����ΏۊO
        Skip_Flg = False
        If Not IsNumeric(ZAIKO_QTY) Then
            Skip_Flg = True
        Else
            If CLng(ZAIKO_QTY) = 0 Then
                Skip_Flg = True
            End If
        End If
        
'        If Len(Trim(HIN_NAI)) = 0 Then     2004.06.15
'            Skip_Flg = True
'        End If
        
        If Not Skip_Flg Then
            '�L���f�[�^�̃`�F�b�N�������O�̊l��
            Skip_Flg = True
            For i = 0 To UBound(JGYOBU_T)
                If JGYOBU = JGYOBU_T(i).CODE Then
                    For j = 0 To UBound(SOKO_T, 2)
                        If SYUSI_CODE = SOKO_T(i, j).HS_SOKO Then
                            Skip_Flg = False
                            Exit For
                        End If
                    Next j
                    Exit For
                End If
            Next i
            If Not Skip_Flg Then
                '�Ώۃf�[�^
                Call UniCode_Conv(K0_SUMZ.JGYOBU, JGYOBU)
                Call UniCode_Conv(K0_SUMZ.NAIGAI, SOKO_T(i, j).NAIGAI)
                Call UniCode_Conv(K0_SUMZ.HIN_GAI, HIN_GAI)
        
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                    Select Case sts
                        Case BtNoErr
                            Upd_com = BtOpUpdate
                            Exit Do
                        Case BtErrKeyNotFound
                            Upd_com = BtOpInsert
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�݌ɏW�v�f�[�^")
                            Exit Function
                    End Select
                Loop
            
                If Upd_com = BtOpInsert Then
                '�V�K�ǉ����A�i�ڃ}�X�^���z�X�g�I�Ԋl��
                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, SOKO_T(i, j).NAIGAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                        '���肦�Ȃ����X���[
                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                            Call UniCode_Conv(ITEMREC.ST_RETU, "")
                            Call UniCode_Conv(ITEMREC.ST_REN, "")
                            Call UniCode_Conv(ITEMREC.ST_DAN, "")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
        
        
                    Call UniCode_Conv(SUMZREC.JGYOBU, JGYOBU)
                    Call UniCode_Conv(SUMZREC.NAIGAI, SOKO_T(i, j).NAIGAI)
                    Call UniCode_Conv(SUMZREC.HIN_GAI, HIN_GAI)
                    Call UniCode_Conv(SUMZREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Call UniCode_Conv(SUMZREC.T_Zai_Qty, "00000000")
                    Call UniCode_Conv(SUMZREC.ZEN_Zai_Qty, "00000000")
                    Call UniCode_Conv(SUMZREC.SYK_E_QTY, "00000000")
                    Call UniCode_Conv(SUMZREC.NYUKA_YQTY, "00000000")
                    Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                    Call UniCode_Conv(SUMZREC.ZEN_HS_ZAIQTY, "00000000")
                    Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")
                    
'2007.05.17                   Call UniCode_Conv(SUMZREC.SUM_DT, Format(Date, "yyyymmdd"))
                    
                    Call UniCode_Conv(SUMZREC.BU_ZAI_QTY, "00000000")
                    Call UniCode_Conv(SUMZREC.PPSC_ZAI_QTY, "00000000")
                    
                    Call UniCode_Conv(SUMZREC.FILLER, "")
                    
                End If
                
        
                Call UniCode_Conv(SUMZREC.SUM_DT, Format(Date, "yyyymmdd"))     '2007.05.17
                Call UniCode_Conv(SUMZREC.BU_ZAI_QTY, Format(CLng(StrConv(SUMZREC.BU_ZAI_QTY, vbUnicode)) + CLng(ZAIKO_QTY), "00000000"))
        
''                SAI_QTY = CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) - CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
''
''                If SAI_QTY >= 0 Then
''                    Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "00000000"))
''                Else
''                    Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "0000000"))
''                End If
        
                Do
                
                    sts = BTRV(Upd_com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, Upd_com, "�݌ɏW�v�f�[�^")
                            Exit Function
                    End Select
            
                Loop
            End If
        
        End If
                                    
    Loop
    
    
    
    
    
    
    SumZ_Update = False

End Function
Private Function HS_ZAI_MAIN() As Integer
'----------------------------------------------------------------------------
'                   �u�݌ɏW�v�f�[�^�v�z�X�g�݌ɍX�V����
'----------------------------------------------------------------------------
Dim Ret         As String

Dim fileName    As String

Dim i           As Integer
Dim ans         As Integer



    HS_ZAI_MAIN = True

    For i = 0 To UBound(JGYOBU_T)
        
        
'''        If JGYOBU_T(i).CODE = SENTAKU Then
        
        
            FileNo = FreeFile
            fileName = HS_ZAI
    
            Ret = InStr(1, Trim(fileName), ".") - 1
            fileName = Left(Trim(fileName), Ret) & "_" & JGYOBU_T(i).CODE & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
            
            On Error GoTo Error_Proc
            
            Open fileName For Input As #FileNo
        
            On Error GoTo 0
        
        
            If SumZ_Update(JGYOBU_T(i).CODE) Then   '�u�݌ɏW�v�f�[�^�v�z�X�g�݌ɍX�V����
    
                Exit Function
            End If
        
        
            Close #FileNo
    
'''        End If
    
    Next i


    If SumZ_Total_Proc() Then
        Exit Function
    End If

    HS_ZAI_MAIN = False

    Exit Function

Error_Proc:
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case ErrDiskNotReady
            Beep
            ans = MsgBox("�h���C�u���m�F���ĉ�����", vbYesNo + vbExclamation + vbDefaultButton1, "�m�F����")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable
            Beep
            ans = MsgBox("�h���C�u��������܂���" & fileName, vbExclamation)
        Case ErrNotFound
            Beep
            ans = MsgBox("�t�@�C����������܂���" & fileName, vbExclamation)
        Case 76
            Beep
            ans = MsgBox("�t�@�C���p�X��������܂���" & fileName, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[ [HS_ZAI Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select
End Function

Private Sub Form_Activate()

Dim ans As Integer
Dim c   As String * 128


    If ZENKAI_YMD = Format(Now, "YYYY/MM/DD") Then
        ans = MsgBox("�{���̍݌Ɏ�荞�ݏ����͏I�����Ă��܂��B���s���܂����H", vbYesNo, "�m�F����")
    Else
        ans = MsgBox("�u�݌ɍ��ك`�F�b�N�����v���s���܂����H", vbYesNo, "�m�F����")
    End If
    
    If ans = vbNo Then
        Unload Me
    End If

        
    F1090251.MousePointer = vbHourglass
    Label1(1).Visible = True
    
    
'''    If SumZ_Init() Then             '���كf�[�^������
'''        Unload Me
'''    End If
    
    If New_SumZ_Init() Then             '���كf�[�^������
        Unload Me
    End If
    
    
    
    Label1(1).Visible = False
    Label1(2).Visible = True
    
    
    If HS_ZAI_MAIN() Then           '���ك`�F�b�N����(���C�����[�v)
        Unload Me
    End If
                                    '�h�m�h�������t�o��
    If WriteIni(App.EXEName, "ZENZENKAI_YMD", "SYS", ZENKAI_YMD) Then
        Beep
        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & " ZENZENKAI_YMD")
        Unload Me
    End If

    If WriteIni(App.EXEName, "ZENKAI_YMD", "SYS", Format(Now, "YYYY/MM/DD")) Then
        Beep
        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & " ZENKAI_YMD")
        Unload Me
    End If



    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer
Dim i       As Integer
Dim j       As Integer
    
Dim Max_Soko    As Integer
    
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                                 
    If JGYOB_TB_Set(1) Then      '���ƕ��̊l��
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
                                
                                
                                
                                
                                '�q�ɍő吔����荞��
                                
    If GetIni(App.EXEName, "MAX_SOKO", "SYS", c) Then
        Max_Soko = 1
    Else
        If Not IsNumeric(RTrim(c)) Then
            Max_Soko = 1
        Else
            Max_Soko = CInt(RTrim(c))
        End If
    End If
                                
                                '�݌Ɏ�荞�ݗp�e�[�u���쐬
    ReDim SOKO_T(0 To UBound(JGYOBU_T), 0 To Max_Soko - 1)
                                '�q�ɏ���荞��
    For i = 0 To UBound(JGYOBU_T)
        j = 0
        Do
                                '�L���q�Ɋl��
            If GetIni(App.EXEName, "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0"), "SYS", c) Then
                Beep
                MsgBox "�q�ɏ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
                End
            End If
    
            If Trim(c) = "**" Then  '�q�Ɏw��I��
                Exit Do
            End If
    
    
'            ReDim Preserve SOKO_T(0 To i, 0 To j)
            SOKO_T(i, j).HS_SOKO = Trim(c)
                                '�����O���l��
            If GetIni(App.EXEName, "NAIG" & JGYOBU_T(i).CODE & Format(j + 1, "0"), "SYS", c) Then
                Beep
                MsgBox "�����O���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
                End
            End If
            
            SOKO_T(i, j).NAIGAI = Trim(c)
            j = j + 1
        Loop
    
    Next i
                                '�݌Ƀt�@�C�����̊l��
    If GetIni("FILE", "HS_ZAI", "SYS", c) Then
        Beep
        MsgBox "�݌Ƀt�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    HS_ZAI = RTrim(c)
                                
                                '�O�񏈗����̊l��
    If GetIni(App.EXEName, "ZENZENKAI_YMD", "SYS", c) Then
        ZENZENKAI_YMD = ""
    Else
        ZENZENKAI_YMD = RTrim(c)
    End If
                                
                                '�O�X�񏈗����̊l��
    If GetIni(App.EXEName, "ZENKAI_YMD", "SYS", c) Then
        ZENKAI_YMD = ""
    Else
        ZENKAI_YMD = RTrim(c)
    End If
                                
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                '�݌ɏW�v�f�[�^�n�o�d�m
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�݌ɏW�v�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌ɏW�v�f�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1090251 = Nothing

    End
End Sub

Private Function SumZ_Total_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u�݌ɏW�v�f�[�^�v�a�t�݌�+PPSC�݌ɂ̏W�v�X�V����
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
    
    
Dim SAI_QTY     As Long
    
Dim ans         As Integer
    
    SumZ_Total_Proc = True
    
    com = BtOpGetFirst
    
    Do
                
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    
                    
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        Loop
        
        If sts Then
            Exit Do
        End If
        
        
        Call UniCode_Conv(SUMZREC.HS_ZAIQTY, Format(CLng(StrConv(SUMZREC.BU_ZAI_QTY, vbUnicode)) + CLng(StrConv(SUMZREC.PPSC_ZAI_QTY, vbUnicode)), "00000000"))
        
        SAI_QTY = CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) - CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
                
If SAI_QTY <> 0 Then
    Debug.Print
End If
        If SAI_QTY >= 0 Then
            Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "00000000"))
        Else
            Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "0000000"))
        End If
        
        
        
        
        
        
        
        
        
        
        Do
            sts = BTRV(BtOpUpdate, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
        
        DoEvents
    Loop

End Function
