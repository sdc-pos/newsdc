VERSION 5.00
Begin VB.Form F3010101 
   BackColor       =   &H00C0C0C0&
   Caption         =   "GLICS �݌Ɏ�荞��V1.00"
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
      Caption         =   "�f�[�^�W�v��"
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
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "GLICS�݌Ɏ�荞�ݏ���"
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
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   5040
   End
End
Attribute VB_Name = "F3010101"
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
Private Function HS_ZAI_INIT() As Integer


Dim In_Rec      As String
Dim In_Text     As Variant
Dim c           As String * 128
Dim FileNo      As Integer
Dim fileName    As String
    
Dim ans         As Integer
    
Dim com         As Integer
Dim sts         As Integer
    
Dim upd_com     As Integer
Dim T_Zai_Qty       As Long
Dim SAI_QTY         As Long
    
    
Dim Skip_Flg    As Boolean
    
Dim i           As Integer
Dim j           As Integer
    
    HS_ZAI_INIT = True
    
                                '�����N���A�[
    com = BtOpGetFirst
    Do
        DoEvents
                
        Do
            sts = BTRV(com + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrEOF
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                HS_ZAI_INIT = True
            End Select
        Loop


        If sts = BtErrEOF Then
            Exit Do
        End If

        Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "00000000")
        Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "00000000")
        
        Do
            sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                HS_ZAI_INIT = True
            End Select
        Loop
                                
        com = BtOpGetNext
    Loop
                                
                                
                                
                                
    com = BtOpGetGreater

    Call UniCode_Conv(K0_SUMZ.JGYOBU, SENTAKU)
    Call UniCode_Conv(K0_SUMZ.NAIGAI, "")
    Call UniCode_Conv(K0_SUMZ.HIN_GAI, "")

    Do
                
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(SUMZREC.JGYOBU, vbUnicode) <> SENTAKU Then
                        sts = BtErrEOF
                    End If
                    
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
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                '�t�@�C������荞��
    If GetIni("FILE", "HS_NEW_ZAI", "SYS", c) Then
        Beep
        MsgBox "�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    fileName = RTrim(c)
    
    
    
    FileNo = FreeFile
        
    On Error GoTo HS_SIJ_Op_Err    '�װ�ׯ��ON
    Open fileName For Input As #FileNo
    On Error GoTo 0
        
        
        
    Do While Not EOF(FileNo)
        
        DoEvents
        Line Input #FileNo, In_Rec
        In_Text = Split(In_Rec, vbTab, -1)
    
        If CStr(In_Text(0)) = "00023100" Then
    
            Call UniCode_Conv(K0_ITEM.JGYOBU, SENTAKU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
            Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(In_Text(1)))
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    HS_ZAI_INIT = True
                End Select
            Loop
    
    
            If sts = BtNoErr Then
                Select Case In_Text(2)
                    Case "S2"
                        If IsNumeric(In_Text(3)) Then
                            Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, Format(CLng(In_Text(3)), "00000000"))
                        Else
                            Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "00000000")
                        End If
                    Case "P2"
                        If IsNumeric(In_Text(3)) Then
                            Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, Format(CLng(In_Text(3)), "00000000"))
                        Else
                            Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "00000000")
                        End If
                End Select
                Call UniCode_Conv(ITEMREC.S_TANTO, CStr(In_Text(7)))
    
    
                Do
                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                        HS_ZAI_INIT = True
                    End Select
                Loop
    
    
            End If
    
    
            Skip_Flg = False
            If Not IsNumeric(In_Text(3)) Then
                Skip_Flg = True
            End If
        
            
            If Not Skip_Flg Then
''                Skip_Flg = True
''                For i = 0 To UBound(JGYOBU_T)
''                    If SENTAKU = JGYOBU_T(i).CODE Then
''                        For j = 0 To UBound(SOKO_T, 2)
''                            If In_Text(2) = SOKO_T(i, j).HS_SOKO Then
''                                Skip_Flg = False
''                                Exit For
''                            End If
''                        Next j
''                        Exit For
''                    End If
''                Next i
        
                If Not Skip_Flg Then
                    '�Ώۃf�[�^
                    Call UniCode_Conv(K0_SUMZ.JGYOBU, SENTAKU)
                    Call UniCode_Conv(K0_SUMZ.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_SUMZ.HIN_GAI, CStr(In_Text(1)))
            
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                        Select Case sts
                            Case BtNoErr
                                upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                upd_com = BtOpInsert
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
                
                    If upd_com = BtOpInsert Then
                    '�V�K�ǉ����A�i�ڃ}�X�^���z�X�g�I�Ԋl��
                    
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SENTAKU)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, SOKO_T(i, j).NAIGAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(In_Text(1)))
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
            
            
                        Call UniCode_Conv(SUMZREC.JGYOBU, SENTAKU)
                        Call UniCode_Conv(SUMZREC.NAIGAI, SOKO_T(i, j).NAIGAI)
                        Call UniCode_Conv(SUMZREC.HIN_GAI, CStr(In_Text(1)))
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
                        Call UniCode_Conv(SUMZREC.FILLER, "")
                        
                    End If
                    
                    Call UniCode_Conv(SUMZREC.SUM_DT, Format(Date, "yyyymmdd"))
            
                    Call UniCode_Conv(SUMZREC.HS_ZAIQTY, Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) + CLng(CStr(In_Text(3))), "00000000"))
            
                    SAI_QTY = CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) - CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
                    
                    If SAI_QTY >= 0 Then
                        Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "00000000"))
                    Else
                        Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "0000000"))
                    End If
            
                    Do
                    
                        sts = BTRV(upd_com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
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
                                Call File_Error(sts, upd_com, "�݌ɏW�v�f�[�^")
                                Exit Function
                        End Select
                
                    Loop
                End If
            End If
        End If
    
    
    Loop

    Close #FileNo
    Exit Function
HS_SIJ_Op_Err:     '�װ����ٰ��
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
            ans = MsgBox("�h���C�u�܂��̓p�X��������܂���" & fileName, vbExclamation)
        Case ErrNotFound
            Beep
            ans = MsgBox("�t�@�C����������܂���" & fileName, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[ [HS_SIJ Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select
End Function

Private Sub Form_Activate()

Dim ans As Integer

'    ans = MsgBox("�uGLICS�݌Ɏ�荞�ݏ����v���s���܂����H", vbYesNo, "�m�F����")
'
'    If ans = vbNo Then
'        Unload Me
'    End If

        
    F3010101.MousePointer = vbHourglass
    Label1(1).Visible = True
    
    If HS_ZAI_INIT() Then           '���ك`�F�b�N����
        Unload Me
    End If
    
    Label1(1).Visible = False
    

    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()
Dim c   As String * 128
Dim sts As Integer
Dim i   As Integer
Dim j   As Integer

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
                                
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If

                                '�i�ڃ}�X�^�n�o�d�m
    If SUMZ_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
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
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F3010101 = Nothing

    End
End Sub


