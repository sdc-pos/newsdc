VERSION 5.00
Begin VB.Form F1030001 
   Caption         =   "���Ɂ^�I�ԃ`�F�b�N���X�g�@(F103000 2011.07.14 12:00)"
   ClientHeight    =   3312
   ClientLeft      =   2028
   ClientTop       =   2268
   ClientWidth     =   10932
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3312
   ScaleWidth      =   10932
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1030001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim NormalFont      As New StdFont              '����t�H���g
Dim NormalBoldFont  As New StdFont              '����t�H���g
Dim SmallFont       As New StdFont              '����t�H���g
Dim SmallBoldFont   As New StdFont              '����t�H���g
Dim LargeFont       As New StdFont              '����t�H���g
Dim LargeUnderFont  As New StdFont              '����t�H���g

Private Const MGN_L% = 2                        '���]���i�����F�P����j
Private Const MGN_U% = 2                        '��]���i�s���F�P����j
Private Const LMax% = 6



Private Sub Form_Activate()

    If Print_Proc() Then
        Unload Me
    End If


    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)


                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

                                '���׃f�[�^�t�@�C��OPEN
    If Y_NYU_Open(0) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�t�@�C��OPEN
    If ZAIKO_Open(0) Then
        Unload Me
    End If
                                '�����σf�[�^�t�@�C��OPEN
    If AVE_SYUKA_Open(0) Then
        Unload Me
    End If
    Show


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '���׃f�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���׃f�[�^�t�@�C��")
        End If
    End If
                                            '�݌Ƀf�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^�t�@�C��")
        End If
    End If
                                            '�����Ϗo�ׂb�k�n�r�d
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����Ϗo��")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1030001 = Nothing

    End
End Sub


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1030001.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030001)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030001)


    F1030001.MousePointer = vbDefault

End Sub
Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   �������
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim i           As Integer

Dim Lcnt        As Integer
Dim Pcnt        As Integer
    
Dim Print_Now   As String
    
    
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
    
    
Dim Read_Next   As Integer
    
    
    Print_Proc = True

    Print_Now = Format(Now, "YYYY/MM/DD HH:MM")
    
    Printer.Orientation = vbPRORLandscape
    Lcnt = 99
    Pcnt = 0
    
    With NormalFont
        .NAME = F1030001.Font.NAME
        .Size = 11
        .Bold = False
    
    End With
    With NormalBoldFont
        .NAME = F1030001.Font.NAME
        .Size = 11
        .Bold = True
    End With
    
    With SmallFont
        .NAME = F1030001.Font.NAME
        .Size = 9
        .Bold = False
    End With
    
    With SmallBoldFont
        .NAME = F1030001.Font.NAME
        .Size = 9
        .Bold = True
    End With
    
    
    With LargeFont
        .NAME = F1030001.Font.NAME
        .Size = 14
        .Bold = True
    End With
    
    With LargeUnderFont
        .NAME = F1030001.Font.NAME
        .Size = 14
        .Bold = True
        .Underline = True
    End With
    
    
    
    com = BtOpGetGreater


    Call UniCode_Conv(K4_Y_NYU.LIST_OUT_END_F, "0")
    Call UniCode_Conv(K4_Y_NYU.JGYOBU, "")
    Call UniCode_Conv(K4_Y_NYU.NAIGAI, "")
    Call UniCode_Conv(K4_Y_NYU.HIN_NO, "")



    Do
        DoEvents
        
        sts = BTRV(com, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
        
        
        Select Case sts
            Case BtNoErr

                If StrConv(Y_NYUREC.LIST_OUT_END_F, vbUnicode) <> "0" Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
        
        
            Case Else
                Call File_Error(sts, com, "���ח\��f�[�^", 0)
                Exit Function

        End Select
        
        
        If StrConv(Y_NYUREC.LIST_NYU_CHECK_F, vbUnicode) <> "0" Then
        Else
        
        
            '2010.12.17
            Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "03000")
            Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
            '2010.12.17
        
        
            
            If Head_Print_Proc(Print_Now, Lcnt, Pcnt) Then
                Unload Me
            End If
            
            
            Set Printer.Font = SmallFont
                
                
            Printer.Print Tab(MGN_L);
            Printer.Print StrConv(Y_NYUREC.JGYOBU, vbUnicode);
            Printer.Print Tab(MGN_L + 3);
            Printer.Print StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode);
            Printer.Print Tab(MGN_L + 14);
            Printer.Print Mid(StrConv(Y_NYUREC.HIN_NO, vbUnicode), 1, 13);
            
            
            Printer.Print Tab(MGN_L + 27);
            Printer.Print Mid(StrConv(Y_NYUREC.HIN_NAI, vbUnicode), 1, 13);
            
            
            Printer.Print Tab(MGN_L + 40);
            Printer.Print StrConv(Y_NYUREC.HIN_NAME, vbUnicode);
            Printer.Print Tab(MGN_L + 65);
            Printer.Print Space(7 - Len(Format(CLng(StrConv(Y_NYUREC.SURYO, vbUnicode)), "#,##0"))) & _
                            Format(CLng(StrConv(Y_NYUREC.SURYO, vbUnicode)), "#,##0");
            Printer.Print Tab(MGN_L + 77);
            Printer.Print StrConv(Y_NYUREC.DEN_NO, vbUnicode);
            Printer.Print Tab(MGN_L + 87);
            Printer.Print StrConv(Y_NYUREC.HTANABAN, vbUnicode);
    
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(Y_NYUREC.JGYOBU, vbUnicode), _
                                    StrConv(Y_NYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_NYUREC.HIN_NO, vbUnicode)) Then
                Exit Function
            End If
            Printer.Print Tab(MGN_L + 97);
            Printer.Print Space(10 - Len(Format(MI_QTY, "#,##0"))) & Format(MI_QTY, "#,##0");
            Printer.Print Tab(MGN_L + 107);
            Printer.Print Space(10 - Len(Format(SUMI_QTY, "#,##0"))) & Format(SUMI_QTY, "#,##0");
    
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(Y_NYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(Y_NYUREC.HIN_NO, vbUnicode))
            
            sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            
            
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(AVE_SYUKAREC.AVE_SYUKA, "0")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א�", 0)
                    Exit Function
        
            End Select
            Printer.Print Tab(MGN_L + 117);
            Printer.Print Space(10 - Len(Format(CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode)), "#,0"))) & Format(CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode)), "#,0");
            Printer.Print Tab(MGN_L + 129);
            Printer.Print StrConv(Y_NYUREC.GENSANKOKU, vbUnicode);
            Printer.Print Tab(MGN_L + 152);
            If StrConv(Y_NYUREC.LIST_NYU_KANRI_F, vbUnicode) = "0" Or StrConv(Y_NYUREC.LIST_NYU_KANRI_F, vbUnicode) = "9" Then
                Printer.Print "�ؑ�";
            End If
            Printer.Print Tab(MGN_L + 157);
            Printer.Print StrConv(Y_NYUREC.SHIIRE_WORK_CENTER, vbUnicode);
            Printer.Print Tab(MGN_L + 167);
            Printer.Print StrConv(Y_NYUREC.AITESAKI_CODE, vbUnicode)
    
    
            Lcnt = Lcnt + 1

        
        
        
        
        End If



                
        Call UniCode_Conv(Y_NYUREC.LIST_NYU_CHECK_F, "9")
        
        
        
        
        
        If (StrConv(Y_NYUREC.LIST_NYU_KANRI_F, vbUnicode) = "9" Or StrConv(Y_NYUREC.LIST_NYU_KANRI_F, vbUnicode) = "8") And StrConv(Y_NYUREC.LIST_NYU_CHECK_F, vbUnicode) = "9" Then
            
            
            
            Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "9")
        
        
        
        
            Read_Next = BtOpGetGreater
        
        
        
        Else
        
            Read_Next = BtOpGetNext
        
        End If
        
        
        
        
        
        sts = BTRV(BtOpUpdate, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
        
        If sts <> BtNoErr Then
        
            Call File_Error(sts, BtOpUpdate, "���ח\��f�[�^", 0)
            Exit Function

        End If



        If Read_Next = BtOpGetGreater Then
            Call UniCode_Conv(K4_Y_NYU.LIST_OUT_END_F, "0")
            Call UniCode_Conv(K4_Y_NYU.JGYOBU, "")
            Call UniCode_Conv(K4_Y_NYU.NAIGAI, "")
            Call UniCode_Conv(K4_Y_NYU.HIN_NO, "")
        End If
        com = Read_Next


    Loop

    If Lcnt <> 99 Then
        Set Printer.Font = SmallFont
        Printer.Print Tab(MGN_L);
        Printer.Print String(90, "��")
    End If

    Printer.EndDoc

    Print_Proc = False

End Function


Private Function Head_Print_Proc(Print_Now As String, Lcnt As Integer, Pcnt As Integer) As Integer
'----------------------------------------------------------------------------
'                   �w�b�_�[����
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

    
    Head_Print_Proc = True


    If Lcnt < 50 Then
        Head_Print_Proc = False
        Exit Function
    End If

    Pcnt = Pcnt + 1
    If Lcnt = 99 Then
    Else
        Set Printer.Font = SmallFont
        Printer.Print Tab(MGN_L);
        Printer.Print String(90, "��")
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    
    '------------------------------------   1�s��
    Set Printer.Font = LargeFont
    Printer.Print Tab(MGN_L);
    Printer.Print "���Ɂ^�I�ԃ`�F�b�N���X�g";

    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 150);
    Printer.Print Print_Now;
    Printer.Print " Page." & Format(Pcnt, "#0")
    
    
    '------------------------------------   6�s��
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L);
    Printer.Print String(90, "��")
    
    
    '------------------------------------   6�s��
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 99);
    Printer.Print "�݌ɐ�";
    
    Printer.Print Tab(MGN_L + 108);
    Printer.Print " �݌ɐ�";
    
    
    Printer.Print Tab(MGN_L + 119);
    Printer.Print "������";
    
    
    Printer.Print Tab(MGN_L + 157);
    Printer.Print "�d����";
    
    Printer.Print Tab(MGN_L + 167);
    Printer.Print "�����"
    
    '------------------------------------   6�s��
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L);
    Printer.Print "��";
    Printer.Print Tab(MGN_L + 3);
    Printer.Print "���ɓ�";
    Printer.Print Tab(MGN_L + 14);
    Printer.Print "�i��";
    Printer.Print Tab(MGN_L + 27);
    Printer.Print "�Γ��i��";
    Printer.Print Tab(MGN_L + 40);
    Printer.Print "�i��";
    Printer.Print Tab(MGN_L + 65);
    Printer.Print "����";
    Printer.Print Tab(MGN_L + 77);
    Printer.Print "�`�[��";
    Printer.Print Tab(MGN_L + 87);
    Printer.Print "�W���I��";
    
    Printer.Print Tab(MGN_L + 99);
    Printer.Print "�����i";
    
    Printer.Print Tab(MGN_L + 108);
    Printer.Print "���i����";
    
    
    Printer.Print Tab(MGN_L + 119);
    Printer.Print "�o�א�";
    
    Printer.Print Tab(MGN_L + 129);
    Printer.Print "���Y��";
    
    Printer.Print Tab(MGN_L + 152);
    Printer.Print "�ؑ�";
    
    Printer.Print Tab(MGN_L + 157);
    Printer.Print "W/C";
    
    Printer.Print Tab(MGN_L + 167);
    Printer.Print "�R�[�h"
    '------------------------------------   6�s��
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L);
    Printer.Print String(90, "��")

    Lcnt = 0

    Head_Print_Proc = False

End Function
