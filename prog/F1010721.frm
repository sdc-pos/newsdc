VERSION 5.00
Begin VB.Form F1010721 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�X�L���i���j���[�Z�b�g�A�b�v����"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   10095
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
   ScaleHeight     =   7230
   ScaleWidth      =   10095
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "�I ��"
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   12
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   6
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "ERR��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   11
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Err_Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "ERR��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   9
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Err_Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label In_Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�S���ҕʃ��j���[��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label In_Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "���j���[�O���[�v��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�X�L���i���j���[�Z�b�g�A�b�v����"
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
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   7680
   End
End
Attribute VB_Name = "F1010721"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Special_Code    As Variant  '2008.08.08
Private Special_Code_F  As Boolean  '2008.08.08




Private Function Update_Proc(Mode As Integer) As Integer

Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer
Dim upd_com         As Integer


Dim In_Count        As Long
Dim Out_Count       As Long
Dim Err_Count       As Long

Dim DISP_INTERVAL   As Long

Dim fileName        As String
Dim FileNo          As Integer

Dim c               As String * 128

        
Dim i               As Integer
        
        
Dim M_JGYOBU        As String
Dim M_NAIGAI        As String
Dim M_MENU_NO       As String
Dim M_DISP          As String
Dim M_YOIN          As String
Dim M_MTS           As String * 8
Dim M_SS            As String * 8
Dim M_YOIN_DISP     As String
Dim M_LOG           As String
        
        
Dim T_TANTO         As String
Dim T_JGYOBU        As String
Dim T_NAIGAI        As String
Dim T_MENU_NO       As String
        
Dim Err_Flg         As Boolean
        

    Update_Proc = True
    Me.MousePointer = vbHourglass
    
    Select Case Mode
        Case 0
'---------------------------------------------  ���j���[�O���[�v�Z�b�g�A�b�v
            MsgLab(1) = "���j���[�O���[�v�Z�b�g�A�b�v�������I�I"
                                                        '���j���[�O���[�v�t���p�X�捞��
            sts = GetIni("FILE", "P_MENU_CSV", "SYS", c)
            If sts <> False Then
                Call LOG_OUT(LOG_F, "SYS.INI [P_MENU_CSV]�ǂݍ��݃G���[ ")
                Exit Function
            End If
            fileName = Trim(c)
            
                
            
            On Error GoTo Error_Proc
                
            FileNo = FreeFile
            Open fileName For Input As #FileNo
            
            On Error GoTo 0
            
            
            
            In_Count = 0
            In_Cnt(0).Caption = Format(In_Count, "#0")
                                                
            Err_Count = 0
            Err_Cnt(0).Caption = Format(Err_Count, "#0")
                                                
                                                
            Do Until EOF(FileNo)

                
                DoEvents
                    
                    
                On Error GoTo Error_Proc
                Input #FileNo, M_JGYOBU, M_NAIGAI, M_MENU_NO, M_DISP, M_YOIN, M_MTS, M_SS, M_YOIN_DISP, M_LOG
                
                On Error Resume Next
                
                In_Count = In_Count + 1
                In_Cnt(0).Caption = Format(In_Count, "#0")
                
                
                
                Err_Flg = False
                
                
                Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(M_YOIN, 1))
                Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(M_YOIN, 1))
                
                sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                    
                        
                        Err_Flg = True
                        
                                                
                        If Special_Code_F Then                          '2008.08.08
                            For i = 0 To UBound(Special_Code)           '2008.08.08
                                If Special_Code(i) = M_YOIN Then        '2008.08.08
                                    Err_Flg = False                     '2008.08.08
                                End If                                  '2008.08.08
                            Next i                                      '2008.08.08
                        End If                                          '2008.08.08
                        
                        
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�v��Ͻ�")
                        Exit Function
                End Select
                
                                
                If Not Err_Flg Then
                
                
                    Call UniCode_Conv(K0_P_MENU.JGYOBU, Trim(M_JGYOBU))
                    Call UniCode_Conv(K0_P_MENU.NAIGAI, Trim(M_NAIGAI))
                    Call UniCode_Conv(K0_P_MENU.MENU_NO, Format(CInt(Trim(M_MENU_NO)), "00"))
                    
                    sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
                    Select Case sts
                        Case BtNoErr
                            upd_com = BtOpUpdate
                        Case BtErrKeyNotFound
                            upd_com = BtOpInsert
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�V�ƭ�Ǘ�Ͻ�")
                            Exit Function
                    End Select
                
                    If upd_com = BtOpInsert Then
                    
                        Call UniCode_Conv(P_MENUREC.JGYOBU, Trim(M_JGYOBU))
                        Call UniCode_Conv(P_MENUREC.NAIGAI, Trim(M_NAIGAI))
                        Call UniCode_Conv(P_MENUREC.MENU_NO, Format(CInt(Trim(M_MENU_NO)), "00"))
                        Call UniCode_Conv(P_MENUREC.MENU_DSP, Trim(M_DISP))
                                            
                        For i = 0 To 19
                        
                            Call UniCode_Conv(P_MENUREC.SAGYO(i).YOIN, "")
                            Call UniCode_Conv(P_MENUREC.SAGYO(i).PARAM, "")
                            Call UniCode_Conv(P_MENUREC.SAGYO(i).Disp, "")
                        
                        Next i
                                            
                        Call UniCode_Conv(P_MENUREC.FILLER, "")
                    
                    End If
                                        
                    For i = 0 To 19
                    
                    
                        If Trim(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)) = Trim(M_YOIN) And _
                            StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode) = M_MTS & M_SS Then
                            Exit For
                        End If
                    
                        If Trim(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)) = "" Then
                            Exit For
                        End If
                            
                    Next i
                
                
                
                    If i > 19 Then
                        Err_Flg = True
                    Else
                                
                        Call UniCode_Conv(P_MENUREC.SAGYO(i).YOIN, Trim(M_YOIN))
                        Call UniCode_Conv(P_MENUREC.SAGYO(i).PARAM, M_MTS & M_SS)
                        Call UniCode_Conv(P_MENUREC.SAGYO(i).Disp, Trim(M_YOIN_DISP))
                
                        If Trim(M_LOG) = "*" Then
                            Call UniCode_Conv(P_MENUREC.SAGYO(i).LOG_OUT, "0")
                        Else
                            Call UniCode_Conv(P_MENUREC.SAGYO(i).LOG_OUT, "1")
                        End If
                        
                        sts = BTRV(upd_com, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
                        Select Case sts
                            Case BtNoErr
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "�V�ƭ��Ǘ�Ͻ�")
                                Exit Function
                        End Select
                
                    End If
            
                End If
            
                If Err_Flg Then
                    Err_Count = Err_Count + 1
                    Err_Cnt(0).Caption = Format(Err_Count, "#0")
                End If
            Loop
        
            MsgBox "����I�����܂���"
    
    
    
        Case 1
'---------------------------------------------  �S���ҕʃ��j���[�Z�b�g�A�b�v
            MsgLab(1) = "�S���ҕʃ��j���[�Z�b�g�A�b�v�������I�I"
                                                        '���j���[�O���[�v�t���p�X�捞��
            sts = GetIni("FILE", "P_TANTOMENU_CSV", "SYS", c)
            If sts <> False Then
                Call LOG_OUT(LOG_F, "SYS.INI [P_TANTOMENU_CSV]�ǂݍ��݃G���[ ")
                Exit Function
            End If
            fileName = Trim(c)
            
                
            
            On Error GoTo Error_Proc
                
            FileNo = FreeFile
            Open fileName For Input As #FileNo
            
            On Error GoTo 0
            
            
            
            In_Count = 0
            In_Cnt(1).Caption = Format(In_Count, "#0")
                                                
            Err_Count = 0
            Err_Cnt(1).Caption = Format(Err_Count, "#0")
                                                
                                                
            Do Until EOF(FileNo)

                
                DoEvents
                    
                    
                On Error GoTo Error_Proc
                Input #FileNo, T_TANTO, T_JGYOBU, T_NAIGAI, T_MENU_NO
                On Error GoTo 0
                
                In_Count = In_Count + 1
                In_Cnt(1).Caption = Format(In_Count, "#0")
                
                Err_Flg = False
                
                
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, T_TANTO)
                
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                    
                        Err_Flg = True
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�S����Ͻ�")
                        Exit Function
                End Select
                
                Call UniCode_Conv(K0_P_MENU.JGYOBU, T_JGYOBU)
                Call UniCode_Conv(K0_P_MENU.NAIGAI, T_NAIGAI)
                Call UniCode_Conv(K0_P_MENU.MENU_NO, Format(CInt(T_MENU_NO), "00"))
                
                
                sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                    
                        Err_Flg = True
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�ƭ�Ǘ�Ͻ�")
                        Exit Function
                End Select
                
                
                
                                
                If Not Err_Flg Then
                
                
                    Call UniCode_Conv(K0_P_TMENU.TANTO_CODE, Trim(T_TANTO))
                    
                    sts = BTRV(BtOpGetEqual, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
                    Select Case sts
                        Case BtNoErr
                            upd_com = BtOpUpdate
                        Case BtErrKeyNotFound
                            upd_com = BtOpInsert
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�V�S���ҕ��ƭ�")
                            Exit Function
                    End Select
                
                    If upd_com = BtOpInsert Then
                        Call UniCode_Conv(P_TMENUREC.TANTO_CODE, Trim(T_TANTO))
                                            
                        For i = 0 To 179
                        
                            Call UniCode_Conv(P_TMENUREC.MENU_T(i).JGYOBU, "")
                            Call UniCode_Conv(P_TMENUREC.MENU_T(i).NAIGAI, "")
                            Call UniCode_Conv(P_TMENUREC.MENU_T(i).MENU_NO, "")
                        
                        Next i
                                            
                        Call UniCode_Conv(P_TMENUREC.FILLER, "")
                    
                    End If
                                        
                    For i = 0 To 179
                    
                        If Trim(StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)) = "" Then
                            Exit For
                        End If
                            
                    Next i
                
                
                
                    If i > 179 Then
                        Err_Flg = True
                    Else
                        Call UniCode_Conv(P_TMENUREC.MENU_T(i).JGYOBU, Trim(T_JGYOBU))
                        Call UniCode_Conv(P_TMENUREC.MENU_T(i).NAIGAI, Trim(T_NAIGAI))
                        Call UniCode_Conv(P_TMENUREC.MENU_T(i).MENU_NO, Trim(T_MENU_NO))
                
                        sts = BTRV(upd_com, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
                        Select Case sts
                            Case BtNoErr
                            Case Else
                                Call File_Error(sts, upd_com, "�V�S���ҕ��ƭ�")
                                Exit Function
                        End Select
                
                    End If
            
                End If
            
                If Err_Flg Then
                    Err_Count = Err_Count + 1
                    Err_Cnt(1).Caption = Format(Err_Count, "#0")
                End If
            Loop
        
            MsgBox "����I�����܂���"
    
    
    
    
    End Select
'---------------------------------------------  �I��
    Me.MousePointer = vbDefault

    Update_Proc = False
    
    Exit Function

Error_Proc:
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case 62
            MsgBox "����I�����܂���"
            Update_Proc = False
            Exit Function
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
            ans = MsgBox("�G���[ [PACKING_CSV Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select

End Function


Private Sub Command1_Click(Index As Integer)


    Select Case Index
        Case 0, 1
            If Update_Proc(Index) Then
                Unload Me
            End If
        Case 2
            Unload Me
    End Select
End Sub

Private Sub Form_DblClick()
    PrintForm
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
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�V�ƭ�Ǘ��}�X�^OPEN
    If P_MENU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�V�S���ҕ��ƭ�Ǘ��}�X�^OPEN
    If P_TMENU_Open(BtOpenNomal) Then
        Unload Me
    End If

    
    Special_Code_F = False                                  '2008.08.08
    If GetIni(App.EXEName, "Special_CD", "SYS", c) Then     '2008.08.08
    Else                                                    '2008.08.08
        Special_Code_F = True                               '2008.08.08
        Special_Code = Split(Trim(c), ",", -1)              '2008.08.08
    End If                                                  '2008.08.08


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
    
                                            '�V�ƭ��Ǘ��}�X�^CLOSE
    sts = BTRV(BtOpClose, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�V�ƭ��Ǘ�")
        End If
    End If
                                            '�V�S���ҕ��ƭ�CLOSE
    sts = BTRV(BtOpClose, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�V�ƭ��Ǘ�")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1010721 = Nothing

    End
End Sub

