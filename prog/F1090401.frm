VERSION 5.00
Begin VB.Form F1090401 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�O�؃f�[�^�폜"
   ClientHeight    =   10770
   ClientLeft      =   2130
   ClientTop       =   3135
   ClientWidth     =   17040
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
   ScaleHeight     =   10770
   ScaleWidth      =   17040
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ListBox List1 
      Height          =   8220
      Index           =   0
      ItemData        =   "F1090401.frx":0000
      Left            =   1200
      List            =   "F1090401.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   10095
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   0
      Top             =   360
      Width           =   1692
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�\ ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� �� "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�d����"
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   19
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���t"
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   18
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   17
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i��"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   16
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label LabJIGYO 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   240
      Left            =   240
      TabIndex        =   15
      Top             =   10440
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i��"
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   14
      Top             =   480
      Width           =   732
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1090401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxHin_Gai% = 0    '�i��

Private Const Text_Max% = 0

Private Const pLstHin_Gai% = 0      '�I���

'Private Const LAST_UPDATE_DAY$ = "[F109040] 2018.11.21 10:45"
'Private Const LAST_UPDATE_DAY$ = "[F109040] 2019.01.11 16:00"
'Private Const LAST_UPDATE_DAY$ = "�O�؃f�[�^�폜 [F109040] 2019.06.21 15:15" '��ʊg�� �^�C�g���o�[�ҏW
Private Const LAST_UPDATE_DAY$ = "�O�؃f�[�^�폜 [F109040] 2019.10.31 9:15 �d���捀�ڒǉ�"

Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim sts     As Integer
    
    Select Case Index
        Case 0
            
            Beep
            yn = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                sts = Delete_Proc
                Select Case sts
                    Case False
                    Case True
                        Unload Me
                    Case SYS_CANCEL
                End Select
            
            
            End If
            
            
        Case 4
            
            
            If List_Disp_Proc() Then
                Unload Me
            End If
                    
        
        Case 11
            Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    If List_Disp_Proc() Then
        Unload Me
    End If

End Sub

Private Sub Form_DblClick()
    PrintForm
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
Dim c       As String * 128
Dim sts     As Integer
Dim Work    As String
Dim i       As Integer

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
                                '���������n�o�d�m
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If


                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
        
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            Me.Caption = "(" + RTrim(JGYOBU_T(i).NAME) + ")" & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                
    Unload SubMenu(i)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'2019.01.11    If UnloadMode = vbFormControlMenu Then Cancel = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '���������b�k�n�r�d
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "��������")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    
    
    sts = BTRV(BtOpReset, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
    End If
    
    Set F1090401 = Nothing

    End

End Sub

Private Sub List1_DblClick(Index As Integer)

Dim Edit        As String
Dim ListIndex   As Long
    

    Text(ptxHin_Gai).Text = List1(pLstHin_Gai).List(List1(pLstHin_Gai).ListIndex)
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If
    
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
'   ���ƕ��؂�ւ�
    F1090401.Caption = "�O�؂�f�[�^�폜�i" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)


    If List_Disp_Proc() Then
        Unload Me
    End If


End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub


Private Function List_Disp_Proc() As Integer

Dim com     As Integer
Dim sts     As Integer
Dim Edit    As String

    List_Disp_Proc = True
    
    List1(pLstHin_Gai).Clear
    
    
    '�i�ڃ}�X�^OPEN 2019/10/30
    If ITEM_Open(BtOpenNomal) Then
        Beep
'        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
'        Unload Me
        Exit Function
    End If
    

    Call UniCode_Conv(K0_J_NYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_J_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_J_NYU.HIN_GAI, "")
    
    com = BtOpGetGreater
    Do
        sts = BTRV(com, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
        
        Select Case sts
            Case BtNoErr
            
                If StrConv(J_NYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(J_NYUREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
                      
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "��������")
                Exit Function
        End Select
    
    '2013/9/19 ���t�Ɛ��ʂ�ǉ�
    
        Edit = StrConv(J_NYUREC.HIN_GAI, vbUnicode) _
        & Right("        " & CDec(StrConv(J_NYUREC.JITU_QTY, vbUnicode)), 8) _
        & Space(10) _
        & StrConv(J_NYUREC.INS_DATE, vbUnicode)                    '2018.11.12
                
        'item��jnyuka�̕R�Â� ���ƕ��E�����O�E�i�� 2019/10/30
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(J_NYUREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(J_NYUREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(J_NYUREC.HIN_GAI, vbUnicode))
        
        '�i�ڃ}�X�^GET 2019/10/30
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
        '�d���捀�ڒǉ� 2019/10/30
        If sts = BtNoErr Then
            Edit = Edit & "    " & Trim(StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode))
        End If
        List1(pLstHin_Gai).AddItem Edit
     
        com = BtOpGetNext
     Loop


    '�i�ڃ}�X�^CLOSE 2019/10/30
   
        Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^", 0)
     
    List_Disp_Proc = False

End Function


Private Function Delete_Proc() As Integer

Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer
                                            
    Delete_Proc = True
                                        
    
                                        
    Call UniCode_Conv(K0_J_NYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_J_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_J_NYU.HIN_GAI, Text(ptxHin_Gai).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Delete_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "��������")
                Exit Function
        End Select
    
    
                
    Loop

    Do
        sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Delete_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "��������")
                Exit Function
            End Select
    Loop
        
    Text(ptxHin_Gai).Text = ""
        
    If List_Disp_Proc() Then
        Exit Function
    End If
    
    
    

    Delete_Proc = False
End Function

