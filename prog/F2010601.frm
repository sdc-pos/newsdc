VERSION 5.00
Begin VB.Form F2010601 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�݌ɏƍ������폜"
   ClientHeight    =   7485
   ClientLeft      =   2130
   ClientTop       =   2835
   ClientWidth     =   12810
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
   ScaleHeight     =   7485
   ScaleWidth      =   12810
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   7920
      MaxLength       =   13
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   6120
      MaxLength       =   2
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   4860
      Index           =   0
      ItemData        =   "F2010601.frx":0000
      Left            =   360
      List            =   "F2010601.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   12015
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   0
      Top             =   600
      Width           =   615
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�Ɓ@��"
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���@�@�@��"
      Height          =   255
      Index           =   10
      Left            =   5400
      TabIndex        =   30
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi�O���j"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   29
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�Ɓ@���@���@��"
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   28
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�@��"
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   27
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   26
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   25
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   24
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���`"
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   23
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   22
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   21
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���t�͈�"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   20
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "F2010601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxStart_YY = 0           '�J�n�@�N
Private Const ptxStart_MM = 1           '�J�n�@��
Private Const ptxStart_DD = 2           '�J�n�@��

Private Const ptxEnd_YY = 3             '�I���@�N
Private Const ptxEnd_MM = 4             '�I���@��
Private Const ptxEnd_DD = 5             '�I���@��

Private Const ptxHin_Gai% = 6           '�i��

Private Const Text_Max% = 6

Private Const pLstRireki% = 0        '����

Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim sts     As Integer
Dim i       As Integer
    
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
            
            List1(pLstRireki).SetFocus
        Case 4
        
            For i = ptxStart_YY To ptxEnd_YY
            
                If Text(i).Text = "" Then
                Else
            
                    If Not IsNumeric(Text(i).Text) Then
                        MsgBox "���͂������ڂ̓G���[�ł��B"
                        Text(i).SetFocus
                        Exit Sub
                    Else
                        If Text(i).MaxLength = 2 Then
                            Text(i).Text = Format(CInt(Text(i).Text), "00")
                        End If
                    End If
                End If
            Next i
        
            
            If Len(Trim(Text(ptxEnd_YY).Text & Text(ptxEnd_MM).Text & Text(ptxEnd_DD).Text)) <> 0 Then
                
                If Text(ptxStart_YY).Text & Text(ptxStart_MM).Text & Text(ptxStart_DD).Text > Text(ptxEnd_YY).Text & Text(ptxEnd_MM).Text & Text(ptxEnd_DD).Text Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Exit Sub
                End If
            End If
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
            List1(pLstRireki).SetFocus
        
        Case 11
            Unload Me
    End Select
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
Dim c As String * 128
Dim sts As Integer
Dim Work As String


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
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    
    
    sts = BTRV(BtOpReset, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
    End If
    
    Set F2010601 = Nothing

    End

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub


Private Function List_Disp_Proc() As Integer

Dim com         As Integer
Dim sts         As Integer
Dim Edit        As String

Dim K0_Flg      As Boolean
Dim Skip_Flg    As Boolean



    List_Disp_Proc = True
    
    List1(pLstRireki).Clear


    If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
        K0_Flg = True
        Call UniCode_Conv(K0_IDO.JGYOBU, SENTAKU)
        Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxStart_YY).Text & Text(ptxStart_MM).Text & Text(ptxStart_DD).Text)
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
    Else
        K0_Flg = False
        Call UniCode_Conv(K1_IDO.JGYOBU, SENTAKU)
        Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K1_IDO.HIN_GAI, Text(ptxHin_Gai).Text)
        Call UniCode_Conv(K1_IDO.JITU_DT, "")
        Call UniCode_Conv(K1_IDO.JITU_TM, "")
    End If
    
    F2010601.MousePointer = vbHourglass
    
    com = BtOpGetGreaterEqual
    Do
        
'        DoEvents
        
        If K0_Flg Then
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Else
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        End If
        Select Case sts
            Case BtNoErr
                Skip_Flg = False
                If K0_Flg Then
                    If Len(Trim(Text(ptxEnd_YY).Text & Text(ptxEnd_MM).Text & Text(ptxEnd_DD).Text)) <> 0 Then
                        If (Text(ptxEnd_YY).Text & Text(ptxEnd_MM).Text & Text(ptxEnd_DD).Text) < StrConv(IDOREC.JITU_DT, vbUnicode) Then
                            Exit Do
                        End If
                    End If
                Else
                    If Len(Trim(Text(ptxStart_YY).Text & Text(ptxStart_MM).Text & Text(ptxStart_DD).Text)) <> 0 Then
                        If (Text(ptxStart_YY).Text & Text(ptxStart_MM).Text & Text(ptxStart_DD).Text) > StrConv(IDOREC.JITU_DT, vbUnicode) Then
                            Skip_Flg = True
                        End If
                    End If
                    If Len(Trim(Text(ptxEnd_YY).Text & Text(ptxEnd_MM).Text & Text(ptxEnd_DD).Text)) <> 0 Then
                        If (Text(ptxEnd_YY).Text & Text(ptxEnd_MM).Text & Text(ptxEnd_DD).Text) < StrConv(IDOREC.JITU_DT, vbUnicode) Then
                            Skip_Flg = True
                        End If
                    End If
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌Ɉړ���")
                Exit Function
        End Select
    
    
        If Skip_Flg Then
        Else
            If StrConv(IDOREC.RIRK_ID, vbUnicode) = "B2" Then
        
               Edit = Left(StrConv(IDOREC.JITU_DT, vbUnicode), 4) & "/" & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" & Right(StrConv(IDOREC.JITU_DT, vbUnicode), 2) & " "
               Edit = Edit & Left(StrConv(IDOREC.JITU_TM, vbUnicode), 2) & ":" & Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 3, 2) & ":" & Right(StrConv(IDOREC.JITU_TM, vbUnicode), 2) & " "
               Edit = Edit & StrConv(IDOREC.JGYOBU, vbUnicode) & " " & StrConv(IDOREC.NAIGAI, vbUnicode) & " " & StrConv(IDOREC.HIN_GAI, vbUnicode) & " "
               Edit = Edit & StrConv(IDOREC.MEMO, vbUnicode)
               List1(pLstRireki).AddItem Edit
            
            End If
        End If
        
        com = BtOpGetNext
    Loop
    
    F2010601.MousePointer = vbNormal

    List_Disp_Proc = False

End Function


Private Function Delete_Proc() As Integer

Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer
                                            
    Delete_Proc = True
                                        
    Call UniCode_Conv(K1_IDO.JGYOBU, SENTAKU)
    Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI_NAI)
                                        
    Call UniCode_Conv(K1_IDO.HIN_GAI, Mid(List1(pLstRireki).List(List1(pLstRireki).ListIndex), 25, 13))
    Call UniCode_Conv(K1_IDO.JITU_DT, Mid(List1(pLstRireki).List(List1(pLstRireki).ListIndex), 1, 4) & Mid(List1(pLstRireki).List(List1(pLstRireki).ListIndex), 6, 2) & Mid(List1(pLstRireki).List(List1(pLstRireki).ListIndex), 9, 2))
    Call UniCode_Conv(K1_IDO.JITU_TM, Mid(List1(pLstRireki).List(List1(pLstRireki).ListIndex), 12, 2) & Mid(List1(pLstRireki).List(List1(pLstRireki).ListIndex), 15, 2) & Mid(List1(pLstRireki).List(List1(pLstRireki).ListIndex), 18, 2))
                                        
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        Select Case sts
            Case BtNoErr
                
                Do
                    sts = BTRV(BtOpDelete, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<idoreki.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Delete_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "�݌Ɉړ���")
                            Exit Function
                        End Select
                Loop
                
                Exit Do
            Case BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Delete_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�݌Ɉړ���")
                Exit Function
        End Select
    
    
                
    Loop
    
    
    
        
    If List_Disp_Proc() Then
        Exit Function
    End If
    
    
    Delete_Proc = False
End Function

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim i   As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            
            Select Case Index
                Case ptxStart_MM, ptxStart_DD, ptxEnd_MM, ptxEnd_DD
                
                    If Text(Index).Text = "" Then
                    Else
                        If Not IsNumeric(Text(Index).Text) Then
                            MsgBox "���͂������ڂ̓G���[�ł��B"
                            Exit Sub
                        Else
                            Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                        End If
                
                    End If
            
            
                    If Index = ptxEnd_DD Then
                        If Len(Trim(Text(ptxEnd_YY).Text & Text(ptxEnd_MM).Text & Text(ptxEnd_DD).Text)) = 0 Then
                        Else
                            If Text(ptxStart_YY).Text & Text(ptxStart_MM).Text & Text(ptxStart_DD).Text > Text(ptxEnd_YY).Text & Text(ptxEnd_MM).Text & Text(ptxEnd_DD).Text Then
                                MsgBox "���͂������ڂ̓G���[�ł��B"
                                Exit Sub
                            End If
                        End If
                    End If
            End Select
    
            If Index < Text_Max Then
                For i = Index + 1 To Text_Max
                    If Text(i).Enabled Then
                        Text(i).SetFocus
                        Exit For
                    End If
                Next i
            End If
    End Select
            

End Sub
