VERSION 5.00
Begin VB.Form ODR30105 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "��]�[���@�ꊇ�o�^"
   ClientHeight    =   2115
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   5445
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   0
      Top             =   225
      Width           =   1485
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   75
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I�@��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3300
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�X�@�V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1350
      TabIndex        =   1
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Lab_Dsp 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   450
      TabIndex        =   5
      Top             =   675
      Width           =   3255
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "��]�[��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   675
      TabIndex        =   4
      Top             =   225
      Width           =   960
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "�����I��"
      Begin VB.Menu SHORI 
         Caption         =   "�X�V"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "�I��"
         Index           =   1
      End
   End
End
Attribute VB_Name = "ODR30105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�R���{�p�Y��
'Private Const pcmbHBUN = 0

'��]�[���̍ő��s��
Private Const Max_Day = 12


'�e�L�X�g�p�Y��
Private Const ptxTOP% = 0
Private Const ptxLAST% = 0

Private Const ptxKIBOU_DT% = 0

'���x���p�Y��
Private Const plabMSG% = 0

'�R�}���h�{�^���p�Y��
Private Const FuncCOR% = 0       '�X�V
Private Const FuncEND% = 1       '�I��

'ListBox�Y��
'Private Const plst_DISP% = 0     '�\���p�f�[�^�@Sort����Key



Dim Init_F      As Integer


Private Function ERR_CHK(Index As Integer)
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String
Dim W_Date      As String

Dim W_Day       As Long

    ERR_CHK = True
    
                        '���͕������`�F�b�N
    'If LenB(StrConv(Text1(Index), vbFromUnicode)) > Text1(Index).MaxLength Then
    '    MsgBox "���͂������ڂ́i�����ӂ�G���[�j�ł��B", vbExclamation
    '    Exit Function
    'End If
    
    Select Case Index
        Case ptxKIBOU_DT%
            Lab_Dsp(plabMSG%) = ""
            W_STR = Trim(Text1(Index))
            
            If W_STR = "" Then
                MsgBox "��]�[���@���ݒ�I", vbExclamation
                Exit Function
            
            Else
            
                If Not IsDate(W_STR) Then
                    MsgBox "���t�G���[�I", vbExclamation
                    Exit Function
                
                End If
                
                W_STR = Format(Trim(Text1(Index)), "yyyy/mm/dd")
                Text1(Index) = W_STR
                DoEvents
                W_Date = Format(Date, "yyyy/mm/dd")
                
                If W_STR < W_Date Then
                    MsgBox "��]�[�� �� �{���G���[�I", vbExclamation
                    Exit Function
                
                End If
                
                W_Day = DateDiff("m", W_Date, W_STR)
                
                If W_Day > Max_Day Then
                
                    MsgBox Max_Day & "�P���ȏ��G���[�I", vbExclamation
                    
                    Exit Function
                End If
            
        End If
            
            
    End Select
    
    
    ERR_CHK = False
End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    ODR30105.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR30105)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR30105)


    ODR30105.MousePointer = vbDefault

End Sub

Private Sub Command1_Click(Index As Integer)
Dim yn      As Integer
Dim X_i     As Integer
Dim W_After     As String

    Select Case Index
    
        Case FuncCOR%
            
            If ERR_CHK(ptxKIBOU_DT) Then
                Text1(ptxKIBOU_DT).SetFocus
                Call Text1_GotFocus(ptxKIBOU_DT)
                Exit Sub
            End If
            
            
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbDefaultButton1 + vbQuestion, "�m�F����")
            'yn = vbYes
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            '�X�V����
            KIBOU_DT = Text1(ptxKIBOU_DT)
            
            Init_F = 0
            ODR30105_Return = False                '�m�F��� �X�V���I��
            Me.Visible = False
            Exit Sub
            
        Case FuncEND%
            If ODR30105_Return = True Then
                'yn = MsgBox("�I�����܂����H", vbYesNo + vbDefaultButton1 + vbQuestion, "�m�F����")
                yn = vbYes
                
                If yn = vbNo Then
                
                    Exit Sub
                End If
            End If
            
            Init_F = 0
            ODR30105_Return = True                '�m�F��ʷ�ݾُI��
            Me.Visible = False
    
    End Select

End Sub

Private Sub Form_Activate()
Dim X_i As Integer

    If Init_F <> 0 Then Exit Sub
    
    ODR30105.Top = ODR30101.Top + (ODR30101.Height - ODR30105.Height) / 2
    
    
    
    ODR30105.Left = ODR30101.Left + (ODR30101.Width - ODR30105.Width) / 2
    
    
    
    
    Text1(ptxKIBOU_DT).SetFocus
    Call Text1_GotFocus(ptxKIBOU_DT)
    
    ODR30105_Return = True
    
    Init_F = 1
    
End Sub

Private Sub Form_Load()
Dim cc As tagINITCOMMONCONTROLSEX
'Dim PanePos(2) As Long

Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer

Dim sBuffer As String * 255
Dim com     As String

Dim W_Date  As String




'�R�����R���g���[��������������
cc.dwSize = Len(cc)
cc.dwICC = ICC_BAR_CLASSES

'�X�e�[�^�X�E�B���h�E���쐬����
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "��]�[���o�^", Me.hwnd, 0)
'�y�C���������
'�Ō�̗v�f��-1�ɂ����
'�e�E�B���h�E�̑S�̂̕��̎c��̕���
'�����I�Ɋ��蓖�Ă�
'PanePos(0) = 200
'PanePos(1) = 300
'PanePos(2) = -1
'Call SendMessageAny(hStatusWnd, SB_SETPARTS, 3, PanePos(0))
Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)


'��ʏ�������
    'Show
    
    'Text1(ptxTANTO_CD).SetFocus
    'Max_Row = 25000
    
    Init_F = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim yn      As Integer

    If UnloadMode <> 0 Then Exit Sub
    'If UnloadMode = 1 Then Exit Sub
    
    yn = MsgBox("�I�����܂����H", vbYesNo + vbDefaultButton1 + vbQuestion, "�m�F����")
    If yn = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    Me.Visible = False
    
End Sub

Private Sub SHORI_Click(Index As Integer)
Dim yn      As Integer


    Select Case Index
    
        Case 0      '�X�V
            Call Command1_Click(FuncCOR)
        
        
        Case 1       '�I��
            Call Command1_Click(FuncEND)
    
    End Select


End Sub


Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index))
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index))
    End If
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim sts As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Text1(Index).Locked = True Then      '���b�N�����ڂȂ珈�����Ȃ�
        Call Tab_Ctrl(Shift)    '�ړ�
        Exit Sub
    End If
                        '���͕������`�F�b�N
    If ERR_CHK(Index) Then
        Call Text1_GotFocus(Index)
        Text1(Index).SetFocus
        Exit Sub
    End If
    
    
    Call Tab_Ctrl(Shift)    '�ړ�
    
End Sub

