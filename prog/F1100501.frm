VERSION 5.00
Begin VB.Form F1100501 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�V�X�e���I������"
   ClientHeight    =   4728
   ClientLeft      =   1908
   ClientTop       =   2424
   ClientWidth     =   7344
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4728
   ScaleWidth      =   7344
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "�V�X�e���I�����������s���܂��B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   22.2
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   732
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6852
   End
End
Attribute VB_Name = "F1100501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WS_NO       As String * 2           '���[���ԍ�
Dim SERVER_ID   As String * 2           '�T�[�o�[�h�c
                                                

Private Sub Form_Activate()
        
Dim sts As Integer
Dim ans As Integer
        
        
    Beep
    MsgBox "�S��Ƃ̏I�����^�X�N�o�[�Ŋm�F���Ă��������B"
    
    Beep
    MsgBox "�u�N���C�A���g�o�b�v�̓d���n�e�e���m�F���Ă��������B", vbSystemModal

    Beep
    ans = MsgBox("�݌ɏW�v���������s���܂����H", vbYesNo + vbSystemModal)
    If ans = vbYes Then
                                    '�݌ɏW�v����̃o�b�`
        sts = Shell("..\exe\F1100501.bat", vbNormalFocus)
        If sts = ZERO Then
            Beep
            MsgBox "[F1100501.bat]���������N���Ɏ��s���܂����B"
            Call Log_Out(LOG_F, "[F1100501.bat]���������N���Ɏ��s���܂����B")
        End If
    Else
                                    '�݌ɏW�v�Ȃ��̃o�b�`
        sts = Shell("..\exe\F1100502.bat", vbNormalFocus)
        If sts = ZERO Then
            Beep
            MsgBox "[F1100502.bat]���������N���Ɏ��s���܂����B"
            Call Log_Out(LOG_F, "[F1100502.bat]���������N���Ɏ��s���܂����B")
        End If
    End If

    Unload Me
End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()

Dim c As String * 128
Dim sts As Integer
    
Dim sBuffer     As String * 255
Dim com         As String
    
    Show
'���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
'���[���ԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> ZERO Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)

'�T�[�o�[�h�c��荞��
    If GetIni("SYSTEM", "SERVER_ID", "SYS", c) Then
        Beep
        MsgBox "�T�[�o�[�h�c�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [SERVER_ID] READ ERROR")
        End
    End If
    SERVER_ID = RTrim(c)
    
    
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set F1100501 = Nothing

    End
End Sub
