VERSION 5.00
Begin VB.Form F9001001 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�ڊǗp���Ɍ��i�[���"
   ClientHeight    =   3690
   ClientLeft      =   2025
   ClientTop       =   2940
   ClientWidth     =   11295
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
   ScaleHeight     =   3690
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   35
      Top             =   480
      Width           =   405
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   7035
      MaxLength       =   2
      TabIndex        =   33
      Top             =   480
      Width           =   405
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   6510
      MaxLength       =   2
      TabIndex        =   31
      Top             =   480
      Width           =   405
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   5985
      MaxLength       =   2
      TabIndex        =   29
      Top             =   480
      Width           =   405
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   5250
      MaxLength       =   2
      TabIndex        =   27
      Top             =   480
      Width           =   405
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   4725
      MaxLength       =   2
      TabIndex        =   25
      Top             =   480
      Width           =   405
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   23
      Top             =   480
      Width           =   405
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   2565
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   21
      Top             =   960
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�p���I��"
      Height          =   975
      Left            =   360
      TabIndex        =   18
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A4"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A5"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   7965
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   17
      Top             =   960
      Width           =   732
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   3675
      MaxLength       =   2
      TabIndex        =   0
      Top             =   480
      Width           =   405
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
      Left            =   10425
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   9585
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   8745
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
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
      Left            =   7905
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   6585
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   5745
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   4905
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
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
      Left            =   4065
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   2745
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   1905
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   1065
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
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
      Index           =   0
      Left            =   225
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   7
      Left            =   7455
      TabIndex        =   34
      Top             =   600
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   6
      Left            =   6930
      TabIndex        =   32
      Top             =   600
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   6405
      TabIndex        =   30
      Top             =   600
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   3
      Left            =   5670
      TabIndex        =   28
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   5145
      TabIndex        =   26
      Top             =   600
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   1
      Left            =   4620
      TabIndex        =   24
      Top             =   600
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   0
      Left            =   4095
      TabIndex        =   22
      Top             =   600
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�������v"
      Height          =   255
      Index           =   12
      Left            =   6885
      TabIndex        =   16
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�I�Ԕ͈�"
      Height          =   255
      Index           =   4
      Left            =   2625
      TabIndex        =   15
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   225
      TabIndex        =   14
      Top             =   2640
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F9001001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NormalFont As New StdFont           '����t�H���g
Dim Code39Font As New StdFont           '����t�H���g


Private Type Print_tbl_tag              '����p�e�[�u��
    NAIGAI          As String * 2
    HIN_GAI         As String * 13
    HIN_NAI         As String * 13
    HIN_NAME        As String * 25
    IRI_QTY         As String * 8
    ST_SOKO         As String * 2
    ST_SOKO_NAME    As String * 5
    ST_RETU         As String * 2
    ST_REN          As String * 2
    ST_DAN          As String * 2
    BIKOU           As String * 15
End Type

Dim Print_tbl(0 To 6, 0 To 1) _
                    As Print_tbl_tag



Dim JGYOBU_NAME As String

Dim Printer_tbl() As String
Dim Max_Gyo     As Integer


Private Const Update_day$ = "2009.03.12"


Private Function Print_Proc(svTanaban As String, SvNaigai As String, SvHin_Gai As String, svNyuka_DT As String, Zaiko_Qty As Long, Gyo As Integer, Retu As Integer) As Integer

Dim Maisu       As Integer
Dim sts         As Integer
Dim flg         As Boolean





    Print_Proc = False






        
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, SvNaigai)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, SvHin_Gai)
    flg = False
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            flg = True
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
        
            
            
    If SvNaigai = NAIGAI_NAI Then
        Print_tbl(Gyo, Retu).NAIGAI = NAIGAI1
    Else
        Print_tbl(Gyo, Retu).NAIGAI = NAIGAI2
    End If
    Print_tbl(Gyo, Retu).HIN_GAI = SvHin_Gai
    If Not flg Then
        Print_tbl(Gyo, Retu).HIN_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
        Print_tbl(Gyo, Retu).HIN_NAME = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Print_tbl(Gyo, Retu).ST_SOKO = Mid(svTanaban, 1, 2)
        Print_tbl(Gyo, Retu).ST_RETU = Mid(svTanaban, 3, 2)
        Print_tbl(Gyo, Retu).ST_REN = Mid(svTanaban, 5, 2)
        Print_tbl(Gyo, Retu).ST_DAN = Mid(svTanaban, 7, 2)

        Call UniCode_Conv(K0_SOKO.Soko_No, Mid(svTanaban, 1, 2))
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                Print_tbl(Gyo, Retu).ST_SOKO_NAME = Left(StrConv(SOKOREC.SOKO_NAME, vbUnicode), 5)
            Case BtErrKeyNotFound
                Print_tbl(Gyo, Retu).ST_SOKO_NAME = " "
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                Beep
                MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
                Unload Me
        End Select
    Else
        Print_tbl(Gyo, Retu).HIN_NAI = " "
        Print_tbl(Gyo, Retu).HIN_NAME = " "
        Print_tbl(Gyo, Retu).ST_SOKO = " "
        Print_tbl(Gyo, Retu).ST_RETU = " "
        Print_tbl(Gyo, Retu).ST_REN = " "
        Print_tbl(Gyo, Retu).ST_DAN = " "
        Print_tbl(Gyo, Retu).ST_SOKO_NAME = " "
    End If

    Print_tbl(Gyo, Retu).IRI_QTY = Zaiko_Qty
    Print_tbl(Gyo, Retu).BIKOU = ""


    
    
        
End Function
                                    
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field()
Dim i As Integer
    
    For i = 0 To 4
        Text(i).Text = ""
    Next i
    Text(8).Text = ""

    Text(9).Text = "0"
    Text(10).Text = "0"
End Sub





Private Sub Command_Click(Index As Integer)

Dim yn              As Integer
Dim sts             As Integer
Dim i               As Integer




Select Case Index
        
        
        
        Case 4
        
            For i = 0 To 7
                Select Case i
                    Case 0, 1, 2, 3
                    
                    
                        If Trim(Text(i).Text) = "" Then
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                    
                    Case 4, 5, 6, 7
                
                        If Trim(Text(i).Text) = "" Then
                            Text(i).Text = "zz"
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                
                
                
                End Select
            Next i
        
            Beep
            yn = MsgBox("�����������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                sts = Maisu_keisan_Proc()
                
                If sts Then
                    Unload Me
                End If
        
            End If
        
            Text(0).SetFocus
        
        
        Case 8                              '���
            
            
            For i = 0 To 7
                Select Case i
                    Case 0, 1, 2, 3
                    
                    
                        If Trim(Text(i).Text) = "" Then
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                    
                    Case 4, 5, 6, 7
                
                        If Trim(Text(i).Text) = "" Then
                            Text(i).Text = "zz"
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                
                
                
                End Select
            Next i
            
            
            
            
            
            Beep
            yn = MsgBox("������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                
                
                
                
                If Print_Main_Proc() Then
                    Unload Me
                End If
                
                
                Printer.EndDoc
            
            
            End If
            
            Text(0).SetFocus
            
        Case 11                             '�I��
            Unload Me
            
        Case Else
            Beep
    End Select
    
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer
Dim Pri_Name    As Printer
Dim DEF         As String
    
    
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�ڊǎ������", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
    
    
    
    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
    
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F9001001.Caption = "�ڊǗp���Ɍ��i�[���(" + RTrim(JGYOBU_T(i).NAME) + ")" & " " & Update_day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i

                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(0) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(0) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(0) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                
                                
                                
                                '�f�t�H���g�p���T�C�Y��荞��
    If GetIni(App.EXEName, "DEF", App.EXEName, c) Then
        c = ""
    End If
    DEF = RTrim(c)
                                
                                '����t�H���g�ݒ�
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
    Set Printer.Font = Code39Font
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F9001001.FontName
        .Size = F9001001.FontSize
    End With
    Set Printer.Font = NormalFont
                                
                                '��ʏ����ݒ�
    
    If DEF = Trim(Option1(0).Caption) Then
        Option1(0).Value = True
        Option1(1).Value = False
    Else
        If DEF = Trim(Option1(1).Caption) Then
            Option1(0).Value = False
            Option1(1).Value = True
        Else
            Option1(0).Value = True
            Option1(1).Value = False
        End If
    End If
    
    
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            Combo1.AddItem Pri_Name.DeviceName
        End If
    Next
    
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName <> Printer.DriverName Then
            Combo1.AddItem Pri_Name.DeviceName
        End If
    Next
    
    
    Combo1.ListIndex = 0
    
    Text(0).SetFocus
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    
    
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "�I�}�X�^")
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
    End If

    End
End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F9001001.Caption = "�ڊǗp���Ɍ��i�[���(" + RTrim(JGYOBU_T(i).NAME) + ")" & " " & Update_day
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub Text_GotFocus(Index As Integer)
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf      As String
Dim i           As Integer
Dim sts         As Integer
Dim sts_QTY     As Integer

    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            Select Case Index
                Case 0, 1, 2, 3
                
                
                    If Trim(Text(Index).Text) = "" Then
                    Else
                        If IsNumeric(Text(Index).Text) Then
                            Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                        End If
                    End If
                
                Case 4, 5, 6, 7
            
                    If Trim(Text(Index).Text) = "" Then
                        Text(Index).Text = "zz"
                    Else
                        If IsNumeric(Text(Index).Text) Then
                            Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                        End If
                    End If
            
            
            
            End Select
        
        
            For i = Index + 1 To 0 Step -1
                If Text(i).Enabled And Not Text(i).Locked Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        
        
        Case vbKeyUp
            For i = Index - 1 To 0 Step -1
                If Text(i).Enabled And Not Text(i).Locked Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        Case vbKeyF5
            Command(4).Value = True
        Case vbKeyF9
            Command(8).Value = True
        Case vbKeyF12
            Command(11).Value = True
    End Select
End Sub


Private Sub Print_Sub_Proc()
                                            
Dim Gyo         As Integer
Dim wk_IRI_QTY  As String * 5
                                            
                                            
                                            
'    Printer.NewPage
                                            
    On Error GoTo Err_Proc
                                            
    For Gyo = 0 To 5
                                            
                                            
        If Len(Trim(Print_tbl(Gyo, 0).HIN_GAI)) = 0 Then
            Exit For
        End If
'------------------------------------------------   1�s��   ------------------
        Set Printer.Font = Code39Font
        Printer.Print Tab(2);
        Printer.Print "*" + Print_tbl(Gyo, 0).HIN_GAI + "*";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(20);
            Printer.Print "*" + Print_tbl(Gyo, 1).HIN_GAI + "*"
        End If
'------------------------------------------------   2�s��   ------------------
        With NormalFont
            .NAME = F9001001.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(4);
        Printer.Print "[" & Trim(JGYOBU_NAME) & "]";
        With NormalFont
            .NAME = F9001001.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(18);
        Printer.Print "[" & Print_tbl(Gyo, 0).NAIGAI & "]";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(53);
            With NormalFont
                .NAME = F9001001.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print "[" & Trim(JGYOBU_NAME) & "]";
            With NormalFont
                .NAME = F9001001.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(67);
            Printer.Print "[" & Print_tbl(Gyo, 1).NAIGAI & "]"
        End If
        Printer.Print
'------------------------------------------------   3�s��   ------------------
        Printer.Print Tab(4);
        Printer.Print "[���Ɍ��i�[]" & "          ";
'        Printer.Print Text(5).Text & "/" & Text(6).Text & "/" & Text(7).Text;
        Printer.Print "          ";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(53);
            Printer.Print "[���Ɍ��i�[]" & "          ";
'            Printer.Print Text(5).Text & "/" & Text(6).Text & "/" & Text(7).Text
            Printer.Print "          "
        End If
'------------------------------------------------   4�s��   ------------------
        With NormalFont
            .NAME = F9001001.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(4);
        Printer.Print "�i��" & "  ";
        Printer.Print Print_tbl(Gyo, 0).HIN_GAI & " (";
        Printer.Print Print_tbl(Gyo, 0).HIN_NAI & ")";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(46);
            Printer.Print "�i��" & "  ";
            Printer.Print Print_tbl(Gyo, 1).HIN_GAI & " (";
            Printer.Print Print_tbl(Gyo, 1).HIN_NAI & ")"
        End If
'------------------------------------------------   5�s��   ------------------
        With NormalFont
            .NAME = F9001001.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(4);
        Printer.Print "�i��  ";
        Printer.Print Print_tbl(Gyo, 0).HIN_NAME;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(53);
            Printer.Print "�i��  ";
            Printer.Print Print_tbl(Gyo, 1).HIN_NAME
        End If
'------------------------------------------------   6�s��   ------------------
        Printer.Print Tab(13);
        Printer.Print "�����F";
        If IsNumeric(Print_tbl(Gyo, 0).IRI_QTY) Then
            wk_IRI_QTY = Right(Format(CLng(Print_tbl(Gyo, 0).IRI_QTY), "###0"), 5)
            wk_IRI_QTY = Space(Len(wk_IRI_QTY) - Len(Trim(wk_IRI_QTY))) & Trim(wk_IRI_QTY)
            
            Printer.Print StrConv(wk_IRI_QTY, vbWide);
        Else
            Printer.Print "�@�@�@�@�@";
        End If
        Printer.Print "  " & Print_tbl(Gyo, 0).BIKOU;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(62);
            Printer.Print "�����F";
            If IsNumeric(Print_tbl(Gyo, 1).IRI_QTY) Then
                wk_IRI_QTY = Right(Format(CLng(Print_tbl(Gyo, 1).IRI_QTY), "###0"), 5)
                wk_IRI_QTY = Space(Len(wk_IRI_QTY) - Len(Trim(wk_IRI_QTY))) & Trim(wk_IRI_QTY)
            
                Printer.Print StrConv(wk_IRI_QTY, vbWide);
            Else
                Printer.Print "�@�@�@�@�@";
            End If
            Printer.Print "  " & Print_tbl(Gyo, 0).BIKOU
        End If
'------------------------------------------------   6�s��   ------------------
        Printer.Print Tab(4);
        Printer.Print "�W�����ɒI  ";
        Printer.Print Print_tbl(Gyo, 0).ST_SOKO & ":";
        Printer.Print Print_tbl(Gyo, 0).ST_SOKO_NAME;
        Printer.Print Tab(37);
        Printer.Print Print_tbl(Gyo, 0).ST_RETU & "-" & Print_tbl(Gyo, 0).ST_REN & "-" & Print_tbl(Gyo, 0).ST_DAN;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(53);
            Printer.Print "�W�����ɒI  ";
            Printer.Print Print_tbl(Gyo, 1).ST_SOKO & ":";
            Printer.Print Print_tbl(Gyo, 1).ST_SOKO_NAME;
            Printer.Print Tab(86);
            Printer.Print Print_tbl(Gyo, 1).ST_RETU & "-" & Print_tbl(Gyo, 1).ST_REN & "-" & Print_tbl(Gyo, 1).ST_DAN
        End If
'------------------------------------------------   7�s��   ------------------
        
        If Gyo <> Max_Gyo Then
        
            With NormalFont
                .NAME = F9001001.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print
            With NormalFont
                .NAME = F9001001.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print
        End If
    Next Gyo

    Exit Sub

Err_Proc:

    If Err.Number = 482 Then
        MsgBox "�v�����^�[�G���[���������܂����B"
    Else
        MsgBox "���s���G���[�F" & Err.Number
    End If
End Sub


Private Function Maisu_keisan_Proc() As Integer


Dim com         As Integer
Dim sts         As Integer

Dim svTanaban   As String * 8
Dim SvNaigai    As String * 1
Dim SvHin_Gai   As String * 20
Dim Maisu       As Integer



    Maisu_keisan_Proc = True

    Call Input_Lock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "��������@�W�v��", Me.hwnd, 0)



    Call UniCode_Conv(K0_ZAIKO.Soko_No, Text(0).Text)
    Call UniCode_Conv(K0_ZAIKO.Retu, Text(1).Text)
    Call UniCode_Conv(K0_ZAIKO.Ren, Text(2).Text)
    Call UniCode_Conv(K0_ZAIKO.Dan, Text(3).Text)
    Call UniCode_Conv(K0_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ZAIKO.NAIGAI, "")
    Call UniCode_Conv(K0_ZAIKO.HIN_GAI, "")
    Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
    Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")


    com = BtOpGetGreaterEqual

    Maisu = 0

    SvNaigai = ""


    Do
        DoEvents
    
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) > _
                    Text(4).Text & Text(5).Text & Text(6).Text & Text(7).Text Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "�݌Ƀf�[�^")
                Exit Function
        End Select
    
    

        If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
        Else
    
    
    
            If Trim(SvNaigai) = "" Then
                svTanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                SvNaigai = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                SvHin_Gai = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
    
                Maisu = Maisu + 1
        
            End If
        
            If svTanaban <> StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Or _
                SvNaigai <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                Trim(SvHin_Gai) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
        
                Maisu = Maisu + 1
                
                svTanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                SvNaigai = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                SvHin_Gai = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
        
            End If
        End If
    
        com = BtOpGetNext
    
    
    Loop

    Text(8).Text = Format(Maisu, "#0")


    Call Input_UnLock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "��������@�W�v�I��", Me.hwnd, 0)

    Maisu_keisan_Proc = False



End Function


Private Function Print_Main_Proc() As Integer


Dim com         As Integer
Dim sts         As Integer

Dim svTanaban   As String * 8
Dim SvNaigai    As String * 1
Dim SvHin_Gai   As String * 20
Dim svNyuka_DT  As String * 8

Dim Zaiko_Qty   As Long

Dim Wk_Printer As Printer


Dim Gyo         As Integer


Dim Retu        As Integer

Dim wk_LOOP      As Integer

    Print_Main_Proc = True

    Call Input_Lock

'�w�蒠�[�p�v�����^���擾
    For Each Wk_Printer In Printers
        If RTrim(Wk_Printer.DeviceName) = RTrim(Combo1.Text) Then
                Set Printer = Wk_Printer
                Exit For
        End If
    Next

    If Option1(0).Value = True Then
        Printer.PaperSize = vbPRPSA5
        Printer.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��
        Max_Gyo = 2
    Else
        Printer.PaperSize = vbPRPSA4
        Printer.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��
        Max_Gyo = 5
    End If



    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "��������@�����", Me.hwnd, 0)


    For Gyo = 0 To UBound(Print_tbl)
        For Retu = 0 To 1
        
            Print_tbl(Gyo, Retu).HIN_GAI = " "
        
        Next Retu
    Next Gyo

    Gyo = 0
    Retu = 0


    For wk_LOOP = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(wk_LOOP).CODE = Last_JGYOBU Then
            JGYOBU_NAME = JGYOBU_T(wk_LOOP).NAME
            Exit For
        End If
    Next wk_LOOP



    Call UniCode_Conv(K0_ZAIKO.Soko_No, Text(0).Text)
    Call UniCode_Conv(K0_ZAIKO.Retu, Text(1).Text)
    Call UniCode_Conv(K0_ZAIKO.Ren, Text(2).Text)
    Call UniCode_Conv(K0_ZAIKO.Dan, Text(3).Text)
    Call UniCode_Conv(K0_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ZAIKO.NAIGAI, "")
    Call UniCode_Conv(K0_ZAIKO.HIN_GAI, "")
    Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
    Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")


    com = BtOpGetGreaterEqual


    SvNaigai = ""


    Do
        DoEvents
    
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) > _
                    Text(4).Text & Text(5).Text & Text(6).Text & Text(7).Text Then
                    
                    
                    
                    
                    If Trim(SvNaigai) <> "" Then
                    
                        If Print_Proc(svTanaban, SvNaigai, SvHin_Gai, svNyuka_DT, Zaiko_Qty, Gyo, Retu) Then
                            Exit Function
                        End If
                        
                        Call Print_Sub_Proc
                    End If
                    
                    
                    
                    
                    
                    
                    
                    Exit Do
                End If
            
            
            Case BtErrEOF
                
                If Trim(SvNaigai) <> "" Then
                
                    If Print_Proc(svTanaban, SvNaigai, SvHin_Gai, svNyuka_DT, Zaiko_Qty, Gyo, Retu) Then
                        Exit Function
                    End If
                    
                    Call Print_Sub_Proc
                End If
                
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "�݌Ƀf�[�^")
                Exit Function
        End Select
    
    

        If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
        Else
    
    
    
            If Trim(SvNaigai) = "" Then
                svTanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                SvNaigai = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                SvHin_Gai = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                svNyuka_DT = ""
                Zaiko_Qty = 0
                                        
                        
        
            End If
        
            If svTanaban <> StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Or _
                SvNaigai <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                Trim(SvHin_Gai) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
        
                
                If Print_Proc(svTanaban, SvNaigai, SvHin_Gai, svNyuka_DT, Zaiko_Qty, Gyo, Retu) Then
                    Exit Function
                End If
                
                
                Retu = Retu + 1
                If Retu > 1 Then
                    Gyo = Gyo + 1
                    If Gyo > Max_Gyo Then
                        Call Print_Sub_Proc
                        Printer.NewPage
                        For Gyo = 0 To Max_Gyo
                            For Retu = 0 To 1
            
                                Print_tbl(Gyo, Retu).HIN_GAI = " "
            
                            Next Retu
                        Next Gyo
            
                        Gyo = 0
                    End If
                    Retu = 0
                End If
                
                
                svTanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                SvNaigai = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                SvHin_Gai = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                svNyuka_DT = ""
                Zaiko_Qty = 0
        
            End If
        
            If Not IsNumeric(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
            Else
                Zaiko_Qty = Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            End If
        
        End If
    
        com = BtOpGetNext
    
    
    Loop



    Call Input_UnLock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "��������@����I��", Me.hwnd, 0)

    Print_Main_Proc = False



End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F9001001.MousePointer = vbHourglass

    Call Ctrl_Lock(F9001001)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F9001001)


    F9001001.MousePointer = vbDefault

End Sub

