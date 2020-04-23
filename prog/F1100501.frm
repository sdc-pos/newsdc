VERSION 5.00
Begin VB.Form F1100501 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "システム終了処理"
   ClientHeight    =   4728
   ClientLeft      =   1908
   ClientTop       =   2424
   ClientWidth     =   7344
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
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
   StartUpPosition =   2  '画面の中央
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "システム終了処理を実行します。"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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

Dim WS_NO       As String * 2           '自端末番号
Dim SERVER_ID   As String * 2           'サーバーＩＤ
                                                

Private Sub Form_Activate()
        
Dim sts As Integer
Dim ans As Integer
        
        
    Beep
    MsgBox "全作業の終了をタスクバーで確認してください。"
    
    Beep
    MsgBox "「クライアントＰＣ」の電源ＯＦＦを確認してください。", vbSystemModal

    Beep
    ans = MsgBox("在庫集計処理を実行しますか？", vbYesNo + vbSystemModal)
    If ans = vbYes Then
                                    '在庫集計ありのバッチ
        sts = Shell("..\exe\F1100501.bat", vbNormalFocus)
        If sts = ZERO Then
            Beep
            MsgBox "[F1100501.bat]日次処理起動に失敗しました。"
            Call Log_Out(LOG_F, "[F1100501.bat]日次処理起動に失敗しました。")
        End If
    Else
                                    '在庫集計なしのバッチ
        sts = Shell("..\exe\F1100502.bat", vbNormalFocus)
        If sts = ZERO Then
            Beep
            MsgBox "[F1100502.bat]日次処理起動に失敗しました。"
            Call Log_Out(LOG_F, "[F1100502.bat]日次処理起動に失敗しました。")
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
'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
'自端末番号取り込み
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> ZERO Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)

'サーバーＩＤ取り込み
    If GetIni("SYSTEM", "SERVER_ID", "SYS", c) Then
        Beep
        MsgBox "サーバーＩＤの獲得に失敗しました。処理を中止して下さい。"
        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [SERVER_ID] READ ERROR")
        End
    End If
    SERVER_ID = RTrim(c)
    
    
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set F1100501 = Nothing

    End
End Sub
