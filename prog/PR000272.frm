VERSION 5.00
Object = "{E95678BE-E45E-471F-9287-59E8911E479E}#1.5#0"; "ArViewer15j.ocx"
Begin VB.Form PR000272 
   ClientHeight    =   11625
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   16620
   LinkTopic       =   "Form1"
   ScaleHeight     =   11625
   ScaleWidth      =   16620
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   WindowState     =   2  'ç≈ëÂâª
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   11655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16605
      _ExtentX        =   29289
      _ExtentY        =   20558
      SectionData     =   "PR000272.frx":0000
   End
End
Attribute VB_Name = "PR000272"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub RunReport(rpt As Object)
    Set ARViewer1.ReportSource = rpt
End Sub

