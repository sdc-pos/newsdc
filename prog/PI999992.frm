VERSION 5.00
Object = "{E95678BE-E45E-471F-9287-59E8911E479E}#1.5#0"; "ArViewer15j.ocx"
Begin VB.Form PI999992 
   ClientHeight    =   11115
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15855
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   19711
      SectionData     =   "PI999992.frx":0000
   End
End
Attribute VB_Name = "PI999992"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub RunReport(rpt As Object)
    Set ARViewer1.ReportSource = rpt
End Sub

