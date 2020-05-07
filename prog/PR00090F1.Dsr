VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00090F1 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   18075
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   31882
   _ExtentY        =   19420
   SectionData     =   "PR00090F1.dsx":0000
End
Attribute VB_Name = "PR00090F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private DET_com As Integer




Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "HAKKO_DT"            '発行日
    Me.Fields.Add "SHIMUKE_CODE"        '仕向け先
    Me.Fields.Add "SHIJI_NO"            '指図票№
    Me.Fields.Add "HIN_GAI"             '品番
    Me.Fields.Add "SHIJI_QTY"           '数量
    Me.Fields.Add "DOUKON"              '同梱件数
    Me.Fields.Add "KAN_DT"              '完了日
    
    
    
    
    


End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)




    If DET_com > SSHIJI.UpperBound(1) Then
        Exit Sub
    End If
    

    '発行日
    Me.Fields("HAKKO_DT") = SSHIJI(DET_com, colHAKKO_DT)
    '仕向先
    Me.Fields("SHIMUKE_CODE") = SSHIJI(DET_com, colSHIMUKE_CODE)
    '指図票№
    Me.Fields("SHIJI_NO") = Format(CLng(SSHIJI(DET_com, colSHIJI_NO)), "00000000")
    '品番
    Me.Fields("HIN_GAI") = SSHIJI(DET_com, colHIN_GAI)
    '指示数
    Me.Fields("SHIJI_QTY") = Format(CLng(SSHIJI(DET_com, colSHIJI_QTY)), "#0")
    '同梱数
    Me.Fields("DOUKON") = Format(CLng(SSHIJI(DET_com, colDOUKON)), "#0")
    '完了日
    Me.Fields("KAN_DT") = SSHIJI(DET_com, colKAN_DT)
    
    
    DET_com = DET_com + 1
    
            
    eof = False
    
    

End Sub

Private Sub ActiveReport_Initialize()



    
    
        
    
    DET_com = 1


End Sub


Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = ddOPortrait
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 25
    Me.PageTopMargin = 25
    Me.PageLeftMargin = 25
    Me.PageRightMargin = 25

    Me.documentName = "指図票実績確認"


    Me.Now_Date = Format(Now, "YYYY/MM/DD HH:MM:SS") & " 現在"

    DoEvents

End Sub

