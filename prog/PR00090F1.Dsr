VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00090F1 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   18075
   StartUpPosition =   3  'Windows �̊���l
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

    Me.Fields.Add "HAKKO_DT"            '���s��
    Me.Fields.Add "SHIMUKE_CODE"        '�d������
    Me.Fields.Add "SHIJI_NO"            '�w�}�[��
    Me.Fields.Add "HIN_GAI"             '�i��
    Me.Fields.Add "SHIJI_QTY"           '����
    Me.Fields.Add "DOUKON"              '��������
    Me.Fields.Add "KAN_DT"              '������
    
    
    
    
    


End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)




    If DET_com > SSHIJI.UpperBound(1) Then
        Exit Sub
    End If
    

    '���s��
    Me.Fields("HAKKO_DT") = SSHIJI(DET_com, colHAKKO_DT)
    '�d����
    Me.Fields("SHIMUKE_CODE") = SSHIJI(DET_com, colSHIMUKE_CODE)
    '�w�}�[��
    Me.Fields("SHIJI_NO") = Format(CLng(SSHIJI(DET_com, colSHIJI_NO)), "00000000")
    '�i��
    Me.Fields("HIN_GAI") = SSHIJI(DET_com, colHIN_GAI)
    '�w����
    Me.Fields("SHIJI_QTY") = Format(CLng(SSHIJI(DET_com, colSHIJI_QTY)), "#0")
    '������
    Me.Fields("DOUKON") = Format(CLng(SSHIJI(DET_com, colDOUKON)), "#0")
    '������
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

    Me.documentName = "�w�}�[���ъm�F"


    Me.Now_Date = Format(Now, "YYYY/MM/DD HH:MM:SS") & " ����"

    DoEvents

End Sub

