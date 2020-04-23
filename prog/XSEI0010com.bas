Attribute VB_Name = "XSEI0010com"
Option Explicit

Private Declare Function ExtFloodFill Lib "gdi32" _
    (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
     ByVal crColor As Long, ByVal wFillType As Long) As Long


