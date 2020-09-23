Attribute VB_Name = "DictMode"
Public Sub dictmodule()
Open frmMain.txtDictFile.Text For Input As #2
While Not EOF(2)
    Input #2, comb
                Select Case OChoice
                Case 1:
                    result = result & vbCrLf & comb
                Case 2:
                    Print #1, comb
                Case 3:
                    SendKeys comb: BFInitialise
                Case 4:
                    processcode (comb)
                
                End Select
Wend
Close #2
frmMain.txtStatus.Text = result

End Sub

