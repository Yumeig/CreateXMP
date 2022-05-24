Attribute VB_Name = "CreateVideo"
Sub CreateVideo()

    Dim file As String
    
    file = VBA.Split(ActivePresentation.FullName, ".")(0) & ".mp4"
    
    ActivePresentation.CreateVideo file, False, 5, 1080, 60, 100
    
End Sub
