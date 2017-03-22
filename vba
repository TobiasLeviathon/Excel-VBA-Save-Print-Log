nSub SaveandPrint()

Dim FileName As String
Dim Path As String

ActiveSheet.PrintOut
Application.DisplayAlerts = False

Path = "C:\destination\" 'Change the directory path here where you want to save the file
FileName = Range("C18").Value & " " & Range("A7").Value & ".xlsm" 'Change extension here

ActiveWorkbook.SaveAs Path & FileName, xlOpenXMLWorkbookMacroEnabled 'Change the format here which matches with the extention above. 
                                            'Choose from the following link http://msdn.microsoft.com/en-us/library/office/ff198017.aspx

Application.DisplayAlerts = True

Dim fName As String
    fName = Range("C18").Value & " " & Range("A7").Value

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
            "C:\destination\" & fName, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

ActiveWorkbook.Close

End Sub
