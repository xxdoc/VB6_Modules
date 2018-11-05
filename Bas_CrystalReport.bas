Attribute VB_Name = "Bas_CrystalReport"
Public Sub ExportReport(ByVal ReportPath As String, _
                                                    ByVal DestinationType As Integer, _
                                                    ByVal FormatType As Integer, _
                                                    ByVal DiskFileName As String)

    Dim CRXApplication As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report
    
    Set Report = CRXApplication.OpenReport(ReportPath)
    
    'for all property constant, please refer to the crystal report reference
        
    With Report.ExportOptions
        
        'Disk file
        .DestinationType = crEDTDiskFile
            
        'Excel 8.0
        .FormatType = crEFTExcel80

        'Export filename
        .DiskFileName = DiskFileName
        
    End With
    
    Report.DiscardSavedData
    
    Report.ReadRecords
    
    'Export without prompt user
    Report.Export False
        
    Set Report = Nothing
        
End Sub

