Attribute VB_Name = "Module1"
Attribute VB_Name = "Module1"
' =============================================
' DOWNLOAD HELPER - Add this once
' =============================================
Private Sub DownloadLatestTemplate(ByVal LocalPath As String, ByVal URL As String)
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", URL, False
    http.send
    
    If http.Status = 200 Then
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1                  ' Binary
        stream.Write http.responseBody
        stream.SaveToFile LocalPath, 2   ' Overwrite
        stream.Close
        Debug.Print "? Latest template downloaded from GitHub"
    Else
        MsgBox "? Failed to download the latest template from GitHub." & vbCrLf & _
               "Status: " & http.Status & vbCrLf & _
               "Please check your internet connection and try again.", vbCritical
        Err.Raise vbObjectError + 1, , "Download failed"
    End If
End Sub

' =============================================
' YOUR MAIN MACRO - Updated version
' =============================================
Sub GeneratePowerPointFromTemplate_v1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Main Calculator")
 
    ' ==================== NEW PART: Download from GitHub ====================
    Dim templatePath As String
    templatePath = Environ("TEMP") & "\TP_Template_ATAP_Small_below_7.14.pptx"   ' Temporary local copy
    
    Dim templateURL As String
    templateURL = "https://raw.githubusercontent.com/alextan7219/ppt-template/main/TP%20Template_ATAP_Small_below_7.14.pptx"
    
    Call DownloadLatestTemplate(templatePath, templateURL)
    ' ======================================================================
 
    Dim saveFolder As String
    saveFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
 
    ' Generate file name
    Dim fileNameBase As String
    fileNameBase = "Solar_Proposal_" & ws.Range("C1").Text & "_" & ws.Range("C2").Text & "_" & ws.Range("C4").Text & "_" & ws.Range("C5").Text & "_" & ws.Range("C10").Text
    fileNameBase = Replace(fileNameBase, "/", "-")
 
    Dim pptPath As String
    pptPath = saveFolder & fileNameBase & ".pptx"
 
    Dim pptApp As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
 
    Dim pptPres As Object
    Set pptPres = pptApp.Presentations.Open(templatePath)
 
    ' Slide 1 replacements
    Dim slide1 As Object
    Set slide1 = pptPres.Slides(1)
  
    Dim shp As Object
    For Each shp In slide1.Shapes
        Select Case shp.Name
            Case "first_page_name"
                shp.TextFrame.TextRange.Text = ws.Range("C2").Text
            Case "first_page_address"
                shp.TextFrame.TextRange.Text = ws.Range("C3").Text
            Case "tnb_bill"
                shp.TextFrame.TextRange.Text = ws.Range("C4").Text
            Case "system_size"
                shp.TextFrame.TextRange.Text = ws.Range("C10").Text
            Case "panel_size"
                shp.TextFrame.TextRange.Text = ws.Range("C27").Text
            Case "solar_savings"
                shp.TextFrame.TextRange.Text = ws.Range("C23").Text
            Case "new_tnb_bill"
                shp.TextFrame.TextRange.Text = ws.Range("C22").Text
        End Select
    Next shp
  
    ' Slide 3 replacements (unchanged)
    Dim slide3 As Object
    Set slide3 = pptPres.Slides(3)
  
    For Each shp In slide3.Shapes
        Select Case shp.Name
            Case "system_size"
                shp.TextFrame.TextRange.Text = ws.Range("C10").Text
            Case "panel_size"
                shp.TextFrame.TextRange.Text = ws.Range("C27").Text
            Case "inverter_size"
                shp.TextFrame.TextRange.Text = ws.Range("C29").Text
            Case "solar_generation_1"
                shp.TextFrame.TextRange.Text = ws.Range("C19").Text
            Case "old_bill"
                shp.TextFrame.TextRange.Text = ws.Range("C4").Text
            Case "old_kwh"
                shp.TextFrame.TextRange.Text = ws.Range("C13").Text
            Case "new_bill"
                shp.TextFrame.TextRange.Text = ws.Range("C22").Text
            Case "new_kwh"
                shp.TextFrame.TextRange.Text = ws.Range("C30").Text
            Case "monthly_savings"
                shp.TextFrame.TextRange.Text = ws.Range("C23").Text
            Case "savings_percent"
                shp.TextFrame.TextRange.Text = ws.Range("C24").Text
            Case "10y_savings"
                shp.TextFrame.TextRange.Text = ws.Range("C25").Text
            Case "payback"
                shp.TextFrame.TextRange.Text = ws.Range("C26").Text
            Case "op_price_1"
                shp.TextFrame.TextRange.Text = ws.Range("C14").Text
            Case "op_price_2"
                shp.TextFrame.TextRange.Text = ws.Range("C14").Text
            Case "5y_price"
                shp.TextFrame.TextRange.Text = ws.Range("C17").Text
            Case "10y_full"
                shp.TextFrame.TextRange.Text = ws.Range("D18").Text
            Case "10y_price"
                shp.TextFrame.TextRange.Text = ws.Range("C18").Text
            Case "solar_generation_2"
                shp.TextFrame.TextRange.Text = ws.Range("C19").Text
            Case "daytime_usage"
                shp.TextFrame.TextRange.Text = ws.Range("C20").Text
        End Select
    Next shp

    ' Slide 4 replacements (unchanged - I kept everything exactly as you had)
    Dim slide4 As Object
    Set slide4 = pptPres.Slides(4)
  
    For Each shp In slide4.Shapes
        Select Case shp.Name
            Case "monthly_saving"
                shp.TextFrame.TextRange.Text = ws.Range("C23").Text
            Case "old_kwh_1"
                shp.TextFrame.TextRange.Text = ws.Range("B58").Text
            Case "solar_kwh"
                shp.TextFrame.TextRange.Text = ws.Range("D58").Text
            Case "daytime_kwh_1"
                shp.TextFrame.TextRange.Text = ws.Range("F58").Text
            Case "savings_percent"
                shp.TextFrame.TextRange.Text = ws.Range("G58").Text
            Case "export_kwh_1"
                shp.TextFrame.TextRange.Text = ws.Range("I58").Text
            Case "old_kwh_2"
                shp.TextFrame.TextRange.Text = ws.Range("D63").Text
            Case "new_kwh"
                shp.TextFrame.TextRange.Text = ws.Range("F63").Text
            Case "daytime_kwh_2"
                shp.TextFrame.TextRange.Text = ws.Range("F65").Text
            Case "before_total_kwh"
                shp.TextFrame.TextRange.Text = ws.Range("D67").Text
            Case "after_total_kwh"
                shp.TextFrame.TextRange.Text = ws.Range("F67").Text
            Case "old_energy"
                shp.TextFrame.TextRange.Text = ws.Range("D80").Text
            Case "old_capacity"
                shp.TextFrame.TextRange.Text = ws.Range("D82").Text
            Case "old_network"
                shp.TextFrame.TextRange.Text = ws.Range("D84").Text
            Case "old_afa"
                shp.TextFrame.TextRange.Text = ws.Range("D86").Text
            Case "old_eei"
                shp.TextFrame.TextRange.Text = ws.Range("D88").Text
            Case "old_retail"
                shp.TextFrame.TextRange.Text = ws.Range("D90").Text
            Case "old_kwtbb"
                shp.TextFrame.TextRange.Text = ws.Range("D92").Text
            Case "old_sst"
                shp.TextFrame.TextRange.Text = ws.Range("D94").Text
            Case "new_energy"
                shp.TextFrame.TextRange.Text = ws.Range("F80").Text
            Case "new_capacity"
                shp.TextFrame.TextRange.Text = ws.Range("F82").Text
            Case "new_network"
                shp.TextFrame.TextRange.Text = ws.Range("F84").Text
            Case "new_afa"
                shp.TextFrame.TextRange.Text = ws.Range("F86").Text
            Case "new_eei"
                shp.TextFrame.TextRange.Text = ws.Range("F88").Text
            Case "new_retail"
                shp.TextFrame.TextRange.Text = ws.Range("F90").Text
            Case "new_kwtbb"
                shp.TextFrame.TextRange.Text = ws.Range("F92").Text
            Case "new_sst"
                shp.TextFrame.TextRange.Text = ws.Range("F94").Text
            Case "old_total_charges"
                shp.TextFrame.TextRange.Text = ws.Range("D96").Text
            Case "new_total_charges"
                shp.TextFrame.TextRange.Text = ws.Range("F96").Text
            Case "solar_export_credit"
                shp.TextFrame.TextRange.Text = ws.Range("F98").Text
            Case "old_total_bill"
                shp.TextFrame.TextRange.Text = ws.Range("D102").Text
            Case "new_total_bill"
                shp.TextFrame.TextRange.Text = ws.Range("F102").Text
            Case "daytime_savings"
                shp.TextFrame.TextRange.Text = ws.Range("F106").Text
            Case "export_savings"
                shp.TextFrame.TextRange.Text = ws.Range("F108").Text
            Case "total_savings"
                shp.TextFrame.TextRange.Text = ws.Range("F110").Text
            Case "eei_adjust"
                shp.TextFrame.TextRange.Text = ws.Range("F100").Text
            Case "export_kwh_2"
                shp.TextFrame.TextRange.Text = ws.Range("I58").Text
        End Select
    Next shp
  
    ' Slide 9 replacements
    Dim slide9 As Object
    Set slide9 = pptPres.Slides(9)
   
    For Each shp In slide9.Shapes
        Select Case shp.Name
            Case "size_1"
                shp.TextFrame.TextRange.Text = ws.Range("B50").Text
            Case "solar_kwh_1"
                shp.TextFrame.TextRange.Text = ws.Range("C50").Text
            Case "old_tnb_1"
                shp.TextFrame.TextRange.Text = ws.Range("D50").Text
            Case "old_kwh_1"
                shp.TextFrame.TextRange.Text = ws.Range("E50").Text
            Case "daytime_percent_1"
                shp.TextFrame.TextRange.Text = ws.Range("F50").Text
            Case "daytime_kwh_1"
                shp.TextFrame.TextRange.Text = ws.Range("G50").Text
            Case "new_tnb_1"
                shp.TextFrame.TextRange.Text = ws.Range("H50").Text
            Case "solar_savings_1"
                shp.TextFrame.TextRange.Text = ws.Range("I50").Text
            Case "solar_savings_percent_1"
                shp.TextFrame.TextRange.Text = ws.Range("J50").Text
            Case "5y_1"
                shp.TextFrame.TextRange.Text = ws.Range("K50").Text
            Case "op_1"
                shp.TextFrame.TextRange.Text = ws.Range("L50").Text
            Case "payback_1"
                shp.TextFrame.TextRange.Text = ws.Range("M50").Text
            Case "size_2"
                shp.TextFrame.TextRange.Text = ws.Range("B51").Text
            Case "solar_kwh_2"
                shp.TextFrame.TextRange.Text = ws.Range("C51").Text
            Case "old_tnb_2"
                shp.TextFrame.TextRange.Text = ws.Range("D51").Text
            Case "old_kwh_2"
                shp.TextFrame.TextRange.Text = ws.Range("E51").Text
            Case "daytime_percent_2"
                shp.TextFrame.TextRange.Text = ws.Range("F51").Text
            Case "daytime_kwh_2"
                shp.TextFrame.TextRange.Text = ws.Range("G51").Text
            Case "new_tnb_2"
                shp.TextFrame.TextRange.Text = ws.Range("H51").Text
            Case "solar_savings_2"
                shp.TextFrame.TextRange.Text = ws.Range("I51").Text
            Case "solar_savings_percent_2"
                shp.TextFrame.TextRange.Text = ws.Range("J51").Text
            Case "5y_2"
                shp.TextFrame.TextRange.Text = ws.Range("K51").Text
            Case "op_2"
                shp.TextFrame.TextRange.Text = ws.Range("L51").Text
            Case "payback_2"
                shp.TextFrame.TextRange.Text = ws.Range("M51").Text
            Case "size_3"
                shp.TextFrame.TextRange.Text = ws.Range("B52").Text
            Case "solar_kwh_3"
                shp.TextFrame.TextRange.Text = ws.Range("C52").Text
            Case "old_tnb_3"
                shp.TextFrame.TextRange.Text = ws.Range("D52").Text
            Case "old_kwh_3"
                shp.TextFrame.TextRange.Text = ws.Range("E52").Text
            Case "daytime_percent_3"
                shp.TextFrame.TextRange.Text = ws.Range("F52").Text
            Case "daytime_kwh_3"
                shp.TextFrame.TextRange.Text = ws.Range("G52").Text
            Case "new_tnb_3"
                shp.TextFrame.TextRange.Text = ws.Range("H52").Text
            Case "solar_savings_3"
                shp.TextFrame.TextRange.Text = ws.Range("I52").Text
            Case "solar_savings_percent_3"
                shp.TextFrame.TextRange.Text = ws.Range("J52").Text
            Case "5y_3"
                shp.TextFrame.TextRange.Text = ws.Range("K52").Text
            Case "op_3"
                shp.TextFrame.TextRange.Text = ws.Range("L52").Text
            Case "payback_3"
                shp.TextFrame.TextRange.Text = ws.Range("M52").Text
            Case "size_4"
                shp.TextFrame.TextRange.Text = ws.Range("B53").Text
            Case "solar_kwh_4"
                shp.TextFrame.TextRange.Text = ws.Range("C53").Text
            Case "old_tnb_4"
                shp.TextFrame.TextRange.Text = ws.Range("D53").Text
            Case "old_kwh_4"
                shp.TextFrame.TextRange.Text = ws.Range("E53").Text
            Case "daytime_percent_4"
                shp.TextFrame.TextRange.Text = ws.Range("F53").Text
            Case "daytime_kwh_4"
                shp.TextFrame.TextRange.Text = ws.Range("G53").Text
            Case "new_tnb_4"
                shp.TextFrame.TextRange.Text = ws.Range("H53").Text
            Case "solar_savings_4"
                shp.TextFrame.TextRange.Text = ws.Range("I53").Text
            Case "solar_savings_percent_4"
                shp.TextFrame.TextRange.Text = ws.Range("J53").Text
            Case "5y_4"
                shp.TextFrame.TextRange.Text = ws.Range("K53").Text
            Case "op_4"
                shp.TextFrame.TextRange.Text = ws.Range("L53").Text
            Case "payback_4"
                shp.TextFrame.TextRange.Text = ws.Range("M53").Text
            Case "size_5"
                shp.TextFrame.TextRange.Text = ws.Range("B54").Text
            Case "solar_kwh_5"
                shp.TextFrame.TextRange.Text = ws.Range("C54").Text
            Case "old_tnb_5"
                shp.TextFrame.TextRange.Text = ws.Range("D54").Text
            Case "old_kwh_5"
                shp.TextFrame.TextRange.Text = ws.Range("E54").Text
            Case "daytime_percent_5"
                shp.TextFrame.TextRange.Text = ws.Range("F54").Text
            Case "daytime_kwh_5"
                shp.TextFrame.TextRange.Text = ws.Range("G54").Text
            Case "new_tnb_5"
                shp.TextFrame.TextRange.Text = ws.Range("H54").Text
            Case "solar_savings_5"
                shp.TextFrame.TextRange.Text = ws.Range("I54").Text
            Case "solar_savings_percent_5"
                shp.TextFrame.TextRange.Text = ws.Range("J54").Text
            Case "5y_5"
                shp.TextFrame.TextRange.Text = ws.Range("K54").Text
            Case "op_5"
                shp.TextFrame.TextRange.Text = ws.Range("L54").Text
            Case "payback_5"
                shp.TextFrame.TextRange.Text = ws.Range("M54").Text
        End Select
    Next shp
  
    ' Slide 18 replacements
    Dim slide18 As Object
    Set slide18 = pptPres.Slides(18)
   
    For Each shp In slide18.Shapes
        Select Case shp.Name
            Case "panel_size"
                shp.TextFrame.TextRange.Text = ws.Range("C27").Text
            Case "inverter"
                shp.TextFrame.TextRange.Text = ws.Range("C28").Text
        End Select
    Next shp
 
    ' Save as new PPTX
    pptPres.SaveAs pptPath
 
    ' Close the presentation
    pptPres.Close
 
    ' Quit PowerPoint app
    pptApp.Quit
 
    Set pptPres = Nothing
    Set pptApp = Nothing
    Set ws = Nothing
 
    MsgBox "PowerPoint generated successfully!" & vbCrLf & _
           "Saved as: " & pptPath & vbCrLf & vbCrLf & _
           "Template was downloaded fresh from GitHub."
End Sub

