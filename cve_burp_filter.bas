Sub HighlightDuplicates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cveDict As Object
    Dim burpDict As Object
    Dim currentCVE As String
    Dim currentComponent As String
    Dim currentTool As String
    Dim currentVulnTitle As String
    Dim currentDetails As String
    Dim ipPort As String
    Dim nessusR7Matches As Long
    Dim burpDuplicates As Long
    
    ' Set worksheet
    Set ws = ActiveSheet
    
    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Create dictionaries
    Set cveDict = CreateObject("Scripting.Dictionary")  ' For Nessus/R7
    Set burpDict = CreateObject("Scripting.Dictionary") ' For Burp duplicates
    ' Make dictionaries case-insensitive
    cveDict.CompareMode = 1 ' vbTextCompare
    burpDict.CompareMode = 1 ' vbTextCompare
    
    ' Initialize counters
    nessusR7Matches = 0
    burpDuplicates = 0
    
    ' First pass: Store R7 CVE-Component-Details combinations and process Burp entries
    For i = 2 To lastRow ' Assuming row 1 is header
        currentTool = Trim(ws.Cells(i, 1).Value) ' Column A - Tools
        currentComponent = Trim(ws.Cells(i, 2).Value) ' Column B - Component
        currentCVE = Trim(ws.Cells(i, 13).Value) ' Column M - CVE ID
        currentVulnTitle = Trim(ws.Cells(i, 3).Value) ' Column C - Vulnerability Title
        
        ' Check for error in Column O before processing
        If IsError(ws.Cells(i, 15).Value) Then
            currentDetails = "" ' Set to empty string if error found
            Debug.Print "Row " & i & " skipped: Error in Column O (All Details)"
        Else
            currentDetails = Trim(ws.Cells(i, 15).Value) ' Column O - All Details
        End If
        
        ' Handle R7 entries
        If UCase(currentTool) = "R7" And currentCVE <> "" Then
            ipPort = GetIpPort(currentComponent)
            If ipPort <> "" And currentDetails <> "" Then ' Only process if Details is not empty (i.e., no error)
                ' Store R7 entries with CVE, IP:Port, and Details
                Dim r7Key As String
                r7Key = currentCVE & "|" & ipPort & "|" & NormalizeString(currentDetails)
                If Not cveDict.Exists(r7Key) Then
                    cveDict.Add r7Key, i
                    Debug.Print "Row " & i & " (R7) added to CVE dictionary: " & r7Key
                End If
            End If
        End If
        
        ' Handle Burp entries
        If UCase(currentTool) = "BURP" Then
            If currentComponent <> "" And currentVulnTitle <> "" Then
                ' Normalize the strings
                Dim normalizedComponent As String
                Dim normalizedVulnTitle As String
                normalizedComponent = NormalizeString(currentComponent)
                normalizedVulnTitle = NormalizeString(currentVulnTitle)
                
                ' Extract Component key excluding Instance
                Dim componentKey As String
                componentKey = GetComponentKeyWithoutInstance(normalizedComponent)
                
                ' Use the Component key and full Vulnerability Title for matching
                Dim key As String
                key = componentKey & "|" & normalizedVulnTitle
                
                Debug.Print "Row " & i & " (Burp) Key: " & key
                Debug.Print "Row " & i & " Component: [" & normalizedComponent & "]"
                Debug.Print "Row " & i & " Component Key: [" & componentKey & "]"
                Debug.Print "Row " & i & " VulnTitle: [" & normalizedVulnTitle & "]"
                
                If Not burpDict.Exists(key) Then
                    ' Store first occurrence with row number
                    burpDict.Add key, i
                    Debug.Print "Row " & i & " (Burp) added to dictionary (first occurrence, not highlighted)"
                Else
                    ' Highlight only the current duplicate (not the first occurrence)
                    Dim firstRow As Long
                    firstRow = Abs(burpDict(key))
                    Debug.Print "Duplicate found at Row " & i & ", first occurrence at Row " & firstRow
                    ' Clear existing formatting for the current row
                    ws.Range("A" & i & ":G" & i).Interior.ColorIndex = xlNone
                    ' Highlight the current duplicate in orange
                    ws.Range("A" & i & ":G" & i).Interior.Color = RGB(255, 165, 0)
                    Debug.Print "Highlighted Row " & i & " in orange (Burp duplicate)"
                    
                    ' Update dictionary to ensure we track the first row
                    burpDict(key) = -firstRow ' Mark as processed
                    burpDuplicates = burpDuplicates + 1
                End If
            Else
                Debug.Print "Row " & i & " (Burp) skipped: Empty Component or Vulnerability Title"
            End If
        End If
    Next i
    
    ' Second pass: Check Nessus entries against R7 entries
    For i = 2 To lastRow
        currentTool = Trim(ws.Cells(i, 1).Value) ' Column A - Tools
        currentCVE = Trim(ws.Cells(i, 13).Value) ' Column M - CVE ID
        currentComponent = ws.Cells(i, 2).Value ' Column B - Component
        
        ' Check for error in Column O before processing
        If IsError(ws.Cells(i, 15).Value) Then
            currentDetails = "" ' Set to empty string if error found
            Debug.Print "Row " & i & " skipped: Error in Column O (All Details)"
        Else
            currentDetails = Trim(ws.Cells(i, 15).Value) ' Column O - All Details
        End If
        
        If UCase(currentTool) = "NESSUS" And currentCVE <> "" Then
            ipPort = GetIpPort(currentComponent)
            If ipPort <> "" And currentDetails <> "" Then ' Only process if Details is not empty (i.e., no error)
                ' Check if this CVE-IP:Port-Details combination exists in R7 entries
                Dim nessusKey As String
                nessusKey = currentCVE & "|" & ipPort & "|" & NormalizeString(currentDetails)
                If cveDict.Exists(nessusKey) Then
                    ' Clear existing formatting
                    ws.Range("A" & i & ":O" & i).Interior.ColorIndex = xlNone
                    ' Highlight entire Nessus row red
                    ws.Range("A" & i & ":O" & i).Interior.Color = vbRed
                    Debug.Print "Highlighted Row " & i & " in red (Nessus matches R7)"
                    nessusR7Matches = nessusR7Matches + 1
                End If
            End If
        End If
    Next i
    
    ' Clean up
    Set cveDict = Nothing
    Set burpDict = Nothing
    
    ' Display completion message with match counts
    MsgBox "Duplicate check complete!" & vbCrLf & vbCrLf & _
           "I compared CVE Number, IP Address, and Details." & vbCrLf & _
           "Found " & nessusR7Matches & " matching R7 and Nessus findings." & vbCrLf & _
           "I compared Vulnerability Name and IP Address." & vbCrLf & _
           "Found " & burpDuplicates & " multiple instances of Burp findings.", vbInformation
End Sub

' Function to normalize strings (remove extra spaces, standardize line endings)
Function NormalizeString(inputStr As String) As String
    Dim lines As Variant
    Dim i As Long
    Dim result As String
    
    ' Replace all line endings with vbLf
    inputStr = Replace(Replace(inputStr, vbCrLf, vbLf), vbCr, vbLf)
    
    ' Split into lines and trim each line
    lines = Split(inputStr, vbLf)
    For i = LBound(lines) To UBound(lines)
        lines(i) = Trim(lines(i))
    Next i
    
    ' Join lines back together
    result = Join(lines, vbLf)
    
    ' Remove any empty lines at the end
    While Right(result, Len(vbLf)) = vbLf
        result = Left(result, Len(result) - Len(vbLf))
    Wend
    
    NormalizeString = result
End Function

' Function to extract Component key excluding Instance (for Burp)
Function GetComponentKeyWithoutInstance(component As String) As String
    Dim lines As Variant
    Dim result As String
    
    lines = Split(component, vbLf)
    If UBound(lines) >= 2 Then ' Need at least 3 lines (URL, Protocol, Port)
        ' Combine URL, Protocol, and Port lines (exclude Instance)
        result = lines(0) & vbLf & lines(1) & vbLf & lines(2)
        GetComponentKeyWithoutInstance = result
    Else
        GetComponentKeyWithoutInstance = ""
    End If
End Function

' Function to extract IP and Port from Component (for Nessus/R7)
Function GetIpPort(component As String) As String
    Dim ip As String
    Dim port As String
    Dim lines As Variant
    
    lines = Split(component, vbLf)
    If UBound(lines) >= 1 Then
        ip = Trim(lines(0)) ' First line is IP
        If InStr(lines(1), "Port:") > 0 Then
            port = Trim(Replace(Replace(lines(1), "Port:", ""), "/udp", ""))
            GetIpPort = ip & ":" & port
        Else
            GetIpPort = ""
        End If
    Else
        GetIpPort = ""
    End If
End Function

