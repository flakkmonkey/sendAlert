Sub SendAlert()
    ' 1. MEMORY & VARIABLE DECLARATIONS
    Dim OutApp As Object        ' Declares a variable to hold the Outlook application object
    Dim OutMail As Object       ' Declares a variable to hold the Outlook email item object
    Dim RngRow As Long          ' Stores the row number of the cell the user currently has selected
    Dim ProjectName As String   ' Stores the project name grabbed from Column A
    Dim PartNumber As String    ' Stores the part number grabbed from Column B
    Dim StatusType As String    ' Stores a text description of which status column was triggered
    Dim TriggerDate As String   ' Stores the date value entered in the active trigger cell
    Dim CheckCol As Long        ' Counting variable used to loop through columns C, D, and E
    Dim FoundAlert As Boolean   ' True/False switch to verify whether an active date trigger was found
    
    ' 2. ROW CONTEXT VALIDATION
    ' Looks at whatever row the user's active cell is currently resting on
    RngRow = ActiveCell.Row
    
    ' Safety check: prevents running the macro if the user is clicked on Row 1 (header row)
    If RngRow = 1 Then
        MsgBox "Please click on a specific part row before sending an alert.", vbExclamation
        Exit Sub
    End If
    
    ' 3. DATA EXTRACTION & SANITIZATION
    ProjectName = Cells(RngRow, 1).Value ' Grabs text from Column A (Project) on that row
    PartNumber = Cells(RngRow, 2).Value  ' Grabs text from Column B (Part Number) on that row
    
    ' Safety check: prevents firing an alert if a completely empty row is selected
    If ProjectName = "" And PartNumber = "" Then
        MsgBox "The selected row appears to be empty.", vbExclamation
        Exit Sub
    End If
    
    FoundAlert = False ' Initializes our switch to False before scanning
    
    ' 4. BOUNDED RANGE AUDIT (Columns C, D, and E)
    ' Limits the search strictly to Columns C (3), D (4), and E (5)
    For CheckCol = 3 To 5
        If Cells(RngRow, CheckCol).Value <> "" Then
            TriggerDate = Cells(RngRow, CheckCol).Value ' Grabs the date value from the cell
            
            ' Matches the column number to the appropriate status description
            Select Case CheckCol
                Case 3
                    StatusType = "Out of Stock in Warehouse"
                Case 4
                    StatusType = "ONE BIN LEFT in Tent"
                Case 5
                    StatusType = "OUT OF STOCK in Tent"
            End Select
            
            FoundAlert = True ' Marks that a valid trigger date was successfully found
            Exit For          ' Stops looping immediately once the trigger is located
        End If
    Next CheckCol
    
    ' Safety check: if no date was typed into Columns C, D, or E on that row
    If Not FoundAlert Then
        MsgBox "No date trigger found in Columns C, D, or E for this row.", vbExclamation
        Exit Sub
    End If
    
    ' 5. ERROR BOUNDING & OUTLOOK HANDOFF
    On Error Resume Next ' Temporarily intercepts runtime errors if Outlook fails to initialize
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    ' Dead-man switch: catches if Outlook is not installed or available on the local machine
    If OutApp Is Nothing Then
        MsgBox "Could not connect to Outlook. Ensure Microsoft Outlook is installed and open.", vbCritical
        Exit Sub
    End If
    
    ' 6. EMAIL CONSTRUCTION & DISPATCH
    With OutMail
        .To = "purchasing@yourdomain.com" ' Update this to your target purchasing email address
        .Subject = "URGENT [ " & ProjectName & " ] - Part " & PartNumber & " Alert"
        .Body = "Attention Purchasing," & vbCrLf & vbCrLf & _
                "An inventory trigger has been logged:" & vbCrLf & _
                "• Project: " & ProjectName & vbCrLf & _
                "• Part Number: " & PartNumber & vbCrLf & _
                "• Status: " & StatusType & vbCrLf & _
                "• Logged Date: " & TriggerDate & vbCrLf & vbCrLf & _
                "Please review the master sheet for immediate action."
        .Send ' Automatically fires the email in the background
    End With
    
    ' 7. MEMORY CLEANUP & RESET
    Set OutMail = Nothing       ' Clears the email object from RAM to prevent leaks
    Set OutApp = Nothing        ' Clears the Outlook application object from RAM
    On Error GoTo 0             ' Restores standard error handling behavior
    
    ' Confirmation prompt for the user
    MsgBox "Purchasing alert sent successfully for Part " & PartNumber & "!", vbInformation
End Sub
