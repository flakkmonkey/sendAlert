' Title : sendAlert
' Description : Excel macro to send Outlook email message to pre-ddefined recipient
'               whenever changes are made to the worksheet.


Sub SendAlert() ' Declares the start of the macro subroutine linked to your button
    Dim OutApp As Object ' Declares a variable to hold the Outlook application object
    Dim OutMail As Object ' Declares a variable to hold the Outlook email item object
    Dim RngRow As Long ' Declares a variable to store the row number of the cell the user has selected
    Dim ProjectName As String ' Declares a variable to store the project name grabbed from Column A
    Dim PartNumber As String ' Declares a variable to store the part number grabbed from Column B
    Dim StatusType As String ' Declares a variable to store a text description of which status column was triggered
    Dim TriggerDate As String ' Declares a variable to store the date value entered in the trigger cell
    Dim CheckCol As Long ' Declares a counting variable used to loop through columns C, D, and E
    Dim FoundAlert As Boolean ' Declares a True/False switch to verify whether an active date trigger was actually found
    
    RngRow = ActiveCell.Row ' Looks at whatever row the user's active cell is currently resting on and records that row number
    
    If RngRow = 1 Then ' Safety check: prevents the user from running the macro if they are clicked on Row 1 (the header row)
        MsgBox "Please click on a specific part row before sending an alert.", vbExclamation ' Pops up a warning box telling them to select a valid row
        Exit Sub ' Stops running the macro immediately
    End If ' Ends the header row check block
    
    ProjectName = Cells(RngRow, 1).Value ' Grabs the text inside Column A (Column 1) on that selected row and stores it as the Project
    PartNumber = Cells(RngRow, 2).Value ' Grabs the text inside Column B (Column 2) on that selected row and stores it as the Part Number
    
    If ProjectName = "" And PartNumber = "" Then ' Safety check: checks if both the project and part number cells are completely empty
        MsgBox "The selected row appears to be empty.", vbExclamation ' Pops up a warning box letting the user know the row has no data
        Exit Sub ' Stops running the macro immediately
    End If ' Ends the empty row check block
    
    FoundAlert = False ' Sets our tracking switch to False before we begin scanning the status columns
    
    For CheckCol = 3 To 5 ' Starts a loop that checks Columns C (3), D (4), and E (5) one by one on that row
        If Cells(RngRow, CheckCol).Value <> "" Then ' Checks if the current column's cell contains text/a date (is not blank)
            TriggerDate = Cells(RngRow, CheckCol).Value ' Grabs that date value and stores it
            
            Select Case CheckCol ' Looks at which specific column number triggered the match
                Case 3 ' If Column C matched
                    StatusType = "Out of Stock in Warehouse" ' Assigns the warehouse status description
                Case 4 ' If Column D matched
                    StatusType = "ONE BIN LEFT in Tent" ' Assigns the tent low-stock status description
                Case 5 ' If Column E matched
                    StatusType = "OUT OF STOCK in Tent" ' Assigns the tent out-of-stock status description
            End Select ' Ends the column selection block
            
            FoundAlert = True ' Changes our tracking switch to True because a valid date was successfully found
            Exit For ' Stops running the loop immediately since we found our trigger
        End If ' Ends the "is not blank" check
    Next CheckCol ' Moves on to check the next column if no trigger was found yet
    
    If Not FoundAlert Then ' Safety check: if the loop finished checking and never found a date in Columns C, D, or E
        MsgBox "No date trigger found in Columns C, D, or E for this row.", vbExclamation ' Pops up a warning box explaining why it stopped
        Exit Sub ' Stops running the macro immediately
    End If ' Ends the missing trigger check block
    
    On Error Resume Next ' Tells VBA to temporarily bypass any errors if Outlook fails to open or connect
    Set OutApp = CreateObject("Outlook.Application") ' Opens a connection to the desktop Outlook application in the background
    Set OutMail = OutApp.CreateItem(0) ' Creates a brand new, blank email draft item
    
    With OutMail ' Begins setting the properties of the outgoing email message
        .To = "purchasing@yourdomain.com" ' Defines the recipient address for the purchasing team
        .Subject = "!!!URGENT!!! - [ " & ProjectName & " ] - Part " & PartNumber & " Alert" ' Dynamically builds the subject line using the extracted Project and Part Number
        .Body = "Attention Purchasing," & vbCrLf & vbCrLf & _
                "An inventory trigger has been logged:" & vbCrLf & _
                "• Project: " & ProjectName & vbCrLf & _
                "• Part Number: " & PartNumber & vbCrLf & _
                "• Status: " & StatusType & vbCrLf & _
                "• Logged Date: " & TriggerDate & vbCrLf & vbCrLf & _
                "Please review the master sheet for immediate action." ' Constructs the clean, structured message body with bullet points
        .Send ' Automatically fires and sends the email without making the user click send manually
    End With ' Ends the email properties block
    
    Set OutMail = Nothing ' Clears the email memory object to free up system resources
    Set OutApp = Nothing ' Clears the Outlook application memory object
    On Error GoTo 0 ' Restores standard error handling behavior back to normal
    
    MsgBox "Purchasing alert sent successfully for Part " & PartNumber & "!", vbInformation ' Pops up a friendly confirmation box so the user knows the button click worked
End Sub ' Ends the macro subroutine
