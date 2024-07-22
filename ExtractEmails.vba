Sub ExtractEmails()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.Namespace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim olMail As Outlook.MailItem
    Dim subjectPhrase As String
    Dim fromAddress As String
    Dim folderPath As String
    Dim fileName As String
    Dim filePath As String
    Dim emailBody As String
    Dim textStream As Object
    
    ' Initialize Outlook application and namespace
    Set olApp = Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Define the specific phrase and email address
    subjectPhrase = "SPECIFIC PHRASE"
    fromAddress = "address@example.com"
    
    ' Loop through each item in the folder
    For Each olItem In olFolder.Items
        If TypeOf olItem Is Outlook.MailItem Then
            Set olMail = olItem
            
            ' Check if the email subject contains the phrase and is from the specific email address
            If InStr(olMail.Subject, subjectPhrase) > 0 And olMail.SenderEmailAddress = fromAddress Then
                ' Create file name based on the received date
                fileName = Format(olMail.ReceivedTime, "yyyy-mm-dd") & ".txt"
                
                ' Define the folder path where the file will be saved
                folderPath = "C:\path\to\your\folder\" ' Change this to your desired folder path
                
                ' Create the file path
                filePath = folderPath & fileName
                
                ' Extract the email body text
                emailBody = olMail.Body ' Use .Body to get plain text
                
                ' Write the email subject and body to the text file
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set textStream = fso.CreateTextFile(filePath, True)
                
                textStream.WriteLine olMail.Subject
                textStream.WriteLine emailBody
                
                textStream.Close
            End If
        End If
    Next olItem
    
    ' Clean up
    Set olMail = Nothing
    Set olItem = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub
