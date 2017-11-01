Attribute VB_Name = "Module1"
Sub CreateImportFileforNXT()
'-------------------------------------------------------------
' Description:  Extracts XML attahcments from emails which the user highlights.  Then creates temporary
' xml files, one for each attachment.  Then loops through these and extracts the contact preferences
' information into arrays.  Opens new CSV file and writes out the data in the arrays.  Leaves the file open for the
' user to look at and save where he wants.
' Inputs:  Emails highlighted in Outlook.  Each email should have an xml attachment but is ignored if not.
' Output:  .csv file with headers that should be meaningful to RE/NXT so that it can be imported.
'          a log file that can be viewed using notepad is also created.
' Notes
' a.  It is assumed that the user will tidy up the emails after they have been extracted
'
' Changes
' 09/08/2017 James Rowley - Initial version pending clarification of RE/NXT required format.
' 24/08/2017 James Rowley - To add case reference to be able to handle Services data entry.
' 20/09/2017 James Rowley - to create strings to assist with manual input into RE/NXT, namely:
'                           strings to hold the channels opted in or out
' 21/09/2017 James Rowley - to set the strings included yesterday to special values if the individual
'                           claims to have provided contact preferences already
'
'------------------------------------------------------------------------------

' These declarations are base on the DRM spreadsheetes Neal has provided.
 Dim appOL As Outlook.Application
 Dim appExcel As Excel.Application
 Dim myOlExp As Outlook.Explorer
 Dim myOlSel As Outlook.Selection
 Dim mySender As Outlook.AddressEntry
 Dim oMail As Outlook.MailItem
 Dim oAppt As Outlook.AppointmentItem
 Dim oPA As Outlook.PropertyAccessor
 Dim objAttachment As Attachment
 Dim objItem As Object
 Dim strSenderID As String
 Dim dtReceived As Date
 Dim MsgTxt As String
 Dim x As Long
 Dim iFileCount As Integer
 Dim lACount As Long
 Dim sBody As String
 Dim sFile As String
 Dim sFullPathName(1000) As String
 Dim sExt As String
 Dim gsActiveWorkbookPath As String
 Dim strTargetFile As String
 Dim sTemp As String
 
' New declarations
 Dim NumFileCount As Integer
 Dim iLastColumn As Long
 Dim iNumCPXMLFile As Long
 Dim iCPXMLFile As Long
 Dim iCol As Long
 Dim sAddressLine1Preferences1(1000) As String
 Dim sAddressLine2Preferences1(1000) As String
 Dim sAddressLine3Preferences1(1000) As String
 Dim sAffectedByMissing(1000) As String
 Dim sByEmail1(1000) As String
 Dim sByPhone1(1000) As String
 Dim sByPost1(1000) As String
 Dim sByText1(1000) As String
 Dim sEmailAddressPreference1(1000) As String
 Dim sFirstNamePreference1(1000) As String
 Dim sLastNamePreference1(1000) As String
 Dim sOptInMessage1(1000) As String
 Dim sMobilePreference1(1000) As String
 Dim sPhonePreference1(1000) As String
 Dim sPostcodePreference1(1000) As String
 Dim sTitlePreference1(1000) As String
 Dim sTownPreference1(1000) As String
 Dim sCPDoneAlready1(1000) As String
 Dim sField As String
 Dim sFieldParts() As String
 Dim iNumFieldParts As Long
 Dim dtEmailReceived(1000) As Date
 Dim dtEmailXMLReceived(1000) As Date
 Dim iFormID(1000) As Long
 Dim sCaseReference(1000) As String
 
 Dim sOutputWorkbook As String
 Dim gsLogFileName As String
 Dim sOutMessage As String
 Dim sOptInComment1, sOptOutComment1 As String
 Dim sTemp1 As String
 Dim sEmailYes, sEmailNo As String
 Dim sPostYes, sPostNo As String
 Dim sTextYes, sTextNo As String
 Dim sPhoneYes, sPhoneNo As String
 
 On Error GoTo Errorhandler
 
 gsActiveWorkbookPath = Cells(1, 2).Value
 
 'gsActiveWorkbookPath = "H:\Webform Reader\MASHImports\"
 gsLogFileName = gsActiveWorkbookPath & Format$(Now, "yyyy-mm-dd hh-mm-ss ") & "LOG"
 
 sOutMessage = "Log File"
 Call LogInformation(gsLogFileName, sOutMessage)
 
 Set appOL = Outlook.Application
 Set myOlExp = appOL.ActiveExplorer
 Set myOlSel = myOlExp.Selection
  
 iFileCount = 0
 ' Loop over the emails and create temporary xml files.
 For x = 1 To myOlSel.Count
 
    Set objItem = appOL.ActiveExplorer.Selection.Item(x)
 
    With objItem
    'An email can have a number of attachments.  There should only be one but even so we have
    'a loop to check each attachment for an email.
        lACount = .Attachments.Count
        dtReceived = .ReceivedTime
        sBody = .Body
              
        For Each objAttachment In .Attachments
            'check the file is an xml file
            sExt = LCase$(Right$(objAttachment.Filename, 4))
            If sExt = ".xml" Or sExt = "xml" Then
'               gsImportFile = gsActiveWorkbookPath & "\MASHImports\" & sFile
'               objAttachment.SaveAsFile sFullPathName
'               SetAttr sFullPathName, vbNormal
                'MsgBox "File is an XML file"
                
                iFileCount = iFileCount + 1
                'Now save the file as a temporary spreadsheet
                sFile = Format$(Now, "yyyy-mm-dd hh-mm-ss ") & iFileCount & " " & objAttachment.Filename
                sFullPathName(iFileCount) = gsActiveWorkbookPath & sFile
                'gsImportFile = gsActiveWorkbookPath & "\MASHImports\" & sFile
                objAttachment.SaveAsFile sFullPathName(iFileCount)
                SetAttr sFullPathName(iFileCount), vbNormal
                'set variable to hold the email received time as the create time is not included in the
                'XML file itself - an issue with RSForms
                dtEmailReceived(iFileCount) = dtReceived
                
            Else
                MsgBox "The file attached to the email is not an XML file. No action taken." & objAttachment.Filename
                sOutMessage = "The file attached to the email is not an XML file. No action taken." & objAttachment.Filename
                Call LogInformation(gsLogFileName, sOutMessage)
            End If
        
        Next objAttachment
    
    End With
  
 Next x
 Debug.Print MsgTxt
 
 'We have finished with Outlook now.  We want to use Excel.
 Set appExcel = Excel.Application
 
 'Loop over the saved temporary XML files

 NumFileCount = iFileCount
 iNumCPXMLFile = 0
 iCPXMLFile = 0
 For iFileCount = 1 To NumFileCount
  
     'Application.DisplayAlerts = False

     ' This opens a new speadsheet named according to strTargetFile and provides some field names above the
     ' data.  These can be parsed by code.
     Workbooks.OpenXML Filename:=sFullPathName(iFileCount), LoadOption:=xlXmlLoadOpenXml
     'Application.DisplayAlerts = True
     
     'MsgBox ActiveSheet.Name & " " & ActiveSheet.Cells(1, 1).Value
     sOutMessage = ActiveSheet.Name & " " & ActiveSheet.Cells(1, 1).Value
     Call LogInformation(gsLogFileName, sOutMessage)
     'Note that cell A1 holds the XML tag1 as set in RSforms XML spec for the particular webform.  This is used to
     'make sure only XML files with contact preference data is considered.
     'Row 2 holds the field names prefixed by the XML tag2.  This is tag is also not relevant because the field
     'names themselves refer to contact preferences.  All forms need to refer to contact preferences in their address,
     'email, phone fields etc.  In other words all forms need to be based on the standard basic contact preference form.
     
     sField = ActiveSheet.Cells(1, 1).Value
     If sField = "/contactpreferencesXFer" Then
     'This is the correct XML type of file
        iCPXMLFile = iCPXMLFile + 1
        'Now parse for the contact preference fields and write to the relevant arrays.
        iLastColumn = ActiveSheet.Cells(2, Columns.Count).End(xlToLeft).Column
        'MsgBox iLastColumn
        dtEmailXMLReceived(iCPXMLFile) = dtEmailReceived(iFileCount)
              
        For iCol = 1 To iLastColumn
            sField = ActiveSheet.Cells(2, iCol).Value
            sFieldParts() = Split(sField, "/")
            iNumFieldParts = UBound(sFieldParts)
            'MsgBox iNumFieldParts
        
            'Warning:  Note that the naming convention for the fields is somewhat variable with "preference" and "preferences" being used.
            Select Case sFieldParts(iNumFieldParts)
        
            Case "AddressLine1Preferences1"
                sAddressLine1Preferences1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "AddressLine2Preferences1"
                sAddressLine2Preferences1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "AddressLine3Preferences1"
                sAddressLine3Preferences1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "AffectedByMissing"
                sAffectedByMissing(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "ByEmail1"
                sByEmail1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "ByPhone1"
                sByPhone1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "ByPost1"
                sByPost1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "ByText1"
                sByText1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "EmailAddressPreference1"
                sEmailAddressPreference1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "FirstNamePreferences1"
                sFirstNamePreference1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "LastNamePreferences1"
                sLastNamePreference1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "OptInMessage1"
                sOptInMessage1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "MobilePreference1"
                sMobilePreference1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "PhonePreference1"
                sPhonePreference1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "PostcodePreferences1"
                sPostcodePreference1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "TitlePreferences1"
                sTitlePreference1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "TownPreferences1"
                sTownPreference1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "CPDoneAlready"
                sCPDoneAlready1(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "formId"
                iFormID(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "CaseReference"
                sCaseReference(iCPXMLFile) = ActiveSheet.Cells(3, iCol).Value
            Case "Capcha"
                'parse standard field that is of no interest to avoid it being highlighted on screen or in a log
            Case Else
                'MsgBox "Field not parsed " & sFieldParts(iNumFieldParts)
                sOutMessage = "Field not parsed " & sFieldParts(iNumFieldParts)
                Call LogInformation(gsLogFileName, sOutMessage)
                
            End Select
                
        Next
        
    Else
        MsgBox "An incorrect XML file has been found and ignored."
        sOutMessage = "An incorrect XML file has been found and ignored."
        Call LogInformation(gsLogFileName, sOutMessage)
        
    End If
    
'    sTemp = sFullPathName(iFileCount)
'    Workbooks(sTemp).Close SaveChanges:=False
    ActiveWorkbook.Close (False)
  
 Next
 
 iNumCPXMLFile = iCPXMLFile
 
'Now write the arrays to the output csv file.  NB 3/8/17:  the order and titles of the columns need to be
'agreed with the RE/NXT team.

 'sOutputWorkbook = gsPathName & gsFileName
 sOutputWorkbook = gsActiveWorkbookPath & Format$(Now, "yyyy-mm-dd hh-mm-ss ")

 sOutMessage = "Output work book name: " & sOutputWorkbook
 MsgBox sOutMessage
 Call LogInformation(gsLogFileName, sOutMessage)
    
 Workbooks.Add
 
 'MsgBox ActiveWorkbook.Name & " " & ActiveSheet.Name
 
 ActiveWorkbook.SaveAs Filename:=sOutputWorkbook, _
 FileFormat:=xlCSV, CreateBackup:=False
 
 'Write the column headers
 ActiveSheet.Cells(1, 1).Value = "DateSubmitted"
 ActiveSheet.Cells(1, 5).Value = "AddressLine1Preferences1"
 ActiveSheet.Cells(1, 6).Value = "AddressLine2Preferences1"
 ActiveSheet.Cells(1, 7).Value = "AddressLine3Preferences1"
 ActiveSheet.Cells(1, 17).Value = "AffectedByMissing"
 ActiveSheet.Cells(1, 10).Value = "ByEmail1"
 ActiveSheet.Cells(1, 12).Value = "ByPhone1"
 ActiveSheet.Cells(1, 16).Value = "ByPost1"
 ActiveSheet.Cells(1, 14).Value = "ByText1"
 ActiveSheet.Cells(1, 11).Value = "EmailAddressPreference1"
 ActiveSheet.Cells(1, 3).Value = "FirstNamePreferences1"
 ActiveSheet.Cells(1, 4).Value = "LastNamePreferences1"
 ActiveSheet.Cells(1, 18).Value = "OptInMessage1"
 ActiveSheet.Cells(1, 15).Value = "MobilePreference1"
 ActiveSheet.Cells(1, 13).Value = "PhonePreference1"
 ActiveSheet.Cells(1, 9).Value = "PostcodePreferences1"
 ActiveSheet.Cells(1, 2).Value = "TitlePreferences1"
 ActiveSheet.Cells(1, 8).Value = "TownPreferences1"
 ActiveSheet.Cells(1, 19).Value = "CPDoneAlready"
 ActiveSheet.Cells(1, 20).Value = "FormIDOnRSForms"
 ActiveSheet.Cells(1, 21).Value = "CaseReference"
 ActiveSheet.Cells(1, 22).Value = "OptInComment1"
 ActiveSheet.Cells(1, 23).Value = "OptOutComment1"

 For iCPXMLFile = 1 To iNumCPXMLFile
    ActiveSheet.Cells(iCPXMLFile + 1, 1).Value = dtEmailXMLReceived(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 5).Value = sAddressLine1Preferences1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 6).Value = sAddressLine2Preferences1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 7).Value = sAddressLine3Preferences1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 17).Value = sAffectedByMissing(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 10).Value = sByEmail1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 12).Value = sByPhone1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 16).Value = sByPost1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 14).Value = sByText1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 11).Value = sEmailAddressPreference1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 3).Value = sFirstNamePreference1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 4).Value = sLastNamePreference1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 18).Value = sOptInMessage1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 15).Value = sMobilePreference1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 13).Value = sPhonePreference1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 9).Value = sPostcodePreference1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 2).Value = sTitlePreference1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 8).Value = sTownPreference1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 19).Value = sCPDoneAlready1(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 20).Value = iFormID(iCPXMLFile)
    ActiveSheet.Cells(iCPXMLFile + 1, 21).Value = sCaseReference(iCPXMLFile)
    
 'JRR 20/9/17:  Now ceate the strings to be copied into the Comments fields of the Opt-in and opt-out attributes
    If sCPDoneAlready1(iCPXMLFile) <> "" Then
        sOptInComment1 = "Claims to have given CP before" & "/" & iFormID(iCPXMLFile)
        sOptOutComment1 = "Claims to have given CP before" & "/" & iFormID(iCPXMLFile)
    Else
        If sByEmail1(iCPXMLFile) = "Yes" Then
            SEmai1Yes = "Email,"
            sEmai1No = ""
        Else
            sEmai1No = "Email,"
            SEmai1Yes = ""
        End If
        If sByPhone1(iCPXMLFile) = "Yes" Then
            sPhoneYes = "Phone,"
            sPhoneNo = ""
        Else
            sPhoneNo = "Phone,"
            sPhoneYes = ""
        End If
        If sByPost1(iCPXMLFile) = "Yes" Then
            sPostYes = "Post,"
            sPostNo = ""
        Else
            sPostNo = "Post,"
            sPostYes = ""
        End If
        If sByText1(iCPXMLFile) = "Yes" Then
            sTextYes = "SMS,"
            sTextNo = ""
        Else
            sTextNo = "SMS,"
            sTextYes = ""
        End If
        
        sTemp1 = Left(sOptInMessage1(iCPXMLFile), 5)
        sOptInComment1 = SEmai1Yes & sTextYes & sPostYes & sPhoneYes & "/" & sTemp1 & "/" & iFormID(iCPXMLFile)
        sOptOutComment1 = sEmai1No & sTextNo & sPostNo & sPhoneNo & "/" & sTemp1 & "/" & iFormID(iCPXMLFile)
    End If
    ActiveSheet.Cells(iCPXMLFile + 1, 22).Value = sOptInComment1
    ActiveSheet.Cells(iCPXMLFile + 1, 23).Value = sOptOutComment1
    
 Next

 sOutMessage = iNumCPXMLFile & " XML files have been written to the csv file"
 MsgBox sOutMessage
 Call LogInformation(gsLogFileName, sOutMessage)
 
' ActiveWorkbook.Save
' ActiveWorkbook.Close
 
 End
 
Errorhandler:
 sOutMessage = "Error Number = " & Err.Number & "  " & Err.Description
 MsgBox sOutMessage
 Call LogInformation(gsLogFileName, sOutMessage)

 
 End Sub

Sub LogInformation(sLogFileName As String, sLogMessage As String)
'-------------------------------------------------------------
' Description:  Creates, if it does not already exist, a log file in the location
' that is uincluded in the input file name and then appends a simple text message.
' Inputs:  Name of the log file.  Message to be added.
' Output:  Updated log file
' Notes
' a.  The log file is best opened with NotePad.
' b.  A useful standalone routine
'
' Changes
' 09/08/2017 James Rowley - Initial version
'
'------------------------------------------------------------------------------
Dim iFileNum As Integer
iFileNum = FreeFile ' next file number
Open sLogFileName For Append As #iFileNum ' creates the file if it doesn't exist
Print #iFileNum, sLogMessage ' write information at the end of the text file
Close #iFileNum ' close the file
End Sub





























