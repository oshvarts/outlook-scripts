Sub CountInboxEmailsbySender()
    Dim objDictionary As Object
    Dim objSizeDictionary As Object
    
    Dim objInbox As Outlook.Folder
    Dim i As Long
    Dim objMail As Outlook.MailItem
    Dim strSender As String
    Dim intSize As Long

    Dim objExcelApp As Excel.Application
    Dim objExcelWorkbook As Excel.Workbook
    Dim objExcelWorksheet As Excel.Worksheet
    Dim varSenders As Variant
    Dim varItemCounts As Variant
    Dim varItemSizes As Variant
    Dim nLastRow As Integer
 
    Set objDictionary = CreateObject("Scripting.Dictionary")
    Set objSizeDictionary = CreateObject("Scripting.Dictionary")
    rem Defaults to deleted Items, change to Inbox or another folder here.
    Set objInbox = Outlook.Application.Session.GetDefaultFolder(olFolderDeletedItems)
    MsgBox (objInbox.Items.Count & " Total items in the folder")
    
    On Error Resume Next
    For i = objInbox.Items.Count To 1 Step -1
        If objInbox.Items(i).Class = olMail And objInbox.Items(i).MessageClass <> "IPM.Outlook.Recall" Then
           Set objMail = objInbox.Items(i)
           
           strSender = objMail.SenderEmailAddress
           intSize = objMail.Size
 
           If objDictionary.Exists(strSender) Then
              objDictionary.Item(strSender) = objDictionary.Item(strSender) + 1
           Else
              objDictionary.Add strSender, 1
           End If
           
           If objSizeDictionary.Exists(strSender) Then
              objSizeDictionary.Item(strSender) = objSizeDictionary.Item(strSender) + intSize
           Else
              objSizeDictionary.Add strSender, intSize
           End If
           
        End If
    Next

    Set objExcelApp = CreateObject("Excel.Application")
    objExcelApp.Visible = True
    Set objExcelWorkbook = objExcelApp.Workbooks.Add
    Set objExcelWorksheet = objExcelWorkbook.Sheets(1)
 
    With objExcelWorksheet
         .Cells(1, 1) = "Sender"
         .Cells(1, 2) = "Count"
         .Cells(1, 3) = "Size (kB)"
    End With
 
    varSenders = objDictionary.Keys
    varItemCounts = objDictionary.Items
 
    For i = LBound(varSenders) To UBound(varSenders)
        nLastRow = objExcelWorksheet.Range("A" & objExcelWorksheet.Rows.Count).End(xlUp).Row + 1
        With objExcelWorksheet
             .Cells(nLastRow, 1) = varSenders(i)
             .Cells(nLastRow, 2) = varItemCounts(i)
             .Cells(nLastRow, 3) = objSizeDictionary.Item(varSenders(i)) / 1000
        End With
    Next
 
    objExcelWorksheet.Columns("A:B").AutoFit
End Sub
