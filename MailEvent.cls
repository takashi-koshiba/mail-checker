VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Item As Outlook.mailItem
Attribute Item.VB_VarHelpID = -1
Private mailArray() As Outlook.mailItem

Private form1 As UserForm1

Public Sub setMail(form As UserForm1)
   ' mailArray = mail
    Set form1 = form
End Sub

Private Sub item_PropertyChange(ByVal Name As String)
    If Name = "FlagRequest" Or Name = "UnRead" Then 'FlagRequest UnRead
  
    
        Dim i As Integer

    
        For i = 0 To UBound(mailArray)
            If mailArray(i).ConversationID = Item.ConversationID Then
                mailArray(i) = Item
         End If
        
        Next i
    
    
        Dim lastUpdate As Date
        lastUpdate = GetLastReceivedTime(mailArray)
    
        Call setMailToList(lastUpdate, mailArray(0).Parent)
        
    End If
    
    
End Sub

Public Sub setMailArrEnd(m() As mailItem)
    mailArray = m

End Sub

Private Function GetLastReceivedTime(mail() As Outlook.mailItem) As Date
    Dim lastUpdated As Date
    lastUpdated = DateValue("2000/01/01")
    Dim mailIndex
    mailIndex = UBound(mail)
    Dim i As Integer
    For i = 0 To mailIndex - 1
        Call lastReceivedTimeOfItem(mail(i), lastUpdated)
    Next i
    
    GetLastReceivedTime = lastUpdated
End Function


Public Sub lastReceivedTimeOfItem(ByVal Item As Variant, listTime As Date)
    
    If listTime < Item.ReceivedTime Then
         listTime = Item.ReceivedTime
    End If
    
End Sub


Private Sub setMailToList(ByVal lastUpdated As Date, tgtFolder As folder)
    Dim arrSubItems As Variant
    Dim count_new_mail As Integer
    Dim count_unreadItem  As Integer
    Dim count_noFlag_new_mail  As Integer
    Dim count_noFlagItem As Integer
    Dim lastReceivedTimeOfNewMail As Date
    Dim lastReceivedTimeOfUnRead As Date
    Dim rootFolder As String
    
    rootFolder = GetReceivedItemRoot(tgtFolder).Name
    count_new_mail = 0
    count_unreadItem = 0
    count_noFlag_new_mail = 0
    count_noFlagItem = 0
    lastReceivedTimeOfNewMail = lastUpdated
    lastReceivedTimeOfUnRead = lastUpdated
    arrSubItems = form.GetSubItems(tgtFolder)
    
    Dim i As Integer
    Dim arr As Variant
    
    If Not mailArray(0) Is Nothing Then
    
        For Each arr In mailArray
            Call form1.countMail(rootFolder, arr, arrSubItems, count_new_mail, count_unreadItem, count_noFlag_new_mail, count_noFlagItem, lastReceivedTimeOfNewMail, lastReceivedTimeOfUnRead)

         Next arr
    End If
    

    
                
   Call form1.ReplaceListItem(form1.unRead_item_listview, tgtFolder.Parent, tgtFolder, CStr(count_unreadItem), CStr(count_noFlagItem), lastReceivedTimeOfUnRead, arrSubItems, False)
            
   Call form1.ReplaceListItem(form1.newMailListView, tgtFolder.Parent, tgtFolder, CStr(count_new_mail), CStr(count_noFlag_new_mail), lastReceivedTimeOfUnRead, arrSubItems, True)


 form1.unRead_item_listview.Sorted = True
    form1.newMailListView.Sorted = True
    
    
    Call form1.changeBackGroundColor

End Sub

