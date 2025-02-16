VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "subfolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Private WithEvents tgtCtrl As Outlook.Items
Attribute tgtCtrl.VB_VarHelpID = -1
Private WithEvents tgtFolder As Outlook.folder
Attribute tgtFolder.VB_VarHelpID = -1
Private form1 As UserForm1

Private mailIndex As Integer
Private mail() As Outlook.mailItem



Private WithEvents myOlExp  As Outlook.Explorer
Attribute myOlExp.VB_VarHelpID = -1

Private mailEvent() As mailEvent

Public Sub setForm(form As UserForm1)
    mailIndex = 0
    Set form1 = form
End Sub

Public Sub SetCtrl(ByVal new_item As Outlook.Items, folder As folder, exp As Outlook.Explorer)
  Set tgtCtrl = new_item
  
  Set tgtFolder = folder
  Set myOlExp = exp
  

End Sub

Public Sub setMail(mailItem As Variant)
    ReDim Preserve mail(mailIndex)
    ReDim Preserve mailEvent(mailIndex)
    
    
    Set mail(mailIndex) = mailItem
    Set mailEvent(mailIndex) = New mailEvent
   Set mailEvent(mailIndex).Item = mailItem
     mailEvent(mailIndex).setMail form1
    mailIndex = mailIndex + 1
    
End Sub


Public Sub setMailArr()
    Dim i As Integer
    On Error Resume Next
    
    For i = 0 To UBound(mail)
          mailEvent(i).setMailArrEnd mail

    Next i
    
   
On Error GoTo 0

End Sub


Private Sub myOlExp_SelectionChange()
 '   MsgBox 0
End Sub

Private Sub tgtCtrl_ItemAdd(ByVal Item As Object)
    Call setMail(Item)
    
    Dim lastUpdate As Date
    
    lastUpdate = GetLastReceivedTime(mail)
    
    
    
    Call setMailToList(lastUpdate)
End Sub


Private Function GetLastReceivedTime(mail() As Outlook.mailItem) As Date
    Dim lastUpdated As Date
    lastUpdated = DateValue("2000/01/01")
    
    Dim i As Integer
    For i = 0 To mailIndex - 1
        Call form1.lastReceivedTimeOfItem(mail(i), lastUpdated)
    Next i
    
    GetLastReceivedTime = lastUpdated
End Function






Private Sub setMailToList(ByVal lastUpdated As Date)
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
    
    If Not mail(0) Is Nothing Then
    
        For Each arr In mail
            Call form1.countMail(rootFolder, arr, arrSubItems, count_new_mail, count_unreadItem, count_noFlag_new_mail, count_noFlagItem, lastReceivedTimeOfNewMail, lastReceivedTimeOfUnRead)
    
         Next arr
    End If
    

    
                
   Call form1.ReplaceListItem(form1.unRead_item_listview, tgtFolder.Parent, tgtFolder, CStr(count_unreadItem), CStr(count_noFlagItem), lastReceivedTimeOfUnRead, arrSubItems, False)
            
   Call form1.ReplaceListItem(form1.newMailListView, tgtFolder.Parent, tgtFolder, CStr(count_new_mail), CStr(count_noFlag_new_mail), lastReceivedTimeOfUnRead, arrSubItems, True)


 form1.unRead_item_listview.Sorted = True
    form1.newMailListView.Sorted = True
    
    
    Call form1.changeBackGroundColor

End Sub


Private Sub tgtCtrl_ItemChange(ByVal Item As Object)
  Dim i As Integer
     
    On Error Resume Next
    Dim index As Integer
    index = UBound(mail)
    If index > 0 Then
    
        For i = 0 To UBound(mail)
            If mail(i).ConversationID = Item.ConversationID Then
                mail(i) = Item
            End If
         mailEvent(i).setMail form1
         Next i
    
    
        Dim lastUpdate As Date
        lastUpdate = GetLastReceivedTime(mail)
        
        Call setMailToList(lastUpdate)
    End If
    
    
     On Error GoTo 0

    
End Sub

Private Sub tgtFolder_BeforeItemMove(ByVal Item As Object, ByVal MoveTo As MAPIFolder, Cancel As Boolean)
    Dim i As Integer
    Dim ii As Integer
    For i = 0 To UBound(mail)
        If mail(i).ConversationID <> Item.ConversationID Then
            mail(ii) = mail(i)
            ii = ii + 1
        
        End If
    Next i
    
    If UBound(mail) = 0 Then
        ReDim mail(0)
        mailIndex = 0
        Call setMailToList(DateValue("2000/01/01"))
    ElseIf i <> ii Then
        ReDim Preserve mail(UBound(mail) - 1)
        mailIndex = UBound(mail) + 1
        Call setMailToList(DateValue("2000/01/01"))
    End If
    

    
  
End Sub
