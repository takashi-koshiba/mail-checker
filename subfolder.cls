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




Public Sub SetCtrl(ByVal new_item As Outlook.Items, form As UserForm, folder As folder, exp As Outlook.Explorer)
  Set tgtCtrl = new_item
  Set form1 = form
  Set tgtFolder = folder
  Set myOlExp = exp
  

End Sub

Public Sub setmail(mailItem As Variant)
    ReDim Preserve mail(mailIndex)
    Set mail(mailIndex) = mailItem
    mailIndex = mailIndex + 1
    
End Sub




Private Sub myOlExp_SelectionChange()
 '   MsgBox 0
End Sub

Private Sub tgtCtrl_ItemAdd(ByVal Item As Object)
    Call setmail(Item)
    Call setMailToList
End Sub

Private Sub tgtCtrl_ItemChange(ByVal Item As Object)

    Dim i As Integer
    For i = 0 To UBound(mail)
        If mail(i).ConversationID = Item.ConversationID Then
            mail(i) = Item
        End If
        
    Next i
    
    
    Call setMailToList
    

End Sub



Private Sub setMailToList()
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
    lastReceivedTimeOfNewMail = DateValue("2000/01/01")
    lastReceivedTimeOfUnRead = DateValue("2000/01/01")
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
        Call setMailToList
    ElseIf i <> ii Then
        ReDim Preserve mail(UBound(mail) - 1)
        mailIndex = UBound(mail) + 1
        Call setMailToList
    End If
    

    
  
End Sub
