VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10485
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents InboxItems As Outlook.Items
Private WithEvents myOlExp As Explorer
Attribute myOlExp.VB_VarHelpID = -1

Private OutlookApp As Outlook.Application
Private namespace As Outlook.namespace


Private subfolderCtl() As New subfolder
Private xml As msxml

Private defaultColor As color2
Private unReadColor  As color2
Private newMailColor As color2
Private receivedMailColor As color2
Private ReadColor As color2
Private newMailReadColor As color2

Private selectedColorButton As String

Private selectedColorLabel As label

Private colorStatus As Integer


Private Sub Add_allowedFolder_Click()

    
    
    Call xml.AddtoList(list_of_exclusion, allowed_folderList, False)
    'ReDim subfolderCtl(allowed_folderList.ListItems.count)
    'Call Add_mail(False)
    
    Call SelectCancelOfList(allowed_folderList)
    Call SelectCancelOfList(list_of_exclusion)

    

    
End Sub






Private Sub allowed_folderList_Click()
    Dim i As Integer
    i = 0
    For i = 1 To allowed_folderList.ListItems.count
        
        If allowed_folderList.ListItems.Item(i).Selected Then
        
            Dim IsGetNewMailListValue As Boolean
            IsGetNewMailListValue = allowed_folderList.ListItems.Item(i).ListSubItems(1).text
        
        
            If IsGetNewMailListValue Then
                ToggleButtonNewMail.Caption = "新規判定する"
            Else
                ToggleButtonNewMail.Caption = "新規判定しない"
            End If
        
            Dim IsGetNoflagMail As Boolean
            IsGetNoflagMail = allowed_folderList.ListItems.Item(i).ListSubItems(2).text
         
            If IsGetNoflagMail Then
                FlagToggleButton.Caption = "フラグを確認する"
            Else
                FlagToggleButton.Caption = "フラグを確認しない"
            End If
        End If
    Next i
End Sub

Private Sub blueslider_Change()
    Call changeColor
End Sub

Private Sub CheckBox1_Change()
    Call Add_mail(False)
    
    unRead_item_listview.Sorted = True
    newMailListView.Sorted = True
   
    Call changeBackGroundColor
End Sub

Public Sub CommandButton1_Click()
    Call Add_mail(False)
    
   unRead_item_listview.Sorted = True
   newMailListView.Sorted = True
   
   Call changeBackGroundColor
End Sub

Private Sub defaultColorButton_Click()

    
    Call setColorButton("default", defaultColor, Label2)
End Sub

Private Sub setColorButton(selectedButton As String, selectedSlider As color2, label As label)
    redslider.Visible = True
    greenslider.Visible = True
    blueslider.Visible = True
    
    If Not selectedButton = "default" Then
        enableButton.Visible = True
    Else
        enableButton.Visible = False
        
    End If
    
    selectedColorButton = selectedButton
    selectedSlider.SetSlider
    Set selectedColorLabel = label

End Sub

Private Sub del_allowed_folder_button_Click()

   
    Call xml.AddtoList(allowed_folderList, list_of_exclusion, True)
    'ReDim subfolderCtl(allowed_folderList.ListItems.count)
   ' Call Add_mail(True)
    
    Call SelectCancelOfList(allowed_folderList)
    Call SelectCancelOfList(list_of_exclusion)


End Sub


   
Private Sub enableButton_Click()
    If IsEnableColor(selectedColorLabel) Then
       selectedColorLabel.ForeColor = 255
    Else
        selectedColorLabel.ForeColor = -2147483630
    
    End If
    
End Sub

Private Function IsEnableColor(label As label)
    Dim result As Boolean
    
    If label.ForeColor = 255 Then
       result = False
    Else
        result = True
    
    End If
    IsEnableColor = result

End Function

Private Sub exclusion_of_sender_mail_button_Click()

    Dim subject
     subject = exclusion_of_sender_subject_textbox.text
    
    If subject = "" Then
        MsgBox "文字列を入力してください"
    
    Else
        Dim arr As Variant
    
        If exclusion_of__mailSubject_list.ListItems.count = 0 Then
            arr = Array("")
            
        
        Else
            Dim listCount As Integer
            listCount = exclusion_of__mailSubject_list.ListItems.count
            ReDim arr(listCount)
            
            
            Dim i As Integer
            For i = 0 To listCount - 1
                arr(i) = exclusion_of__mailSubject_list.ListItems.Item(i + 1).text
            Next i
            
        End If
    

    
        Dim ans As Integer
        ans = MsgBox("""" & subject & """" & "を追加しますか", vbYesNo)
        If ans = vbYes Then

            If IsExist(subject, arr) Then
                MsgBox """" & subject & """" & "は既に存在しています。"
    
            Else
            
                exclusion_of__mailSubject_list.ListItems.Add , , subject
                exclusion_of_sender_subject_textbox.text = ""
      
            End If
        
        End If
    
    End If
   
   Call SelectCancelOfList(exclusion_of__mailSubject_list)
   

End Sub




Private Sub exclusion_of_sender_mail_del_button_Click()
Dim list As Variant
    
    Dim i As Integer
    Dim ans As Integer
    Dim result As Boolean
    result = False
    
    If exclusion_of__mailSubject_list.ListItems.count > 0 Then
        For i = 1 To exclusion_of__mailSubject_list.ListItems.count
            If exclusion_of__mailSubject_list.ListItems.Item(i).Selected Then
            ans = MsgBox("""" & exclusion_of__mailSubject_list.ListItems.Item(i).text & """" & "を削除しますか？", vbYesNo)
           
                If ans = vbYes Then
                    
                    exclusion_of__mailSubject_list.ListItems.Remove (i)
                    
                  
                End If
                
                result = True
                Exit For
            End If
        Next
    End If
    If Not result Then MsgBox "削除する項目を選択してください。"
    Call SelectCancelOfList(exclusion_of__mailSubject_list)


End Sub


Private Sub FlagToggleButton_Click()
    
    
    Call ToggleButton(allowed_folderList, FlagToggleButton, 2)

    Call SelectCancelOfList(allowed_folderList)
    Call SelectCancelOfList(list_of_exclusion)
    
End Sub

Private Sub greenslider_Change()
    
    Call changeColor
End Sub


Private Sub MultiPage1_Change()
    Call changeListTab(Me.MultiPage1, Me.MultiPage1.value)
    Call changeListTab(Me.MultiPage2, Me.MultiPage2.value)
End Sub

Private Sub MultiPage2_Change()
    
    Call changeListTab(Me.MultiPage2, Me.MultiPage2.value)
    
   
End Sub



'メールを選択したとき
Private Sub myOlExp_SelectionChange()
    Dim selectedItems As Selection
    Dim currentItem As Object
    
    '選択したメールのアイテムを取得
    Set selectedItems = myOlExp.Selection
    
    If selectedItems.count > 0 Then
        Dim folderPath As String
        
        '選択したアイテムのフォルダパスを取得
        folderPath = selectedItems.Item(1).Parent.folderPath
        
        'リストの文字を太字から元に戻す
        Call toggleFontWeightOfListItems(unRead_item_listview, folderPath, False)
        Call toggleFontWeightOfListItems(newMailListView, folderPath, False)
        
        Call changeBackGroundColor
    End If
End Sub

Private Sub newMailReadColorButton_Click()

    
     Call setColorButton("newMailRead", newMailReadColor, Label4)

End Sub

Private Sub newMailUnReadColorButton_Click()

   
   Call setColorButton("newMail", newMailColor, Label3)
   
End Sub

Private Sub OtherReadMailColorButton_Click()

    
    Call setColorButton("Read", ReadColor, Label5)

End Sub

Private Sub receivedMailColorButton_Click()

   
    
    Call setColorButton("receivedMail", receivedMailColor, labe4)

    
    
End Sub


Private Sub redslider_Change()
    Call changeColor
    

    
End Sub

Private Sub saveButton_Click()
    Call Add_mail(True)
    Call xml.AddSubItemToRoot(allowed_folderList, "IsNotNewMail", "check", 1)
    Call xml.AddSubItemToRoot(allowed_folderList, "IsNoFlag", "check", 2)
        
    Call receivedMailColor.changeColorXml("receivedMailColor", xml)
    Call newMailColor.changeColorXml("newMailColor", xml)
    Call unReadColor.changeColorXml("unReadColor", xml)
    Call defaultColor.changeColorXml("defaultColor", xml)
    Call ReadColor.changeColorXml("ReadColor", xml)
    Call newMailReadColor.changeColorXml("newMailRead", xml)
    
    Call receivedMailColor.changeColorStatusXml("receivedMailEnable", xml)
    Call newMailColor.changeColorStatusXml("newMailEnable", xml)
    Call unReadColor.changeColorStatusXml("unReadEnable", xml)
    Call defaultColor.changeColorStatusXml("defaultEnable", xml)
    Call ReadColor.changeColorStatusXml("ReadEnable", xml)
    Call newMailReadColor.changeColorStatusXml("newMailReadEnable", xml)
    
    Call xml.AddItemToRoot(exclusion_of__mailSubject_list, "exclusionMailSubject", "mailsubject")
    Call xml.saveXml(allowed_folderList)
    

     Call changeListTab(Me.MultiPage1, Me.MultiPage1.value)
      Call changeListTab(Me.MultiPage2, Me.MultiPage2.value)
      
      MsgBox "設定を保存しました。"
      
    Call SelectCancelOfList(allowed_folderList)
    Call SelectCancelOfList(list_of_exclusion)
End Sub

Private Sub ToggleButtonNewMail_Click()
  
    
    Call ToggleButton(allowed_folderList, ToggleButtonNewMail, 1)

    
    Call SelectCancelOfList(allowed_folderList)
    Call SelectCancelOfList(list_of_exclusion)
    
End Sub


Private Sub ToggleButton(list As listview, button As CommandButton, subItemIndex As Integer)
 'リストが選択されているか
    If Not list.selectedItem Is Nothing Then
        Dim value As Boolean
        value = list.selectedItem.ListSubItems(subItemIndex).text
        
        Dim ans As Integer
        ans = MsgBox(button.Caption & "に変更しますか。", vbYesNo)
    
        If ans = vbYes Then
            If value = 0 Then
                list.selectedItem.ListSubItems(subItemIndex).text = True
            Else
                list.selectedItem.ListSubItems(subItemIndex).text = False
            End If
        End If
    Else
        MsgBox "リストが選択されていません"
    End If
    
    Call Add_mail(False)
    allowed_folderList.selectedItem = Nothing
    
End Sub
Private Sub OtherRunReadMailColorButton_Click()
  
    Call setColorButton("unRead", unReadColor, label)

End Sub



Private Sub UserForm_Initialize()
    
    colorStatus = 0
    
    enableButton.Visible = False
    
    
    Set defaultColor = New color2
    Set unReadColor = New color2
    Set newMailColor = New color2
    Set receivedMailColor = New color2
    Set ReadColor = New color2
    Set newMailReadColor = New color2
    
    Set xml = New msxml
    xml.initXml
    
    
    Call defaultColor.SetItem(defaultColorButton, redslider, greenslider, blueslider, "defaultColor", xml, Label2, "defaultEnable")
    Call unReadColor.SetItem(OtherRunReadMailColorButton, redslider, greenslider, blueslider, "unReadColor", xml, label, "unReadEnable")
    Call newMailColor.SetItem(newMailUnReadColorButton, redslider, greenslider, blueslider, "newMailColor", xml, Label3, "newMailEnable")
    Call receivedMailColor.SetItem(receivedMailColorButton, redslider, greenslider, blueslider, "receivedMailColor", xml, labe4, "receivedMailEnable")
    Call ReadColor.SetItem(OtherReadMailColorButton, redslider, greenslider, blueslider, "ReadColor", xml, Label5, "ReadEnable")
    Call newMailReadColor.SetItem(newMailReadColorButton, redslider, greenslider, blueslider, "newMailRead", xml, Label4, "newMailReadEnable")
    
    
    '設定ファイルから色を取得し、ボタンの色を設定する
    defaultColorButton.BackColor = defaultColor.getColorCode
    newMailUnReadColorButton.BackColor = newMailColor.getColorCode
    OtherRunReadMailColorButton.BackColor = unReadColor.getColorCode
    receivedMailColorButton.BackColor = receivedMailColor.getColorCode
    newMailReadColorButton.BackColor = newMailReadColor.getColorCode
    OtherReadMailColorButton.BackColor = ReadColor.getColorCode
     
   

     
     
    Set myOlExp = Application.ActiveExplorer


          'Outlookのインスタンスを生成

    Set OutlookApp = New Outlook.Application
    Set namespace = OutlookApp.GetNamespace("MAPI")
  

    Dim i As Integer
    
    Dim listItem As MSComctlLib.listItem
    
    With Me.unRead_item_listview
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        
        ' カラムの数を指定
        .ColumnHeaders.Clear
        
       
        .ColumnHeaders.Add , , "親フォルダ", 100
        .ColumnHeaders.Add , , "子フォルダ", 100
        .ColumnHeaders.Add , , "24時間以内の未読数", 50
        .ColumnHeaders.Add , , "対応中", 140
        .ColumnHeaders.Add , , "変更前の未読数", 0
        .ColumnHeaders.Add , , "変更前の対応中の数", 0
        .ColumnHeaders.Add , , "更新日時", 0
        
        .SortKey = 6
        .Sorted = True
        .SortOrder = lvwDescending
        
    End With

    With Me.newMailListView
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        
        ' カラムの数を指定
        .ColumnHeaders.Clear
        
       
        .ColumnHeaders.Add , , "親フォルダ", 100
        .ColumnHeaders.Add , , "子フォルダ", 100
        .ColumnHeaders.Add , , "新規未読数", 50
        .ColumnHeaders.Add , , "対応中", 140
        .ColumnHeaders.Add , , "変更前の新規未読数", 0
        .ColumnHeaders.Add , , "変更前の対応中の数", 0
        .ColumnHeaders.Add , , "更新日時", 0
      
        .SortKey = 6
        .Sorted = True
        .SortOrder = lvwDescending
        
    End With
    
    
    With allowed_folderList
        .View = lvwReport
        .GridLines = False
        .FullRowSelect = True
        
        ' カラムの数を指定
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "件数を取得するフォルダ", allowed_folderList.Width * 0.8
        .ColumnHeaders.Add , , "新規なし", allowed_folderList.Width * 0.1
        .ColumnHeaders.Add , , "フラグなしをカウントしない", allowed_folderList.Width * 0.1
       
       
        
    End With
    
    With list_of_exclusion
        .View = lvwReport
        .GridLines = False
        .FullRowSelect = True
        
        ' カラムの数を指定
        .ColumnHeaders.Clear
         .ColumnHeaders.Add , , "除外したフォルダ", list_of_exclusion.Width
       

        
    End With
        
    With exclusion_of__mailSubject_list
        .View = lvwReport
        .GridLines = False
        .FullRowSelect = True
        .ColumnHeaders.Clear ' 既存のカラムをクリア
        .ColumnHeaders.Add , , "除外する件名", exclusion_of__mailSubject_list.Width
       
    End With

   Call AddConfigToList(allowed_folderList)

  
   
   Call Add_mail(True)
  
   

    Call changeBackGroundColor
    
    '設定ファイルの値をリストに入れる
    Dim arrayExclusionMailSubject As Variant
    arrayExclusionMailSubject = xml.getXml("exclusionMailSubject")
    Dim Item
    For Each Item In arrayExclusionMailSubject
        If Item = "" Then
        Else
            exclusion_of__mailSubject_list.ListItems.Add , , Item
        End If
        
    Next Item
    
    
    
    Call changeListTab(Me.MultiPage1, Me.MultiPage1.value)
    Call changeListTab(Me.MultiPage2, Me.MultiPage2.value)
    
    


End Sub

'メールを受信すると発火


Private Function AddBoldItemToArray(arr As Variant, list As listview) As Variant
        
        arr = Array("")
        Dim Item As Variant
        
        For Each Item In list.ListItems
            If Item.Bold Then
                ReDim arr(UBound(arr))
                arr(UBound(arr)) = Item.text & "\" & Item.ListSubItems(1).text
            End If
        Next Item
        
        AddBoldItemToArray = arr
End Function



'すべてのメーリングリストをループ
Public Sub Add_mail(list_of_exclusionMode As Boolean)
    ReDim subfolderCtl(allowed_folderList.ListItems.count)
    Dim result

    Dim countNotExistFolder As Integer
                               
    
        Dim arrUnRead As Variant
        arrUnRead = AddBoldItemToArray(arrUnRead, unRead_item_listview)
        
        Dim arrNewMail As Variant
        arrNewMail = AddBoldItemToArray(arrNewMail, newMailListView)
        
        unRead_item_listview.ListItems.Clear
        
        newMailListView.ListItems.Clear
        If list_of_exclusionMode = True Then
            list_of_exclusion.ListItems.Clear
        End If
     


        Dim arr() As String
    
        
        Dim existAllowFolderArr As Variant
        
        If allowed_folderList.ListItems.count > 0 Then
        
            ReDim existAllowFolderArr(allowed_folderList.ListItems.count - 1)
        
            Dim MailBox As Outlook.MAPIFolder
            Dim listIndex As Integer
            listIndex = 0
            Dim listIndex2 As Integer
            listIndex2 = 0
            For Each MailBox In namespace.Folders  'アカウントの数だけ繰り返す
               
                Dim MailBox_Item As Outlook.MAPIFolder
                
                Call run_subfolder(MailBox, list_of_exclusionMode, MailBox.Name, countNotExistFolder, listIndex, existAllowFolderArr)  ' 関数を実行
            
            
              
                listIndex2 = listIndex2 + 1
            Next MailBox
        

            'existAllowFolderArrがリストに存在しない場合はその項目の色を赤にする
            Dim Item As Variant
            For Each Item In allowed_folderList.ListItems
                Dim i As Integer
                i = 1
            
                Dim ExistItem As Variant
                For Each ExistItem In existAllowFolderArr
                    Dim notExist As Boolean
                    
                    
                    If IsEmpty(ExistItem) Then
                        allowed_folderList.ListItems.Item(i).ForeColor = 255
                        notExist = True
                    End If
                    i = i + 1
                Next ExistItem
            
            Next Item
        
            If notExist Then
                MsgBox "存在しないフォルダが指定されています。", vbInformation
            End If
        
        
            'リストの太字を元に戻す
            Dim ii As Integer
        
            For ii = 0 To UBound(arrUnRead)
                Call toggleFontWeightOfListItems(Me.unRead_item_listview, arrUnRead(ii), True)
            Next ii
        
            For ii = 0 To UBound(arrNewMail)
                Call toggleFontWeightOfListItems(Me.newMailListView, arrNewMail(ii), True)
            Next ii
      
        
        End If
        Call changeBackGroundColor
End Sub



'サブフォルダの未読数を再帰処理で実行
Public Sub run_subfolder(folder As Outlook.MAPIFolder, list_of_exclusionMode As Boolean, _
                             rootFolder As String, countNotExistFolder As Integer, _
                             list_index As Integer, existAllowFolderArr As Variant)
   
   Dim subfolder As Outlook.MAPIFolder

   Dim countNotMail
   countNotMail = 0
   For Each subfolder In folder.Folders
        
         '再帰処理
         Call run_subfolder(subfolder, list_of_exclusionMode, rootFolder, countNotExistFolder, list_index, existAllowFolderArr)  '再起処理
          
         
        '指定したフォルダのみ取得
        '必要がなければ変更があったフォルダのみ取得
        If IsExist3(subfolder.Parent.fullFolderPath & "\", subfolder.Name, allowed_folderList.ListItems, existAllowFolderArr) = True Then
           
            
            Dim count_unreadItem
            count_unreadItem = 0
            Dim count_noFlagItem
            count_noFlagItem = 0
            Dim count_new_mail
            count_new_mail = 0

            Dim count_noFlag_new_mail
            count_noFlag_new_mail = 0

           
            Dim sortmail_item As Outlook.Items
            Set sortmail_item = subfolder.Items

            
            'メールの未読数などをリストに追加する
            Call AddNumberOfMailToList(0, 0, 0, 0, sortmail_item, unRead_item_listview, newMailListView, list_index, Me, subfolder, rootFolder, folder, False)
       
            
        Else
            
            If list_of_exclusionMode = True And addExclusionFolder(subfolder) = False Then
                 
                 '予定表などメール以外のフォルダを除外
                 If canSortMail(subfolder.Items) And subfolder <> "送信トレイ" Then
                    list_of_exclusion.ListItems.Add , , subfolder.Parent.fullFolderPath & "\" & subfolder.Name
                 End If
            End If
       
        End If
    
    Next subfolder
    
    

End Sub



'xmlから取得してリストに追加する
Private Sub AddConfigToList(list As listview)
    Dim allowedFolder As Variant
    Dim IsNotNewMail As Variant
    Dim IsNoFlag As Variant
    
    Dim str As Variant
    allowedFolder = xml.getXml("allowedFolder")
    IsNotNewMail = xml.getXml("IsNotNewMail")
    IsNoFlag = xml.getXml("IsNoFlag")
    
    Dim listItem As MSComctlLib.listItem
    list.ListItems.Clear
    
    Dim i As Integer
    i = 1
    '設定ファイルの要素数が一致するか
    If UBound(allowedFolder) = UBound(IsNotNewMail) And UBound(IsNotNewMail) = UBound(IsNoFlag) Then
    
        
        For Each str In allowedFolder
            Call list.ListItems.Add(, , str)
            list.ListItems.Item(i).SubItems(1) = IsNotNewMail(i - 1)
            list.ListItems.Item(i).SubItems(2) = IsNoFlag(i - 1)
         
            i = i + 1
        Next str
    Else
        
        For Each str In allowedFolder
            Call list.ListItems.Add(, , str)
            list.ListItems.Item(i).SubItems(1) = False
            list.ListItems.Item(i).SubItems(2) = False
         
            i = i + 1
        Next str
        
        MsgBox "設定ファイルに問題があります。"
    End If
    

       
    

  
    
End Sub


'許可していないフォルダーを除外フォルダーに入れる
Private Function addExclusionFolder(subfolder As Outlook.MAPIFolder) As Boolean

    Dim Item As Variant
    Dim result As Boolean
    Dim i As Long
    
    result = False
    
    If Not (allowed_folderList Is Nothing) Then
        If IsArray(allowed_folderList.ListItems) Then
            For Each Item In allowed_folderList.ListItems
                If subfolder.Parent.fullFolderPath & "\" & subfolder.Name = Item Then
                    result = True
                    Exit For
                End If
            Next Item
        End If
    End If
    
    addExclusionFolder = result
        
End Function





Private Sub changeColor()
    
    If selectedColorButton = "default" Then
        Call defaultColor.SetRGB(redslider.value, greenslider.value, blueslider.value)
        defaultColorButton.BackColor = defaultColor.getColorCode
        
        
    ElseIf selectedColorButton = "unRead" Then
        Call unReadColor.SetRGB(redslider.value, greenslider.value, blueslider.value)
        OtherRunReadMailColorButton.BackColor = unReadColor.getColorCode
        
    ElseIf selectedColorButton = "newMail" Then
        Call newMailColor.SetRGB(redslider.value, greenslider.value, blueslider.value)
        newMailUnReadColorButton.BackColor = newMailColor.getColorCode
     
     ElseIf selectedColorButton = "Read" Then
        Call ReadColor.SetRGB(redslider.value, greenslider.value, blueslider.value)
        OtherReadMailColorButton.BackColor = ReadColor.getColorCode
     
     ElseIf selectedColorButton = "newMailRead" Then
        Call newMailReadColor.SetRGB(redslider.value, greenslider.value, blueslider.value)
        newMailReadColorButton.BackColor = newMailReadColor.getColorCode
    Else
        Call receivedMailColor.SetRGB(redslider.value, greenslider.value, blueslider.value)
        receivedMailColorButton.BackColor = receivedMailColor.getColorCode
        
    End If
    
    Call changeBackGroundColor
End Sub


Public Function GetSubItems(ByVal folder As folder) As Variant
    Dim Item As Variant
    For Each Item In allowed_folderList.ListItems
        If folder.folderPath = Item Then
            GetSubItems = Array(Item.SubItems(1), Item.SubItems(2))
            Exit For
            
        End If
        
    Next Item

End Function


'メールの未読数などをリストに追加
Public Sub AddNumberOfMailToList(count_unreadItem As Integer, count_noFlagItem As Integer, count_new_mail As Integer, _
                    count_noFlag_new_mail As Integer, sortmail_item As Outlook.Items, _
                    unRead_item_listview As listview, newMailListView As listview, list_index As Integer, form As UserForm, _
                     subfolder As folder, rootFolder As String, folder As folder, IsReplaceItems As Boolean)
    
    
    'リストから新規判定にしないのと、フラグなしを取得
    Dim arrSubItems As Variant
    arrSubItems = GetSubItems(subfolder)
    If canSortMail(sortmail_item) Then
                
        '受信日時で並び替え
        sortmail_item.sort "[ReceivedTime]", True
        
        Dim lastReceivedTimeOfNewMail As Date
        Dim lastReceivedTimeOfUnRead As Date
        lastReceivedTimeOfNewMail = DateValue("2000/01/01")
        lastReceivedTimeOfUnRead = DateValue("2000/01/01")
               
        Dim mail_item As Variant
        For Each mail_item In sortmail_item
            
            If IsExistPropatyOfReceivedTime(mail_item) Then
            
                '受信してから24時間以内のメールを取得
                If (DateDiff("d", mail_item.ReceivedTime, Now)) <= 1 Then
                            
                    '差出人と受取先が同じなら何もしない
                    '件名が除外する件名なら何もしない
                    If rootFolder = mail_item.SenderEmailAddress Or IsExistExclusion_of__mailSubject(mail_item.subject) Then
                        '何もしない
                    Else
                        
                        'メールを配列にいれる
                        
                        Call subfolderCtl(list_index).setmail(mail_item)
                        
                        Call countMail(rootFolder, mail_item, arrSubItems, count_new_mail, count_unreadItem, count_noFlag_new_mail, count_noFlagItem, lastReceivedTimeOfNewMail, lastReceivedTimeOfUnRead)
                    End If
                            
                            
                            
                
                Else
                    '24時間以上経過のメールがあったらループを抜ける
                    Exit For
                    
                End If
                
            End If
            

        Next mail_item
                
        'リストを追加せず置き換えるか
        If IsReplaceItems Then
            
            
            Call ReplaceListItem(unRead_item_listview, folder, subfolder, CStr(count_unreadItem), CStr(count_noFlagItem), lastReceivedTimeOfUnRead, arrSubItems, False)
            
            Call ReplaceListItem(newMailListView, folder, subfolder, CStr(count_new_mail), CStr(count_noFlag_new_mail), lastReceivedTimeOfUnRead, arrSubItems, True)
           
            
        Else

            
            Call AddNewItemToList(unRead_item_listview, folder.folderPath, subfolder.Name, CStr(count_unreadItem), CStr(count_noFlagItem), lastReceivedTimeOfUnRead, arrSubItems, False)
             
            Call AddNewItemToList(newMailListView, folder.folderPath, subfolder.Name, CStr(count_new_mail), CStr(count_noFlag_new_mail), lastReceivedTimeOfNewMail, arrSubItems, True)
                
            'メールアイテムにイベントを追加
            Call subfolderCtl(list_index).SetCtrl(subfolder.Items, form, subfolder, Application.ActiveExplorer)
            'list_index
        End If
        
        
                
    End If
        
    list_index = list_index + 1




End Sub


'引数で受け取ったフォルダパスがリストにあればその項目の文字の太字を切り替える
Private Sub toggleFontWeightOfListItems(list As listview, ByVal folderPath As String, ByVal IsBold As Boolean)
        Dim Item As Variant
        For Each Item In list.ListItems
            If Item.text & "\" & Item.ListSubItems(1).text = folderPath Then
               Call ChangeItemBold(Item, IsBold)
            End If
        Next Item
End Sub



'指定したリストの項目の太字を切り替える
Private Sub ChangeItemBold(Item As Variant, ByVal IsBold)
    Item.Bold = IsBold
    
    Dim count As Integer
    Dim i As Integer
    count = Item.ListSubItems.count
    Item.Bold = IsBold
    'サブアイテムの数だけループ
    For i = 1 To count
        Item.ListSubItems.Item(i).Bold = IsBold
    Next i
    
    'リストを選択しないと太字が戻らないので
    'リストを選択して太字から元に戻す

    Item.Selected = True
    Item.Selected = False
 
    
    
End Sub


'最終更新日時を書き換える
Private Sub lastReceivedTimeOfItem(ByVal Item As Variant, listTime As Date)
    
    If listTime < Item.ReceivedTime Then
        listTime = Item.ReceivedTime
    End If
    
End Sub

'同じリストがあれば追加せず書き換える
Public Function ReplaceListItem(list As listview, folder As folder, subfolder As folder, _
                           ByVal count_item As String, count_noFlagItem As String, lastReceivedTimeOfUnRead As Date, arrSubItems As Variant, IsNewMailList As Boolean) As Boolean
    Dim index As Integer
    Dim IsExistListItem As Boolean
    IsExistListItem = False
    
        '同じリストが存在するかの確認
        For index = 1 To list.ListItems.count
            Dim ListStr As String
            ListStr = list.ListItems(index).text & list.ListItems(index).SubItems(1)
                
            If ListStr = folder.folderPath & subfolder.Name Then
                list.ListItems.Item(index) = folder.folderPath
                list.ListItems.Item(index).SubItems(1) = subfolder.Name
                
                '新規判定しない場合は置き換える
                If IsNewMailList And arrSubItems(0) = True Then
                    count_item = "-"
                Else
                End If
                
                'フラグなしの場合は「ー」に置き換える
                If arrSubItems(1) = True Then
                    count_noFlagItem = "-"
                End If
                
                list.ListItems.Item(index).SubItems(2) = count_item
                list.ListItems.Item(index).SubItems(3) = count_noFlagItem
                    
                '前回よりも未読が増えたとき
                If count_item > list.ListItems.Item(index).SubItems(4) Then
 
                    Call ChangeItemBold(list.ListItems.Item(index), True)
                        
                    '更新日時を入れる
                    list.ListItems.Item(index).SubItems(6) = lastReceivedTimeOfUnRead
                        
                End If
                
                
                '比較用に数を書き込む
                list.ListItems.Item(index).SubItems(4) = count_item
                list.ListItems.Item(index).SubItems(5) = count_noFlagItem
                
                IsExistListItem = True
                
                'チェックなしで0件なら削除
                If Me.CheckBox1.value = False And (count_item = "0" Or count_item = "-") And (count_noFlagItem = "0" Or count_noFlagItem = "-") Then
                    list.ListItems.Remove (index)
                    Exit For
                End If
                
            End If
        Next index
    
    'リストが見つからない場合はリストに追加する
    If Not IsExistListItem Then
        Call AddNewItemToList(list, folder.folderPath, subfolder.Name, count_item, count_noFlagItem, lastReceivedTimeOfUnRead, arrSubItems, IsNewMailList)
            
    End If
    
    ReplaceListItem = IsExistListItem
End Function

'新規を判定する
Private Function IsNewMail(subject As String) As Boolean
    IsNewMail = InStr(UCase(subject), "RE:") = 0
End Function


'リストにアイテムを追加する
Private Sub AddNewItemToList(listview As listview, ByVal folderPath As String, subfolderName As String, _
                             ByVal count_unreadItem As String, ByVal count_noFlagItem As String, lastReceivedTimeOfUnRead As Date, ByVal arrSubItems As Variant, ByVal IsNewMailList As Boolean)
    
    '新規判定しない場合は置き換える
    If IsNewMailList And arrSubItems(0) = True Then
        count_unreadItem = "-"
    Else
    End If
                
    'フラグなしの場合は「ー」に置き換える
    If arrSubItems(1) = True Then
        count_noFlagItem = "-"
    End If
    
    '未読数をリストに追加
    Dim list As listItem
    Set list = listview.ListItems.Add(, , folderPath) ' 最初のアイテムを追加
    list.SubItems(1) = subfolderName ' 2番目の列にサブアイテムを追加
    list.SubItems(2) = count_unreadItem
    list.SubItems(3) = count_noFlagItem
            
    '比較用に数を書き込む
    list.SubItems(4) = count_unreadItem
    list.SubItems(5) = count_noFlagItem
                    
    '更新時間を入れる
    list.SubItems(6) = lastReceivedTimeOfUnRead
    
    'チェックなしで0件なら削除
    If Me.CheckBox1.value = False And (count_unreadItem = "0" Or count_unreadItem = "-") And (count_noFlagItem = "0" Or count_noFlagItem = "-") Then
        listview.ListItems.Remove (listview.ListItems.count)
    End If
End Sub







Public Sub changeBackGroundColor()
    Dim Item
    Dim status As Integer
    
    
    
    status = GetColorStatusFromList(Me.newMailListView, IsEnableColor(Label2), IsEnableColor(Label3), IsEnableColor(Label4))
    If status = 0 Then
        status = GetColorStatusFromList(Me.unRead_item_listview, IsEnableColor(Label2), IsEnableColor(label), IsEnableColor(Label5)) + 4
    End If
    
    Dim color As String
    If status = 3 Or status = 7 Then
        color = receivedMailColorButton.BackColor
    ElseIf status = 2 Then
        color = newMailUnReadColorButton.BackColor
    ElseIf status = 1 Then
        color = newMailReadColorButton.BackColor
    ElseIf status = 6 Then
        color = OtherRunReadMailColorButton.BackColor
    ElseIf status = 5 Then
        color = OtherReadMailColorButton.BackColor
    ElseIf status = 4 Then
        color = defaultColorButton.BackColor
    End If
    
    
    
    
    Me.BackColor = color
'新着　　　　　3,7
'新規未読      2
'新規既読      1
'その他未読    6
'その他既読    5
'デフォ        4
End Sub

'リストからステータスを取得する
Private Function GetColorStatusFromList(list As listview, EnableReceivedMail As Boolean, EnableUnReadMail As Boolean, EnableReadMail) As Integer
    Dim status As Integer
    status = 0
    Dim Item
    Dim i As Integer
    i = 1
    
    For Each Item In list.ListItems
        
        If Item.Bold And EnableReceivedMail Then
            status = 3
        ElseIf Item.SubItems(2) > 0 And status < 2 And EnableUnReadMail Then
            status = 2
        ElseIf Item.SubItems(3) > 0 And status < 1 And EnableReadMail Then
            status = 1
        End If
        i = i + 1
    Next Item
    GetColorStatusFromList = status
    
End Function

'メールの件名が除外する件名と一致するか
Private Function IsExistExclusion_of__mailSubject(mailSubject As String) As Boolean
    Dim Item
    Dim result As Boolean
    result = False
    
    For Each Item In exclusion_of__mailSubject_list.ListItems
        If Item.text = mailSubject Then
            result = True
            Exit For
        End If
    Next Item
    IsExistExclusion_of__mailSubject = result
End Function


'選択状態のリストを選択なしに変更する
Private Sub SelectCancelOfList(list As listview)
    Dim Item As Variant
    For Each Item In list.ListItems
        Item.Selected = False
    Next Item
End Sub

Public Sub countMail(rootFolder As String, mail_item As Variant, arrSubItems As Variant, count_new_mail As Integer, count_unreadItem As Integer, count_noFlag_new_mail As Integer, count_noFlagItem As Integer, lastReceivedTimeOfNewMail As Date, lastReceivedTimeOfUnRead As Date)

                        
                        '未読か
                        If mail_item.unRead = True Then
                                                                             '0以外なら新規判定にしない
                            If IsNewMail(mail_item.subject) And arrSubItems(0) = False Then
                                '新規で未読をカウント
                                count_new_mail = count_new_mail + 1
                                
                                'フォルダの最終更新日時を取得
                                Call lastReceivedTimeOfItem(mail_item, lastReceivedTimeOfNewMail)
                            Else
                                '新規以外の未読をカウント
                                count_unreadItem = count_unreadItem + 1 '未読数をカウント
                                
                                'フォルダの最終更新日時を取得
                                Call lastReceivedTimeOfItem(mail_item, lastReceivedTimeOfUnRead)
                            End If
                        Else
                                
                            'フラグがないか
                            If mail_item.FlagStatus = olNoFlag Then
                            
                                '新規か
                                If IsNewMail(mail_item.subject) And arrSubItems(0) = False Then
                                    count_noFlag_new_mail = count_noFlag_new_mail + 1
                                Else
                                    count_noFlagItem = count_noFlagItem + 1
                                End If
                            End If
                            
                            
                            If IsNewMail(mail_item.subject) Then
                                'フォルダの最終更新日時を取得
                                Call lastReceivedTimeOfItem(mail_item, lastReceivedTimeOfNewMail)
                            Else
                             'フォルダの最終更新日時を取得
                                Call lastReceivedTimeOfItem(mail_item, lastReceivedTimeOfUnRead)
                            End If
                            
                            

                        End If
                    
                  
                            
                         
End Sub

