VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "msxml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private xml As Object
Private fileName As String
Private nodes As IXMLDOMNode
Private IsSaved As Boolean

Private Sub Class_Initialize()
    Call resetXml
    IsSaved = True
    
End Sub

Public Function getXml(element As String) As Variant

    
    Dim elements As Object
    Dim obj As Object
    Dim ii As Integer
    ii = 0
    '�w�肵���v�f��I��
    For Each obj In xml.getElementsByTagName("list").Item(0).ChildNodes
        If obj.BaseName = element Then
            Set elements = obj.ChildNodes
        End If
        
        
        ii = ii + 1
    Next obj
    
    

    Dim result_array As Variant
    Dim i As Integer
    If elements.Length > 0 Then
    
        ReDim result_array(elements.Length - 1)

        For i = 0 To elements.Length - 1
            result_array(i) = elements.Item(i).text
        Next i

        
    Else
         ReDim result_array(0)
    End If
    
        getXml = result_array
        
    

End Function

Public Sub initXml()
 
    
    Set xml = New MSXML2.DOMDocument60
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    fileName = wsh.SpecialFolders("AppData") & "\Microsoft\Outlook\config.xml"
    xml.Load fileName
    
    Dim ixml As IXMLDOMProcessingInstruction
    Set ixml = xml.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    
    ' �ŏ��̍s��XML�錾��ǉ�����
    If xml.ChildNodes.Length = 0 Then
        xml.appendChild ixml
    End If
    
    

    
   ' Call initElements(parentNode) ' xml������������
    
   ' Set nodes = xml.appendChild(xml.createElement(parentNode))
  ' Set nodes = xml.ChildNodes
   
End Sub

Public Function getFileName() As String
    getFileName = fileName

End Function


Public Sub AddSubItemToRoot(list As listview, parentNode As String, element As String, index As Integer)
   ' Call initXml(parentNode)
    Call initElements(parentNode)
    
    
    Dim node As IXMLDOMNode
    Dim str As Variant
    For Each str In list.ListItems
        If IsNull(str) Then Exit For
        
        Set node = xml.createElement(element)
        node.text = str.SubItems(index)
        nodes.appendChild node
    Next str

    
End Sub


Public Sub AddItemToRoot(list As listview, parentNode As String, element As String)
    Call initElements(parentNode)
    
    
    Dim node As IXMLDOMNode
    Dim str As Variant
    For Each str In list.ListItems
        If IsNull(str) Then Exit For
        
        Set node = xml.createElement(element)
        node.text = str
        nodes.appendChild node
    Next str
End Sub

Public Sub writeXmlColor(parentNode As String, element As String, colorCode As Long)
    Call initElements(parentNode) ' xml������������
    
    Dim node As IXMLDOMNode
    Dim str As Variant
   
        
    Set node = xml.createElement(element)
    node.text = colorCode
    nodes.appendChild node
End Sub


Public Sub writeXmlEnableColor(parentNode As String, element As String, flag As Boolean)
    Call initElements(parentNode) ' xml������������
    
    Dim node As IXMLDOMNode
    Dim str As Variant
   
        
    Set node = xml.createElement(element)
    node.text = CStr(flag)
    nodes.appendChild node
End Sub

Public Sub saveXml(list As listview)
    Call AddItemToRoot(list, "allowedFolder", "Address")
                        
    Call AddSubItemToRoot(list, "IsNotNewMail", "check", 1)
    Call AddSubItemToRoot(list, "IsNoFlag", "check", 2)
                        
    

    xml.Save fileName
    IsSaved = True
End Sub


Private Sub initElements(parentNode)
    Dim i As Integer
    For i = 0 To xml.ChildNodes(1).ChildNodes.Length - 1
        If parentNode = xml.ChildNodes(1).ChildNodes(i).BaseName Then
            'xml.RemoveChild xml.ChildNodes(i)
          '  Set nodes = xml.ChildNodes(1).ChildNodes(i)
            'xml.RemoveChild xml.ChildNodes(1).ChildNodes(i)
            xml.ChildNodes(1).RemoveChild xml.ChildNodes(1).ChildNodes(i) '�폜����
            Set nodes = xml.ChildNodes(1).appendChild(xml.createElement(parentNode)) '�č쐬����
            Exit For
        End If
    Next i
End Sub



'���X�g�ɗv�f��ǉ�����
Public Sub AddtoList(del_list As listview, add_list As listview, IsDel As Boolean)
Dim list As Variant
    
    Dim i As Integer
    Dim ans As Integer
    Dim Item As Variant
    Dim flag As Boolean
    flag = False
    
    Dim txt As String
    If IsDel Then
        txt = "�폜"
    Else
        txt = "�ǉ�"
    End If
    
    
    If Not IsNull(del_list) Then
    
        For i = 1 To del_list.ListItems.count
            
        
            If del_list.ListItems.Item(i).Selected Then
                flag = True
                ans = MsgBox("""" & del_list.ListItems.Item(i).text & """" & "��" & txt & "���܂����H", vbYesNo)
           
                If ans = vbYes Then
                    Call add_list.ListItems.Add(, , del_list.ListItems.Item(i).text)
                    
                    If Not IsDel Then
                        add_list.ListItems.Item(add_list.ListItems.count).SubItems(1) = False
                        add_list.ListItems.Item(add_list.ListItems.count).SubItems(2) = False
                    End If
                    
                   ' add_list.AddItem del_list.text
                    del_list.ListItems.Remove (i)
                    'del_list.ListItems.Remove (i)
                    
                    IsSaved = False
                End If
           
                Exit For
           
           
            End If
           

        Next i
    End If
    If Not flag Then
        MsgBox "���X�g�̍��ڂ��I������Ă��܂���B"
    End If
    
End Sub
Public Sub ShowMsg()
    If Not IsSaved Then
        MsgBox "�ݒ肪�ۑ�����Ă��܂���"
    End If
    IsSaved = True
End Sub


Public Sub resetXml()
    Set xml = New MSXML2.DOMDocument60
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    fileName = wsh.SpecialFolders("AppData") & "\Microsoft\Outlook\config.xml"
    xml.Load fileName
    
    IsSaved = True
End Sub

