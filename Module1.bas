Attribute VB_Name = "Module1"

Public form As New UserForm1



Function IsExist(ByVal str As String, ByVal arr As Variant) As Boolean
    For Each folder In arr
        If str = folder Then
            IsExist = True
            Exit Function
        End If
     
    
    Next folder
    IsExit = False
End Function


Function IsExist3(ByVal str As String, ByVal str2 As String, ByVal arr As Variant, _
                         existAllowFolderArr As Variant) As Boolean
    Dim count As Variant
    count = 0
    For Each folder In arr
        
        If str & str2 = folder Then
            Dim countArr As Integer
            existAllowFolderArr(count) = folder '���݂���t�H���_��z��ɓ����
            
            IsExist3 = True
            Exit Function
        End If
     
    count = count + 1
    Next folder
    IsExit3 = False
End Function
'�����񂪃p�^�[���Ƀ}�b�`���邩�ǂ���
Function RegExpTest(ByVal str As String, ByVal pattern As String)
    Dim r As RegExp
    Set r = New RegExp
    
    r.Global = True
    r.IgnoreCase = True '�啶������������ʂ��Ȃ�
    r.pattern = pattern
    
    RegExpTest = r.Test(str)
    

End Function

Function SetColor()
    Dim objColorDialog As Object
    Dim intColor As Long
    
    ' ColorDialog �I�u�W�F�N�g���쐬
    Set objColorDialog = Application.filedialog(3)
End Function


'��M�������[���̃��[�g���擾
'�߂�l�Ƃ��āuexample@ex.com�v �Ȃǂ̃��[���A�h���X���Ԃ���܂��B
Function GetReceivedItemRoot(ByVal Item As Object) As folder
    Dim ItemParent As Object
    Set ItemParent = Item
    Do
        Set ItemParent = ItemParent.Parent
    Loop While Not RegExpTest(ItemParent.Name, "^[A-Za-z\-.\d\+_]+@[A-Za-z\-.\d\+_]+\.[A-Za-z\d]+$")
    
    
    Set GetReceivedItemRoot = ItemParent
End Function


'�\�[�g���s���A�\�[�g�̎��s�ۂ�߂�
Function canSortMail(mail As Outlook.Items) As Boolean
On Error GoTo err ' �G���[������������ Catch �ֈړ�����
    
    
    mail.sort "[ReceivedTime]", True
    
    
    canSortMail = True

    Exit Function
err: ' �G���[�����������炱�����珈�����n�܂�

    canSortMail = False
    

End Function


Function IsExistPropatyOfReceivedTime(ByVal obj As Variant)
    On Error GoTo err ' �G���[������������ Catch �ֈړ�����
    
    Dim time As String
    time = obj.ReceivedTime
    
    IsExistPropatyOfReceivedTime = True

    Exit Function
err: ' �G���[�����������炱�����珈�����n�܂�

    IsExistPropatyOfReceivedTime = False
End Function




Sub changeListTab(multiPage As multiPage, listPage As Long)
    '���X�g�ǉ���Ɉʒu���ς��̂��C��
    Dim index As Integer
        multiPage.value = listPage
    For index = 0 To multiPage.count - 1
        multiPage.value = index
    Next index

    multiPage.value = listPage
End Sub
