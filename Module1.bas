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
            existAllowFolderArr(count) = folder '存在するフォルダを配列に入れる
            
            IsExist3 = True
            Exit Function
        End If
     
    count = count + 1
    Next folder
    IsExit3 = False
End Function
'文字列がパターンにマッチするかどうか
Function RegExpTest(ByVal str As String, ByVal pattern As String)
    Dim r As RegExp
    Set r = New RegExp
    
    r.Global = True
    r.IgnoreCase = True '大文字小文字を区別しない
    r.pattern = pattern
    
    RegExpTest = r.Test(str)
    

End Function

Function SetColor()
    Dim objColorDialog As Object
    Dim intColor As Long
    
    ' ColorDialog オブジェクトを作成
    Set objColorDialog = Application.filedialog(3)
End Function


'受信したメールのルートを取得
'戻り値として「example@ex.com」 などのメールアドレスが返されます。
Function GetReceivedItemRoot(ByVal Item As Object) As folder
    Dim ItemParent As Object
    Set ItemParent = Item
    Do
        Set ItemParent = ItemParent.Parent
    Loop While Not RegExpTest(ItemParent.Name, "^[A-Za-z\-.\d\+_]+@[A-Za-z\-.\d\+_]+\.[A-Za-z\d]+$")
    
    
    Set GetReceivedItemRoot = ItemParent
End Function


'ソートを行い、ソートの実行可否を戻す
Function canSortMail(mail As Outlook.Items) As Boolean
On Error GoTo err ' エラーが発生したら Catch へ移動する
    
    
    mail.sort "[ReceivedTime]", True
    
    
    canSortMail = True

    Exit Function
err: ' エラーが発生したらここから処理が始まる

    canSortMail = False
    

End Function


Function IsExistPropatyOfReceivedTime(ByVal obj As Variant)
    On Error GoTo err ' エラーが発生したら Catch へ移動する
    
    Dim time As String
    time = obj.ReceivedTime
    
    IsExistPropatyOfReceivedTime = True

    Exit Function
err: ' エラーが発生したらここから処理が始まる

    IsExistPropatyOfReceivedTime = False
End Function




Sub changeListTab(multiPage As multiPage, listPage As Long)
    'リスト追加後に位置が変わるのを修正
    Dim index As Integer
        multiPage.value = listPage
    For index = 0 To multiPage.count - 1
        multiPage.value = index
    Next index

    multiPage.value = listPage
End Sub
