VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private form As UserForm1



Private Sub Application_Startup()
    Set form = New UserForm1

    form.Show vbModeless
   
End Sub



Public Sub run2()
On Error GoTo err ' �G���[������������ Catch �ֈړ�����
    Set form = New UserForm1
    form.Show vbModeless
    
     Exit Sub
err: ' �G���[�����������炱�����珈�����n�܂�
    
    MsgBox "�v���I�ȃG���[���������܂����B", vbCritical, "�G���["
    MsgBox "�}�N�����I�����܂��B", vbCritical
End Sub
