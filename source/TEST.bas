Attribute VB_Name = "TEST"
' %UI Button btn_create ´´½¨Í¼¿ò
' %UI Button btn_delete É¾³ýÍ¼¿ò
' %UI Button btn_resize ¸ü¸ÄÍ¼¿ò³ß´ç
' %UI Button btn_update ¸üÐÂÍ¼¿ò

Sub tets()



    Dim oFrm: Set oFrm = New Cls_DynaFrm
    
    
    
    If oFrm.IsCancelled Then
        MsgBox "ÎÔ²Û"
        
    End If
End Sub
