Attribute VB_Name = "TEST"
' %UI Button btn_nihao 你好
' %UI TextBox txt_where 请输入去哪吃
' %UI chk chk_eatshit  是否吃屎

Sub test()
    Dim oFrm: Set oFrm = New Cls_DynaFrm
    If oFrm.IsCancelled Then
        MsgBox "卧槽,你干嘛取消"
    Exit Sub
    End If
    '按钮的使用
    Select Case oFrm.BtnClicked
        Case "btn_nihao": MsgBox "你好！吃了吗"
        Case Else
            MsgBox "没吃，去哪吃？吃点啥"
            '其他控件
            ctrl = "txt_where"
                If frmDic.Exists(ctrl) Then MsgBox "这儿吃去：" And frmDic(ctrl)
            ctrl = "chk_eatshit"
                If frmDic.Exists(ctrl) Then
                    Select Case frmDic(ctrl)
                        Case True: MsgBox "决定吃屎"
                    End Select
                End If
     End Select
End Sub
