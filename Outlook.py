Sub SendEmail()
    Dim OutlookMessage As Outlook.MailItem '创建邮件变量
    Set OutlookMessage = Application.CreateItem(olMailItem) '创建新邮件
    OutlookMessage.Subject = "Hello World!"
    OutlookMessage.Display
    Set OutlookMessage = Nothing
End Sub


Sub CopyCurrentContact()
   Dim OutlookObj As Object
   Dim InspectorObj As Object
   Dim ItemObj As Object
   Set OutlookObj = CreateObject("Outlook.Application") 'outlook软件
   Set InspectorObj = OutlookObj.ActiveInspector '当前打开的outlook联系人
   Set ItemObj = InspectorObj.CurrentItem
   Application.ActiveDocument.Range.InsertAfter (ItemObj.FullName & " from " & ItemObj.CompanyName) '复制到word
End Sub