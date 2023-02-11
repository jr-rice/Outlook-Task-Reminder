Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim response As Integer
    response = MsgBox("Do you want to create a task reminder for this email?", vbYesNo + vbQuestion, "Create Task Reminder")
    If response = vbYes Then
        Dim newTask As Outlook.TaskItem
        Set newTask = Outlook.Application.CreateItem(olTaskItem)
        newTask.Subject = "Follow up on email: " & Item.Subject
        Dim inputDate As String
        inputDate = InputBox("Enter the due date for the task (mm/dd/yyyy):")
        If IsDate(inputDate) Then
            newTask.DueDate = inputDate
            newTask.Save
        Else
            MsgBox "Invalid date format. Please enter the date in the format mm/dd/yyyy"
        End If
    End If
End Sub
