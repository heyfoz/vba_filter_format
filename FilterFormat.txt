Private Sub FilterNFormat()
'
' FilterNFormat Macro Sample
' By Forrest Moulin
'
' Automatically formats and filters Excel file by calling multiple
' custom VBA macros.

  ' Macro file name object creation
  Dim fileName As String
  fileName = VBAProject.ThisWorkBook.Name
  
  'Show custom user form with text input field
  AliasForm.Show
  
  ' Visibility of system status
  MsgBox("Macro running... Workbook with update until the macro is complete!")
  
  ' Add footer with signature/date prompts
  With ActiveSheet.PageSetup
    .FirstPage.LeftFooter.Text = "&""Segoe UI,Bold""Associate Signature"
                                 &Chr(10)&""&Chr(10)&"Date"
    .FirstPage.CenterFooter.Text = "&""Segoe UI,Bold""Manager Signature"
                                   &Chr(10)&""&Chr(10)&"Date"
  ActiveSheet.PageSetup.LeftHeaderPicture.fileName =_
    "C:\Users\UserName\OneDrive\HeaderLogo.jpg"
    
    ' Call several macros from the .xlsm file
    Call Macro2
    Call Macro3
    
    ' Display update to user and close macro file
    MsgBox("Macro complete :)" + fileName)
    Workbooks(fileName).Close
    
End Sub
    
Private Sub SubmitButton_Click()
  
  Dim textBoxString As String
  textBoxString = AliasForm.InputTextBox.Text
  AliasForm.Hide
  Range("C1").Select
  ActiveCell.FormulaR1C1 = textBoxString
 
End Sub
