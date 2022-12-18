Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
    Label3.Visible = True
Else
    Form1.Show
    Form2.Hide
End If
End Sub
