Private Const pi = 3.14159265359
Private Sub Command1_Click()
 RotatePicture Picture1, Picture2, 45
End Sub

Public Sub RotatePicture(ByVal picSource As PictureBox, ByVal picDest As PictureBox, Optional ByVal Angle As Long = 0)
 
  ' cp0 - cp3 = Color of single Pixel
  Dim cp0 As Long, cp1 As Long
  Dim cp2 As Long, cp3 As Long
 
  ' Picture-Dimensions
  Dim w1 As Long, h1 As Long
  Dim w2 As Long, h2 As Long
 
  Dim p1x As Double, p1y As Double
  Dim p2x As Double, p2y As Double
  Dim n As Double, r As Double, a As Double
 
  picSource.AutoSize = True
  picSource.Visible = False
 
 
  picSource.AutoRedraw = True
  picDest.AutoRedraw = True
 
  picSource.ScaleMode = vbPixels
  picDest.ScaleMode = vbPixels
 
 
  Set picDest.Picture = Nothing
  picDest.Cls
 
 …
[11:23 am, 18/12/2022] Pawan Kumar: Form1.frm

Dim i As Integer

Private Sub Command1_Click()
Command1.Enabled = False
If i Mod 2 = 0 Then
    Command1.Caption = "X"
    Label1.Caption = Form2.Text1.Text + "'s turn..."
Else
    Command1.Caption = "O"
    Label1.Caption = Form2.Text2.Text + "'s turn..."
End If
If (Command1.Caption = Command2.Caption And Command2.Caption = Command3.Caption) Or (Command4.Caption = Command5.Caption And Command5.Caption = Command6.Caption) Or (Command7.Caption = Command8.Caption And Command8.Caption = Command9.Caption) Or (Command1.Caption = Command4.Caption And Command4.Caption = Command7.Caption) Or (Command2.Caption = Command5.Caption And Command5.Caption = Command8.Caption) Or (Command3.Caption = Command6.Caption And Command6.Caption = Command9.Caption) Or (Com…
[11:23 am, 18/12/2022] Pawan Kumar: Label1.Caption = "Game Ties..."
        Command10.Visible = True
        Command11.Visible = True
    End If
End If
i = i + 1

End Sub

Private Sub Command5_Click()
Command5.Enabled = False
If i Mod 2 = 0 Then
    Command5.Caption = "X"
    Label1.Caption = Form2.Text1.Text + "'s turn..."
Else
    Command5.Caption = "O"
    Label1.Caption = Form2.Text2.Text + "'s turn..."
End If
If (Command5.Caption = Command2.Caption And Command2.Caption = Command3.Caption) Or (Command4.Caption = Command5.Caption And Command5.Caption = Command6.Caption) Or (Command7.Caption = Command8.Caption And Command8.Caption = Command9.Caption) Or (Command1.Caption = Command4.Caption And Command4.Caption = Command7.Caption) Or (Command2.Caption = Command5.Caption And Command5.Caption = Command8.Caption) Or (Command3.Caption = Command6.Caption And Command6.Caption = Command9.Caption) Or (Command1.Caption = Command5.Caption And Command5.Caption = Command9.Caption) Or (Command3.Caption = Command5.Caption And Command5.Caption = Command7.Caption) Then
    If Command5.Caption = "O" Then
       Label1.Caption = Form2.Text1.Text + "'ve won..."
    Else
        Label1.Caption = Form2.Text2.Text + "'ve won..."
    End If
Else
    If i = 9 Then
        Label1.Caption = "Game Ties..."
        Command10.Visible = True
        Command11.Visible = True
    End If
End If
i = i + 1
End Sub

Private Sub Command6_Click()
Command6.Enabled = False
If i Mod 2 = 0 Then
    Command6.Caption = "X"
    Label1.Caption = Form2.Text1.Text + "'s turn..."
Else
    Command6.Caption = "O"
    Label1.Caption = Form2.Text2.Text + "'s turn..."
End If
If (Command6.Caption = Command2.Caption And Command2.Caption = Command3.Caption) Or (Command4.Caption = Command5.Caption And Command5.Caption = Command6.Caption) Or (Command7.Caption = Command8.Caption And Command8.Caption = Command9.Caption) Or (Command1.Caption = Command4.Caption And Command4.Caption = Command7.Caption) Or (Command2.Caption = Command5.Caption And Command5.Caption = Command8.Caption) Or (Command3.Caption = Command6.Caption And Command6.Caption = Command9.Caption) Or (Command1.Caption = Command5.Caption And Command5.Caption = Command9.Caption) Or (Command3.Caption = Command5.Caption And Command5.Caption = Command7.Caption) Then
    If Command6.Caption = "O" Then
       Label1.Caption = Form2.Text1.Text + "'ve won..."
    Else
        Label1.Caption = Form2.Text2.Text + "'ve won..."
    End If
Else
    If i = 9 Then
        Label1.Caption = "Game Ties..."
        Command10.Visible = True
        Command11.Visible = True
    End If
End If
i = i + 1
End Sub

Private Sub Command7_Click()
Command7.Enabled = False
If i Mod 2 = 0 Then
    Command7.Caption = "X"
    Label1.Caption = Form2.Text1.Text + "'s turn..."
Else
    Command7.Caption = "O"
    Label1.Caption = Form2.Text2.Text + "'s turn..."
End If
If (Command7.Caption = Command2.Caption And Command2.Caption = Command3.Caption) Or (Command4.Caption = Command5.Caption And Command5.Caption = Command6.Caption) Or (Command7.Caption = Command8.Caption And Command8.Caption = Command9.Caption) Or (Command1.Caption = Command4.Caption And Command4.Caption = Command7.Caption) Or (Command2.Caption = Command5.Caption And Command5.Caption = Command8.Caption) Or (Command3.Caption = Command6.Caption And Command6.Caption = Command9.Caption) Or (Command1.Caption = Command5.Caption And Command5.Caption = Command9.Caption) Or (Command3.Caption = Command5.Caption And Command5.Caption = Command7.Caption) Then
    If Command7.Caption = "O" Then
       Label1.Caption = Form2.Text1.Text + "'ve won..."
    Else
        Label1.Caption = Form2.Text2.Text + "'ve won..."
    End If
Else
    If i = 9 Then
        Label1.Caption = "Game Ties..."
        Command10.Visible = True
        Command11.Visible = True
    End If
End If
i = i + 1
End Sub

Private Sub Command8_Click()
Command8.Enabled = False
If i Mod 2 = 0 Then
    Command8.Caption = "X"
    Label1.Caption = Form2.Text1.Text + "'s turn..."
Else
    Command8.Caption = "O"
    Label1.Caption = Form2.Text2.Text + "'s turn..."
End If
If (Command8.Caption = Command2.Caption And Command2.Caption = Command3.Caption) Or (Command4.Caption = Command5.Caption And Command5.Caption = Command6.Caption) Or (Command7.Caption = Command8.Caption And Command8.Caption = Command9.Caption) Or (Command1.Caption = Command4.Caption And Command4.Caption = Command7.Caption) Or (Command2.Caption = Command5.Caption And Command5.Caption = Command8.Caption) Or (Command3.Caption = Command6.Caption And Command6.Caption = Command9.Caption) Or (Command1.Caption = Command5.Caption And Command5.Caption = Command9.Caption) Or (Command3.Caption = Command5.Caption And Command5.Caption = Command7.Caption) Then
    If Command8.Caption = "O" Then
       Label1.Caption = Form2.Text1.Text + "'ve won..."
    Else
        Label1.Caption = Form2.Text2.Text + "'ve won..."
    End If
Else
    If i = 9 Then
        Label1.Caption = "Game Ties..."
        Command10.Visible = True
        Command11.Visible = True
    End If
End If
i = i + 1
End Sub

Private Sub Command9_Click()
Command9.Enabled = False
If i Mod 2 = 0 Then
    Command9.Caption = "X"
    Label1.Caption = Form2.Text1.Text + "'s turn..."
Else
    Command9.Caption = "O"
    Label1.Caption = Form2.Text2.Text + "'s turn..."
End If
If (Command9.Caption = Command2.Caption And Command2.Caption = Command3.Caption) Or (Command4.Caption = Command5.Caption And Command5.Caption = Command6.Caption) Or (Command7.Caption = Command8.Caption And Command8.Caption = Command9.Caption) Or (Command1.Caption = Command4.Caption And Command4.Caption = Command7.Caption) Or (Command2.Caption = Command5.Caption And Command5.Caption = Command8.Caption) Or (Command3.Caption = Command6.Caption And Command6.Caption = Command9.Caption) Or (Command1.Caption = Command5.Caption And Command5.Caption = Command9.Caption) Or (Command3.Caption = Command5.Caption And Command5.Caption = Command7.Caption) Then
    If Command9.Caption = "O" Then
       Label1.Caption = Form2.Text1.Text + "'ve won..."
    Else
        Label1.Caption = Form2.Text2.Text + "'ve won..."
    End If
Else
    If i = 9 Then
        Label1.Caption = "Game Ties..."
        Command10.Visible = True
        Command11.Visible = True
    End If
End If
i = i + 1
End Sub

Private Sub Form_Load()
i = 1
Label1.Caption = Form2.Text1.Text + "'s turn..."
End Sub
