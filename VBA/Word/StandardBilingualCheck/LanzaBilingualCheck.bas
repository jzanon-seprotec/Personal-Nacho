Attribute VB_Name = "LanzaBilingualCheck"
Sub VeABilingualCheckForm()
Load BilingualCheckForm
  
BilingualCheckForm.StartUpPosition = 0

BilingualCheckForm.Top = Application.Top + 250
BilingualCheckForm.Left = Application.Left + Application.Width - BilingualCheckForm.Width - 50
BilingualCheckForm.Show vbModeless


End Sub

