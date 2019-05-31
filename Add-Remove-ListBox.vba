Private Sub btnCadastro_Click()

     ListBox2.AddItem ListBox1.List(ListBox1.ListIndex)
   
End Sub

Private Sub btnRemove_Click()
    ListBox2.RemoveItem (ListBox2.ListIndex)
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Activate()
    ListBox1.AddItem Worksheets("Plan1").Range("A2")
    ListBox1.AddItem Worksheets("Plan1").Range("A3")
    ListBox1.AddItem Worksheets("Plan1").Range("A4")

    
End Sub
