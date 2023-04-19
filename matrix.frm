VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} macierzy 
   Caption         =   "Kalkulator macierzy"
   ClientHeight    =   6552
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9816.001
   OleObjectBlob   =   "macierzy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "macierzy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then
        CheckBox2.Enabled = False
        TextBox31.Visible = False
        TextBox35.Visible = False
        TextBox36.Visible = False
        TextBox37.Visible = False
        TextBox34.Visible = False
        TextBox29.Visible = False
        TextBox25.Visible = False
        
        TextBox14.Visible = False
        TextBox18.Visible = False
        TextBox19.Visible = False
        TextBox20.Visible = False
        TextBox17.Visible = False
        TextBox12.Visible = False
        TextBox8.Visible = False
        
        TextBox48.Visible = False
        TextBox52.Visible = False
        TextBox53.Visible = False
        TextBox54.Visible = False
        TextBox51.Visible = False
        TextBox46.Visible = False
        TextBox42.Visible = False
        
    ElseIf CheckBox1.Value = False Then
        CheckBox2.Enabled = True
    End If
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2.Value = True And CheckBox3.Value = False Then
        CheckBox1.Enabled = False
        TextBox31.Visible = True
        TextBox35.Visible = True
        TextBox36.Visible = True
        TextBox37.Visible = True
        TextBox34.Visible = True
        TextBox29.Visible = True
        TextBox25.Visible = True
        
        TextBox14.Visible = True
        TextBox18.Visible = True
        TextBox19.Visible = True
        TextBox20.Visible = True
        TextBox17.Visible = True
        TextBox12.Visible = True
        TextBox8.Visible = True
        
        TextBox48.Visible = True
        TextBox52.Visible = True
        TextBox53.Visible = True
        TextBox54.Visible = True
        TextBox51.Visible = True
        TextBox46.Visible = True
        TextBox42.Visible = True
        
    ElseIf CheckBox2.Value = False Then
        CheckBox1.Enabled = True
        
    ElseIf CheckBox2.Value = True And CheckBox3.Value = True Then
        CheckBox1.Enabled = False
        TextBox31.Visible = False
        TextBox35.Visible = False
        TextBox36.Visible = False
        TextBox37.Visible = False
        TextBox34.Visible = False
        TextBox29.Visible = False
        TextBox25.Visible = False
        
        TextBox14.Visible = True
        TextBox18.Visible = True
        TextBox19.Visible = True
        TextBox20.Visible = True
        TextBox17.Visible = True
        TextBox12.Visible = True
        TextBox8.Visible = True
        
        TextBox48.Visible = True
        TextBox52.Visible = True
        TextBox53.Visible = True
        TextBox54.Visible = True
        TextBox51.Visible = True
        TextBox46.Visible = True
        TextBox42.Visible = True
    
    End If
End Sub

Private Sub CheckBox3_Click()
    If CheckBox3.Value = True And (CheckBox1.Value = True Or CheckBox2.Value = True) Then
        CheckBox4.Enabled = False
        CheckBox5.Enabled = False
        TextBox21.Visible = False
        TextBox22.Visible = False
        TextBox23.Visible = False
        TextBox24.Visible = False
        TextBox25.Visible = False
        TextBox26.Visible = False
        TextBox27.Visible = False
        TextBox28.Visible = False
        TextBox29.Visible = False
        TextBox30.Visible = False
        TextBox31.Visible = False
        TextBox32.Visible = False
        TextBox33.Visible = False
        TextBox34.Visible = False
        TextBox35.Visible = False
        TextBox36.Visible = False
        TextBox37.Visible = False
        
    ElseIf CheckBox3.Value = False And CheckBox2.Value = True Then
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        TextBox21.Visible = True
        TextBox22.Visible = True
        TextBox23.Visible = True
        TextBox24.Visible = True
        TextBox25.Visible = True
        TextBox26.Visible = True
        TextBox27.Visible = True
        TextBox28.Visible = True
        TextBox29.Visible = True
        TextBox30.Visible = True
        TextBox31.Visible = True
        TextBox32.Visible = True
        TextBox33.Visible = True
        TextBox34.Visible = True
        TextBox35.Visible = True
        TextBox36.Visible = True
        TextBox37.Visible = True
        
    ElseIf CheckBox3.Value = False And CheckBox1.Value = True Then
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        TextBox21.Visible = True
        TextBox22.Visible = True
        TextBox23.Visible = True
        TextBox24.Visible = True
        TextBox25.Visible = False
        TextBox26.Visible = True
        TextBox27.Visible = True
        TextBox28.Visible = True
        TextBox29.Visible = False
        TextBox30.Visible = True
        TextBox31.Visible = False
        TextBox32.Visible = True
        TextBox33.Visible = True
        TextBox34.Visible = False
        TextBox35.Visible = False
        TextBox36.Visible = False
        TextBox37.Visible = False
        
    ElseIf CheckBox3.Value = False Then
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        TextBox21.Visible = True
        TextBox22.Visible = True
        TextBox23.Visible = True
        TextBox24.Visible = True
        TextBox25.Visible = True
        TextBox26.Visible = True
        TextBox27.Visible = True
        TextBox28.Visible = True
        TextBox29.Visible = True
        TextBox30.Visible = True
        TextBox31.Visible = True
        TextBox32.Visible = True
        TextBox33.Visible = True
        TextBox34.Visible = True
        TextBox35.Visible = True
        TextBox36.Visible = True
        TextBox37.Visible = True
    Else
        CheckBox4.Enabled = False
        CheckBox5.Enabled = False
        TextBox21.Visible = False
        TextBox22.Visible = False
        TextBox23.Visible = False
        TextBox24.Visible = False
        TextBox25.Visible = False
        TextBox26.Visible = False
        TextBox27.Visible = False
        TextBox28.Visible = False
        TextBox29.Visible = False
        TextBox30.Visible = False
        TextBox31.Visible = False
        TextBox32.Visible = False
        TextBox33.Visible = False
        TextBox34.Visible = False
        TextBox35.Visible = False
        TextBox36.Visible = False
        TextBox37.Visible = False
        
    End If
End Sub

Private Sub CheckBox4_Click()
    If CheckBox4.Value = True Then
        CheckBox3.Enabled = False
        CheckBox5.Enabled = False
    ElseIf CheckBox4.Value = False Then
        CheckBox3.Enabled = True
        CheckBox5.Enabled = True
    End If
End Sub

Private Sub CheckBox5_Click()
    If CheckBox5.Value = True Then
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
    ElseIf CheckBox5.Value = False Then
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
    End If
End Sub

Private Sub CommandButton1_Click()
    If CheckBox1.Value = True And CheckBox3.Value = True Then
        TextBox48.Visible = False
        TextBox52.Visible = False
        TextBox53.Visible = False
        TextBox54.Visible = False
        TextBox51.Visible = False
        TextBox46.Visible = False
        TextBox42.Visible = False
        
        TextBox39.Text = TextBox5.Text
        TextBox43.Text = TextBox6.Text
        TextBox47.Text = TextBox7.Text
        TextBox40.Text = TextBox9.Text
        TextBox44.Text = TextBox10.Text
        TextBox49.Text = TextBox11.Text
        TextBox41.Text = TextBox13.Text
        TextBox45.Text = TextBox15.Text
        TextBox50.Text = TextBox16.Text
        
    ElseIf CheckBox2.Value = True And CheckBox3.Value = True Then
        TextBox48.Visible = True
        TextBox52.Visible = True
        TextBox53.Visible = True
        TextBox54.Visible = True
        TextBox51.Visible = True
        TextBox46.Visible = True
        TextBox42.Visible = True
        
        TextBox39.Text = TextBox5.Text
        TextBox43.Text = TextBox6.Text
        TextBox47.Text = TextBox7.Text
        TextBox40.Text = TextBox9.Text
        TextBox44.Text = TextBox10.Text
        TextBox49.Text = TextBox11.Text
        TextBox41.Text = TextBox13.Text
        TextBox45.Text = TextBox15.Text
        TextBox50.Text = TextBox16.Text
        TextBox42.Text = TextBox14.Text
        TextBox46.Text = TextBox18.Text
        TextBox51.Text = TextBox19.Text
        TextBox54.Text = TextBox20.Text
        TextBox48.Text = TextBox8.Text
        TextBox52.Text = TextBox12.Text
        TextBox53.Text = TextBox17.Text
        
    ElseIf CheckBox1.Value = True And CheckBox5.Value = True Then
        TextBox39.Text = TextBox5.Text - TextBox22.Text
        TextBox43.Text = TextBox9.Text - TextBox26.Text
        TextBox47.Text = TextBox13.Text - TextBox30.Text
        TextBox40.Text = TextBox6.Text - TextBox23.Text
        TextBox44.Text = TextBox10.Text - TextBox27.Text
        TextBox49.Text = TextBox15.Text - TextBox32.Text
        TextBox41.Text = TextBox7.Text - TextBox24.Text
        TextBox45.Text = TextBox11.Text - TextBox28.Text
        TextBox50.Text = TextBox16.Text - TextBox33.Text
        
    ElseIf CheckBox2.Value = True And CheckBox5.Value = True Then
        TextBox39.Text = TextBox5.Text - TextBox22.Text
        TextBox43.Text = TextBox9.Text - TextBox26.Text
        TextBox47.Text = TextBox13.Text - TextBox30.Text
        TextBox40.Text = TextBox6.Text - TextBox23.Text
        TextBox44.Text = TextBox10.Text - TextBox27.Text
        TextBox49.Text = TextBox15.Text - TextBox32.Text
        TextBox41.Text = TextBox7.Text - TextBox24.Text
        TextBox45.Text = TextBox11.Text - TextBox28.Text
        TextBox50.Text = TextBox16.Text - TextBox33.Text
        TextBox48.Text = TextBox14.Text - TextBox31.Text
        TextBox52.Text = TextBox18.Text - TextBox35.Text
        TextBox53.Text = TextBox19.Text - TextBox36.Text
        TextBox54.Text = TextBox20.Text - TextBox37.Text
        TextBox42.Text = TextBox8.Text - TextBox25.Text
        TextBox46.Text = TextBox12.Text - TextBox29.Text
        TextBox51.Text = TextBox17.Text - TextBox34.Text
    
    ElseIf CheckBox1.Value = True And CheckBox4.Value = True Then
        TextBox39.Text = CDec(TextBox5.Text) + CDec(TextBox22.Text)
        TextBox43.Text = CDec(TextBox9.Text) + CDec(TextBox26.Text)
        TextBox47.Text = CDec(TextBox13.Text) + CDec(TextBox30.Text)
        TextBox40.Text = CDec(TextBox6.Text) + CDec(TextBox23.Text)
        TextBox44.Text = CDec(TextBox10.Text) + CDec(TextBox27.Text)
        TextBox49.Text = CDec(TextBox15.Text) + CDec(TextBox32.Text)
        TextBox41.Text = CDec(TextBox7.Text) + CDec(TextBox24.Text)
        TextBox45.Text = CDec(TextBox11.Text) + CDec(TextBox28.Text)
        TextBox50.Text = CDec(TextBox16.Text) + CDec(TextBox33.Text)
        
    ElseIf CheckBox2.Value = True And CheckBox4.Value = True Then
        TextBox39.Text = CDec(TextBox5.Text) + CDec(TextBox22.Text)
        TextBox43.Text = CDec(TextBox9.Text) + CDec(TextBox26.Text)
        TextBox47.Text = CDec(TextBox13.Text) + CDec(TextBox30.Text)
        TextBox40.Text = CDec(TextBox6.Text) + CDec(TextBox23.Text)
        TextBox44.Text = CDec(TextBox10.Text) + CDec(TextBox27.Text)
        TextBox49.Text = CDec(TextBox15.Text) + CDec(TextBox32.Text)
        TextBox41.Text = CDec(TextBox7.Text) + CDec(TextBox24.Text)
        TextBox45.Text = CDec(TextBox11.Text) + CDec(TextBox28.Text)
        TextBox50.Text = CDec(TextBox16.Text) + CDec(TextBox33.Text)
        TextBox48.Text = CDec(TextBox14.Text) + CDec(TextBox31.Text)
        TextBox52.Text = CDec(TextBox18.Text) + CDec(TextBox35.Text)
        TextBox53.Text = CDec(TextBox19.Text) + CDec(TextBox36.Text)
        TextBox54.Text = CDec(TextBox20.Text) + CDec(TextBox37.Text)
        TextBox42.Text = CDec(TextBox8.Text) + CDec(TextBox25.Text)
        TextBox46.Text = CDec(TextBox12.Text) + CDec(TextBox29.Text)
        TextBox51.Text = CDec(TextBox17.Text) + CDec(TextBox34.Text)
    End If
End Sub

Private Sub CommandButton2_Click()
    Dim cControl As Control
    For Each cControl In Me.Controls
        If cControl.Name Like "Text*" Then cControl = vbNullString
    Next
End Sub

Private Sub UserForm_Click()

End Sub
