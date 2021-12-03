VERSION 5.00
Begin VB.Form func_delay 
   Caption         =   "Delay"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox text_g4 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo de delay (s):"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "func_delay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim G4 As Coords


Private Sub Command1_Click()
    
    
    If text_g4.Text = "" Or Not IsNumeric(text_g4.Text) Then
        MsgBox ("Entre com um valor válido!")
        text_g4 = ""
    Else
    
        txtinsere = "/Delay"
        'txtinsere = Replace(txtinsere, ",", ".")
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        
        G4.X = CDbl(text_g4.Text)
        TempString = "G4 X" & CStr(G4.X)
        ListIndex = frmVisu.lst_codcnc.ListIndex
        frmVisu.lst_codcnc.AddItem TempString
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & frmVisu.NumLinhas * 10
        'frmVisu.lst_codcnc.ListIndex = ListIndex + 1
        'frmVisu.codcnc.Text = frmVisu.codcnc.Text + vbNewLine + TempString
        func_delay.Enabled = False
        func_delay.Visible = False
    End If
    
    
End Sub
