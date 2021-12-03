VERSION 5.00
Begin VB.Form ins_cabe 
   Caption         =   "Inserir Cabeçalho"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4005
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4005
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inserir"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Cabeçalho"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "ins_cabe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        ListIndex = frmVisu.lst_codcnc.ListIndex
        txtinsere = ins_cabe.Text1.Text
        frmVisu.lst_codcnc.AddItem txtinsere, ListIndex + 1
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & frmVisu.NumLinhas * 10, frmVisu.NumLinhas
        frmVisu.lst_codcnc.ListIndex = ListIndex + 1
        txtinsere = ""
        Text1.Text = ""
    
End Sub

Private Sub Command2_Click()
    ins_cabe.Visible = False
    ins_cabe.Enabled = False
    Text1.Text = ""
End Sub
