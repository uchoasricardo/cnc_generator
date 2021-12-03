VERSION 5.00
Begin VB.Form ent_param 
   Caption         =   "Parâmetros Básicos"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4185
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   735
      Left            =   1320
      TabIndex        =   17
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3000
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Nome do programa:"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Número do programa:"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Avanço de acabamento (mm/rot):"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Ferramenta:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Diâmetro bruto (mm):"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Comprimento bruto (mm):"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Avanço de desbaste (mm/rot):"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Velocidade de corte (m/min):"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label ent_param 
      Caption         =   "Parâmetros do programa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "ent_param"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    If Not IsNumeric(Text1.Text) Then
            MsgBox "Valor inválido para Diâmetro Bruto!"
            Text1.Text = ""
    ElseIf Not IsNumeric(Text2.Text) Then
            MsgBox "Valor inválido para Comprimento Bruto!"
            Text2.Text = ""
    ElseIf Not IsNumeric(Text3.Text) Then
            MsgBox "Valor inválido para Avanço!"
            Text3.Text = ""
    ElseIf Not IsNumeric(Text4.Text) Then
            MsgBox "Valor inválido para Avanço!"
            Text4.Text = ""
    ElseIf Not IsNumeric(Text5.Text) Then
            MsgBox "Valor inválido para Velocidade de Corte!"
            Text5.Text = ""
    ElseIf Not IsNumeric(Text6.Text) Or Text6.Text < 1 Or Text6.Text > 12 Then
            MsgBox "Valor inválido para Ferramenta!"
            Text6.Text = ""
    ElseIf Not IsNumeric(Text7.Text) Or CInt(Round(Text7.Text, 0)) < 1 Or CInt(Round(Text7.Text, 0)) > 7999 Then
        MsgBox "Valor inválido para Número do Programa. Insira um número entre 1 e 7999."
        Text7.Text = ""
    Else
        Text1.Text = Replace(Text1.Text, ".", ",")
        Text2.Text = Replace(Text2.Text, ".", ",")
        Text3.Text = Replace(Text3.Text, ".", ",")
        Text4.Text = Replace(Text4.Text, ".", ",")
        Text5.Text = Replace(Text5.Text, ".", ",")
        Text6.Text = Replace(Text6.Text, ".", ",")
        Text7.Text = Replace(Text7.Text, ".", ",")
        
        frmVisu.DiametroBruto = CDbl(Text1.Text)
        frmVisu.ComprimentoBruto = CDbl(Text2.Text)
        frmVisu.VelAvanco = CDbl(Text3.Text)
        frmVisu.VelAvancoAcab = CDbl(Text4.Text)
        frmVisu.VelocidadeCorte = CDbl(Text5.Text)
        frmVisu.ParametrosDefinidos = True
        frmVisu.Ferramenta = CInt(Round(Text6.Text, 0))
        frmVisu.NumeroPrograma = CInt(Round(Text7.Text, 0))
        frmVisu.NomePrograma = CStr(Text8.Text)
        
        frmVisu.btn_cabecalho.Enabled = True
        frmVisu.btn_zerar.Enabled = True
        
        Unload Me
    End If
    
End Sub

