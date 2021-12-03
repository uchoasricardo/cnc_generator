VERSION 5.00
Begin VB.Form func_face 
   Caption         =   "Faceamento"
   ClientHeight    =   2820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Ferramenta:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Profundidade do passe (mm):"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "func_face"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodeFace As String
Dim txtzero As String
Dim tipo As String
Public pZ As Double
Public sZ As Double
Public pX As Double
Public sX As Double
Dim NumeroPasses As Double
Dim UltimoPasse As Double

Dim poschar As Integer

Private Sub Command1_Click()
 
    If Not IsNumeric(Text1.Text) Then
            MsgBox "Valor inválido para Profundidade de Faceamento!"
            Text1.Text = ""
    ElseIf CDbl(Text1.Text) > 5 Then
            MsgBox "Não é recomendado utilizar profundidade de corte superior a 5mm."
            Text1.Text = ""
    Else
            If frmVisu.FaceamentoCompFinal Then
                    'faceamento de comprimento final
                    frmVisu.ProfundidadeFaceamento = CDbl(Text1.Text)
                    NumeroPasses = ((frmVisu.ComprimentoBruto - frmVisu.ComprimentoFinal) \ frmVisu.ProfundidadeFaceamento)
                    UltimoPasse = (frmVisu.ComprimentoBruto - frmVisu.ComprimentoFinal) Mod frmVisu.ProfundidadeFaceamento
                    If Not (NumeroPasses > 0) Then
                        MsgBox "Esta entidade não tem a distância apropriada para o comprimento bruto informado! Defina um comprimento bruto apropriado!"
                        ent_param.Enabled = True
                        ent_param.Visible = True
                        Unload func_face
                    End If
                    
                    CodeFace = "/Ciclo de Faceamento - Diam./Comp. Bruto"
                    frmVisu.lst_codcnc.AddItem CodeFace
                    
                    CodeFace = "G00 X" & (frmVisu.DiametroBruto + 5) & " Z" & (frmVisu.ComprimentoBruto - frmVisu.ComprimentoFinal + 5)
                    CodeFace = Replace(CodeFace, ",", ".")
                    frmVisu.lst_codcnc.AddItem CodeFace
                    
                    CodeFace = "G01 X" & (frmVisu.DiametroBruto + 2) & " Z" & Round((frmVisu.ComprimentoBruto - frmVisu.ComprimentoFinal - frmVisu.ProfundidadeFaceamento), 4)
                    CodeFace = Replace(CodeFace, ",", ".")
                    frmVisu.lst_codcnc.AddItem CodeFace
                    
                    CodeFace = "G75 X-1 Z" & Round((func_face.pZ - frmVisu.ZOffset) + 1, 4) & " P20000 Q" & frmVisu.ProfundidadeFaceamento * 1000 & " R" & Round((frmVisu.ProfundidadeFaceamento), 0) & " F" & frmVisu.VelocidadeCorte
                    CodeFace = Replace(CodeFace, ",", ".")
                    frmVisu.lst_codcnc.AddItem CodeFace
            Else
                    'faceamento de comprimento comum
                    CodeFace = "/Ciclo de Faceamento"
                    frmVisu.ProfundidadeFaceamento = CDbl(Text1.Text)
                    frmVisu.lst_codcnc.AddItem CodeFace
                    
                    CodeFace = "G00 X" & ((func_face.sX - frmVisu.XOffset) * 2 + 5) & " Z" & (func_face.pZ - frmVisu.ZOffset + 5) & " T" & frmVisu.Ferramenta
                    CodeFace = Replace(CodeFace, ",", ".")
                    frmVisu.lst_codcnc.AddItem CodeFace
                    
                    CodeFace = "G01 X" & Round(((func_face.sX - frmVisu.XOffset) * 2 + 2), 4) & " Z" & Round((func_face.pZ - frmVisu.ZOffset - frmVisu.ProfundidadeFaceamento), 4)
                    CodeFace = Replace(CodeFace, ",", ".")
                    frmVisu.lst_codcnc.AddItem CodeFace
                    
                    CodeFace = "G75 X" & Round(((func_face.pX - frmVisu.XOffset) * 2), 4) & " Z" & Round((func_face.pZ - frmVisu.ZOffset), 4) & " P20000 Q" & frmVisu.ProfundidadeFaceamento * 1000 & " R" & Round((1000 * frmVisu.ProfundidadeFaceamento), 0) & " F" & frmVisu.VelocidadeCorte
                    CodeFace = Replace(CodeFace, ",", ".")
                    frmVisu.lst_codcnc.AddItem CodeFace
            End If

            
            Unload Me
    End If
End Sub

Private Sub Form_Load()

    
    
            Text2.Text = CStr(frmVisu.Ferramenta)
            txtzero = frmVisu.lst_entidades.Text
            poschar = InStr(1, txtzero, ";", vbTextCompare)
            tipo = Mid(txtzero, 1, (poschar - 1))
            txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
    
            If tipo = "LINE" Then
            poschar = InStr(1, txtzero, ";", vbTextCompare)
            pZ = CDbl(Mid(txtzero, 1, (poschar - 1)))
            txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
            poschar = InStr(1, txtzero, ";", vbTextCompare)
            pX = CDbl(Mid(txtzero, 1, (poschar - 1)))
            txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
            poschar = InStr(1, txtzero, ";", vbTextCompare)
            sZ = CDbl(Mid(txtzero, 1, (poschar - 1)))
            txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
            sX = CDbl(Mid(txtzero, 1, (poschar - 1)))
                           
                
        Else
            MsgBox ("Entidade inválida!")
        End If
        
        If Not pZ = sZ Then
            MsgBox "O Faceamento deve ser feito em uma superfície perpendicular ao eixo da peça."
            Exit Sub
            'Unload Me
        ElseIf Not pZ = frmVisu.ZOffset And frmVisu.FaceamentoCompFinal Then
            MsgBox "A linha de faceamento do comprimento final deve estar posicionada em Z = 0"
            Exit Sub
            Unload Me
        Else
            NumeroPasses = (ComprimentoBruto - pZ)
            CodeFace = ""
inseresobremetal:
        texto = InputBox("Insira o valor para sobremetal (mm).")
        If IsNumeric(texto) Then
            frmVisu.FaceamentoSobremetal = CDbl(texto)
        Else
            MsgBox ("Valor Inválido!")
            GoTo inseresobremetal
        End If
        End If
        
    Label1.Caption = "Diâmetro Bruto = " & CStr(Round(frmVisu.DiametroBruto, 4)) & "mm"
    Label2.Caption = "Comprimento Bruto = " & CStr(Round(frmVisu.ComprimentoBruto, 4)) & "mm"
End Sub

