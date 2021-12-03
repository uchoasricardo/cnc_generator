VERSION 5.00
Begin VB.Form func_desb 
   Caption         =   "Desbaste"
   ClientHeight    =   1935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4290
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Profundidade de desbaste (mm):"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Sobremetal (mm):"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "func_desb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim txtinsere As String

                If Not IsNumeric(Text1.Text) Then
                    MsgBox "Valor inválido para Sobremetal!"
                    Text1.Text = ""
                ElseIf Not IsNumeric(Text2.Text) Then
                    MsgBox "Valor inválido para Profundidade de Desbaste!"
                    Text2.Text = ""
                Else
                     Text1.Text = Replace(Text1.Text, ".", ",")
                     Text2.Text = Replace(Text2.Text, ".", ",")
                     
                     frmVisu.Sobremetal = CDbl(Text1.Text)
                     frmVisu.ProfundidadeDesbaste = CDbl(Text2.Text)
                     
                     
                    txtinsere = "G00 X" & CStr(frmVisu.DiametroBruto) & " Z" & CStr(frmVisu.ComprimentoBruto - frmVisu.ComprimentoFinal + 1)
                    txtinsere = Replace(txtinsere, ",", ".")
                    frmVisu.lst_codcnc.AddItem txtinsere
                    frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
                    frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
            
            
                    txtinsere = "/Ciclo de Desbaste"
                    'txtinsere = Replace(txtinsere, ",", ".")
                    frmVisu.lst_codcnc.AddItem txtinsere
                    frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
                    frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        
        
                     txtinsere = "G71 U" & CStr(frmVisu.ProfundidadeDesbaste) & " R2"
                     txtinsere = Replace(txtinsere, ",", ".")
                     frmVisu.lst_codcnc.AddItem txtinsere
                     frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
                     frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
                     frmVisu.PosG71 = frmVisu.lst_codcnc.ListCount
                     
                     
                                 frmVisu.Command4.Enabled = True
                                frmVisu.Command5.Enabled = True
                                
                                frmVisu.btn_veloc.Enabled = False
                                frmVisu.btn_zerar.Enabled = False
                                frmVisu.btn_cabecalho.Enabled = False
                                frmVisu.btn_desbastar.Enabled = False
                                frmVisu.btn_acabar.Enabled = False
                                frmVisu.btn_rotacaoligar.Enabled = False
                                frmVisu.btn_rotacaodesligar.Enabled = False
                                frmVisu.btn_fluidoligar.Enabled = False
                                frmVisu.btn_fluidodesligar.Enabled = False
                                frmVisu.btn_delay.Enabled = False
                                frmVisu.Command3.Enabled = False
                                
                                func_desb.Visible = False
                                func_desb.Enabled = False
                End If
            


    
End Sub

'Private Sub Form_Load()
'
'    Dim txtzero As String
'    Dim tipo As String
'    Dim pZ As Double
'    Dim sZ As Double
'    Dim pX As Double
'    Dim sX As Double
'    Dim poschar As Integer
'
'
'
'            txtzero = frmVisu.lst_entidades.Text
'            poschar = InStr(1, txtzero, ";", vbTextCompare)
'            tipo = Mid(txtzero, 1, (poschar - 1))
'            txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
'
'            If tipo = "LINE" Then
'            poschar = InStr(1, txtzero, ";", vbTextCompare)
'            pZ = CDbl(Mid(txtzero, 1, (poschar - 1)))
'            txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
'            poschar = InStr(1, txtzero, ";", vbTextCompare)
'            pX = CDbl(Mid(txtzero, 1, (poschar - 1)))
'            txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
'            poschar = InStr(1, txtzero, ";", vbTextCompare)
'            sZ = CDbl(Mid(txtzero, 1, (poschar - 1)))
'            txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
'            sX = CDbl(Mid(txtzero, 1, (poschar - 1)))
'
''            If pZ <= sZ Then
' '               ZOffset = pZ
''            Else
''                ZOffset = sZ
''            End If
''
''            If pX <= sX Then
''                XOffset = pX
''            Else
''                XOffset = sX
''            End If
'
'            'SistemaReferenciado = True
'            'frmVisu.lbl_offset.Caption = "Offset X =" + CStr(Round(ZOffset, 4)) + "/ Offset Y =" + CStr(Round(XOffset, 4))
'
'
'        Else
'            MsgBox ("Entidade inválida!")
'        End If
'
'    Label1.Caption = "Z1: " & CStr(Round(pZ - frmVisu.ZOffset, 4))
'    Label2.Caption = "Z2: " & CStr(Round(sZ - frmVisu.ZOffset, 4))
'    Label3.Caption = "X1: " & CStr(Round(pX - frmVisu.XOffset, 4))
'    Label4.Caption = "X2: " & CStr(Round(sX - frmVisu.XOffset, 4))
'
'End Sub

Private Sub Label1_Click()

End Sub
