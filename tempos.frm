VERSION 5.00
Begin VB.Form tempos 
   Caption         =   "Estimativa de tempo de produção"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox List3 
      Height          =   5910
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular tempos"
      Height          =   975
      Left            =   3840
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.ListBox List2 
      Height          =   5910
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   5910
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "tempos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        For i = 0 To (List1.ListCount - 2)
            txt = frmVisu.lst_codcnc.List(i)
            
            poschar = InStr(1, txt, "/", vbTextCompare)
            
            If poschar = 0 Then
                poschar = InStr(1, txt, " ", vbTextCompare)
                If poschar = 0 Then
                    comando = txt
                Else
                    comando = Mid(txt, 1, (poschar - 1))
                    txt = Mid(txt, poschar + 1, Len(txt))
                    Do While Not Len(txt) = 0
                        poschar = InStr(1, txt, " ", vbTextCompare)
                        Param = Mid(txt, 1, 1)
                        txt = Mid(txt, 2, Len(txt))
                        poschar = InStr(1, txt, " ", vbTextCompare)
                        
                        If poschar = 0 Then
                            ValorParam = Mid(txt, 1, Len(txt))
                            txt = ""
                        Else
                            ValorParam = Mid(txt, 1, (poschar - 1))
                            txt = Mid(txt, poschar + 1, Len(txt))
                        End If
                        
                        Select Case Param
                            Case "X"
                                paramX = ValorParam
                            Case "Z"
                                paramZ = ValorParam
                            Case "F"
                                paramF = ValorParam
                            Case "U"
                                paramU = ValorParam
                            Case "R"
                                paramR = ValorParam
                            Case "P"
                                paramP = ValorParam
                            Case "Q"
                                paramQ = ValorParam
                            Case "W"
                                paramW = ValorParam

                        End Select
                        
                    Loop
                
                End If
            
                Select Case comando
                
                Case "G71"
                'faz um for para pegar o caminho
                If Not paramP = 0 Then
                    mov = (paramP / 10) + 2
                    For ii = mov To ((paramQ / 10) - 1)
                        txt = List1.List(ii)
                        poschar = InStr(1, txt, " ", vbTextCompare)
                        comando = Mid(txt, 1, (poschar - 1))
                        txt = Mid(txt, poschar + 1, Len(txt))
                        Do While Not Len(txt) = 0
                            poschar = InStr(1, txt, " ", vbTextCompare)
                            Param = Mid(txt, 1, 1)
                            txt = Mid(txt, 2, Len(txt))
                            poschar = InStr(1, txt, " ", vbTextCompare)
                            
                            If poschar = 0 Then
                                ValorParam = Mid(txt, 1, Len(txt))
                                txt = ""
                            Else
                                ValorParam = Mid(txt, 1, (poschar - 1))
                                txt = Mid(txt, poschar + 1, Len(txt))
                            End If
                            
                            Select Case Param
                                Case "X"
                                    paramX = ValorParam
                                Case "Z"
                                    paramZ = ValorParam
    
                            End Select
                            
                        Xdesb = (paramX / 2) - frmVisu.Sobremetal
                        Zdesb = frmVisu.ComprimentoBruto - frmVisu.ComprimentoFinal - paramZ - frmVisu.Sobremetal
                        Loop
                    Next ii
                    
                    'ALTERAR para soma dos tempos
                    ItemTempo = 1
                    List3.AddItem ItemTempo, i
                    
                    For K = (i + 1) To ((paramQ / 10) - 1)
                        ItemTempo = 0
                        List3.AddItem ItemTempo, K
                    Next K
                    
                    i = (paramQ / 10) - 1
                    paramU = 0
                    paramR = 0
                Else
                    profG71 = paramU
                    recuoG71 = paramR
                    ItemTempo = 0
                    List3.AddItem ItemTempo, i
                End If

                    
                
                Case Else
                'ocorre quando não é G00, G01, G02, G03, G70, G71
                    ItemTempo = 0
                    List3.AddItem ItemTempo, i
                
                End Select
            
            Else
                ItemTempo = 0
                List3.AddItem ItemTempo, i
            
            End If
            
            paramX = 0
            paramZ = 0
            paramF = 0
            paramU = 0
            paramR = 0
            paramP = 0
            paramQ = 0
            paramW = 0
        Next i
End Sub

Private Sub Form_Load()
    
        For i = 0 To frmVisu.lst_codcnc.ListCount
            txtinsere = frmVisu.lst_codcnc.List(i)
            List1.AddItem txtinsere
        Next
        
        For i = 0 To frmVisu.lst_codcnc.ListCount
            txtinsere = frmVisu.lst_n.List(i)
            List2.AddItem txtinsere
        Next
    
End Sub

