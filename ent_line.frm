VERSION 5.00
Begin VB.Form ent_code 
   Caption         =   "MecArm CNC Generator"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Iniciar movimento"
      Height          =   975
      Left            =   360
      TabIndex        =   47
      Top             =   4560
      Width           =   9135
      Begin VB.OptionButton Option7 
         Caption         =   "No menor valor de Y"
         Height          =   195
         Index           =   1
         Left            =   4920
         TabIndex        =   53
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Option7 
         Caption         =   "No maior valor de Y"
         Height          =   195
         Index           =   0
         Left            =   4920
         TabIndex        =   52
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option5 
         Caption         =   "No menor valor de X"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   51
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Option5 
         Caption         =   "No maior valor de X"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   50
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Movimento"
      Height          =   975
      Left            =   360
      TabIndex        =   46
      Top             =   3480
      Width           =   9135
      Begin VB.OptionButton Option4 
         Caption         =   "Anti-Horário"
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   49
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Horário"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   48
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Apoio"
      Height          =   3135
      Left            =   6240
      TabIndex        =   41
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton Option6 
         Caption         =   "Ligado se Y < que"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   55
         Top             =   2520
         Width           =   1815
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Ligado se Y > que"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   54
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   2160
         TabIndex        =   45
         Top             =   2400
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Desligado"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   44
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ligado"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   43
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Dinâmico"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   42
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Velocidade"
      Height          =   3135
      Left            =   3240
      TabIndex        =   21
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton Option2 
         Caption         =   "Dinâmico:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   32
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Estático:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   31
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Manter a anterior"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   600
         TabIndex        =   29
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   600
         TabIndex        =   28
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   600
         TabIndex        =   27
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "P1"
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label15 
         Caption         =   "P2"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "P3"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "F1"
         Height          =   255
         Left            =   1440
         TabIndex        =   37
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "F2"
         Height          =   255
         Left            =   1440
         TabIndex        =   36
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "F3"
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "F4"
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "F"
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "MecRep"
      Height          =   3135
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Text0 
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Top             =   1680
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Desligado"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Estático:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Dinâmico:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "M"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "M4"
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "M3"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "M2"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "M1"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "P3"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "P2"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "P1"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   6360
      Width           =   1935
   End
End
Attribute VB_Name = "ent_code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim inX As Double
    Dim inY As Double
    Dim fimX As Double
    Dim fimY As Double
    Dim Aux As Integer
    Dim AuxS As String
    Dim AuxString As String
    Dim rad As Double
    Dim ang1 As Double
    Dim ang2 As Double
    Dim radConv As Double
    Dim crX As Double
    Dim crY As Double
    
    
    
    
    


Private Sub Command1_Click()

    Dim MecRep As String
    Dim Velocidade As String
    Dim LineV As Boolean
    Dim LineA As Boolean

    
    
    
    If Option1(1) And Not IsNumeric(Text0.Text) Then
        MsgBox "M inválido!", , "Alerta"
    ElseIf Option1(2) And (Not IsNumeric(Text1.Text) Or Not IsNumeric(Text2.Text) Or Not IsNumeric(Text3.Text) Or Not IsNumeric(Text4.Text) Or Not IsNumeric(Text5.Text) Or Not IsNumeric(Text6.Text) Or Not IsNumeric(Text7.Text)) Then
        MsgBox "Parâmetros de MecRep inválidos!", , "Alerta"
    ElseIf Option2(1) And Not IsNumeric(Text8.Text) Then
        MsgBox "F inválido", , "Alerta"
    ElseIf Option2(2) And (Not IsNumeric(Text9.Text) Or Not IsNumeric(Text10.Text) Or Not IsNumeric(Text15.Text) Or Not IsNumeric(Text11.Text) Or Not IsNumeric(Text12.Text) Or Not IsNumeric(Text13.Text) Or Not IsNumeric(Text14.Text)) Then
        MsgBox "Parâmetros de Velocidade inválidos!", , "Alerta"
    ElseIf Option3(2) And (Not IsNumeric(Text16)) Then
    ElseIf entrada.dr = "LINE" And Not Option5(0) And Not Option5(1) And Not Option7(0) And Not Option7(1) Then
        MsgBox "Selecione o início do movimento!", , "Alerta"
    
    'REVER AQUI------------------------------------------------------
    'ElseIf entrada.dr = "ARC"
    
    

    Else
'INÍCIO DE CODIFICAÇÃO DE GERAÇÃO DE CNC________________________________________________________________
    
    ListIndex = frmVisu.lst_codcnc.ListIndex
    ListCount = frmVisu.lst_codcnc.ListCount
    
    'MecRep
    If Option1(0) Then
        MecRep = "M20"
    ElseIf Option1(1) Then
        MecRep = "M21 M" + CStr(Text0.Text)
    Else
        MecRep = "M22" + " P1 " + CStr(Text1.Text) + " P2 " + CStr(Text2.Text) + " P3 " + CStr(Text3.Text) + " M1 " + CStr(Text4.Text) + " M2 " + CStr(Text5.Text) + " M3 " + CStr(Text6.Text) + " M4 " + CStr(Text7.Text)
    End If
    
    
    'Velocidade
    If Option2(0) Then
        Velocidade = ""
        LineV = False
    ElseIf Option2(1) Then
        Velocidade = "F " + CStr(Text8.Text)
        LineV = True
    ElseIf Option2(2) Then
        Velocidade = "M30" + " P1 " + CStr(Text15.Text) + " P2 " + CStr(Text14.Text) + " P3 " + CStr(Text13.Text) + " F1 " + CStr(Text12.Text) + " F2 " + CStr(Text11.Text) + " F3 " + CStr(Text10.Text) + " F4 " + CStr(Text9.Text)
        LineV = True
    End If
    
    
    'Apoio
    If Option3(0) Then
        Apoio = ""
        LineA = False
    ElseIf Option3(1) Then
        Apoio = "M27"
        LineA = True
    ElseIf Option3(2) Then
        If Option6(0) Then
            Apoio = "M28" + "M=1 " + "P1 " + CStr(Text16.Text)
        ElseIf Option6(1) Then
            Apoio = "M28" + "M=0 " + "P1 " + CStr(Text16.Text)
        End If
        LineV = True
    End If
    
    
    'Inserção
    frmVisu.lst_codcnc.AddItem MecRep, ListIndex + 1
    ListIndex = ListIndex + 1
    
    If LineV Then
        frmVisu.lst_codcnc.AddItem Velocidade, ListIndex + 1
        ListIndex = ListIndex + 1
    End If
    
    If LineA Then
        frmVisu.lst_codcnc.AddItem Apoio, ListIndex + 1
        ListIndex = ListIndex + 1
    End If

    If entrada.dr = "LINE" Then
        
        If Option5(0) Then
            If inX > fimX Then
                AuxS = "G0 X" + CStr(inX) + " Y" + CStr(inY)
                AuxS2 = "G1 X" + CStr(fimX) + " Y" + CStr(fimY)
            Else
                AuxS = "G0 X" + CStr(fimX) + " Y" + CStr(fimY)
                AuxS2 = "G1 X" + CStr(inX) + " Y" + CStr(inY)
            End If
        ElseIf Option5(1) Then
            If inX < fimX Then
                AuxS = "G0 X" + CStr(inX) + " Y" + CStr(inY)
                AuxS2 = "G1 X" + CStr(fimX) + " Y" + CStr(fimY)
            Else
                AuxS = "G0 X" + CStr(fimX) + " Y" + CStr(fimY)
                AuxS2 = "G1 X" + CStr(inX) + " Y" + CStr(inY)
            End If
        ElseIf Option7(0) Then
            If inY > fimY Then
                AuxS = "G0 X" + CStr(inX) + " Y" + CStr(inY)
                AuxS2 = "G1 X" + CStr(fimX) + " Y" + CStr(fimY)
            Else
                AuxS = "G0 X" + CStr(fimX) + " Y" + CStr(fimY)
                AuxS2 = "G1 X" + CStr(inX) + " Y" + CStr(inY)
            End If
        ElseIf Option7(1) Then
            If inY < fimY Then
                AuxS = "G0 X" + CStr(inX) + " Y" + CStr(inY)
                AuxS2 = "G1 X" + CStr(fimX) + " Y" + CStr(fimY)
            Else
                AuxS = "G0 X" + CStr(fimX) + " Y" + CStr(fimY)
                AuxS2 = "G1 X" + CStr(inX) + " Y" + CStr(inY)
            End If
        End If
        
    End If
    
    If entrada.dr = "ARC" Then
        
        If Option4(0) Then 'horário
            'If ang2 > ang1 Then 'inverte in e fim
                AuxS = "G0 X" + CStr(Round(fimX, 4)) + " Y" + CStr(Round(fimY, 4))
                If crX > inX Then
                    AuxS2 = "G02 X" + CStr(Round(inX, 4)) + " Y" + CStr(Round(inY, 4)) + " I" + CStr(Round(crX - inX))
                Else
                    AuxS2 = "G02 X" + CStr(inX) + " Y" + CStr(inY) + " I" + CStr(inX - crX)
                End If

                If crY > inY Then
                    AuxS2 = AuxS2 + " J" + (crY - inY)
                Else
                    AuxS2 = AuxS2 + " J" + CStr(Round((inY - crY), 4))
                End If
            'Else 'não inverte
'                AuxS = "G0 X" + CStr(Round(inX, 4)) + " Y" + CStr(Round(inY, 4))
'                If crX > fimX Then
'                    AuxS2 = "G02 X" + CStr(Round(fimX, 4)) + " Y" + CStr(Round(fimY, 4)) + " I" + CStr(Round(crX - fimX))
'                Else
'                    AuxS2 = "G02 X" + fimX + " Y" + fimY + " I" + (fimX - crX)
'                End If
'
'                If crY > fimY Then
'                    AuxS2 = AuxS2 + " J" + (crY - fimY)
'                Else
'                    AuxS2 = AuxS2 + " J" + CStr(Round((fimY - crY), 4))
'                End If
'            End If

'''''        If Angle1 > Angle2 Then
'''''            Anglef = Angle2
'''''            Anglei = Angle1
'''''        ElseIf Angle1 < Angle2 Then
'''''            Anglef = Angle1
'''''            Anglei = Angle2
'''''        Else
'''''            Exit Sub
'''''        End If
'''''        AuxS = "G0 X" + CStr(fimX) + " Y" + CStr(fimY)
'''''        AuxS2 = "G02 X" + CStr(fimX) + " Y" + CStr(fimY) + " I" + CStr(crX - fimX)
        
        
        ElseIf Option4(1) Then 'anti horario
'            If ang2 < ang1 Then 'inverte in e fim
'                AuxS = "G0 X" + CStr(Round(fimX, 4)) + " Y" + CStr(Round(fimY, 4))
'                If crX > inX Then
'                    AuxS2 = "G02 X" + CStr(Round(inX, 4)) + " Y" + CStr(Round(inY, 4)) + " I" + CStr(Round(crX - inX))
'                Else
'                    AuxS2 = "G02 X" + inX + " Y" + inY + " I" + (inX - crX)
'                End If
'
'                If crY > inY Then
'                    AuxS2 = AuxS2 + " J" + (crY - inY)
'                Else
'                    AuxS2 = AuxS2 + " J" + CStr(Round((inY - crY), 4))
'                End If
'            Else 'não inverte
                AuxS = "G0 X" + CStr(Round(inX, 4)) + " Y" + CStr(Round(inY, 4))
                If crX > fimX Then
                    AuxS2 = "G03 X" + CStr(Round(fimX, 4)) + " Y" + CStr(Round(fimY, 4)) + " I" + CStr(Round(crX - fimX))
                Else
                    AuxS2 = "G03 X" + fimX + " Y" + fimY + " I" + (fimX - crX)
                End If
                
                If crY > fimY Then
                    AuxS2 = AuxS2 + " J" + CStr(crY - fimY)
                Else
                    AuxS2 = AuxS2 + " J" + CStr(Round((fimY - crY), 4))
                End If
           'End If
        Else
            
        End If
        
    End If
    
    frmVisu.lst_codcnc.AddItem AuxS, ListIndex + 1
    ListIndex = ListIndex + 1
    frmVisu.lst_codcnc.AddItem AuxS2, ListIndex + 1
    ListIndex = ListIndex + 1
    
    
    
    
    ent_code.Visible = False
    ent_code.Enabled = False
    Unload ent_code
    
    End If
    

    
End Sub

Private Sub Form_Load()
    'If Not Option1(2) Then
        Text0.Enabled = False
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
        Text6.Enabled = False
        Text7.Enabled = False
        Text8.Enabled = False
        Text9.Enabled = False
        Text10.Enabled = False
        Text11.Enabled = False
        Text12.Enabled = False
        Text13.Enabled = False
        Text14.Enabled = False
        Text15.Enabled = False
        Text16.Enabled = False
        
        Option6(0).Enabled = False
        Option6(1).Enabled = False
        
    'End If
    
    If entrada.dr = "LINE" Then
        Frame5.Enabled = True
        Frame4.Enabled = False
            
            AuxString = entrada.dados
                
            Aux = Len(AuxString)
            AuxString = Mid(AuxString, 6, Aux)
            
            Aux = InStr(1, AuxString, ";")
            AuxS = Mid(AuxString, 1, Aux - 1)
            inX = CDbl(AuxS)
            AuxString = Mid(AuxString, Aux + 1, Len(AuxString) - Aux)
            
            Aux = InStr(1, AuxString, ";")
            AuxS = Mid(AuxString, 1, Aux - 1)
            inY = CDbl(AuxS)
            AuxString = Mid(AuxString, Aux + 1, Len(AuxString) - Aux)
            
            Aux = InStr(1, AuxString, ";")
            AuxS = Mid(AuxString, 1, Aux - 1)
            fimX = CDbl(AuxS)
            AuxString = Mid(AuxString, Aux + 1, Len(AuxString) - Aux)
            
            fimY = CDbl(AuxString)
            
            If inX > fimX Then
                maiorX = inX
                Option5(0).Enabled = True
                Option5(1).Enabled = True
            ElseIf inX < fimX Then
                maiorX = fimX
                Option5(0).Enabled = True
                Option5(1).Enabled = True
            Else
                Option5(0).Enabled = False
                Option5(1).Enabled = False
            End If
            
            
            If inY > fimY Then
                maiorY = inY
                Option7(0).Enabled = True
                Option7(1).Enabled = True
            ElseIf inY < fimY Then
                maiorY = fimY
                Option7(0).Enabled = True
                Option7(1).Enabled = True
            Else
                Option7(0).Enabled = False
                Option7(1).Enabled = False
            End If
            
            

    ElseIf entrada.dr = "ARC" Then
        Frame4.Enabled = True
        Frame5.Enabled = False
        
            radConv = Round(3.141592654 / 180, 7)
'
'            cosa =
'            X1 = Xc + rad * Cos(Angle1 * radConv)
'            sena =
'            Y1 = Yc + rad * Sin(Angle1 * radConv)
'            Xf = Xc + rad * Cos(Angle2 * radConv)
'            Yf = Yc + rad * Sin(Angle2 * radConv)

        AuxString = entrada.dados
        
        Aux = Len(AuxString)
        AuxString = Mid(AuxString, 5, Aux)
        
        Aux = InStr(1, AuxString, ";")
        AuxS = Mid(AuxString, 1, Aux - 1)
        crX = CDbl(AuxS)
        AuxString = Mid(AuxString, Aux + 1, Len(AuxString) - Aux)
        
        Aux = InStr(1, AuxString, ";")
        AuxS = Mid(AuxString, 1, Aux - 1)
        crY = CDbl(AuxS)
        AuxString = Mid(AuxString, Aux + 1, Len(AuxString) - Aux)
        
        Aux = InStr(1, AuxString, ";")
        AuxS = Mid(AuxString, 1, Aux - 1)
        rad = CDbl(AuxS)
        AuxString = Mid(AuxString, Aux + 1, Len(AuxString) - Aux)
        
        Aux = InStr(1, AuxString, ";")
        AuxS = Mid(AuxString, 1, Aux - 1)
        ang1 = CDbl(AuxS)
        AuxString = Mid(AuxString, Aux + 1, Len(AuxString) - Aux)
        
        Aux = InStr(1, AuxString, ";")
        AuxS = AuxString
        ang2 = CDbl(AuxS)
        'AuxString = Mid(AuxString, Aux + 1, Len(AuxString) - Aux)
        
        cosa = Cos(ang1 * radConv)
        inX = crX + rad * cosa
        sena = Sin(ang1 * radConv)
        inY = crY + rad * sena
        
        cosa = Cos(ang2 * radConv)
        fimX = crX + rad * cosa
        sena = Sin(ang2 * radConv)
        fimY = crY + rad * sena
        
        
    End If
    
    
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(1) Then
        Text0.Enabled = True
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
        Text6Enabled = False
        Text7.Enabled = False
    ElseIf Option1(2) Then
        Text0.Enabled = False
        Text1.Enabled = True
        Text2.Enabled = True
        Text3.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
        Text6.Enabled = True
        Text7.Enabled = True
    End If
End Sub

Private Sub Option2_Click(Index As Integer)
    If Option2(1) Then
        Text8.Enabled = True
        Text9.Enabled = False
        Text10.Enabled = False
        Text11.Enabled = False
        Text12.Enabled = False
        Text13.Enabled = False
        Text14.Enabled = False
        Text15.Enabled = False
    ElseIf Option2(2) Then
        Text8.Enabled = False
        Text9.Enabled = True
        Text10.Enabled = True
        Text11.Enabled = True
        Text12.Enabled = True
        Text13.Enabled = True
        Text14.Enabled = True
        Text15.Enabled = True
    Else
        Text8.Enabled = False
        Text9.Enabled = False
        Text10.Enabled = False
        Text11.Enabled = False
        Text12.Enabled = False
        Text13.Enabled = False
        Text14.Enabled = False
        Text15.Enabled = False
    End If
End Sub

Private Sub Text17_Change()

End Sub

Private Sub Option3_Click(Index As Integer)
    If Option3(0) Then
        Option6(0).Enabled = False
        Option6(1).Enabled = False
        Text16.Enabled = False
    ElseIf Option3(1) Then
        Option6(0).Enabled = False
        Option6(1).Enabled = False
        Text16.Enabled = False
    Else
        Option6(0).Enabled = True
        Option6(1).Enabled = True
        Text16.Enabled = True
    End If
End Sub

Private Sub Option5_Click(Index As Integer)
    Option7(0).Value = False
    Option7(1).Value = False
End Sub

Private Sub Option7_Click(Index As Integer)
    Option5(0).Value = False
    Option5(1).Value = False
    
End Sub
