VERSION 5.00
Begin VB.Form frmVisu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador CNC"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   17970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   9405
   ScaleWidth      =   17970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "Ok"
      Height          =   375
      Left            =   10440
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10440
      TabIndex        =   36
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Estimar tempo de produção"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   35
      Top             =   7320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Linha Final de Acabamento"
      Enabled         =   0   'False
      Height          =   735
      Left            =   10440
      TabIndex        =   34
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Linha Inicial de Acabamento"
      Enabled         =   0   'False
      Height          =   735
      Left            =   10440
      TabIndex        =   33
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Editar Linha"
      Height          =   735
      Left            =   16920
      TabIndex        =   32
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Perfil inserido"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9240
      TabIndex        =   31
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ListBox lst_n 
      Height          =   7665
      Left            =   12600
      TabIndex        =   30
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Inserir Perfil"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   29
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Finalizar programa"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   28
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton btn_veloc 
      Caption         =   "Alterar parâmetros"
      Height          =   615
      Left            =   8040
      TabIndex        =   25
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton btn_entant 
      Caption         =   "<<"
      Height          =   615
      Left            =   6600
      TabIndex        =   23
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton btn_proxent 
      Caption         =   ">>"
      Height          =   615
      Left            =   7320
      TabIndex        =   22
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton btn_rotacaodesligar 
      Caption         =   "Desligar rotação"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9240
      TabIndex        =   21
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton btn_rotacaoligar 
      Caption         =   "Ligar rotação"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   20
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton btn_fluidodesligar 
      Caption         =   "Desligar fluido"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9240
      TabIndex        =   19
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton btn_fluidoligar 
      Caption         =   "Ligar fluido"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   18
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton btn_acabar 
      Caption         =   "Acabamento"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   17
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton btn_desbastar 
      Caption         =   "Desbastar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   16
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton btn_zerar 
      Caption         =   "Referência (zero)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   15
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton btn_cabecalho 
      Caption         =   "Inserir cabeçalho"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   14
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "\/"
      Enabled         =   0   'False
      Height          =   375
      Left            =   16920
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "/\"
      Enabled         =   0   'False
      Height          =   375
      Left            =   16920
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdZoomOut 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2040
      Picture         =   "principal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton cmdZoomIn 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   1440
      Picture         =   "principal.frx":0067
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6720
      Width           =   615
   End
   Begin VB.OptionButton optMouse 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   1
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "principal.frx":00CF
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Magnify"
      Top             =   6720
      Width           =   615
   End
   Begin VB.OptionButton optMouse 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   0
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "principal.frx":0135
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Pan"
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton btn_salvasai 
      Caption         =   "Salvar e Sair"
      Height          =   735
      Left            =   13320
      TabIndex        =   7
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton btn_sai 
      Caption         =   "Sair"
      Height          =   735
      Left            =   15120
      TabIndex        =   6
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton btn_excluilinha 
      Caption         =   "Excluir Linha"
      Height          =   735
      Left            =   16920
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton btn_inserelinha 
      Caption         =   "Inserir Linha"
      Height          =   735
      Left            =   16920
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.ListBox lst_codcnc 
      Height          =   7665
      Left            =   13320
      TabIndex        =   3
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton btn_delay 
      Caption         =   "Delay"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   2
      Top             =   5880
      Width           =   2295
   End
   Begin VB.ListBox lst_entidades 
      Enabled         =   0   'False
      Height          =   645
      Left            =   240
      OLEDragMode     =   1  'Automatic
      TabIndex        =   1
      Top             =   7440
      Width           =   7695
   End
   Begin VB.PictureBox VisuDXF 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000012&
      Height          =   6495
      Left            =   240
      ScaleHeight     =   6435
      ScaleMode       =   0  'User
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label Label3 
      Caption         =   "Linha"
      Height          =   255
      Left            =   12720
      TabIndex        =   40
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Programa CNC"
      Height          =   255
      Left            =   13440
      TabIndex        =   39
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Ferramenta de Acabamento:"
      Height          =   495
      Left            =   10440
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbl_unit 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10440
      TabIndex        =   27
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lbl_vel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   26
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lbl_offset 
      Caption         =   "Label1"
      Height          =   615
      Left            =   360
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Menu Arquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu abrir 
         Caption         =   "Abrir DXF..."
         Index           =   0
      End
      Begin VB.Menu salvar 
         Caption         =   "Salvar CNC"
         Index           =   1
      End
      Begin VB.Menu sair 
         Caption         =   "Sair"
         Index           =   2
      End
   End
   Begin VB.Menu ajuda 
      Caption         =   "Ajuda"
      Index           =   0
   End
End
Attribute VB_Name = "frmVisu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ListIndex As Integer
Dim ListCount As Integer
Dim TempString As String
Dim Aux As Integer
Dim Plot As Boolean
Dim Entite As String
Dim Position As Integer
Dim Lenght As Integer
Dim Sel As String
Dim Sele As Boolean
Dim dr As String
Dim div As Double
Dim SaveName As String
Dim pos As Integer
Dim cont As String
Dim Import As Boolean
Dim txtinsere As String
Public SistemaReferenciado As Boolean
Public ParametrosDefinidos As Boolean
'Dim Velocidade As Integer
Public ZOffset As Double
Public XOffset As Double
Public DiametroBruto As Double
Public ComprimentoBruto As Double
Public ComprimentoFinal As Double
Public VelocidadeCorte As Double
Public Rotacao As Double
Public VelAvanco As Double
Public VelAvancoAcab As Double
Public ProfundidadeFaceamento As Double
Public FaceamentoSobremetal As Double
Public ProfundidadeDesbaste As Double
Public Ferramenta As Integer
Public FaceamentoCompFinal As Boolean
Dim FlagProgramando As Boolean
Public NumLinhas As Integer
Public Sobremetal As Double
Public Linha As Integer
Public PosG71 As Integer
Public InAcab As Integer
Public FiAcab As Integer


Public UltimoX As Double
Public UltimoZ As Double

Public NomePrograma As String
Public NumeroPrograma As Integer



  


Dim X1 As Double
Dim X2 As Double
Dim Y1 As Double
Dim Y2 As Double

Dim rad As Double
Dim Angle1 As Double
Dim Angle2 As Double




Dim i As Integer




'Base de données corrrespondant au contenu du fichier DXF

Dim BdDXF As DXFDonnee

Dim DragX As Long
Dim DragY As Long

Dim SelGroup As RECT
Dim Deplace As Boolean
Dim Zoom As Boolean
Dim SvgDeplace As Boolean
Dim SvgZoom As Boolean


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

       Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
         
        Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
    Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

       Private Type OPENFILENAME
         lStructSize As Long
         hwndOwner As Long
         hInstance As Long
         lpstrFilter As String
         lpstrCustomFilter As String
         nMaxCustFilter As Long
         nFilterIndex As Long
         lpstrFile As String
         nMaxFile As Long
         lpstrFileTitle As String
         nMaxFileTitle As Long
         lpstrInitialDir As String
         lpstrTitle As String
         flags As Long
         nFileOffset As Integer
         nFileExtension As Integer
         lpstrDefExt As String
         lCustData As Long
         lpfnHook As Long
         lpTemplateName As String
       End Type
Dim fso As New FileSystemObject
Dim arqtxt As TextStream


Public Sub abrir_Click(Index As Integer)
         Dim salvar As Integer
         
         
         Dim OpenFile As OPENFILENAME
         Dim lReturn As Long
         Dim sFilter As String
         
         
        'Verifica se já há código
        If FlagProgramando = True Then
            salvar = MsgBox("Deseja salvar o código gerado?", vbYesNoCancel, "Salvar código gerado...")
            
            If salvar = 6 Then
                    MsgBox "Salve o código CNC, e em seguida abra um novo DXF..."
                    salvar_Click 0
            ElseIf salvar = 7 Then
                    lst_codcnc.Clear
                    MsgBox "Código descartado! Abra um novo arquivo DXF..."
            ElseIf salvar = 2 Then
                    Exit Sub
            End If
        End If

         
         
         
         
         OpenFile.lStructSize = Len(OpenFile)
         OpenFile.hwndOwner = frmVisu.hWnd
         OpenFile.hInstance = App.hInstance
         sFilter = "Batch Files (*.dxf)" & Chr(0) & "*.dxf" & Chr(0)
         OpenFile.lpstrFilter = sFilter
         OpenFile.nFilterIndex = 1
         OpenFile.lpstrFile = String(257, 0)
         OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
         OpenFile.lpstrFileTitle = OpenFile.lpstrFile
         OpenFile.nMaxFileTitle = OpenFile.nMaxFile
         OpenFile.lpstrInitialDir = "C:\"
         OpenFile.lpstrTitle = "Abrir DXF..."
         OpenFile.flags = 0
         lReturn = GetOpenFileName(OpenFile)
         If lReturn = 0 Then
'            MsgBox "The User pressed the Cancel Button"
         Else
            '
            Dim i As Integer
        

        
            ImportDXF Trim(OpenFile.lpstrFile), BdDXF
            
            lst_entidades.Clear
            Plot = True
            
            Import = True
            Call VisuDXF_DblClick
            Import = False
            FlagProgramando = True
            
            SistemaReferenciado = False
            ZOffset = 0
            XOffset = 0
            ent_param.Enabled = True
            ent_param.Visible = True
            
            'On Error Resume Next
            'TreeViewLayer.Nodes.Clear
            'TreeViewLayer.Nodes.Add , tvwLast, "LAYER", "Affichage des Layers "
            'TreeViewLayer.Nodes.Item(1).Expanded = True
            
            
            ' liste les layers
            'For I = 0 To UBound(MonLayer)
            'TreeViewLayer.Nodes.Add "LAYER", tvwChild, "Layer" & I, MonLayer(I).Nom
            'Next I
            '
            
            'MsgBox "The user Chose " & Trim(OpenFile.lpstrFile)
         End If
         
         If frmVisu.lst_entidades.ListCount > 0 Then
            frmVisu.lst_entidades.ListIndex = 0
         End If
         
         SistemaReferenciado = False
         ParametrosDefinidos = False
         
         

         
    

End Sub

Private Sub ajuda_Click(Index As Integer)
    form_ajuda.Visible = True
    form_ajuda.Enabled = True
    
End Sub

Private Sub btn_acabar_Click()
        
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

        frmVisu.Label1.Visible = True
        frmVisu.Text1.Visible = True
        frmVisu.Text1.Enabled = True
        Command10.Visible = True
        frmVisu.Command7.Visible = True
        frmVisu.Command8.Visible = True
End Sub

Private Sub btn_cabecalho_Click()
    Dim ListIndex As Integer
        'ListIndex = frmVisu.lst_codcnc.ListIndex
        'If ListIndex = -1 Then
        txtinsere = "/Cabeçalho"
        frmVisu.lst_codcnc.AddItem txtinsere, 0
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)

        txtinsere = "G21"
        frmVisu.lst_codcnc.AddItem txtinsere, 1
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)

        txtinsere = "G90"
        frmVisu.lst_codcnc.AddItem txtinsere, 2
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)

        txtinsere = "G54"
        frmVisu.lst_codcnc.AddItem txtinsere, 3
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        
        txtinsere = "G00 X350 Z250 T00"
        frmVisu.lst_codcnc.AddItem txtinsere, 4
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)

        txtinsere = "G95"
        frmVisu.lst_codcnc.AddItem txtinsere, 5
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)

        txtinsere = "M04"
        frmVisu.lst_codcnc.AddItem txtinsere, 6
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)

        txtinsere = "M08"
        frmVisu.lst_codcnc.AddItem txtinsere, 7
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)

        txtinsere = "G92 S4000"
        frmVisu.lst_codcnc.AddItem txtinsere, 8
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)

        txtinsere = "G96 S" & CStr(frmVisu.VelocidadeCorte)
        frmVisu.lst_codcnc.AddItem txtinsere, 9
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
       
        If frmVisu.Ferramenta < 10 Then
            txtinsere = "T0" & CStr(frmVisu.Ferramenta) & "0" & CStr(frmVisu.Ferramenta)
            frmVisu.lst_codcnc.AddItem txtinsere, 10
            frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
            frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        Else
            txtinsere = "T" & CStr(frmVisu.Ferramenta) & CStr(frmVisu.Ferramenta)
            frmVisu.lst_codcnc.AddItem txtinsere, 10
            frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
            frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        End If
        
            'frmVisu.btn_veloc.Enabled = True
            'frmVisu.btn_zerar.Enabled = True
            'frmVisu.btn_cabecalho.Enabled = True
            frmVisu.btn_desbastar.Enabled = True
            'frmVisu.btn_acabar.Enabled = True
            frmVisu.btn_rotacaoligar.Enabled = True
            frmVisu.btn_rotacaodesligar.Enabled = True
            frmVisu.btn_fluidoligar.Enabled = True
            frmVisu.btn_fluidodesligar.Enabled = True
            frmVisu.btn_delay.Enabled = True
            frmVisu.Command3.Enabled = True


End Sub

Private Sub btn_delay_Click()
    func_delay.Enabled = True
    func_delay.Visible = True
End Sub

Private Sub btn_desbastar_Click()


        
    If lst_entidades.ListCount = 0 Then
        MsgBox "Carregue um arquivo DXF em Arquivo -> Abrir DXF..."
        
    ElseIf SistemaReferenciado = False Then
        MsgBox "É necessário referenciar o sistema!"
    ElseIf ParametrosDefinidos = False Then
        MsgBox "É necessário definir os parâmetros básicos!"
        ent_param.Enabled = True
        ent_param.Visible = True
    Else
        func_desb.Visible = True
        func_desb.Enabled = True
        
        
    End If

End Sub

Private Sub btn_entant_Click()
    Dim ListPos As Integer
    Dim ListCount As Integer
    
    ListPos = lst_entidades.ListIndex

    If ListPos <= 0 Then
        
    Else
        frmVisu.lst_entidades.ListIndex = ListPos - 1
    End If
    
                    'Redesenha zero
            If frmVisu.SistemaReferenciado Then
                frmVisu.VisuDXF.Circle (ZOffset, -XOffset), 3, vbGreen
                frmVisu.VisuDXF.Line (ZOffset - 6, -XOffset)-(ZOffset + 6, -XOffset), vbGreen
                frmVisu.VisuDXF.Line (ZOffset, -XOffset - 6)-(ZOffset, -XOffset + 6), vbGreen
            End If
    
End Sub

Private Sub btn_excluilinha_Click()
Dim NIndex As Integer
    ListIndex = frmVisu.lst_codcnc.ListIndex
    NIndex = frmVisu.lst_n.ListCount - 1
    If ListIndex = -1 Then
        MsgBox "Selecione alguma linha!"
    Else
        frmVisu.lst_codcnc.RemoveItem ListIndex
        frmVisu.lst_n.RemoveItem NIndex
    End If

    
End Sub

Private Sub btn_facear_Click()
Dim texto As String
        frmVisu.FaceamentoCompFinal = True
        
    If lst_entidades.ListCount = 0 Then
        MsgBox "Carregue um arquivo DXF em Arquivo -> Abrir DXF..."
        
    ElseIf SistemaReferenciado = False Then
        MsgBox "É necessário referenciar o sistema!"
    ElseIf ParametrosDefinidos = False Then
        MsgBox "É necessário definir os parâmetros básicos!"
        ent_param.Enabled = True
        ent_param.Visible = True
    Else
        func_face.Visible = True
        func_face.Enabled = True
        If Not (func_face.sZ = func_face.pZ) Then
            Unload func_face
        End If
        If Not (func_face.sZ = ZOffset) Then
            Unload func_face
        End If
        

        
    End If
    
    
End Sub

Private Sub btn_facearco_Click()
        frmVisu.FaceamentoCompFinal = False
        If lst_entidades.ListCount = 0 Then
        MsgBox "Carregue um arquivo DXF em Arquivo -> Abrir DXF..."
        
    ElseIf SistemaReferenciado = False Then
        MsgBox "É necessário referenciar o sistema!"
    ElseIf ParametrosDefinidos = False Then
        MsgBox "É necessário definir os parâmetros básicos!"
        ent_param.Enabled = True
        ent_param.Visible = True
    Else
        func_face.Visible = True
        func_face.Enabled = True
        If Not (func_face.sZ = func_face.pZ) Then
            Unload func_face
        End If
        If Not (func_face.sZ = ZOffset) And frmVisu.FaceamentoCompFinal Then
            Unload func_face
        End If
        
        
    End If
    
End Sub

Private Sub btn_fluidodesligar_Click()
        ListIndex = frmVisu.lst_codcnc.ListIndex
    'If ListIndex = -1 Then
        
        txtinsere = "/Desligar Fluido"
        'txtinsere = Replace(txtinsere, ",", ".")
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        
        txtinsere = "M09"
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)

End Sub

Private Sub btn_fluidoligar_Click()
        ListIndex = frmVisu.lst_codcnc.ListIndex
    'If ListIndex = -1 Then
        txtinsere = "M08"
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        frmVisu.lst_codcnc.ListIndex = ListIndex + 1
        txtinsere = ""
End Sub

Private Sub btn_inserelinha_Click()
        ListIndex = frmVisu.lst_codcnc.ListIndex
        If ListIndex = -1 Then
            txtinsere = InputBox("Insira o comando CNC:", "Inserir linha")
            frmVisu.lst_codcnc.AddItem txtinsere
        Else
            txtinsere = InputBox("Insira o comando CNC:", "Inserir linha")
            frmVisu.lst_codcnc.AddItem txtinsere, ListIndex + 1
            frmVisu.lst_codcnc.ListIndex = ListIndex + 1
            txtinsere = ""
        End If
        
                frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
                frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)

End Sub

Private Sub btn_ligaspindle_Click()
    ListIndex = frmVisu.lst_codcnc.ListIndex
    ListCount = frmVisu.lst_codcnc.ListCount
    TempString = "M3"
    'If ListIndex = -1 And ListCount = 0 Then
        frmVisu.lst_codcnc.AddItem TempString, ListIndex + 1
        frmVisu.lst_codcnc.ListIndex = ListIndex + 1
    'Else
    '    frmVisu.lst_codcnc.AddItem TempString, ListIndex + 1
    'End If
    
    'frmVisu.codcnc.Text = frmVisu.codcnc.Text + vbNewLine + TempString
End Sub

Private Sub btn_desligaspindle_Click()
    ListIndex = frmVisu.lst_codcnc.ListIndex
    ListCount = frmVisu.lst_codcnc.ListCount
    TempString = "M4"
    'If ListIndex = -1 And ListCount = 0 Then
        frmVisu.lst_codcnc.AddItem TempString, ListIndex + 1
        frmVisu.lst_codcnc.ListIndex = ListIndex + 1
    'Else
    '    frmVisu.lst_codcnc.AddItem TempString, ListIndex + 1
    'End If

End Sub


Private Sub btn_proxent_Click()
    Dim ListPos As Integer
    Dim ListCount As Integer
    
    ListPos = lst_entidades.ListIndex
    ListCount = lst_entidades.ListCount
    
    If ListPos = -1 Then
        
    ElseIf Not ListPos >= (ListCount - 1) Then
            frmVisu.lst_entidades.ListIndex = ListPos + 1
    End If
    
                    'Redesenha zero
            If frmVisu.SistemaReferenciado Then
                frmVisu.VisuDXF.Circle (ZOffset, -XOffset), 3, vbGreen
                frmVisu.VisuDXF.Line (ZOffset - 6, -XOffset)-(ZOffset + 6, -XOffset), vbGreen
                frmVisu.VisuDXF.Line (ZOffset, -XOffset - 6)-(ZOffset, -XOffset + 6), vbGreen
            End If
    
End Sub

Private Sub btn_rotacaodesligar_Click()
        ListIndex = frmVisu.lst_codcnc.ListIndex
    'If ListIndex = -1 Then
    
        txtinsere = "/Desligar Rotação"
        'txtinsere = Replace(txtinsere, ",", ".")
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        
        txtinsere = "M05"
        frmVisu.lst_codcnc.AddItem txtinsere
        'frmVisu.lst_codcnc.ListIndex = ListIndex + 1
        txtinsere = ""
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
End Sub

Private Sub btn_rotacaoligar_Click()
        ListIndex = frmVisu.lst_codcnc.ListIndex
    'If ListIndex = -1 Then
        txtinsere = "M04"
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.lst_codcnc.ListIndex = ListIndex + 1
        txtinsere = ""
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
End Sub

Private Sub btn_sai_Click()
    End
End Sub

Private Sub btn_salvasai_Click()

salvar_Click 0

End
'
'On Error GoTo trataerro
'
'Set arqtxt = fso.CreateTextFile("teste.nc", True)
''gravando no arquivo
'
'With arqtxt
'  '.WriteLine ("Isto é um teste")
'    ListCount = frmVisu.lst_codcnc.ListCount
'    For I = 0 To ListCount
'        TempString = frmVisu.lst_codcnc.List(I)
'        .WriteLine (TempString)
'    Next
'
'    .Close
'End With
'End
'Exit Sub
'
'trataerro:
'MsgBox Err.Description & " - " & Err.Number, vbCritical

End Sub






Private Sub frmVisu_Click()

End Sub

Private Sub CommandClick1_Click()

End Sub





Private Sub btn_veloc_Click()
        If lst_entidades.ListCount = 0 Then
            MsgBox "Carregue um arquivo DXF em Arquivo -> Abrir DXF..."
        Else
            ent_param.Enabled = True
            ent_param.Visible = True
        End If
End Sub

Private Sub btn_zerar_Click()
    Dim txtzero As String
    Dim tipo As String
    Dim pZ As Double
    Dim sZ As Double
    Dim pX As Double
    Dim sX As Double
    Dim poschar As Integer
    Dim texto As String

        
    If lst_entidades.ListCount = 0 Then
        MsgBox "Carregue um arquivo DXF em Arquivo -> Abrir DXF..."
        
    Else
        txtzero = lst_entidades.Text
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
            
            If (pX = sX) And (Not pZ = sZ) Then
                    If Not (pZ <= sZ) Then
                        ZOffset = pZ
                    Else
                        ZOffset = sZ
                    End If
                    
                    If pX <= sX Then
                        XOffset = pX
                    Else
                        XOffset = sX
                    End If
                    
                    SistemaReferenciado = True
                    frmVisu.lbl_offset.Caption = "Offset X =" + CStr(Round(ZOffset, 4)) + "/ Offset Y =" + CStr(Round(XOffset, 4))
                    
                    

                    ReAfficheDXF
                    
                    ComprimentoFinal = sZ - pZ

'inserecomprimento:
'                    texto = InputBox("Insira o comprimento final da peça em milímetros:")
'                    If IsNumeric(texto) Then
'                        If Not texto < frmVisu.ComprimentoBruto Then
'                            MsgBox "O comprimento final não pode ser maior do que o comprimento bruto! (C.Bruto = " & frmVisu.ComprimentoBruto & "mm)"
'                            GoTo inserecomprimento
'                        Else
'                            ComprimentoFinal = CDbl(texto)
'                        End If
'                    Else
'                        MsgBox "Valor inválido. Insira um número!"
'                        GoTo inserecomprimento
'                    End If
                    
                    MsgBox "Sistema referenciado."
                    lst_entidades_Click
                    
                    'Redesenha zero
                    If frmVisu.SistemaReferenciado Then
                        frmVisu.VisuDXF.Circle (ZOffset, -XOffset), 3, vbGreen
                        frmVisu.VisuDXF.Line (ZOffset - 6, -XOffset)-(ZOffset + 6, -XOffset), vbGreen
                        frmVisu.VisuDXF.Line (ZOffset, -XOffset - 6)-(ZOffset, -XOffset + 6), vbGreen
                    End If
            Else
                    MsgBox "Referencie o sistema no eixo principal da peça!"
            End If

            
                
        Else
            MsgBox ("Entidade inválida! Utilize uma linha reta!")
        End If
        
        
    End If
    
End Sub

Private Sub cmdZoomIn_Click()
    VisuDXF.ScaleHeight = 0.75 * VisuDXF.ScaleHeight
    VisuDXF.ScaleWidth = 0.75 * VisuDXF.ScaleWidth

    ReAfficheDXF
        frmVisu.VisuDXF.Line (-10000, 0)-(10000, 0), vbWhite
        frmVisu.VisuDXF.Line (0, -10000)-(0, 10000), vbWhite
    
                    'Redesenha zero
            If frmVisu.SistemaReferenciado Then
                frmVisu.VisuDXF.Circle (ZOffset, -XOffset), 3, vbGreen
                frmVisu.VisuDXF.Line (ZOffset - 6, -XOffset)-(ZOffset + 6, -XOffset), vbGreen
                frmVisu.VisuDXF.Line (ZOffset, -XOffset - 6)-(ZOffset, -XOffset + 6), vbGreen
            End If
End Sub

Private Sub cmdZoomOut_Click()
    VisuDXF.ScaleHeight = 1.25 * VisuDXF.ScaleHeight
    VisuDXF.ScaleWidth = 1.25 * VisuDXF.ScaleWidth

    ReAfficheDXF
            frmVisu.VisuDXF.Line (-10000, 0)-(10000, 0), vbWhite
            frmVisu.VisuDXF.Line (0, -10000)-(0, 10000), vbWhite
    
                'Redesenha zero
            If frmVisu.SistemaReferenciado Then
                frmVisu.VisuDXF.Circle (ZOffset, -XOffset), 3, vbGreen
                frmVisu.VisuDXF.Line (ZOffset - 6, -XOffset)-(ZOffset + 6, -XOffset), vbGreen
                frmVisu.VisuDXF.Line (ZOffset, -XOffset - 6)-(ZOffset, -XOffset + 6), vbGreen
            End If
End Sub

Private Sub Command1_Click()
    Dim indlst As Integer
    Dim contlst As String
    
    indlst = lst_codcnc.ListIndex
    contlst = lst_codcnc.Text
    
    If Not indlst <= 0 Then
        lst_codcnc.RemoveItem (indlst)
        lst_codcnc.AddItem contlst, indlst - 1
        lst_codcnc.ListIndex = indlst - 1
    End If
End Sub

Private Sub Command10_Click()
    If IsNumeric(Text1.Text) Then
        If Not (Text1.Text < 1) And Not (Text1.Text > 12) Then
            frmVisu.Ferramenta = CStr(Round(Text1.Text, 0))
            Command7.Enabled = True
            Label1.Enabled = False
            Text1.Enabled = False
            Command10.Enabled = False
        Else
            MsgBox "Insira um valor de 1 a 12 para a ferramenta!"
        End If
    End If
End Sub

Private Sub Command2_Click()
    Dim indlst As Integer
    Dim contlst As String
    Dim n As Integer
    
    n = lst_codcnc.ListCount
    indlst = lst_codcnc.ListIndex
    contlst = lst_codcnc.Text
    
    If Not indlst >= n - 1 Then
        lst_codcnc.RemoveItem (indlst)
        lst_codcnc.AddItem contlst, indlst + 1
        lst_codcnc.ListIndex = indlst + 1
    End If
End Sub

Private Sub Command3_Click()
        ListIndex = frmVisu.lst_codcnc.ListIndex
    'If ListIndex = -1 Then
    
        txtinsere = "/Fim de Programa"
        'txtinsere = Replace(txtinsere, ",", ".")
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        
        
        txtinsere = "M30"
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        'frmVisu.lst_codcnc.ListIndex = ListIndex + 1
        txtinsere = ""
        
        'Command9.Enabled = True
        
End Sub

   Sub Resize_For_Resolution(ByVal SFX As Single, ByVal SFY As Single, MyForm As Form)
      
      Dim i As Integer
      Dim SFFont As Single

      SFFont = (SFX + SFY) / 2  ' escala média
      ' Tamanho dos controles para a nova resolução
      On Error Resume Next  '
      With MyForm
        For i = 0 To .Count - 1
         If TypeOf .Controls(i) Is ComboBox Then 'Combobox não altera a propriedade Height
           .Controls(i).Left = .Controls(i).Left * SFX
           .Controls(i).Top = .Controls(i).Top * SFY
           .Controls(i).Width = .Controls(i).Width * SFX
         Else
           .Controls(i).Move .Controls(i).Left * SFX, _
            .Controls(i).Top * SFY, .Controls(i).Width * SFX, .Controls(i).Height * SFY
         End If
           ' Redimensiona e reposiciona antes de alterar o tamanho da fonte
           .Controls(i).FontSize = .Controls(i).FontSize * SFFont
        Next i
        'If RePosForm Then
          ' Redimensiona o formulario
          '.Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
        'End If
      End With

End Sub

Private Sub Command4_Click()
Dim texto As String
Dim txtzero As String
Dim poschar As Integer
Dim tipo As String
Dim pZ As Double
Dim sZ As Double
Dim pX As Double
Dim sX As Double
Dim raio As Double
Dim Angulo1 As Double
Dim Angulo2 As Double
Dim Temp As Double

        
    If lst_entidades.ListCount = 0 Then
        MsgBox "Carregue um arquivo DXF em Arquivo -> Abrir DXF..."
        
    ElseIf SistemaReferenciado = False Then
        MsgBox "É necessário referenciar o sistema!"
    ElseIf ParametrosDefinidos = False Then
        MsgBox "É necessário definir os parâmetros básicos!"
        ent_param.Enabled = True
        ent_param.Visible = True
    Else
        
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
            
                pX = pX - XOffset
                sX = sX - XOffset
                pZ = pZ - ZOffset
                sZ = sZ - ZOffset
                
                If pX > sX Then
                    Temp = sX
                    sX = pX
                    pX = Temp
                End If
                If pZ < sZ Then
                    Temp = sZ
                    sZ = pZ
                    pZ = Temp
                End If
            
                If Round(pZ, 0) = Round(frmVisu.UltimoZ, 0) And Round(pX, 0) = Round(UltimoX, 0) Then
                    txtinsere = "G01 X" & CDbl(2 * Round(sX, 4)) & " Z" & CDbl(Round(sZ, 4)) & " F" & frmVisu.VelAvancoAcab
                    txtinsere = Replace(txtinsere, ",", ".")
                    frmVisu.lst_codcnc.AddItem txtinsere
                    frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
                   ' txtinsere =
                    frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
                    
                    btn_proxent_Click
                    
                    UltimoX = sX
                    UltimoZ = sZ
                Else
                    MsgBox "Esta entidade não está ligada à última inserida. Selecione a entidade subsequente!"
                End If
                
    
            ElseIf tipo = "ARC" Then
                    
                    poschar = InStr(1, txtzero, ";", vbTextCompare)
                    pZ = CDbl(Mid(txtzero, 1, (poschar - 1)))
                    txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
                    poschar = InStr(1, txtzero, ";", vbTextCompare)
                    pX = CDbl(Mid(txtzero, 1, (poschar - 1)))
                    txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
                    poschar = InStr(1, txtzero, ";", vbTextCompare)
                    raio = CDbl(Mid(txtzero, 1, (poschar - 1)))
                    txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
                    poschar = InStr(1, txtzero, ";", vbTextCompare)
                    Angulo1 = CDbl(Mid(txtzero, 1, (poschar - 1)))
                    txtzero = Mid(txtzero, poschar + 1, Len(txtzero))
                    Angulo2 = CDbl(Mid(txtzero, 1, (poschar)))
                    
                    pX = pX - XOffset
                    pZ = pZ - ZOffset
                    
                    If Angulo1 = 0 And Angulo2 = 90 Then
                        'G03
                        txtinsere = "G03 X" & CStr(2 * (pX + raio)) & " Z" & CStr(pZ) & " R" & CStr(raio)
                        frmVisu.lst_codcnc.AddItem txtinsere
                        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
                        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
                        UltimoX = (pX + raio)
                        UltimoZ = pZ
                        btn_proxent_Click
                    Else
                        'G02
                        txtinsere = "G02 X" & CStr(2 * (pX)) & " Z" & CStr(pZ - raio) & " R" & CStr(raio)
                        frmVisu.lst_codcnc.AddItem txtinsere
                        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
                        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
                        UltimoX = pX
                        UltimoZ = pZ - raio
                        btn_proxent_Click
                    End If
                                
            End If
    End If
End Sub

Private Sub Command5_Click()
Dim txtinsere As String
            frmVisu.Command4.Enabled = False
            frmVisu.Command5.Enabled = False
    
            frmVisu.btn_veloc.Enabled = True
            frmVisu.btn_zerar.Enabled = True
            frmVisu.btn_cabecalho.Enabled = True
            frmVisu.btn_desbastar.Enabled = True
            frmVisu.btn_acabar.Enabled = True
            frmVisu.btn_rotacaoligar.Enabled = True
            frmVisu.btn_rotacaodesligar.Enabled = True
            frmVisu.btn_fluidoligar.Enabled = True
            frmVisu.btn_fluidodesligar.Enabled = True
            frmVisu.btn_delay.Enabled = True
            frmVisu.Command3.Enabled = True
            
            
            
            txtinsere = "G71 P" & CStr((frmVisu.PosG71 + 2) * 10) & " Q" & CStr((frmVisu.lst_codcnc.ListCount + 4) * 10) & " U" & CStr(frmVisu.Sobremetal) & " W" & CStr(frmVisu.Sobremetal) & " F" & CStr(frmVisu.VelAvanco)
            txtinsere = Replace(txtinsere, ",", ".")
            frmVisu.lst_codcnc.AddItem txtinsere, (frmVisu.PosG71)
            frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
            frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
            
            txtinsere = "G01 X0 Z" & CStr(frmVisu.ComprimentoBruto - frmVisu.ComprimentoFinal + 1) & " F" & CStr(frmVisu.VelAvanco)
            txtinsere = Replace(txtinsere, ",", ".")
            frmVisu.lst_codcnc.AddItem txtinsere, (frmVisu.PosG71 + 1)
            frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
            frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
            
            txtinsere = "G01 X0" & " Z" & CStr(frmVisu.ComprimentoBruto - frmVisu.ComprimentoFinal) & " F" & CStr(frmVisu.VelAvanco)
            txtinsere = Replace(txtinsere, ",", ".")
            frmVisu.lst_codcnc.AddItem txtinsere, (frmVisu.PosG71 + 2)
            frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
            frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
            
            txtinsere = "G01 X0 Z0 F" & CStr(frmVisu.VelAvancoAcab)
            txtinsere = Replace(txtinsere, ",", ".")
            frmVisu.lst_codcnc.AddItem txtinsere, (frmVisu.PosG71 + 3)
            frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
            frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
            
            frmVisu.UltimoX = 0
            frmVisu.UltimoZ = 0
    
End Sub

Private Sub Command6_Click()

    ListIndex = frmVisu.lst_codcnc.ListIndex
    If ListIndex = -1 Then
        MsgBox "Selecione alguma linha!"
    Else
        txtinsere = InputBox("Realize as alterações no campo abaixo!", , frmVisu.lst_codcnc.Text)
        If Not txtinsere = frmVisu.lst_codcnc.Text Then
            If txtinsere = "" Then
                txtinsere = frmVisu.lst_codcnc.Text
            End If
            frmVisu.lst_codcnc.RemoveItem ListIndex
            frmVisu.lst_codcnc.AddItem txtinsere, ListIndex
        End If
        
    End If

End Sub

Private Sub Command7_Click()
Dim inicio As Integer

        If frmVisu.lst_n.ListIndex = -1 Then
            MsgBox "Selecione uma linha na lista à esquerda (N---)."
        Else
            inicio = (frmVisu.lst_n.ListIndex + 1) * 10
            If inicio < 0 Then
                Exit Sub
            Else
                frmVisu.InAcab = inicio
            End If
            
            frmVisu.Command8.Enabled = True
            frmVisu.Command7.Enabled = False
        End If
End Sub

Private Sub Command8_Click()
Dim final As Integer

        
        frmVisu.btn_veloc.Enabled = True
        frmVisu.btn_zerar.Enabled = True
        frmVisu.btn_cabecalho.Enabled = True
        frmVisu.btn_desbastar.Enabled = True
        frmVisu.btn_acabar.Enabled = True
        frmVisu.btn_rotacaoligar.Enabled = True
        frmVisu.btn_rotacaodesligar.Enabled = True
        frmVisu.btn_fluidoligar.Enabled = True
        frmVisu.btn_fluidodesligar.Enabled = True
        frmVisu.btn_delay.Enabled = True
        frmVisu.Command3.Enabled = True
        
        final = (frmVisu.lst_n.ListIndex + 1) * 10
        If final < 0 Then
            Exit Sub
        Else
            frmVisu.FiAcab = final
            If frmVisu.InAcab >= frmVisu.FiAcab Then
                'frmVisu.InAcab = 0
                frmVisu.FiAcab = 0
                MsgBox "Selecione linha posterior à linha para Início do Acabamento!"
                Exit Sub
            End If
        End If
        
        txtinsere = "/Ciclo de Acabamento"
        'txtinsere = Replace(txtinsere, ",", ".")
        
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        
        txtinsere = "G00 X350 Z250 T00"
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        
        If frmVisu.Ferramenta < 10 Then
            txtinsere = "T0" & CStr(frmVisu.Ferramenta) & "0" & CStr(frmVisu.Ferramenta)
            frmVisu.lst_codcnc.AddItem txtinsere
            frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
            frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        Else
            txtinsere = "T" & CStr(frmVisu.Ferramenta) & CStr(frmVisu.Ferramenta)
            frmVisu.lst_codcnc.AddItem txtinsere
            frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
            frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        End If

        
        txtinsere = "G00 X" & CStr(frmVisu.DiametroBruto) & " Z" & CStr(frmVisu.ComprimentoBruto - frmVisu.ComprimentoFinal + 1)
        txtinsere = Replace(txtinsere, ",", ".")
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        
        txtinsere = "G70 P" & CStr(frmVisu.InAcab) & " Q" & CStr(frmVisu.FiAcab)
        'txtinsere = Replace(txtinsere, ",", ".")
        frmVisu.lst_codcnc.AddItem txtinsere
        frmVisu.NumLinhas = frmVisu.lst_codcnc.ListCount
        frmVisu.lst_n.AddItem "N" & CStr(frmVisu.NumLinhas * 10)
        
        frmVisu.InAcab = 0
        frmVisu.FiAcab = 0
        
        frmVisu.Command7.Enabled = False
        frmVisu.Command7.Visible = False
        frmVisu.Command8.Enabled = False
        frmVisu.Command8.Visible = False
        Label1.Enabled = False
        Label1.Visible = False
        Text1.Enabled = False
        Text1.Visible = False
        Command10.Enabled = False
        Command10.Visible = False
        

End Sub

Private Sub Command9_Click()
            
    tempos.Visible = True
    tempos.Enabled = True

        
End Sub

Private Sub Form_Load()
    FlagProgramando = False
    Ferramenta = 1
    UltimoX = 0
    UltimoZ = 0
    frmVisu.NumLinhas = 0
    
End Sub

Private Sub lst_entidades_Click()


    
    If Plot Then
    
    If Sele Then
        frmVisu.VisuDXF.Cls

                    ReAfficheDXF
                    frmVisu.VisuDXF.Line (-10000, 0)-(10000, 0), vbWhite
                    frmVisu.VisuDXF.Line (0, -10000)-(0, 10000), vbWhite
                    
                    'Redesenha zero
                    If frmVisu.SistemaReferenciado Then
                        frmVisu.VisuDXF.Circle (ZOffset, -XOffset), 3, vbGreen
                        frmVisu.VisuDXF.Line (ZOffset - 6, -XOffset)-(ZOffset + 6, -XOffset), vbGreen
                        frmVisu.VisuDXF.Line (ZOffset, -XOffset - 6)-(ZOffset, -XOffset + 6), vbGreen
                    End If


    Else
    
    End If
    
    Sele = True
    
    ListIndex = frmVisu.lst_entidades.ListIndex
    Entite = frmVisu.lst_entidades.Text
    Position = InStr(1, Entite, ";", vbTextCompare)
    
    If Position = 0 Then
        Sel = Entite
    Else
        Sel = Mid(Entite, 1, Position - 1)
    End If
    
    entrada.dr = Sel
    
    Select Case Sel
    
    Case "LINE"
    
    Lenght = Len(Entite)
    Entite = Mid(Entite, Position + 1, Lenght)
    
    Position = InStr(1, Entite, ";", vbTextCompare)
    X1 = CDbl(Mid(Entite, 1, Position - 1))
    Lenght = Len(Entite)
    Entite = Mid(Entite, Position + 1, Lenght)
    
    Position = InStr(1, Entite, ";", vbTextCompare)
    Y1 = CDbl(Mid(Entite, 1, Position - 1))
    Lenght = Len(Entite)
    Entite = Mid(Entite, Position + 1, Lenght)
    
    Position = InStr(1, Entite, ";", vbTextCompare)
    X2 = CDbl(Mid(Entite, 1, Position - 1))
    Lenght = Len(Entite)
    Entite = Mid(Entite, Position + 1, Lenght)
    
    Y2 = CDbl(Entite)
    Lenght = Len(Entite)
    
    frmVisu.VisuDXF.DrawWidth = 5
    frmVisu.VisuDXF.Line (X1, -Y1)-(X2, -Y2), 16777215
    frmVisu.VisuDXF.DrawWidth = 1
    
    Case "ARC"
    
    Lenght = Len(Entite)
    Entite = Mid(Entite, Position + 1, Lenght)
    
    Position = InStr(1, Entite, ";", vbTextCompare)
    X1 = CDbl(Mid(Entite, 1, Position - 1))
    Lenght = Len(Entite)
    Entite = Mid(Entite, Position + 1, Lenght)
    
    Position = InStr(1, Entite, ";", vbTextCompare)
    Y1 = CDbl(Mid(Entite, 1, Position - 1))
    Lenght = Len(Entite)
    Entite = Mid(Entite, Position + 1, Lenght)
    
    Position = InStr(1, Entite, ";", vbTextCompare)
    rad = CDbl(Mid(Entite, 1, Position - 1))
    Lenght = Len(Entite)
    Entite = Mid(Entite, Position + 1, Lenght)
    
    Position = InStr(1, Entite, ";", vbTextCompare)
    Angle1 = CDbl(Mid(Entite, 1, Position - 1))
    Lenght = Len(Entite)
    Entite = Mid(Entite, Position + 1, Lenght)
    
    Angle2 = CDbl(Entite)
    Lenght = Len(Entite)
    
    'frmVisu.VisuDXF.DrawWidth = 5
    'frmVisu.VisuDXF.Line (X1, -Y1)-(X2, -Y2), 16777215
    'frmVisu.VisuDXF.DrawWidth = 1
    
    AfficheDXFArcX X1, Y1, rad, Angle1, Angle2, 16777215
    
    Case Else
    
    End Select
    
    
    
    
    End If
    
End Sub


'Private Sub optMouse_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Select Case Index
'
'Case 0
'    If SvgDeplace Then
'      optMouse(0).Value = False
'      SvgDeplace = False
'    Else
'     SvgDeplace = True
'    End If
'
'
'Case 1
'    If SvgZoom Then
'      optMouse(1).Value = False
'      SvgZoom = False
'    Else
'     SvgZoom = True
'    End If
'End Select
'
'End Sub
'
'Private Sub optMouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.Caption = "DXF Visualisateur -- (" & Format(X, "0.000") & "," & Format(-Y, "0.000") & ")"
'If Deplace Then
'    VisuDXF.ScaleTop = VisuDXF.ScaleTop + (DragY - Y)
'    VisuDXF.ScaleLeft = VisuDXF.ScaleLeft + (DragX - X)
'    VisuDXF.Cls
'    VisuDXF.Picture = LoadPicture()
'    ReAfficheDXF
'    Exit Sub
'End If
'' Mode zoom dessine une fenetre matérialisant la zone de Zoom
'If Zoom Then
'    VisuDXF.DrawMode = 6
'    VisuDXF.DrawStyle = 1
'    VisuDXF.DrawWidth = 1
'    VisuDXF.Line (SelGroup.X1, SelGroup.Y1)-(SelGroup.X2, SelGroup.Y2), vbBlack, B
'    VisuDXF.Line (SelGroup.X1, SelGroup.Y1)-(X, Y), vbBlack, B
'    SelGroup.X2 = X
'    SelGroup.Y2 = Y
'End If
'End Sub

Private Sub sair_Click(Index As Integer)
    End
End Sub

Private Sub salvar_Click(Index As Integer)


         Dim OpenFile As OPENFILENAME
         Dim lReturn As Long
         Dim sFilter As String
         OpenFile.lStructSize = Len(OpenFile)
         OpenFile.hwndOwner = frmVisu.hWnd
         OpenFile.hInstance = App.hInstance
         sFilter = "Batch Files (*.nc)" & Chr(0) & "*.nc" & Chr(0)
         OpenFile.lpstrFilter = sFilter
         OpenFile.nFilterIndex = 1
            OpenFile.lpstrFile = String(277, 0)


         OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
         OpenFile.lpstrFileTitle = OpenFile.lpstrFile
         OpenFile.nMaxFileTitle = 255
         OpenFile.lpstrInitialDir = "C:\"
         OpenFile.flags = 0

        OpenFile.lpstrTitle = "Salvar CNC"
          lReturn = GetSaveFileName(OpenFile)
         If lReturn = 0 Then
         Else
        On Error GoTo trataerro
            SaveName = OpenFile.lpstrFile
            cont = String(1, 0)
            pos = InStr(1, SaveName, cont, vbTextCompare)
            
            
            SaveName = Mid(SaveName, 1, pos - 1)
            If Mid(SaveName, (Len(SaveName) - 3), 1) = "." Then
            Else
                SaveName = SaveName + ".nc"
            End If
            Set arqtxt = fso.CreateTextFile(SaveName, True)
            'gravando no arquivo
            
            With arqtxt
                If frmVisu.NumeroPrograma < 10 Then
                    .WriteLine ("O000" & CStr(frmVisu.NumeroPrograma) & " (" & frmVisu.NomePrograma & ")")
                ElseIf frmVisu.NumeroPrograma < 100 Then
                    .WriteLine ("O00" & CStr(frmVisu.NumeroPrograma) & " (" & frmVisu.NomePrograma & ")")
                ElseIf frmVisu.NumeroPrograma < 1000 Then
                    .WriteLine ("O0" & CStr(frmVisu.NumeroPrograma) & " (" & frmVisu.NomePrograma & ")")
                Else
                    .WriteLine ("O" & CStr(frmVisu.NumeroPrograma) & " (" & frmVisu.NomePrograma & ")")
                End If
                ListCount = frmVisu.lst_codcnc.ListCount
                For i = 0 To (ListCount - 1)
                    TempString = "N" & CStr((i + 1) * 10) & " " & frmVisu.lst_codcnc.List(i) & ";"
                    .WriteLine (TempString)
                Next
                
                .Close
            End With
            Exit Sub
            
trataerro:
            MsgBox Err.Description & " - " & Err.Number, vbCritical
         End If
    


End Sub


'Recherche les valeurs XYZ max pour un centrage auto
Private Sub XYZmax(PoMini As Point2, PoMaxi As Point2, ByVal cX As Double, ByVal cY As Double, ByVal EchelleX As Double, ByVal EchelleY As Double, ByVal Angle As Double)
On Error Resume Next
Dim Couleur As Long
Dim Depart As Long
Dim X1 As Double
Dim Y1 As Double
Dim X2 As Double
Dim Y2 As Double
Dim X3 As Double
Dim Y3 As Double
Dim Angle1 As Double
Dim Angle2 As Double
Dim Angle3 As Double
Dim Ratio As Double
Dim rad As Double
Dim PCount As Integer
Dim Text As String
Dim Size As Double
Dim Nom As String
Dim EndPoly As Boolean

div = 1000000

    PoMaxi.X = -INFINITY
    PoMaxi.Y = -INFINITY
    
    PoMini.X = INFINITY
    PoMini.Y = INFINITY


For Depart = 0 To UBound(BdDXF.Entite)

'If Depart = 0 Then
'    Aux = 0
'End If

Select Case BdDXF.Entite(Depart).Type
    Case "LINE"
        'Récupère les valeurs
        
        'If CDbl(BdDXF.Entite(Depart).Donnee(0).Valeur) - Round(CDbl(BdDXF.Entite(Depart).Donnee(0).Valeur), 3)
        X1 = CDbl(Replace(BdDXF.Entite(Depart).Donnee(0).Valeur, ".", ","))
        Y1 = CDbl(Replace(BdDXF.Entite(Depart).Donnee(1).Valeur, ".", ","))
        X2 = CDbl(Replace(BdDXF.Entite(Depart).Donnee(2).Valeur, ".", ","))
        Y2 = CDbl(Replace(BdDXF.Entite(Depart).Donnee(3).Valeur, ".", ","))
        'Facteur d'échelle  des entitées selon leur origine
        X1 = X1 * EchelleX
        Y1 = Y1 * EchelleY
        X2 = X2 * EchelleX
        Y2 = Y2 * EchelleY

        
        'Rotation des entitées selon leur origine
        If Angle <> 0 Then
            X3 = RotationX(X1, Y1, Angle)
            Y3 = RotationY(X1, Y1, Angle)
            X1 = X3
            Y1 = Y3
            X3 = RotationX(X2, Y2, Angle)
            Y3 = RotationY(X2, Y2, Angle)
            X2 = X3
            Y2 = Y3
        End If
        'Déplace l'origine
        X1 = X1 + cX
        Y1 = Y1 + cY
        X2 = X2 + cX
        Y2 = Y2 + cY
        
        'Test le Maxi et mini pour les zooms
        Call minmax(PoMini, PoMaxi, X1, Y1)
        Call minmax(PoMini, PoMaxi, X2, Y2)
        
        If Import Then
        frmVisu.lst_entidades.AddItem "LINE" + ";" + CStr(Round(X1, 4)) + ";" + CStr(Round(Y1, 4)) + ";" + CStr(Round(X2, 4)) + ";" + CStr(Round(Y2, 4)), Depart
'        Aux = Aux + 1
        End If
        
       
    Case "ARC"
        ' Les cercles et les lignes deviennent automatiquement des ellipses quant on leur applique un facteur d'échelle
        X1 = Replace(BdDXF.Entite(Depart).Donnee(0).Valeur, ".", ",")
        Y1 = Replace(BdDXF.Entite(Depart).Donnee(1).Valeur, ".", ",")
        rad = Replace(BdDXF.Entite(Depart).Donnee(2).Valeur, ".", ",")
        Angle1 = Replace(BdDXF.Entite(Depart).Donnee(3).Valeur, ".", ",")
        Angle2 = Replace(BdDXF.Entite(Depart).Donnee(4).Valeur, ".", ",")
        X1 = X1 * EchelleX
        Y1 = Y1 * EchelleY
        ''Si le facteur d'echelle suivant X est différent de Y alors
        'L 'arc ou le cercle devient une ellipse
        If EchelleX <> 1 Then
            rad = rad * EchelleX
        ElseIf EchelleY <> 1 Then
            rad = rad * EchelleY
        End If
        If Angle <> 0 Then
            X3 = RotationX(X1, Y1, Angle)
            Y3 = RotationY(X1, Y1, Angle)
            X1 = X3
            Y1 = Y3
        End If
        If EchelleX < 0 Or EchelleY < 0 Then
            ' L'arc subie une transformation mirroir
            Intervertir Angle1, Angle2
            Angle1 = 180 - Angle1
            Angle2 = 180 - Angle2
        End If
        Angle1 = Angle1 + (Angle * 180 / PI)
        Angle2 = Angle2 + (Angle * 180 / PI)
        X1 = X1 + cX
        Y1 = Y1 + cY
        
        'Test le Maxi et mini pour les zooms
        Call minmax(PoMini, PoMaxi, X1, Y1)
        
               If Import Then
        If Angle2 = 0 Then
            Angle2 = 360000000
        End If
        TempString = "ARC" + ";" + CStr(Round(X1, 4)) + ";" + CStr(Round(Y1, 4)) + ";" + CStr(Round(rad, 4)) + ";" + CStr(Round(Angle1, 4)) + ";" + CStr(Round(Angle2, 4))
        frmVisu.lst_entidades.AddItem TempString, Depart
        'AfficheDXFArc Pict, X1, Y1, Abs(rad), Angle1, Angle2, 16777215
        'DesenhaArco X1 / 1000000, Y1 / 1000000, rad / 1000000, Angle1 / 1000000, Angle2 / 1000000
        End If
        
    Case "CIRCLE"
        ' Les cercles et les lignes deviennent automatiquement des ellipses quant on leur applique un facteur d'échelle
        X1 = Replace(BdDXF.Entite(Depart).Donnee(0).Valeur, ".", ",")
        Y1 = Replace(BdDXF.Entite(Depart).Donnee(1).Valeur, ".", ",")
        rad = Replace(BdDXF.Entite(Depart).Donnee(2).Valeur, ".", ",")
        X1 = X1 * EchelleX
        Y1 = Y1 * EchelleY
        If EchelleX <> 1 Then
            rad = rad * EchelleX
        ElseIf EchelleY <> 1 Then
            rad = rad * EchelleY
        End If
        If Angle <> 0 Then
            X3 = RotationX(X1, Y1, Angle)
            Y3 = RotationY(X1, Y1, Angle)
            X1 = X3
            Y1 = Y3
        End If
        X1 = X1 + cX
        Y1 = Y1 + cY
        
        'Test le Maxi et mini pour les zooms
        Call minmax(PoMini, PoMaxi, X1, Y1)
        
               If Import Then
        'frmVisu.lst_entidades.AddItem "-", Depart
        End If

    Case "ELLIPSE"
        X1 = BdDXF.Entite(Depart).Donnee(0).Valeur
        Y1 = BdDXF.Entite(Depart).Donnee(1).Valeur
        X2 = BdDXF.Entite(Depart).Donnee(2).Valeur
        Y2 = BdDXF.Entite(Depart).Donnee(3).Valeur
        Ratio = BdDXF.Entite(Depart).Donnee(4).Valeur
        Angle1 = BdDXF.Entite(Depart).Donnee(5).Valeur
        Angle2 = BdDXF.Entite(Depart).Donnee(6).Valeur
        X1 = X1 * EchelleX
        Y1 = Y1 * EchelleY
        X2 = X2 * EchelleX
        Y2 = Y2 * EchelleY
        If Angle <> 0 Then
            X3 = RotationX(X1, Y1, Angle)
            Y3 = RotationY(X1, Y1, Angle)
            X1 = X3
            Y1 = Y3
            X3 = RotationX(X2, Y2, Angle)
            Y3 = RotationY(X2, Y2, Angle)
            X2 = X3
            Y2 = Y3
        End If
        If EchelleX < 0 Or EchelleY < 0 Then Ratio = -Ratio ' L'ELLIPSE est inversée
        X1 = X1 + cX
        Y1 = Y1 + cY
        'Test le Maxi et mini pour les zooms
        Call minmax(PoMini, PoMaxi, X1, Y1)
        
               If Import Then
        'frmVisu.lst_entidades.AddItem "-", Depart
        End If
    
    Case "POLYLINE"
        ' Une POLYLINE est une suite de ligne liée entre elles
        PCount = 1
        EndPoly = False
        Do While Not EndPoly
            X1 = BdDXF.Entite(Depart + PCount).Donnee(0).Valeur
            Y1 = BdDXF.Entite(Depart + PCount).Donnee(1).Valeur
            X2 = BdDXF.Entite(Depart + PCount + 1).Donnee(0).Valeur
            Y2 = BdDXF.Entite(Depart + PCount + 1).Donnee(1).Valeur
            'Facteur d'échelle  des entitées selon leur origine
            X1 = X1 * EchelleX
            X2 = X2 * EchelleX
            Y1 = Y1 * EchelleY
            Y2 = Y2 * EchelleY
            'Rotation des entitées selon leur origine
            If Angle <> 0 Then
                X3 = RotationX(X1, Y1, Angle)
                Y3 = RotationY(X1, Y1, Angle)
                X1 = X3
                Y1 = Y3
                X3 = RotationX(X2, Y2, Angle)
                Y3 = RotationY(X2, Y2, Angle)
                X2 = X3
                Y2 = Y3
            End If
            'Déplace l'origine
            X1 = X1 + cX
            Y1 = Y1 + cY
            X2 = X2 + cX
            Y2 = Y2 + cY

        'Test le Maxi et mini pour les zooms
        Call minmax(PoMini, PoMaxi, X1, Y1)
        Call minmax(PoMini, PoMaxi, X2, Y2)
        
            PCount = PCount + 1
            If Depart + PCount + 1 > UBound(BdDXF.Entite) Then
                EndPoly = True
            ElseIf BdDXF.Entite(Depart + PCount + 1).Type <> "VERTEX" Then
                EndPoly = True
            End If
        Loop
        
               If Import Then
        'frmVisu.lst_entidades.AddItem "-", Depart
        End If
        
    Case "TEXT"
        'pas de facteur d'échelle pour le TEXTE
        X1 = BdDXF.Entite(Depart).Donnee(0).Valeur
        Y1 = BdDXF.Entite(Depart).Donnee(1).Valeur
        Size = BdDXF.Entite(Depart).Donnee(2).Valeur
        Angle1 = BdDXF.Entite(Depart).Donnee(3).Valeur + Angle
        Text = BdDXF.Entite(Depart).Donnee(4).Valeur
        'Déplace l'origine
        X1 = X1 + cX
        Y1 = Y1 + cY
        'Test le Maxi et mini pour les zooms
        Call minmax(PoMini, PoMaxi, X1, Y1)
        
               If Import Then
        'frmVisu.lst_entidades.AddItem "-", Depart
        End If
        
    Case "INSERT"
        'Just a note: Block can not be "Stretched" but if they are mirrored . . that
        'shows up in the "scale" Variable for Block
        Nom = BdDXF.Entite(Depart).Donnee(0).Valeur
        X1 = BdDXF.Entite(Depart).Donnee(1).Valeur
        Y1 = BdDXF.Entite(Depart).Donnee(2).Valeur
        X2 = BdDXF.Entite(Depart).Donnee(3).Valeur
        Y2 = BdDXF.Entite(Depart).Donnee(4).Valeur
        '"0" scale = scale of "1"
        If X2 = 0 Then X2 = 1
        If Y2 = 0 Then Y2 = 1
        Angle1 = BdDXF.Entite(Depart).Donnee(5).Valeur * PI / 180
        'Test le Maxi et mini pour les zooms
        Call minmax(PoMini, PoMaxi, X1, Y1)
        
               If Import Then
        'frmVisu.lst_entidades.AddItem "-", Depart
        End If
    
    Case "DIMENSION"
        'Just a note: Block can not be "Stretched" but if they are mirrored . . that
        'shows up in the "scale" Variable for Block
        Nom = BdDXF.Entite(Depart).Donnee(0).Valeur
        X1 = BdDXF.Entite(Depart).Donnee(1).Valeur
        Y1 = BdDXF.Entite(Depart).Donnee(2).Valeur
        'Test le Maxi et mini pour les zooms
        Call minmax(PoMini, PoMaxi, X1, Y1)
        
               If Import Then
        'frmVisu.lst_entidades.AddItem "-", Depart
        End If
        
    Case Else
               If Import Then
        'frmVisu.lst_entidades.AddItem "-", Depart
        End If
    
End Select

Next Depart

End Sub

Private Sub optMouse_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index

Case 0
    If SvgDeplace Then
      optMouse(0).Value = False
      SvgDeplace = False
    Else
     SvgDeplace = True
    End If


Case 1
    If SvgZoom Then
      optMouse(1).Value = False
      SvgZoom = False
    Else
     SvgZoom = True
    End If
End Select
End Sub


Sub ReAfficheDXF(Optional Layer As String)
' affichage dans le controle ''VisuDXF de la base de données ''VisuDXF

 AfficheDXF VisuDXF, BdDXF, Layer
End Sub

Private Sub minmax(PoMini As Point2, PoMaxi As Point2, X As Double, Y As Double)
        If X < PoMini.X Then PoMini.X = X
        If Y < PoMini.Y Then PoMini.Y = Y

        If X > PoMaxi.X Then PoMaxi.X = X
        If Y > PoMaxi.Y Then PoMaxi.Y = Y
End Sub


'Private Sub VisuDXF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    DragX = X
'    DragY = Y

    
'    If optMouse(0) Then Deplace = True
'    If optMouse(1) And Zoom = False Then
'        Zoom = True
'        SelGroup.X1 = X
'        SelGroup.Y1 = Y
'        SelGroup.X2 = X
'        SelGroup.Y2 = Y
'    End If
    
'End Sub

'Private Sub VisuDXF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''Affiche les coordonnées
'Me.Caption = "DXF Visualisateur -- (" & Format(X, "0.000") & "," & Format(-Y, "0.000") & ")"
'If Deplace Then
'    VisuDXF.ScaleTop = VisuDXF.ScaleTop + (DragY - Y)
'    VisuDXF.ScaleLeft = VisuDXF.ScaleLeft + (DragX - X)
'    VisuDXF.Cls
'    VisuDXF.Picture = LoadPicture()
'    ReAfficheDXF
'    Exit Sub
'End If
'' Mode zoom dessine une fenetre matérialisant la zone de Zoom
'If Zoom Then
'    VisuDXF.DrawMode = 6
'    VisuDXF.DrawStyle = 1
'    VisuDXF.DrawWidth = 1
'    VisuDXF.Line (SelGroup.X1, SelGroup.Y1)-(SelGroup.X2, SelGroup.Y2), vbBlack, B
'    VisuDXF.Line (SelGroup.X1, SelGroup.Y1)-(X, Y), vbBlack, B
'    SelGroup.X2 = X
'    SelGroup.Y2 = Y
'End If
'End Sub

'Private Sub VisuDXF_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim A As Double
'Dim B As Double
'Dim C As Double
'Dim D As Double

'DoEvents
'VisuDXF.DrawMode = 13
'VisuDXF.DrawStyle = 0
'VisuDXF.DrawWidth = 1
'If Zoom Then
'    If SelGroup.X2 < SelGroup.X1 Then Intervertir SelGroup.X1, SelGroup.X2
'    If SelGroup.Y2 < SelGroup.Y1 Then Intervertir SelGroup.Y1, SelGroup.Y2
'    SelGroup.Y2 = SelGroup.Y1 + Abs(SelGroup.X2 - SelGroup.X1)
'    If SelGroup.X2 = SelGroup.X1 Then
'        Exit Sub
'    End If
'    If SelGroup.Y2 = SelGroup.Y1 Then
'        Exit Sub
'    End If
    
'C = Abs(SelGroup.X2 - SelGroup.X1)
'D = Abs(SelGroup.Y1 - SelGroup.Y2)
'If C > D Then
'A = C
'Else
'A = D
'End If

'B = VisuDXF.Height / VisuDXF.Width

    'VisuDXF.ScaleWidth = A
    'VisuDXF.ScaleLeft = P1.X - 15
    'VisuDXF.ScaleHeight = A * B
    'VisuDXF.ScaleTop = -P2.Y - 15
    
'    VisuDXF.ScaleWidth = A 'Abs(SelGroup.X2 - SelGroup.X1)
'    VisuDXF.ScaleLeft = SelGroup.X1
'    VisuDXF.ScaleHeight = A * B ' Abs(SelGroup.Y1 - SelGroup.Y2)
'    VisuDXF.ScaleTop = SelGroup.Y1
'    ReAfficheDXF
'End If
'Deplace = False
'Zoom = False
'End Sub
'Public Function NegritoLinha(Index) As Boolean
'Dim Couleur As Long
'Dim Depart As Long
'Dim X1 As Double
'Dim Y1 As Double
'Dim X2 As Double
'Dim Y2 As Double
'Dim X3 As Double
'Dim Y3 As Double
'Dim Angle1 As Double
'Dim Angle2 As Double
'Dim Angle3 As Double
'Dim Ratio As Double
'Dim rad As Double
'Dim PCount As Integer
'Dim Text As String
'Dim Size As Double
'Dim Nom As String
'Dim EndPoly As Boolean
'
'    Select Case BdDXF.Entite(Index).Type
'
'    Case "LINE"
'        X1 = BdDXF.Entite(Index).Donnee(0).Valeur
'        Y1 = BdDXF.Entite(Index).Donnee(1).Valeur
'        X2 = BdDXF.Entite(Index).Donnee(2).Valeur
'        Y2 = BdDXF.Entite(Index).Donnee(3).Valeur
'        'Facteur d'échelle  des entitées selon leur origine
''        X1 = X1 * EchelleX
''        Y1 = Y1 * EchelleY
''        X2 = X2 * EchelleX
''        Y2 = Y2 * EchelleY
'        'Rotation des entitées selon leur origine
'        If Angle <> 0 Then
'            X3 = RotationX(X1, Y1, Angle)
'            Y3 = RotationY(X1, Y1, Angle)
'            X1 = X3
'            Y1 = Y3
'            X3 = RotationX(X2, Y2, Angle)
'            Y3 = RotationY(X2, Y2, Angle)
'            X2 = X3
'            Y2 = Y3
'        End If
'        'Déplace l'origine
'        X1 = X1 + cX
'        Y1 = Y1 + cY
'        X2 = X2 + cX
'        Y2 = Y2 + cY
'
'        'Test le Maxi et mini pour les zooms
'        Call minmax(PoMini, PoMaxi, X1, Y1)
'        Call minmax(PoMini, PoMaxi, X2, Y2)
'
'        AfficheDXFLigneX Pict, X1, Y1, X2, Y2, Couleur
'
'    Case "ARC"
'        X1 = BdDXF.Entite(Index).Donnee(0).Valeur
'        Y1 = BdDXF.Entite(Index).Donnee(1).Valeur
'        rad = BdDXF.Entite(Index).Donnee(2).Valeur
'        Angle1 = BdDXF.Entite(Index).Donnee(3).Valeur
'        Angle2 = BdDXF.Entite(Index).Donnee(4).Valeur
'        X1 = X1 * EchelleX
'        Y1 = Y1 * EchelleY
'        ''Si le facteur d'echelle suivant X est différent de Y alors
'        'L 'arc ou le cercle devient une ellipse
'        If EchelleX <> 1 Then
'            rad = rad * EchelleX
'        ElseIf EchelleY <> 1 Then
'            rad = rad * EchelleY
'        End If
'        If Angle <> 0 Then
'            X3 = RotationX(X1, Y1, Angle)
'            Y3 = RotationY(X1, Y1, Angle)
'            X1 = X3
'            Y1 = Y3
'        End If
'        If EchelleX < 0 Or EchelleY < 0 Then
'            ' L'arc subie une transformation mirroir
'            Intervertir Angle1, Angle2
'            Angle1 = 180 - Angle1
'            Angle2 = 180 - Angle2
'        End If
'        Angle1 = Angle1 + (Angle * 180 / PI)
'        Angle2 = Angle2 + (Angle * 180 / PI)
'        X1 = X1 + cX
'        Y1 = Y1 + cY
'
'        'Test le Maxi et mini pour les zooms
'        Call minmax(PoMini, PoMaxi, X1, Y1)
'
'    End Select
'End Function

Private Sub VisuDXF_DblClick()

Dim P0 As Point2
Dim P1 As Point2
Dim P2 As Point2


Dim A As Double
Dim B As Double
Dim C As Double
Dim D As Double
 
Deplace = False
Zoom = False


Call XYZmax(P1, P2, 0, 0, 1, 1, 0)


'P0.X = P1.X + (P2.X - P1.X) / 2
'P0.Y = P1.Y + (P2.Y - P1.Y) / 2


'A = Abs(P2.X - P1.X)
'B = Abs(P2.Y - P1.Y)

'A = Distance(P1, P2) '/ 2
'B = Distance(P1, P2) / 2
'C = Distance(P1, P0)
'D = Distance(P0, P2)

'VisuDXF.Scale (-C, -D)-(D, C)

C = Abs(P2.X - P1.X) + 30
D = Abs(P2.Y - P1.Y) + 30

If C > D Then
A = C
Else
A = D
End If

B = VisuDXF.Height / VisuDXF.Width

    VisuDXF.ScaleWidth = A
    VisuDXF.ScaleLeft = P1.X - 15
    VisuDXF.ScaleHeight = A * B
    VisuDXF.ScaleTop = -P2.Y - 15
    'Debug.Print B
    'Debug.Print VisuDXF.ScaleHeight / VisuDXF.ScaleWidth
        
optMouse(0).Value = False
optMouse(1).Value = False


ReAfficheDXF
    frmVisu.VisuDXF.Line (-10000, 0)-(10000, 0), vbWhite
    frmVisu.VisuDXF.Line (0, -10000)-(0, 10000), vbWhite
    
    'Redesenha zero
    If frmVisu.SistemaReferenciado Then
        frmVisu.VisuDXF.Circle (ZOffset, -XOffset), 5, vbWhite
    End If


End Sub

Private Sub VisuDxf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragX = X
    DragY = Y

    
    If optMouse(0) Then
        Deplace = True
    ElseIf optMouse(1) And Zoom = False Then
        Zoom = True
        SelGroup.X1 = X
        SelGroup.Y1 = Y
        SelGroup.X2 = X
        SelGroup.Y2 = Y
    End If
    
End Sub

Private Sub VisuDXF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Deplace Then
    VisuDXF.ScaleTop = VisuDXF.ScaleTop + (DragY - Y)
    VisuDXF.ScaleLeft = VisuDXF.ScaleLeft + (DragX - X)
    VisuDXF.Cls
    VisuDXF.Picture = LoadPicture()
    ReAfficheDXF
        frmVisu.VisuDXF.Line (-10000, 0)-(10000, 0), vbWhite
    frmVisu.VisuDXF.Line (0, -10000)-(0, 10000), vbWhite
    Exit Sub
End If
' Mode zoom dessine une fenetre matérialisant la zone de Zoom
If Zoom Then
    VisuDXF.DrawMode = 6
    VisuDXF.DrawStyle = 1
    VisuDXF.DrawWidth = 1
    frmVisu.VisuDXF.Line (-10000, 0)-(10000, 0), vbWhite
    frmVisu.VisuDXF.Line (0, -10000)-(0, 10000), vbWhite
    VisuDXF.Line (SelGroup.X1, SelGroup.Y1)-(SelGroup.X2, SelGroup.Y2), vbBlack, B
    VisuDXF.Line (SelGroup.X1, SelGroup.Y1)-(X, Y), vbBlack, B
    SelGroup.X2 = X
    SelGroup.Y2 = Y
End If

    'Redesenha zero
    If frmVisu.SistemaReferenciado Then
        frmVisu.VisuDXF.Circle (ZOffset, -XOffset), 3, vbGreen
        frmVisu.VisuDXF.Line (ZOffset - 6, -XOffset)-(ZOffset + 6, -XOffset), vbGreen
        frmVisu.VisuDXF.Line (ZOffset, -XOffset - 6)-(ZOffset, -XOffset + 6), vbGreen
    End If

End Sub

Private Sub VisuDXF_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim A As Double
Dim B As Double
Dim C As Double
Dim D As Double

DoEvents
VisuDXF.DrawMode = 13
VisuDXF.DrawStyle = 0
VisuDXF.DrawWidth = 1
If Zoom Then
    If SelGroup.X2 < SelGroup.X1 Then Intervertir SelGroup.X1, SelGroup.X2
    If SelGroup.Y2 < SelGroup.Y1 Then Intervertir SelGroup.Y1, SelGroup.Y2
    SelGroup.Y2 = SelGroup.Y1 + Abs(SelGroup.X2 - SelGroup.X1)
    If SelGroup.X2 = SelGroup.X1 Then
        Exit Sub
    End If
    If SelGroup.Y2 = SelGroup.Y1 Then
        Exit Sub
    End If
    
C = Abs(SelGroup.X2 - SelGroup.X1)
D = Abs(SelGroup.Y1 - SelGroup.Y2)
If C > D Then
A = C
Else
A = D
End If

B = VisuDXF.Height / VisuDXF.Width

    'VisuDXF.ScaleWidth = A
    'VisuDXF.ScaleLeft = P1.X - 15
    'VisuDXF.ScaleHeight = A * B
    'VisuDXF.ScaleTop = -P2.Y - 15
    
    VisuDXF.ScaleWidth = A 'Abs(SelGroup.X2 - SelGroup.X1)
    VisuDXF.ScaleLeft = SelGroup.X1
    VisuDXF.ScaleHeight = A * B ' Abs(SelGroup.Y1 - SelGroup.Y2)
    VisuDXF.ScaleTop = SelGroup.Y1

    ReAfficheDXF
        frmVisu.VisuDXF.Line (-10000, 0)-(10000, 0), vbWhite
    frmVisu.VisuDXF.Line (0, -10000)-(0, 10000), vbWhite
End If
Deplace = False
Zoom = False



End Sub

Public Function DesenhaArco(Xc, Yc, rad, Angle1, Angle2)

    Dim X1 As Double
    Dim Y1 As Double
    Dim X2 As Double
    Dim Y2 As Double
    Dim Xf As Double
    Dim Yf As Double
    Dim radConv As Double
    Dim constScale As Double
    Dim cosa As Double
    Dim sena As Double
    
    
    Dim i As Integer
    
    
    Dim P0 As Point2
    Dim P1 As Point2
    Dim P2 As Point2
    
    
    Dim A As Double
    Dim B As Double
    Dim C As Double
    Dim D As Double

    
    
'    Call XYZmax(P1, P2, 0, 0, 1, 1, 0)
    
    
    'P0.X = P1.X + (P2.X - P1.X) / 2
    'P0.Y = P1.Y + (P2.Y - P1.Y) / 2
    
    
    'A = Abs(P2.X - P1.X)
    'B = Abs(P2.Y - P1.Y)
    
    'A = Distance(P1, P2) '/ 2
    'B = Distance(P1, P2) / 2
    'C = Distance(P1, P0)
    'D = Distance(P0, P2)
    
    'VisuDXF.Scale (-C, -D)-(D, C)
    
    C = Abs(P2.X - P1.X) + 30
    D = Abs(P2.Y - P1.Y) + 30
    
    If C > D Then
    A = C
    Else
    A = D
    End If
    
    B = VisuDXF.Height / VisuDXF.Width

    VisuDXF.ScaleWidth = A
    VisuDXF.ScaleLeft = P1.X - 15
    VisuDXF.ScaleHeight = A * B
    VisuDXF.ScaleTop = -P2.Y - 15
    
    
    constScale = 1000000
    radConv = Round(3.141592654 / 180, 7)
    
    cosa = Cos(Angle1 * radConv)
    X1 = Xc + rad * cosa
    sena = Sin(Angle1 * radConv)
    Y1 = Yc + rad * sena
    Xf = Xc + rad * Cos(Angle2 * radConv)
    Yf = Yc + rad * Sin(Angle2 * radConv)

    For i = 1 To 360
        
        If Angle1 * radConv + i * radConv > Angle2 * radConv Then
            frmVisu.VisuDXF.Line (X1, -Y1)-(Xf, -Yf), vbWhite
            Exit For
        Else
            cosa = Round(Cos((Angle1 + i) * radConv), 5)
            sena = Round(Sin((Angle1 + i) * radConv), 5)
            X2 = Xc + rad * cosa
            Y2 = Yc + rad * sena
            
            X1 = Round(X1 * constScale, 0)
            X2 = Round(X2 * constScale, 0)
            Y1 = Round(Y1 * constScale, 0)
            Y2 = Round(Y2 * constScale, 0)
            
            VisuDXF.DrawWidth = 1
            
            VisuDXF.Line (X1, -Y1)-(X2, -Y2), vbWhite 'DANDO ERRO
            'AfficheDXFLigne VisuDXF, X1, Y1, X2, Y2, vbWhite
            
            X1 = X2
            Y1 = Y2
        End If
    
    Next
End Function
