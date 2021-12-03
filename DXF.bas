Attribute VB_Name = "DXF"
Option Explicit

'Global variables

Public Const PI = 3.14159265358979
Public Const INFINITY As Double = 9E+99

'---------------------------------------------------------------------------
' Point en 2D
'---------------------------------------------------------------------------
Public Type Point2
    X As Double
    Y As Double
End Type


Type RECT
    X1 As Double
    Y1 As Double
    X2 As Double
    Y2 As Double
End Type

Type cadpoint
    X As Double
    Y As Double
    Z As Double
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceNom(LF_FACESIZE)
  lfFaceNom As String * 33
End Type

'Variable DXF
 ' Constituee d'un numero de Cle
 ' Et de sa valeur de type variant (entier reel, chaine de charactère
Type Variable
    Cle As Integer
    Valeur As Variant
End Type

Type Geometrie
    Type As String
    Couleur_62 As Long     ' Color
    Epaisseur_39 As Long 'Epaisseur ligne
    Layer_8 As String    ' Layer Nom
    Style_6 As Integer    ' Style ligne
                         '0=vbSolid
                         '1=vbDash
                         '2=vbDot
                         '3=vbDashDot
                         '4=vbDashDotDot
                         '5=vbInvisible
                         '6=vbInsideSolid
    Donnee() As Variable
End Type


Type Layer
    Nom As String
    Visible_290    As Boolean
    Couleur_62     As Long    ' Couleur
    Epaisseur_39   As Long      ' Epaisseur ligne
    Type_6         As Integer   ' Style ligne
                                ' 0=vbSolid
                                ' 1=vbDash
                                ' 2=vbDot
                                ' 3=vbDashDot
                                ' 4=vbDashDotDot
                                ' 5=vbInvisible
                                ' 6=vbInsideSolid
    Donnee() As Variable
End Type

Type Block
    Nom As String
    Entite() As Geometrie
End Type

Type DXFDonnee
    Block() As Block
    Entite() As Geometrie
End Type

Dim Section() As String
Public MonLayer() As Layer

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Sub Lire_Cle(FF As Integer, DxfGroupe As Integer, DxfValeur As String)
    Static Ligne$
    Line Input #FF, Ligne$
    DxfGroupe = Val(Ligne$)
    
    Line Input #FF, Ligne$
    DxfValeur = Trim(Ligne$)

    
End Sub


Sub RAZ_Cle(ByRef Geo As Geometrie)
Dim i As Integer
For i = 0 To UBound(Geo.Donnee)
    Geo.Donnee(i).Cle = i
Next i
End Sub

Function Donne_Angle(Angle As Double) As Double
If Angle > 360 Then
    Donne_Angle = Angle - 360
ElseIf Angle < 0 Then
    Donne_Angle = Angle + 360
Else
    Donne_Angle = Angle
End If
End Function
Function TrouveDepartTexte(sTableau() As String, Depart As Long)
Dim i As Long
Dim PosDepart As Long

For i = Depart + 1 To UBound(sTableau) Step 2
    If sTableau(i) = "1" Then
        PosDepart = i
    End If
    Debug.Print sTableau(i)
    If sTableau(i) = "10" And sTableau(i + 2) = "20" Then
        If PosDepart > Depart Then
            TrouveDepartTexte = PosDepart
        Else
            TrouveDepartTexte = i
        End If
        Exit Function
    End If
Next i
TrouveDepartTexte = -1
End Function

Function TrouveDepart(sTableau() As String, Depart As Long)
Dim i As Long
For i = Depart + 1 To UBound(sTableau) Step 2
    If sTableau(i) = "10" And sTableau(i + 2) = "20" Then
        TrouveDepart = i
        Exit Function
    End If
Next i
TrouveDepart = -1
End Function

'Analyse les informations lues pendant le chargement du fichier DXF
'effectue un tri suivant nos besoins et les options supportées par le programme
Sub AnalyseEntitee(ByRef Geo As Geometrie)

Dim i As Integer
Dim tpGeo As Geometrie

        ReDim tpGeo.Donnee(1) As Variable
        tpGeo.Type = Geo.Type
        tpGeo.Couleur_62 = CVal(Geo.Donnee(), 62) 'couleur
        tpGeo.Layer_8 = CVal(Geo.Donnee(), 8) 'layer
        tpGeo.Epaisseur_39 = CVal(Geo.Donnee(), 39) 'Epaisseur
        
        If CVal(Geo.Donnee(), 6) <> "0" Then
        Debug.Print "@" & CVal(Geo.Donnee(), 6) & "@"
        End If
        
      'Style_6 As Integer    ' Style ligne
                         '0=vbSolid
                         '1=vbDash
                         '2=vbDot
                         '3=vbDashDot
                         '4=vbDashDotDot
                         '5=vbInvisible
                         '6=vbInsideSolid
                         
        Select Case LTrim(RTrim(CVal(Geo.Donnee(), 6))) 'type de ligne
        
        
        Case "0" '0=vbSolid
              tpGeo.Style_6 = 0
              
        Case "HIDDEN" '1=vbDash
              tpGeo.Style_6 = 1
        
        Case "CONTINUOUS" '0=vbSolid
              tpGeo.Style_6 = 0
              
        Case "CENTRE", "CENTER" '3=vbDashDot
              tpGeo.Style_6 = 3
        
        Case "DASHED3" '4=vbDashDotDot
              tpGeo.Style_6 = 4
              
        Case "BYBLOCK"
              tpGeo.Style_6 = 1
              
        Case Else
            Debug.Print "!" & CVal(Geo.Donnee(), 6) & "!"
            tpGeo.Style_6 = 0
        
        End Select
        

                         

        
Select Case Geo.Type
    Case "LINE"
        ReDim Preserve tpGeo.Donnee(3) As Variable
        tpGeo.Donnee(0).Valeur = CVal(Geo.Donnee(), 10) 'X
        tpGeo.Donnee(1).Valeur = CVal(Geo.Donnee(), 20) 'Y
        tpGeo.Donnee(2).Valeur = CVal(Geo.Donnee(), 11) 'End point X
        tpGeo.Donnee(3).Valeur = CVal(Geo.Donnee(), 21) 'End point Y
        ReDim Geo.Donnee(3) As Variable
            
    Case "ARC"
        ReDim Preserve tpGeo.Donnee(4) As Variable
        tpGeo.Donnee(0).Valeur = CVal(Geo.Donnee(), 10)
        tpGeo.Donnee(1).Valeur = CVal(Geo.Donnee(), 20)
        tpGeo.Donnee(2).Valeur = CVal(Geo.Donnee(), 40)
        tpGeo.Donnee(3).Valeur = CVal(Geo.Donnee(), 50)
        tpGeo.Donnee(4).Valeur = CVal(Geo.Donnee(), 51)
        ReDim Geo.Donnee(4) As Variable
        
    Case "CIRCLE"
        ReDim Preserve tpGeo.Donnee(2) As Variable
        tpGeo.Donnee(0).Valeur = CVal(Geo.Donnee(), 10)
        tpGeo.Donnee(1).Valeur = CVal(Geo.Donnee(), 20)
        tpGeo.Donnee(2).Valeur = CVal(Geo.Donnee(), 40)
        ReDim Geo.Donnee(2) As Variable

        
    Case "ELLIPSE"
        ReDim Preserve tpGeo.Donnee(6) As Variable
        tpGeo.Donnee(0).Valeur = CVal(Geo.Donnee(), 10)
        tpGeo.Donnee(1).Valeur = CVal(Geo.Donnee(), 20)
        tpGeo.Donnee(2).Valeur = CVal(Geo.Donnee(), 11)
        tpGeo.Donnee(3).Valeur = CVal(Geo.Donnee(), 21)
        tpGeo.Donnee(4).Valeur = CVal(Geo.Donnee(), 40)
        tpGeo.Donnee(5).Valeur = CVal(Geo.Donnee(), 41)
        tpGeo.Donnee(6).Valeur = CVal(Geo.Donnee(), 42)
        ReDim Geo.Donnee(6) As Variable

    Case "VERTEX"
        ReDim Preserve tpGeo.Donnee(1) As Variable
        tpGeo.Donnee(0).Valeur = CVal(Geo.Donnee(), 10)
        tpGeo.Donnee(1).Valeur = CVal(Geo.Donnee(), 20)
        ReDim Geo.Donnee(1) As Variable

    Case "TEXT"
        ReDim Preserve tpGeo.Donnee(4) As Variable
        tpGeo.Donnee(0).Valeur = CVal(Geo.Donnee(), 10)   '10  First alignment point (in OCS) DXF: X Value; APP: 3D point
        tpGeo.Donnee(1).Valeur = CVal(Geo.Donnee(), 20)   '20, 30  DXF: Y and Z Values of first alignment point (in OCS)
        tpGeo.Donnee(2).Valeur = CVal(Geo.Donnee(), 40)   '40  Text height
        tpGeo.Donnee(3).Valeur = CVal(Geo.Donnee(), 50)   '50  Text rotation (optional; default = 0)
        tpGeo.Donnee(4).Valeur = Corrige_texte(CVal(Geo.Donnee(), 1))    '1   Default Value (the string itself)
        ReDim Geo.Donnee(4) As Variable

    
    Case "INSERT"
        ReDim Preserve tpGeo.Donnee(5) As Variable
        tpGeo.Donnee(0).Valeur = CVal(Geo.Donnee(), 2)
        tpGeo.Donnee(1).Valeur = CVal(Geo.Donnee(), 10)
        tpGeo.Donnee(2).Valeur = CVal(Geo.Donnee(), 20)
        tpGeo.Donnee(3).Valeur = CVal(Geo.Donnee(), 41)
        tpGeo.Donnee(4).Valeur = CVal(Geo.Donnee(), 42)
        ReDim Geo.Donnee(5) As Variable

    
    Case "DIMENSION"
        ReDim Preserve tpGeo.Donnee(10) As Variable
        tpGeo.Donnee(0).Valeur = CVal(Geo.Donnee(), 2)
        tpGeo.Donnee(1).Valeur = CVal(Geo.Donnee(), 10)
        tpGeo.Donnee(2).Valeur = CVal(Geo.Donnee(), 20)
        tpGeo.Donnee(3).Valeur = CVal(Geo.Donnee(), 11)
        tpGeo.Donnee(4).Valeur = CVal(Geo.Donnee(), 21)
        tpGeo.Donnee(5).Valeur = CVal(Geo.Donnee(), 12)
        tpGeo.Donnee(6).Valeur = CVal(Geo.Donnee(), 22)
        tpGeo.Donnee(7).Valeur = CVal(Geo.Donnee(), 13)
        tpGeo.Donnee(8).Valeur = CVal(Geo.Donnee(), 23)
        tpGeo.Donnee(9).Valeur = CVal(Geo.Donnee(), 14)
        tpGeo.Donnee(10).Valeur = CVal(Geo.Donnee(), 24)
        ReDim Geo.Donnee(10) As Variable
End Select

Geo = tpGeo
        
RAZ_Cle Geo
End Sub
Function PtAng(X1 As Double, Y1 As Double) As Double
If X1 = 0 Then
    If Y1 >= 0 Then
        PtAng = 90
    Else
        PtAng = 270
    End If
    PtAng = PtAng * PI / 180
    Exit Function
ElseIf Y1 = 0 Then
    If X1 >= 0 Then
        PtAng = 0
    Else
        PtAng = 180
    End If
    PtAng = PtAng * PI / 180
    Exit Function
Else
    PtAng = Atn(Y1 / X1)
    PtAng = PtAng * 180 / PI
    If PtAng < 0 Then PtAng = PtAng + 360
    If PtAng > 360 Then PtAng = PtAng - 360
    '----------Test la direction-(test par quart)-------
    If X1 < 0 Then PtAng = PtAng + 180
    If Y1 < 0 And PtAng < 90 Then PtAng = PtAng + 180
    'If X1 < 0 And PtAng <> 180 Then PtAng = PtAng + 180
    'If Y1 < 0 And PtAng = 90 Then PtAng = PtAng + 180
    
    'One final check
    If PtAng < 0 Then PtAng = PtAng + 360
    If PtAng > 360 Then PtAng = PtAng - 360
    PtAng = PtAng * PI / 180
End If
End Function
'Retourne l'Donne_Hypotenuse
Function Donne_Hypo(X1 As Double, Y1 As Double) As Double
Donne_Hypo = Sqr((X1 * X1) + (Y1 * Y1))
End Function

Function Distance(P1 As Point2, P2 As Point2) As Double
    Distance = Sqr((P2.X - P1.X) ^ 2 + (P2.Y - P1.Y) ^ 2)
End Function


Sub AfficheDXF(Pict As PictureBox, DXF As DXFDonnee, Optional Layer As String)
On Error GoTo Sortie
Pict.Cls
Pict.Picture = LoadPicture()
Dim i As Integer
For i = 0 To UBound(DXF.Entite)
    AfficheDXFGeometrie Pict, DXF, DXF.Entite(), i, 0, 0, 1, 1, 0, Layer
Next i

Pict.Picture = Pict.Image
Sortie:
End Sub

' Subroutine pour affichage d'un block
' Subroutine inutilisée pour le moment
Sub AfficheBlock(Pict As PictureBox, DXF As DXFDonnee, NumBlock As Integer)
On Error GoTo Sortie
Pict.Cls
Pict.Picture = LoadPicture()
Dim i As Integer

For i = 0 To UBound(DXF.Block(NumBlock).Entite)
    AfficheDXFGeometrie Pict, DXF, DXF.Block(NumBlock).Entite(), i, 0, 0, 1, 1, 0
Next i
Pict.Picture = Pict.Image
Sortie:
End Sub
Sub AfficheDXFBlock(Pict As PictureBox, DXF As DXFDonnee, Nom As String, cX As Double, cY As Double, EchelleX As Double, EchelleY As Double, Angle As Double)
Dim i As Integer
Dim bNum As Integer
bNum = GetBlock(DXF, Nom)
For i = 0 To UBound(DXF.Block(bNum).Entite)
    AfficheDXFGeometrie Pict, DXF, DXF.Block(bNum).Entite(), i, cX, cY, EchelleX, EchelleY, Angle
Next i
End Sub
Sub AfficheDXFDImension(Pict As PictureBox, DXF As DXFDonnee, Nom As String)
Dim i As Integer
Dim bNum As Integer
bNum = GetBlock(DXF, Nom)
For i = 0 To UBound(DXF.Block(bNum).Entite)
    AfficheDXFGeometrie Pict, DXF, DXF.Block(bNum).Entite(), i, 0, 0, 1, 1, 0
Next i
End Sub
Sub AfficheDXFLigne(Pict As PictureBox, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Couleur As Long)
    Pict.DrawWidth = 1
    If Couleur = 0 Then
        Pict.Line (X1, -Y1)-(X2, -Y2), 16777215
    Else
        Pict.Line (X1, -Y1)-(X2, -Y2), Couleur
    End If
    
End Sub
Sub AfficheDXFLigneX(Pict As PictureBox, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Couleur As Long)
    Pict.DrawWidth = 3
    Pict.Line (X1, -Y1)-(X2, -Y2), Couleur
End Sub

Sub AfficheDXFText(Pict As PictureBox, X1 As Double, Y1 As Double, Angle As Double, Size As Double, Text As String, Couleur As Long)
Dim f As LOGFONT
Dim hPrevFont As Long
Dim hFont As Long
Dim FontNom As String
Dim XSIZE As Integer
Dim YSIZE As Integer
f.lfEscapement = 10 * Val(Angle) 'Angle de rotation en dixième
FontNom = "Arial Black" + Chr$(0) 'Terminaison Null
f.lfFaceNom = FontNom
XSIZE = Pict.ScaleX(Size, 0, 2)
YSIZE = Pict.ScaleY(Size, 0, 2)
If XSIZE = 0 Then XSIZE = 1
If YSIZE = 0 Then YSIZE = 1
f.lfWidth = (XSIZE * -15) / Screen.TwipsPerPixelY
f.lfHeight = (YSIZE * -20) / Screen.TwipsPerPixelY
hFont = CreateFontIndirect(f)
hPrevFont = SelectObject(Pict.hdc, hFont)
Pict.ForeColor = Couleur
Pict.CurrentX = X1
Pict.CurrentY = -Y1 - Size
Pict.Print Text

'  RAZ, restore la police originale
hFont = SelectObject(Pict.hdc, hPrevFont)
DeleteObject hFont
End Sub

Sub AfficheDXFArc(Pict As PictureBox, X1 As Double, Y1 As Double, rad As Double, Angle1 As Double, Angle2 As Double, Couleur As Long)

Angle1 = Donne_Angle(Angle1)
Angle2 = Donne_Angle(Angle2)
Dim i As Double
Dim interval As Double
If Angle1 > Angle2 Then
    If Angle1 <> 360 Then Pict.Circle (X1, -Y1), rad, Couleur, Angle1 * PI / 180, 2 * PI
    If Angle2 <> 0 Then Pict.Circle (X1, -Y1), rad, Couleur, 0, Angle2 * PI / 180
Else
    'Decoupage du cercle en deux demi arc de cercle
    'Permet déviter des traitements différent suivant si arc ou cercle complet
    interval = (Angle2 - Angle1) / PI
    For i = Angle1 To Angle2 - interval Step interval
        Pict.Circle (X1, -Y1), rad, Couleur, i * PI / 180, (i + interval) * PI / 180
    Next i
    Pict.Circle (X1, -Y1), rad, Couleur, i * PI / 180, (Angle2) * PI / 180
End If
End Sub
Sub AfficheDXFArcX(X1 As Double, Y1 As Double, rad As Double, Angle1 As Double, Angle2 As Double, Couleur As Long)

frmVisu.VisuDXF.DrawWidth = 5
'frmVisu.VisuDXF.Line (X1, -Y1)-(X2, -Y2), 16777215

Angle1 = Donne_Angle(Angle1)
Angle2 = Donne_Angle(Angle2)
Dim i As Double
Dim interval As Double
If Angle1 > Angle2 Then
    If Angle1 <> 360 Then frmVisu.VisuDXF.Circle (X1, -Y1), rad, Couleur, Angle1 * PI / 180, 2 * PI
    If Angle2 <> 0 Then frmVisu.VisuDXF.Circle (X1, -Y1), rad, Couleur, 0, Angle2 * PI / 180
Else
    'Decoupage du cercle en deux demi arc de cercle
    'Permet déviter des traitements différent suivant si arc ou cercle complet
    interval = (Angle2 - Angle1) / PI
    For i = Angle1 To Angle2 - interval Step interval
        frmVisu.VisuDXF.Circle (X1, -Y1), rad, Couleur, i * PI / 180, (i + interval) * PI / 180
    Next i
    frmVisu.VisuDXF.Circle (X1, -Y1), rad, Couleur, i * PI / 180, (Angle2) * PI / 180
End If
End Sub

Sub AfficheDXFCercle(Pict As PictureBox, X1 As Double, Y1 As Double, rad As Double, Couleur As Long)
Pict.Circle (X1, -Y1), rad, Couleur
End Sub
Sub AfficheDXFPoint(Pict As PictureBox, X1 As Double, Y1 As Double, Couleur As Long)
    Pict.DrawWidth = 3
    Pict.PSet (X1, -Y1), Couleur
    Pict.DrawWidth = 1
End Sub
Sub AfficheDXFEllipse(Pict As PictureBox, cX As Double, cY As Double, mX As Double, mY As Double, Ratio As Double, Angle1 As Double, Angle2 As Double, NumPoints As Integer, Couleur)
Dim A As Double, B As Double
Dim RotAngle As Double
Dim a1 As Double, A2 As Double
Dim X1 As Double, Y1 As Double
Dim X2 As Double, Y2 As Double
Dim X3 As Double, Y3 As Double
Dim Hyp As Double
Dim J As Double
Dim U As Double
Dim Count As Integer

A = Sqr((mX * mX) + (mY * mY))
If mX < 0 Then A = -A
B = Ratio * A
If mX = 0 Then
    RotAngle = PI / 2
Else
    RotAngle = Atn(mY / mX)
End If
For U = Angle1 To Angle2 + (PI / (NumPoints * 2)) Step PI / NumPoints
    X1 = A * Cos(U)
    Y1 = B * Sin(U)
    Hyp = Sqr((X1 * X1) + (Y1 * Y1))
    If X1 = 0 Then
        J = PI / 2
    Else
        J = Atn(Y1 / X1)
    End If
    If X1 < 0 Then Hyp = -Hyp
    If (J * 180 / PI) + (RotAngle * 180 / PI) > 360 Then J = J + (2 * PI)
    X2 = (Hyp * Cos(RotAngle + J))
    Y2 = (Hyp * Sin(RotAngle + J))
    If Count > 0 Then Pict.Line (cX + X3, -cY - Y3)-(cX + X2, -cY - Y2), Couleur
    X3 = X2
    Y3 = Y2
    Count = Count + 1
Next U
End Sub
Sub AfficheDXFGeometrie(Pict As PictureBox, DXF As DXFDonnee, Geo() As Geometrie, Depart As Integer, cX As Double, cY As Double, EchelleX As Double, EchelleY As Double, Angle As Double, Optional lay As String)
'Si une Geometrie est modifié par une transformation du plan
'les transformations sont appliquées dans l'ordre suivant
    '--------
    'Facteur d'echelle
    'Rotation
    'Changement origine
    '--------
On Error Resume Next
Dim Couleur As Long
Dim i As Integer
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
        
         If lay <> "" And lay <> Geo(Depart).Layer_8 Then
            Exit Sub
        End If
              
        If Geo(Depart).Couleur_62 < 0 Then
           Debug.Print "ENTITY OFF"
        End If
 

        If Geo(Depart).Couleur_62 Then
            Couleur = ConvCouleur(Geo(Depart).Couleur_62)
        Else
            Couleur = ConvCouleur(Donne_Couleur(Geo(Depart).Layer_8))
        End If
        
        If Geo(Depart).Epaisseur_39 Then
            Pict.DrawWidth = Geo(Depart).Epaisseur_39
        Else
            Pict.DrawWidth = 1
        End If
        
        If Geo(Depart).Style_6 Then
            Pict.DrawStyle = Geo(Depart).Style_6
        Else
            Pict.DrawStyle = vbSolid
        End If
        
        
Select Case Geo(Depart).Type
    Case "LINE"
        'Récupère les valeurs
        X1 = Replace(Geo(Depart).Donnee(0).Valeur, ".", ",")
        Y1 = Replace(Geo(Depart).Donnee(1).Valeur, ".", ",")
        X2 = Replace(Geo(Depart).Donnee(2).Valeur, ".", ",")
        Y2 = Replace(Geo(Depart).Donnee(3).Valeur, ".", ",")
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
        'Dessine la ligne
            AfficheDXFLigne Pict, X1, Y1, X2, Y2, Couleur

    Case "ARC"
        ' Les cercles et les lignes deviennent automatiquement des ellipses quant on leur applique un facteur d'échelle
        X1 = Replace(Geo(Depart).Donnee(0).Valeur, ".", ",")
        Y1 = Replace(Geo(Depart).Donnee(1).Valeur, ".", ",")
        rad = Replace(Geo(Depart).Donnee(2).Valeur, ".", ",")
        Angle1 = Replace(Geo(Depart).Donnee(3).Valeur, ".", ",")
        Angle2 = Replace(Geo(Depart).Donnee(4).Valeur, ".", ",")
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
        AfficheDXFArc Pict, X1, Y1, Abs(rad), Angle1, Angle2, 16777215

    Case "CIRCLE"
        ' Les cercles et les lignes deviennent automatiquement des ellipses quant on leur applique un facteur d'échelle
        X1 = Replace(Geo(Depart).Donnee(0).Valeur, ".", ",")
        Y1 = Replace(Geo(Depart).Donnee(1).Valeur, ".", ",")
        rad = Replace(Geo(Depart).Donnee(2).Valeur, ".", ",")
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
        AfficheDXFCercle Pict, X1, Y1, Abs(rad), Couleur

        
    Case "ELLIPSE"
        X1 = Geo(Depart).Donnee(0).Valeur
        Y1 = Geo(Depart).Donnee(1).Valeur
        X2 = Geo(Depart).Donnee(2).Valeur
        Y2 = Geo(Depart).Donnee(3).Valeur
        Ratio = Geo(Depart).Donnee(4).Valeur
        Angle1 = Geo(Depart).Donnee(5).Valeur
        Angle2 = Geo(Depart).Donnee(6).Valeur
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
        AfficheDXFEllipse Pict, X1, Y1, X2, Y2, Ratio, Angle1, Angle2, 32, Couleur

        
    Case "POLYLINE"
        ' Une POLYLINE est une suite de ligne liée entre elles
        PCount = 1
        EndPoly = False
        Do While Not EndPoly
            X1 = Geo(Depart + PCount).Donnee(0).Valeur
            Y1 = Geo(Depart + PCount).Donnee(1).Valeur
            X2 = Geo(Depart + PCount + 1).Donnee(0).Valeur
            Y2 = Geo(Depart + PCount + 1).Donnee(1).Valeur
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
            'Dessine la ligne
            AfficheDXFLigne Pict, X1, Y1, X2, Y2, Couleur
            
            PCount = PCount + 1
            If Depart + PCount + 1 > UBound(Geo) Then
                EndPoly = True
            ElseIf Geo(Depart + PCount + 1).Type <> "VERTEX" Then
                EndPoly = True
            End If
        Loop
        
    Case "TEXT"
        'pas de facteur d'échelle pour le TEXTE
        X1 = Geo(Depart).Donnee(0).Valeur
        Y1 = Geo(Depart).Donnee(1).Valeur
        Size = Geo(Depart).Donnee(2).Valeur
        Angle1 = Geo(Depart).Donnee(3).Valeur + Angle
        Text = Geo(Depart).Donnee(4).Valeur
        'Déplace l'origine
        X1 = X1 + cX
        Y1 = Y1 + cY
        AfficheDXFText Pict, X1, Y1, Angle1, Size, Text, Couleur

    Case "INSERT"
        'pas de facteur d'échelle pour la INSERT
        Nom = Geo(Depart).Donnee(0).Valeur
        X1 = Geo(Depart).Donnee(1).Valeur
        Y1 = Geo(Depart).Donnee(2).Valeur
        X2 = Geo(Depart).Donnee(3).Valeur
        Y2 = Geo(Depart).Donnee(4).Valeur
        '"0" scale = scale of "1"
        If X2 = 0 Then X2 = 1
        If Y2 = 0 Then Y2 = 1
        Angle1 = Geo(Depart).Donnee(5).Valeur * PI / 180
        AfficheDXFBlock Pict, DXF, Nom, X1, Y1, X2, Y2, Angle1
    
    Case "DIMENSION"
        'pas de facteur d'échelle pour la DIMENSION
        Nom = Geo(Depart).Donnee(0).Valeur
        X1 = Geo(Depart).Donnee(1).Valeur
        Y1 = Geo(Depart).Donnee(2).Valeur
        AfficheDXFDImension Pict, DXF, Nom
End Select
End Sub

Function FindCommand(FileNum As Integer, Command As String)
Dim X As String
On Error GoTo Fin
Do While UCase(Trim(X)) <> UCase(Command)
    Line Input #FileNum, X
Loop
FindCommand = 1
Exit Function
Fin:
FindCommand = 0
Seek #FileNum, 1
End Function

Function GetBlock(DXF As DXFDonnee, Nom As String) As Integer
Dim i As Integer
For i = 0 To UBound(DXF.Block)
    If DXF.Block(i).Nom = Nom Then
        GetBlock = i
        Exit Function
    End If
Next i
End Function

Function GetSection(FichierNum As Integer, Depart As String, Fin As String, FinChaine As String, sTableau() As String) As Boolean
ReDim sTableau(0) As String
Dim Temp As String
Dim i As Long

'recherche le debut de la section
Do While Temp <> Depart
    Line Input #FichierNum, Temp
    Temp = UCase(Trim(Temp))
    If Temp = FinChaine Then
        GetSection = False
        Exit Function
    End If
Loop

'Relecture de la section
Do While Temp <> Fin
    Line Input #FichierNum, Temp
    Temp = UCase(Trim(Temp))
    If Temp <> Fin Then
        ReDim Preserve sTableau(i) As String
        sTableau(i) = Temp
        i = i + 1
    End If
Loop

GetSection = True
End Function

'Importe le fichier DXF
Sub ImportDXF(FileDXF As String, ByRef DXF As DXFDonnee)
Dim FF As Integer
Dim DXFLine As String
Dim bCount As Integer
Dim eCount As Integer
Dim ENDSEC As Boolean


ReDim DXF.Block(0) As Block
ReDim DXF.Entite(0) As Geometrie

FF = FreeFile
Open FileDXF For Input As #FF
'Traitement des layer
If FindCommand(FF, "LAYER") Then
   GetSection FF, "LAYER", "ENDTAB", "ENDTAB", Section()
   FiltreLayer Section(), MonLayer()
End If


'Saute la définitions des header pour aller a la définitions des Block ou si il n'y a pas a ENDSEC

If FindCommand(FF, "BLOCKS") = 0 Then
    FindCommand FF, "ENDSEC"
End If
'---------------------------
'les Block sont des groupes de géométries
'qui sont réutilisables dans le dessin
Do While Not ENDSEC
    'Chargement a línterieur d'une SECTION d'un BLOCK de (BLOCK) à (ENDBLK)
    'Boucle jusqu'a l'instruction  "ENDSEC"
    If GetSection(FF, "BLOCK", "ENDBLK", "ENDSEC", Section()) Then
        'Un "BLOCK" a éte stocké dans le tableau
        'Je redimmensione le tableau pour le "BLOCK" suivant
        ReDim Preserve DXF.Block(bCount) As Block
        ReDim Preserve DXF.Block(bCount).Entite(eCount) As Geometrie
        If FiltreBlock(Section(), DXF.Block(bCount)) Then
            bCount = bCount + 1
            eCount = 0
        End If
    Else
        ENDSEC = True
    End If
Loop

ENDSEC = False
eCount = 0



GetSection FF, "ENTITIES", "ENDSEC", "ENDSEC", Section()
'Permet de relire toutes les entites du fichier comme un super BLOCK
Close #FF 'Fin de traitement fermeture du Fichier

'Filtre et remplissage de la BD dxf
FiltreDB Section(), DXF.Entite()

End Sub
Function InstructionCommand(InText As String)
Select Case UCase(InText)
    'ENTITY COMMANDS traité dans le language DXF
    Case "LINE", "VERTEX", "POLYLINE", "CIRCLE", "ARC", "ELLIPSE", "TEXT", "INSERT", "DIMENSION"
        InstructionCommand = True
    Case Else
        InstructionCommand = False
End Select
End Function

Function CVal(Donnee() As Variable, CleRef As Integer) As Variant
Dim i As Integer
CVal = 0

For i = 0 To UBound(Donnee)
    If Donnee(i).Cle = CleRef Then
        CVal = Donnee(i).Valeur
    End If
Next i

End Function
Function CVal_debug(Donnee() As Variable, Key As Integer) As Variant
Dim i As Integer
For i = 0 To UBound(Donnee)
    If Donnee(i).Cle = Key Then
        Debug.Print Donnee(i).Cle & "->" & Donnee(i).Valeur
        
        CVal_debug = Donnee(i).Valeur
        Exit Function
    End If
Next i
CVal_debug = 0
End Function

Function FiltreBlock(sTableau() As String, ByRef tBlock As Block) As Boolean
'On Local Error GoTo Sortie:
Dim i As Long
Dim J As Long
Dim K As Long
Dim p As Long

'Analyse la section 6
i = TrouveSection(sTableau(), i, "6")
If i = -1 Then
    FiltreBlock = False
    Exit Function
End If
i = TrouveSection(sTableau(), i, "2") + 1
tBlock.Nom = sTableau(i)
For J = i To UBound(sTableau)
    If InstructionCommand(sTableau(J)) Then ' ENTITY COMMAND trouvee
        ReDim Preserve tBlock.Entite(K) As Geometrie
        tBlock.Entite(K).Type = sTableau(J)
        Select Case tBlock.Entite(K).Type
            Case "INSERT", "DIMENSION"
                'CLE "2" Donne le nom du BLOCK
                J = TrouveSection(sTableau(), J, "2")
           Case "TEXT"
               J = TrouveDepartTexte(sTableau(), J)
               
            Case Else
                J = TrouveDepart(sTableau(), J)
                'j = TrouveSection(sTableau(), j, "10")
        End Select
        Do While sTableau(J) <> "0"
            ReDim Preserve tBlock.Entite(K).Donnee(p)
            tBlock.Entite(K).Donnee(p).Cle = sTableau(J)
            tBlock.Entite(K).Donnee(p).Valeur = sTableau(J + 1)
            p = p + 1
            J = J + 2
        Loop
        AnalyseEntitee tBlock.Entite(K)
        K = K + 1
        p = 0
    End If
Next J
FiltreBlock = True
Exit Function
Sortie:
MsgBox "ERROR  " & Err.Description
End Function

Function FiltreLayer(sTableau() As String, ByRef tLayer() As Layer) As Boolean
Dim i As Long
Dim J As Long
Dim K As Long
Dim p As Long
Dim L As Long
Dim exist As Boolean

'i = TrouveSection(sTableau(), j, "LAYER")
 
For J = i To UBound(sTableau)


    If UCase(sTableau(J)) = "LAYER" Then 'presence d'une COMMAND ENTITY
        ReDim Preserve tLayer(K) As Layer

         'CLE "2" sur une commande INSERT domne le nom du BLOCK a insérer
        J = TrouveSection(sTableau(), J, "2")
        tLayer(K).Nom = sTableau(J + 1)
        
        Do While sTableau(J) <> "0"
            ReDim Preserve tLayer(K).Donnee(p)
            tLayer(K).Donnee(p).Cle = sTableau(J)
            tLayer(K).Donnee(p).Valeur = sTableau(J + 1)
            p = p + 1
            J = J + 2
        Loop
        
        exist = False
        For L = 0 To K - 1
            If tLayer(K).Nom = tLayer(L).Nom Then exist = True
        Next L
        
        If Not exist Then
        If UBound(tLayer(K).Donnee) < 1 Then ReDim Preserve tLayer(K).Donnee(1) As Variable
        tLayer(K).Donnee(0).Valeur = CVal(tLayer(K).Donnee(), 6)
        tLayer(K).Donnee(1).Valeur = CVal(tLayer(K).Donnee(), 62)
        ReDim Preserve tLayer(K).Donnee(1) As Variable
        
        K = K + 1
        End If
        p = 0
        
    End If
Next J

FiltreLayer = True
End Function
Function Donne_Couleur(Nom_layer As Variant) As Integer
Dim i As Integer
Dim Val_C As Double

Val_C = 0

For i = 0 To UBound(MonLayer)
If MonLayer(i).Nom = Nom_layer Then
    Val_C = Val(MonLayer(i).Donnee(1).Valeur)
End If
Next i

Donne_Couleur = Val_C
End Function

Function FiltreDB(sTableau() As String, ByRef tGeo() As Geometrie) As Boolean
Dim i As Long
Dim J As Long
Dim K As Long
Dim p As Long
For J = i To UBound(sTableau)


    If InstructionCommand(sTableau(J)) Then 'COMMAND ENTITY trouvee
        ReDim Preserve tGeo(K) As Geometrie
        tGeo(K).Type = sTableau(J)
        J = J + 1

        Do While Trim(sTableau(J)) <> "0"
            ReDim Preserve tGeo(K).Donnee(p)
            tGeo(K).Donnee(p).Cle = sTableau(J)
            tGeo(K).Donnee(p).Valeur = sTableau(J + 1)
            p = p + 1
            J = J + 2
        Loop
        
        AnalyseEntitee tGeo(K)
        K = K + 1
        p = 0
    End If
Next J
FiltreDB = True
End Function
Function RotationX(X1 As Double, Y1 As Double, Angle As Double) As Double
RotationX = Donne_Hypo(X1, Y1) * Cos(PtAng(X1, Y1) + Angle)
End Function

Function RotationY(X1 As Double, Y1 As Double, Angle As Double) As Double
RotationY = Donne_Hypo(X1, Y1) * Sin(PtAng(X1, Y1) + Angle)
End Function
Function TrouveSection(sTableau() As String, Depart As Long, Value As String) As Long
Dim i As Long
For i = Depart To UBound(sTableau)
    If sTableau(i) = Value Then
        TrouveSection = i
        Exit Function
    End If
Next i
TrouveSection = -1
End Function

Function ConvCouleur(Couleur As Variant) As Long

' VB QbColor             AUTOCAD
' =============         ===========
'0   Noir                 7
'1   Bleu                 50
'2   Vert                 41
'3   bleu                 45
'4   Rose                 34
'5   Magenta              55
'6   jaune                38
'7   Blanc sale           9
'8   Gris                 8
'9   Bleu Clair           5
'10  Vert Clair           3
'11  bleu Clair           4
'12  Rose Clair           1
'13  Magenta Clair        6
'14  Jaune Clair          2
'15  Blanc brillant       0


 ' Convertir une Couleur AutoCAD 10-14 en couleur VB
 
 Dim C As Integer
 
 Select Case Val(Couleur)
    Case 7: C = 15
    Case 50: C = 1
    Case 42: C = 2
    Case 45: C = 3
    Case 34: C = 4
    Case 55: C = 5
    Case 38: C = 6
    Case 9: C = 7
    Case 8: C = 8
    Case 5: C = 9
    Case 3: C = 10
    Case 4: C = 11
    Case 1: C = 12
    Case 6: C = 13
    Case 2: C = 14
    Case 7: C = 15
 End Select


ConvCouleur = QBColor(C)

End Function

Sub Intervertir(ByRef A As Variant, ByRef B As Variant)
Dim C As Variant
C = A
A = B
B = C
End Sub

Function Corrige_texte(string_depart As String) As String
Dim Chaine_traitee As String
Chaine_traitee = string_depart


    ' %%C Représente le signe Diametre
    Chaine_traitee = mReplaceInString(string_depart, "%%C", Chr$(216))
    '%%P Repr'esente le signe +-
    Chaine_traitee = mReplaceInString(Chaine_traitee, "%%P", Chr$(177))

Corrige_texte = Chaine_traitee
End Function

'****************************************************************
' Nom: mReplaceInString

Public Function mReplaceInString(ByVal vstrInputString As String, ByVal vstrA As String, ByVal vstrB As String) As String

       Dim intPos As Integer
       intPos = InStr(1, vstrInputString, vstrA)
       '     'Replace vstrA with vstrB

              Do Until intPos = 0
                     vstrInputString = Mid$(vstrInputString, 1, intPos - 1) + vstrB + Mid$(vstrInputString, intPos + Len(vstrA))
                     intPos = InStr(1, vstrInputString, vstrA)
              Loop

       mReplaceInString = vstrInputString
End Function

