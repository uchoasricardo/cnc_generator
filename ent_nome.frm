VERSION 5.00
Begin VB.Form ent_nome 
   Caption         =   "Dados do programa"
   ClientHeight    =   1995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Nome do programa:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Número do programa:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "ent_nome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
