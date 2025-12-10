VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "PDVNovatos"
   ClientHeight    =   7200
   ClientLeft      =   4980
   ClientTop       =   2070
   ClientWidth     =   7635
   LinkTopic       =   "Form2"
   ScaleHeight     =   7200
   ScaleWidth      =   7635
   Begin VB.Image imgLogo 
      Height          =   7680
      Left            =   0
      Picture         =   "frmPrincipal.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   7680
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "Cadastros"
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnuServiços 
         Caption         =   "Serviços"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mnuClientes_Click()
frmClientes.Show vbModal
End Sub
Private Sub mnuServiços_Click()
frmServiços.Show vbModal
End Sub

