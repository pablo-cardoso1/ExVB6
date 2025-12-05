VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   5880
   ClientTop       =   2535
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   7680
   Begin VB.CommandButton cmbLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmbAdicionar 
      Caption         =   "Adicionar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox lstNomes 
      BackColor       =   &H8000000F&
      Height          =   4815
      Left            =   4200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtNome 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Lista de nomes"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
     txtNome.SetFocus ' quando ativar o formulario o foco ira para Nome
End Sub

Private Sub cmbAdicionar_Click()
    
    AdicionarNome

End Sub

Private Sub AdicionarNome()

     Dim sNome As String
    sNome = Trim(txtNome.Text)

    ' Validação: nome não pode estar vazio
    If sNome = "" Then
        MsgBox "Digite um nome antes de adicionar!", vbExclamation, "Aviso"
        txtNome.SetFocus
        Exit Sub
    End If

    ' Se o campo da lista estiver vazio, adiciona o primeiro nome.
    ' Se já houver texto, adiciona uma nova linha antes de incluir.
    If lstNomes.Text = "" Then
        lstNomes.Text = sNome
    Else
        lstNomes.Text = lstNomes.Text & vbCrLf & sNome
    End If

    MsgBox "Nome incluído com sucesso!", vbInformation, "OK"

    ' Limpa o campo de digitação e retorna o foco
    txtNome.Text = ""
    txtNome.SetFocus

End Sub


Private Sub txtNome_KeyPress(KeyAscii As Integer)

    ' Se o usuário pressionou Enter
    If KeyAscii = 13 Then
        
        ' ...chama a funcao que adiciona o nome
        AdicionarNome
        
    End If

End Sub

Private Sub cmbLimpar_Click()

    ' Limpa o textbox
    lstNomes.Text = ""

    ' Devolve o foco para o campo de texto
    txtNome.SetFocus

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Só avisa se houver algo dentro da lista, se não fecha normal
    If Trim(lstNomes.Text) <> "" Then
        
        If MsgBox("Deseja realmente sair? Todos os nomes serão perdidos.", _
                  vbYesNo + vbQuestion, "Confirmar saída") = vbNo Then
            
            Cancel = True
            Exit Sub
        End If

    End If

End Sub




