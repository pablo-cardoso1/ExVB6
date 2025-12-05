VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   6225
   ClientTop       =   2790
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame frmDados 
      Caption         =   "Dados"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar"
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chkEspecial 
         Caption         =   "Especial"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtEndereco 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtIdade 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblEndereco 
         Caption         =   "Endereço completo"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblIdade 
         Caption         =   "Idade"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ListBox lstRegistros 
      Height          =   3180
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0007
      TabIndex        =   0
      Top             =   2880
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variável para saber se estamos incluindo ou editando
Public ModoFormulario As String
Private Sub Form_Load()

    ' Abre a conexão com o banco
    AbreConexao
    
    ' Carrega a lista ao iniciar
    CarregarLista

End Sub

Private Sub Form_Activate()

    ' validações
    txtNome.MaxLength = 50
    txtEndereco.MaxLength = 250
    txtIdade.MaxLength = 3

End Sub
Private Sub CarregarLista()

    Dim rs As New ADODB.Recordset
    lstRegistros.Clear

    ' Carrega registros do banco
    rs.Open "SELECT * FROM Cliente ORDER BY Nome", Conn, adOpenForwardOnly, adLockReadOnly

    Do While Not rs.EOF

        Dim linha As String

        ' Monta o texto exibido
        linha = rs!Id & " - " & rs!Nome & " - " & rs!idade & " - " & rs!Endereco

        If CBool(rs!Especial) Then
            linha = linha & " - Especial"
        End If

        lstRegistros.AddItem linha

        rs.MoveNext
    Loop

    rs.Close

End Sub


Private Sub cmdIncluir_Click()

    ' Define o modo Inclusão
    ModoFormulario = "I"

    ' Limpa campos antes de incluir
    LimparCampos

    ' Mostra o quadro de dados
    frmDados.Visible = True

    ' Oculta lista ao entrar no modo formulário
    lstRegistros.Visible = False

    ' Mostra botões de confirmação
    cmdConfirmar.Visible = True
    cmdCancelar.Visible = True
    
    ' Oculta os botões de cima
    cmdIncluir.Visible = False
    cmdEditar.Visible = False
    cmdExcluir.Visible = False
    

End Sub
Private Sub cmdEditar_Click()

    Dim idSel As Long
    idSel = PegarIdSelecionado

    If idSel = -1 Then
        MsgBox "Selecione um registro para editar.", vbExclamation
        Exit Sub
    End If

    ' Modo edição
    ModoFormulario = "E"

    ' Carrega dados do registro selecionado
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM Cliente WHERE Id=" & idSel, Conn

    txtNome.Text = rs!Nome
    txtIdade.Text = rs!idade
    txtEndereco.Text = rs!Endereco
    If Not IsNull(rs!Especial) And CBool(rs!Especial) Then
        chkEspecial.Value = vbChecked
    Else
        chkEspecial.Value = vbUnchecked
    End If



    rs.Close

    ' Alterna visibilidade
    lstRegistros.Visible = False
    frmDados.Visible = True

    cmdConfirmar.Visible = True
    cmdCancelar.Visible = True
    cmdIncluir.Visible = False
    cmdEditar.Visible = False
    cmdExcluir.Visible = False

End Sub
Private Sub cmdExcluir_Click()

    Dim idSel As Long
    idSel = PegarIdSelecionado

    If idSel = -1 Then
        MsgBox "Selecione um registro para excluir.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Deseja realmente excluir este registro?", _
              vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Conn.Execute "DELETE FROM Cliente WHERE Id=" & idSel

    CarregarLista

End Sub
Private Sub cmdConfirmar_Click()

    ' Verificação obrigatória
    If Trim(txtNome.Text) = "" Or _
       Trim(txtIdade.Text) = "" Or _
       Trim(txtEndereco.Text) = "" Then

        MsgBox "Preencha os campos obrigatórios!", vbExclamation
        Exit Sub
    End If

    ' Validação idade
    Dim idade As Integer
    idade = Val(txtIdade.Text)

    If idade < 0 Or idade > 125 Then
        MsgBox "Idade inválida!", vbExclamation
        Exit Sub
    End If

    ' Checkbox Especial
    Dim esp As Integer
    If chkEspecial.Value = vbChecked Then
        esp = 1
    Else
        esp = 0
    End If

    

    Dim sql As String

    ' Monta INSERT
    If ModoFormulario = "I" Then

        sql = "INSERT INTO Cliente (Nome, Idade, Endereco, Especial) VALUES (" & _
              "'" & Replace(txtNome.Text, "'", "''") & "', " & idade & ", '" & _
              Replace(txtEndereco.Text, "'", "''") & "', " & esp & ")"

    ' Monta UPDATE
    ElseIf ModoFormulario = "E" Then

        Dim idSel As Long
        idSel = PegarIdSelecionado

        sql = "UPDATE Cliente SET " & _
              "Nome='" & Replace(txtNome.Text, "'", "''") & "', " & _
              "Idade=" & idade & ", " & _
              "Endereco='" & Replace(txtEndereco.Text, "'", "''") & "', " & _
              "Especial=" & esp & _
              " WHERE Id=" & idSel

    End If

    Conn.Execute sql

    ' Retorna ao modo lista
    frmDados.Visible = False
    lstRegistros.Visible = True
    cmdConfirmar.Visible = False
    cmdCancelar.Visible = False
    cmdIncluir.Visible = True
    cmdEditar.Visible = True
    cmdExcluir.Visible = True

    CarregarLista

End Sub
Private Sub cmdCancelar_Click()

    frmDados.Visible = False
    lstRegistros.Visible = True

    cmdConfirmar.Visible = False
    cmdCancelar.Visible = False
    cmdIncluir.Visible = True
    cmdEditar.Visible = True
    cmdExcluir.Visible = True

End Sub

Private Function PegarIdSelecionado() As Long

    If lstRegistros.ListIndex < 0 Then
        PegarIdSelecionado = -1
        Exit Function
    End If

    Dim linha As String
    linha = lstRegistros.List(lstRegistros.ListIndex)

    ' Extrai tudo antes do primeiro " - "
    PegarIdSelecionado = CLng(Left(linha, InStr(linha, " - ") - 1))

End Function
Private Sub LimparCampos()

    txtNome.Text = ""
    txtIdade.Text = ""
    txtEndereco.Text = ""
    chkEspecial.Value = 0

End Sub
Private Sub txtIdade_KeyPress(KeyAscii As Integer)

    ' Permite números e Backspace
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

