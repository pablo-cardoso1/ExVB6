VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmClientes 
   Caption         =   "frmClientes"
   ClientHeight    =   6240
   ClientLeft      =   3930
   ClientTop       =   2730
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   1560
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   615
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   2775
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   9135
      _Version        =   458752
      _ExtentX        =   16113
      _ExtentY        =   4895
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   1
      SpreadDesigner  =   "Form1.frx":0000
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame frmDados 
      Caption         =   "Dados"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9135
      Begin VB.Frame frmComunicacao 
         Caption         =   "Tipo de comunicação"
         Height          =   735
         Left            =   3360
         TabIndex        =   20
         Top             =   1200
         Width           =   4575
         Begin VB.OptionButton opbCarta 
            Caption         =   "Carta"
            Height          =   195
            Left            =   3360
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton opbWhats 
            Caption         =   "WhatsApp"
            Height          =   195
            Left            =   2160
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton opbFone 
            Caption         =   "Telefone"
            Height          =   195
            Left            =   1080
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton opbEmail 
            Caption         =   "Email"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   1455
         _Version        =   196608
         _ExtentX        =   2566
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   1
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "18/09/2020"
         DateCalcMethod  =   0
         DateTimeFormat  =   0
         UserDefinedFormat=   ""
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDoubleSingle fpDoubleSingle1 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
         _Version        =   196608
         _ExtentX        =   1931
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         DecimalPlaces   =   -1
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.ComboBox cmbSexo 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox chbEspecial 
         Caption         =   "Especial"
         Height          =   255
         Left            =   8040
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtEndereco 
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtIdade 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3720
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Valor de crédito"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Sexo"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Endereço completo"
         Height          =   255
         Left            =   4320
         TabIndex        =   16
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Idade"
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fmodoFormulario As String
Private fIdSelecionado As Long

Private Enum eSexo
    eSexoMasculino = 1
    eSexoFeminino = 2
    eSexoIndefinido = 3
End Enum

Private Enum eTipoComunicacao
    eEmail = 1
    eTelefone = 2
    eWhatsApp = 3
    eCarta = 4
End Enum
Private Sub Form_Load()
    AbreConexao
    ConfigurarSpread
    CarregarClientes
    fmodoFormulario = ""
    AlternarTelaPrincipal False

    ' Preenche ComboBox Sexo
    cmbSexo.Clear
    cmbSexo.AddItem "Masculino"
    cmbSexo.AddItem "Feminino"
    cmbSexo.AddItem "Indefinido"
    cmbSexo.ListIndex = 2
End Sub
Public Function Nz(Value As Variant, Optional DefaultValue As Variant) As Variant
    If IsNull(Value) Or IsEmpty(Value) Then
        Nz = DefaultValue
    Else
        Nz = Value
    End If
End Function
Private Function CalculaIdade(dataNasc As Date) As Long
    Dim hoje As Date
    hoje = Date
    CalculaIdade = Year(hoje) - Year(dataNasc)
    If Month(hoje) < Month(dataNasc) Or _
       (Month(hoje) = Month(dataNasc) And Day(hoje) < Day(dataNasc)) Then
        CalculaIdade = CalculaIdade - 1
    End If
End Function
Private Function PegarIdSelecionado() As Long
    On Error GoTo Trata
    If fpSpread1.ActiveRow <= 0 Then
        PegarIdSelecionado = -1
        Exit Function
    End If

    fpSpread1.Row = fpSpread1.ActiveRow
    fpSpread1.Col = 1
    Dim sId As String
    sId = Trim$(fpSpread1.Text)
    If Len(sId) = 0 Then
        PegarIdSelecionado = -1
    Else
        PegarIdSelecionado = CLng(Val(sId))
    End If
    Exit Function
Trata:
    PegarIdSelecionado = -1
End Function
Private Sub AlternarTelaPrincipal(modoFormulario As Boolean)
    fpSpread1.Visible = Not modoFormulario
    frmDados.Visible = modoFormulario

    cmdIncluir.Visible = Not modoFormulario
    cmdEditar.Visible = Not modoFormulario
    cmdExcluir.Visible = Not modoFormulario
    cmdConfirmar.Visible = modoFormulario
    cmdCancelar.Visible = modoFormulario
End Sub
Private Sub ConfigurarSpread()
    With fpSpread1
        .MaxCols = 9
        .MaxRows = 1
        .Row = 0
        .Col = 1: .ColWidth(1) = 5:  .Text = "Id"
        .Col = 2: .ColWidth(2) = 10: .Text = "Nome"
        .Col = 3: .ColWidth(3) = 8: .Text = "Data"
        .Col = 4: .ColWidth(4) = 5: .Text = "Idade"
        .Col = 5: .ColWidth(5) = 18: .Text = "Endereço"
        .Col = 6: .ColWidth(6) = 5: .Text = "Sexo"
        .Col = 7: .ColWidth(7) = 5: .Text = "Crédito"
        .Col = 8: .ColWidth(8) = 9: .Text = "Comunicação"
        .Col = 9: .ColWidth(9) = 7: .Text = "Especial"
        
    End With
End Sub
Private Sub CarregarClientes()
    On Error GoTo Trata
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Dim i As Long

    sSql = "SELECT CLI_ID, CLI_NOME, CLI_DATA_NASCIMENTO, CLI_ENDERECO, " & _
           "CLI_SEXO, CLI_VALOR_CREDITO, CLI_TIPO_COMUNICACAO, CLI_ESPECIAL " & _
           "FROM CLIENTE ORDER BY CLI_NOME"

    Set rs = New ADODB.Recordset
    rs.Open sSql, pConn, adOpenForwardOnly, adLockReadOnly

    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    fpSpread1.MaxRows = 0
    

    If rs.EOF Then Exit Sub
    i = 1

    Do While Not rs.EOF
        If fpSpread1.MaxRows < i + 1 Then fpSpread1.MaxRows = fpSpread1.MaxRows + 1
        fpSpread1.Row = i

        ' ID
        fpSpread1.Col = 1: fpSpread1.Text = Nz(rs!CLI_ID, 0)

        ' Nome
        fpSpread1.Col = 2: fpSpread1.Text = Nz(rs!CLI_NOME, "")

        ' Data Nascimento
        fpSpread1.Col = 3
        If IsNull(rs!CLI_DATA_NASCIMENTO) Then
            fpSpread1.Text = ""
        Else
            fpSpread1.Text = Format$(rs!CLI_DATA_NASCIMENTO, "dd/MM/yyyy")
        End If

        ' Idade
        fpSpread1.Col = 4
        If IsNull(rs!CLI_DATA_NASCIMENTO) Then
            fpSpread1.Text = ""
        Else
            fpSpread1.Text = CalculaIdade(rs!CLI_DATA_NASCIMENTO)
        End If

        ' Endereço
        fpSpread1.Col = 5: fpSpread1.Text = Nz(rs!CLI_ENDERECO, "")

        ' Sexo
        fpSpread1.Col = 6
        Select Case Nz(rs!CLI_SEXO, eSexoIndefinido)
            Case eSexoMasculino: fpSpread1.Text = "Masculino"
            Case eSexoFeminino:  fpSpread1.Text = "Feminino"
            Case Else:           fpSpread1.Text = "Indefinido"
        End Select

        ' Crédito
        fpSpread1.Col = 7: fpSpread1.Text = Format$(Nz(rs!CLI_VALOR_CREDITO, 0), "###,##0.00")

        ' Tipo Comunicação
        fpSpread1.Col = 8
        Select Case Nz(rs!CLI_TIPO_COMUNICACAO, eWhatsApp)
            Case eEmail:     fpSpread1.Text = "E-mail"
            Case eTelefone:  fpSpread1.Text = "Telefone"
            Case eWhatsApp:  fpSpread1.Text = "WhatsApp"
            Case eCarta:     fpSpread1.Text = "Carta"
            Case Else:       fpSpread1.Text = "N/D"
        End Select

        ' Especial
        fpSpread1.Col = 9
        If CBool(Nz(rs!CLI_ESPECIAL, 0)) Then
            fpSpread1.Text = "Sim"
        Else
            fpSpread1.Text = "Não"
        End If

        i = i + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Exit Sub
Trata:
    MsgBox "Erro ao carregar clientes:" & vbCrLf & Err.Description, vbCritical, "ERRO"
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Sub
Private Sub CarregarDadosCliente(ByVal idCliente As Long)
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM CLIENTE WHERE CLI_ID=" & idCliente, pConn

    If Not rs.EOF Then
        txtNome.Text = Nz(rs!CLI_NOME, "")
        txtEndereco.Text = Nz(rs!CLI_ENDERECO, "")

        If Not IsNull(rs!CLI_DATA_NASCIMENTO) Then
            fpDateTime1.Text = Format$(rs!CLI_DATA_NASCIMENTO, "dd/MM/yyyy")
            txtIdade.Text = CalculaIdade(rs!CLI_DATA_NASCIMENTO)
        Else
            fpDateTime1.Text = ""
            txtIdade.Text = ""
        End If

        fpDoubleSingle1.Value = Nz(rs!CLI_VALOR_CREDITO, 0)

        Select Case Nz(rs!CLI_SEXO, eSexoIndefinido)
            Case eSexoMasculino: cmbSexo.ListIndex = 0
            Case eSexoFeminino:  cmbSexo.ListIndex = 1
            Case Else:           cmbSexo.ListIndex = 2
        End Select

        Select Case Nz(rs!CLI_TIPO_COMUNICACAO, eWhatsApp)
            Case eEmail:     opbEmail.Value = True
            Case eTelefone:  opbFone.Value = True
            Case eWhatsApp:  opbWhats.Value = True
            Case eCarta:     opbCarta.Value = True
        End Select

        chbEspecial.Value = IIf(CBool(Nz(rs!CLI_ESPECIAL, 0)), vbChecked, vbUnchecked)
    End If

    rs.Close
End Sub
Private Sub LimparCampos()
    txtNome.Text = ""
    txtEndereco.Text = ""
    fpDateTime1.Text = ""
    fpDoubleSingle1.Value = 0
    cmbSexo.ListIndex = 2
    opbEmail.Value = False
    opbFone.Value = False
    opbWhats.Value = True
    opbCarta.Value = False
    chbEspecial.Value = vbUnchecked
    txtIdade.Text = ""
End Sub
Private Sub cmdIncluir_Click()
    fmodoFormulario = "I"
    LimparCampos
    fIdSelecionado = 0
    AlternarTelaPrincipal True
    txtNome.SetFocus
End Sub
Private Sub cmdEditar_Click()
    fIdSelecionado = PegarIdSelecionado()
    If fIdSelecionado <= 0 Then
        MsgBox "Selecione um registro para editar.", vbExclamation
        Exit Sub
    End If
    fmodoFormulario = "E"
    CarregarDadosCliente fIdSelecionado
    AlternarTelaPrincipal True
End Sub
Private Sub cmdCancelar_Click()
    LimparCampos
    fmodoFormulario = ""
    fIdSelecionado = 0
    AlternarTelaPrincipal False
End Sub
Private Sub cmdExcluir_Click()
    fIdSelecionado = PegarIdSelecionado()
    If fIdSelecionado <= 0 Then
        MsgBox "Selecione um registro para excluir.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir este registro?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    pConn.Execute "DELETE FROM CLIENTE WHERE CLI_ID=" & fIdSelecionado
    CarregarClientes
End Sub
Private Sub cmdConfirmar_Click()
    Dim sql As String
    Dim sexo As Integer
    Dim tipoComunicacao As Integer
    Dim valorCredito As Double
    Dim dataNasc As Date
    Dim especial As Integer

    ' Validações
    If Trim(txtNome.Text) = "" Or Trim(txtEndereco.Text) = "" Then
        MsgBox "Preencha os campos obrigatórios!", vbExclamation
        Exit Sub
    End If

    If fpDateTime1.Text = "" Then
        MsgBox "Informe a data de nascimento!", vbExclamation
        Exit Sub
    Else
        dataNasc = CDate(fpDateTime1.Text)
    End If

    valorCredito = fpDoubleSingle1.Value
    If valorCredito < 0 Or valorCredito > 5000 Then
        MsgBox "Valor de crédito inválido!", vbExclamation
        Exit Sub
    End If

    sexo = cmbSexo.ListIndex + 1

    If opbEmail.Value Then tipoComunicacao = eEmail
    If opbFone.Value Then tipoComunicacao = eTelefone
    If opbWhats.Value Then tipoComunicacao = eWhatsApp
    If opbCarta.Value Then tipoComunicacao = eCarta

    especial = IIf(chbEspecial.Value = vbChecked, 1, 0)

    ' SQL
    If fmodoFormulario = "I" Then
        sql = "INSERT INTO CLIENTE (CLI_NOME, CLI_DATA_NASCIMENTO, CLI_ENDERECO, " & _
              "CLI_SEXO, CLI_VALOR_CREDITO, CLI_TIPO_COMUNICACAO, CLI_ESPECIAL) VALUES (" & _
              "'" & Replace(txtNome.Text, "'", "''") & "', " & _
              "'" & Format$(dataNasc, "yyyy-mm-dd") & "', " & _
              "'" & Replace(txtEndereco.Text, "'", "''") & "', " & _
              sexo & ", " & valorCredito & ", " & tipoComunicacao & ", " & especial & ")"
    ElseIf fmodoFormulario = "E" Then
        sql = "UPDATE CLIENTE SET " & _
              "CLI_NOME='" & Replace(txtNome.Text, "'", "''") & "', " & _
              "CLI_DATA_NASCIMENTO='" & Format$(dataNasc, "yyyy-mm-dd") & "', " & _
              "CLI_ENDERECO='" & Replace(txtEndereco.Text, "'", "''") & "', " & _
              "CLI_SEXO=" & sexo & ", " & _
              "CLI_VALOR_CREDITO=" & valorCredito & ", " & _
              "CLI_TIPO_COMUNICACAO=" & tipoComunicacao & ", " & _
              "CLI_ESPECIAL=" & especial & " " & _
              "WHERE CLI_ID=" & fIdSelecionado
    End If

    pConn.Execute sql
    CarregarClientes
    LimparCampos
    fmodoFormulario = ""
    fIdSelecionado = 0
    AlternarTelaPrincipal False
End Sub
Private Sub fpDateTime1_LostFocus()
    If fpDateTime1.Text <> "" Then
        txtIdade.Text = CalculaIdade(CDate(fpDateTime1.Text))
    Else
        txtIdade.Text = ""
    End If
End Sub
Private Sub fpSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
    fIdSelecionado = PegarIdSelecionado()
    If fIdSelecionado <= 0 Then Exit Sub
    fmodoFormulario = "E"
    CarregarDadosCliente fIdSelecionado
    AlternarTelaPrincipal True
End Sub


