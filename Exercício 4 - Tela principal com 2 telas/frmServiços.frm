VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmServiços 
   Caption         =   "Serviços"
   ClientHeight    =   7200
   ClientLeft      =   4995
   ClientTop       =   2100
   ClientWidth     =   7905
   LinkTopic       =   "Form2"
   ScaleHeight     =   7200
   ScaleWidth      =   7905
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   1440
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmDados 
      Caption         =   "Dados"
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   7695
      Begin VB.TextBox txtDescricao 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   1440
         Width           =   7455
      End
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   4575
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
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
         Left            =   6360
         TabIndex        =   3
         Top             =   600
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2143
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
      Begin VB.Label Label1 
         Caption         =   "Data do serviço"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Descrição completa do serviço"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor do serviço"
         Height          =   195
         Left            =   6360
         TabIndex        =   9
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   7695
      _Version        =   458752
      _ExtentX        =   13573
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
      MaxCols         =   4
      MaxRows         =   1
      SpreadDesigner  =   "frmServiços.frx":0000
   End
End
Attribute VB_Name = "frmServiços"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fmodoFormulario As String
Private fIdSelecionado As Long

Private Sub Form_Load()

AbreConexao
ConfigurarSpread
CarregarServicos
fmodoFormulario = ""
AlternarTelaPrincipal False

End Sub
Public Function Nz(Value As Variant, Optional DefaultValue As Variant) As Variant
    If IsNull(Value) Or IsEmpty(Value) Then
        Nz = DefaultValue
    Else
        Nz = Value
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
        .MaxCols = 4
        .MaxRows = 1
        .Row = 0
        .Col = 1: .ColWidth(1) = 5:  .Text = "Id"
        .Col = 2: .ColWidth(2) = 15: .Text = "Data do Serviço"
        .Col = 3: .ColWidth(3) = 15: .Text = "Cliente"
        .Col = 4: .ColWidth(4) = 10: .Text = "Valor"
        
    End With
End Sub
Private Sub CarregarComboCliente()
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT CLI_ID, CLI_NOME FROM CLIENTE ORDER BY CLI_NOME", pConn

    cmbCliente.Clear
    While Not rs.EOF
        Dim sDisplay As String
        sDisplay = CStr(rs!CLI_ID) & " - " & rs!CLI_NOME
        cmbCliente.AddItem sDisplay
        cmbCliente.ItemData(cmbCliente.NewIndex) = rs!CLI_ID
        rs.MoveNext
    Wend

    rs.Close
End Sub
Private Sub CarregarServicos()
    Dim rs As New ADODB.Recordset
    Dim sSql As String

    sSql = "SELECT S.SER_ID, S.SER_DATA_SERVICO, C.CLI_NOME AS CLIENTE, " & _
           "S.SER_VALOR, LEFT(S.SER_DESCRICAO, 50) AS Descricao " & _
           "FROM SERVICO S INNER JOIN CLIENTE C ON C.CLI_ID = S.CLI_ID " & _
           "ORDER BY S.SER_DATA_SERVICO DESC"

    rs.Open sSql, pConn, adOpenForwardOnly, adLockReadOnly

    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    fpSpread1.MaxRows = 0
    

    Dim i As Long: i = 1
    Do While Not rs.EOF
        If fpSpread1.MaxRows < i + 1 Then fpSpread1.MaxRows = fpSpread1.MaxRows + 1

        fpSpread1.Row = i
        fpSpread1.Col = 1: fpSpread1.Text = Nz(rs!SER_ID, 0)
        fpSpread1.Col = 2: fpSpread1.Text = Format$(Nz(rs!SER_DATA_SERVICO, 0), "dd/MM/yyyy")
        fpSpread1.Col = 3: fpSpread1.Text = Nz(rs!Cliente, "")
        fpSpread1.Col = 4: fpSpread1.Text = Format$(Nz(rs!SER_VALOR, 0), "###,##0.00")
        

        i = i + 1
        rs.MoveNext
    Loop

    rs.Close
End Sub
Private Sub cmdIncluir_Click()
    fmodoFormulario = "I"
    fIdSelecionado = 0

    LimparCampos
    CarregarComboCliente

    fpDateTime1.Text = ""
    txtDescricao.Text = ""
    fpDoubleSingle1.Value = 0

    AlternarTelaPrincipal True
End Sub
Private Sub cmdEditar_Click()
    fIdSelecionado = PegarIdSelecionado()
    If fIdSelecionado <= 0 Then
        MsgBox "Selecione um serviço para editar.", vbExclamation
        Exit Sub
    End If

    fmodoFormulario = "E"

    LimparCampos
    CarregarComboCliente
    CarregarDadosServico fIdSelecionado

    AlternarTelaPrincipal True
End Sub
Private Sub cmdExcluir_Click()
    Dim id As Long
    id = PegarIdSelecionado()

    If id <= 0 Then
        MsgBox "Selecione um serviço para excluir.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Deseja realmente excluir este serviço?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    pConn.Execute "DELETE FROM SERVICO WHERE SER_ID=" & id

    MsgBox "Serviço excluído com sucesso!", vbInformation

    CarregarServicos
End Sub
Private Sub cmdConfirmar_Click()
    On Error GoTo Trata

    '--- Validações obrigatórias ---
    If Trim$(cmbCliente.Text) = "" Then
        MsgBox "Selecione um cliente.", vbExclamation
        Exit Sub
    End If

    If Trim$(fpDateTime1.Text) = "" Then
        MsgBox "Informe a data do serviço.", vbExclamation
        Exit Sub
    End If

    If CDate(fpDateTime1.Text) > Date Then
        MsgBox "A data do serviço não pode ser futura.", vbExclamation
        Exit Sub
    End If

    If fpDoubleSingle1.Value <= 0 Then
        MsgBox "Informe o valor do serviço.", vbExclamation
        Exit Sub
    End If

    If Trim$(txtDescricao.Text) = "" Then
        MsgBox "Informe a descrição do serviço.", vbExclamation
        Exit Sub
    End If

    Dim sSql As String
    Dim idCliente As Long

    idCliente = cmbCliente.ItemData(cmbCliente.ListIndex)

    If fmodoFormulario = "I" Then
        
        sSql = "INSERT INTO SERVICO (SER_DATA_SERVICO, CLI_ID, SER_VALOR, SER_DESCRICAO) VALUES (" & _
               "'" & Format$(CDate(fpDateTime1.Text), "yyyy-MM-dd HH:mm:ss") & "', " & _
               idCliente & ", " & _
               Replace(fpDoubleSingle1.Value, ",", ".") & ", " & _
               "'" & Replace(txtDescricao.Text, "'", "''") & "')"

    ElseIf fmodoFormulario = "E" Then

        sSql = "UPDATE SERVICO SET " & _
               "SER_DATA_SERVICO='" & Format$(CDate(fpDateTime1.Text), "yyyy-MM-dd HH:mm:ss") & "', " & _
               "CLI_ID=" & idCliente & ", " & _
               "SER_VALOR=" & Replace(fpDoubleSingle1.Value, ",", ".") & ", " & _
               "SER_DESCRICAO='" & Replace(txtDescricao.Text, "'", "''") & "' " & _
               "WHERE SER_ID=" & fIdSelecionado

    End If

    pConn.Execute sSql

    MsgBox "Registro salvo com sucesso!", vbInformation

    CarregarServicos
    AlternarTelaPrincipal False
    Exit Sub

Trata:
    MsgBox "Erro ao salvar:" & vbCrLf & Err.Description, vbCritical

End Sub
Private Sub cmdCancelar_Click()
    LimparCampos
    AlternarTelaPrincipal False
End Sub
Private Sub LimparCampos()
    cmbCliente.ListIndex = -1
    fpDateTime1.Text = ""
    fpDoubleSingle1.Value = 0
    txtDescricao.Text = ""
End Sub
Private Sub CarregarDadosServico(ByVal idServico As Long)
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM SERVICO WHERE SER_ID=" & idServico, pConn

    If Not rs.EOF Then

        fpDateTime1.Text = Format$(rs!SER_DATA_SERVICO, "dd/MM/yyyy")

        fpDoubleSingle1.Value = Nz(rs!SER_VALOR, 0)

        txtDescricao.Text = Nz(rs!SER_DESCRICAO, "")

        '--- Selecionar cliente no combo ---
        Dim i As Long
        For i = 0 To cmbCliente.ListCount - 1
            If cmbCliente.ItemData(i) = rs!CLI_ID Then
                cmbCliente.ListIndex = i
                Exit For
            End If
        Next i

    End If

    rs.Close
End Sub

Private Sub fpSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
    On Error GoTo Trata

    '--- Linha inválida ou cabeçalho ---
    If Row <= 0 Then Exit Sub

    fpSpread1.Row = Row
    fpSpread1.Col = 1
    Dim sId As String
    sId = Trim$(fpSpread1.Text)

    If Len(sId) = 0 Then Exit Sub

    fIdSelecionado = CLng(Val(sId))
    If fIdSelecionado <= 0 Then Exit Sub

    '--- Abrir modo edição ---
    fmodoFormulario = "E"
    LimparCampos
    CarregarComboCliente
    CarregarDadosServico fIdSelecionado
    AlternarTelaPrincipal True
    Exit Sub

Trata:
    MsgBox "Erro ao selecionar o serviço:" & vbCrLf & Err.Description, vbExclamation
End Sub



