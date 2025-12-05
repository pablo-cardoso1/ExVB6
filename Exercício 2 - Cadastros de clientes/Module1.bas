Attribute VB_Name = "Module1"
Public Conn As ADODB.Connection

Public Sub AbreConexao()

    Set Conn = New ADODB.Connection

    Conn.ConnectionString = _
        "Provider=SQLOLEDB;" & _
        "Data Source=DEV-PABLO\PDVNET;" & _
        "Initial Catalog=Exemplos;" & _
        "User ID=sa;" & _
        "Password=inter#system;" & _
        "Encrypt=Yes;" & _
        "TrustServerCertificate=Yes;"

    Conn.Open

End Sub


