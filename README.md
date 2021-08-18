Private Banco As ADODB.Connection
Private Tabela As ADODB.Recordset

Public Resultado() As Variant

Private Function Conectar() As Boolean

On Error GoTo Erro

    Dim Caminho As String, Arquivo As String
    
    Caminho = ThisWorkbook.Path
    Arquivo = "Banco_SQL.accdb"
    
    Set Banco = New ADODB.Connection
    Banco.Provider = "Microsoft.ACE.OLEDB.12.0"
    Banco.ConnectionString = "Data Source=\\192.168.1.201\CREDSOFT\Banco_SQL.accdb"
    Banco.Open
   
Erro:
    
    If VBA.Err Then
        Conectar = False
        MsgBox "Erro na Função Conectar" & VBA.Chr(10) & VBA.Chr(10) & "Erro: " & _
        VBA.Err.Description, vbCritical, "ATENÇÃO!"
    Else
        Conectar = True
    End If
    
End Function


Public Function ExecutarSQL(Sql As String) As Boolean

On Error GoTo Erro
    ExecutarSQL = False
    
    If Conectar = False Then Exit Function

    Banco.Execute Sql
    
    Call Desconectar
    
    VBA.Err = Empty
    
Erro:
    
    If VBA.Err Then
        MsgBox "Erro Executar SQL!" & VBA.Chr(10) & VBA.Chr(10) & _
        VBA.Err.Description & VBA.Chr(10) & VBA.Chr(10) & "SQL: " & Sql, vbCritical, "ATENÇÃO!!!"
    Else
        ExecutarSQL = True
    End If
    
End Function

Private Sub Desconectar()

    If Not Tabela Is Nothing Then
        If Tabela.State = 1 Then Tabela.Close
        Set Tabela = Nothing
    End If
    
    If Not Banco Is Nothing Then
        If Banco.State = 1 Then Banco.Close
        Set Banco = Nothing
    End If
        
End Sub

Private Function SQL_INSERT(Tabela As String, c() As String, v() As String) As String
    
    Dim Sql As String
    
    Sql = "INSERT INTO " & Tabela & " (" & VBA.Join(c, ", ") & ") "
    Sql = Sql & "VALUES('" & VBA.Join(v, "', '") & "')"
    
    SQL_INSERT = Sql
    
End Function


Public Function Salvar(Tabela As String, c() As String, v() As String, Id As String) As Boolean
    
    Dim Sql As String
    
    If Id = Empty Then
        Sql = SQL_INSERT(Tabela, c, v)
    Else
        Sql = SQL_UPDATE(Tabela, c, v, Id)
    End If
    
    Sql = Replace(Sql, "'null'", "null")
    Sql = Replace(Sql, "'%$#@", Empty)
    Sql = Replace(Sql, "%$#@'", Empty)
        
    Salvar = ExecutarSQL(Sql)
    
End Function


Private Function SQL_UPDATE(Tabela As String, c() As String, v() As String, Id As String) As String
    
    Dim Sql As String
    
    Dim Total As Integer, i As Integer
    Total = UBound(c, 1)
        
    For i = 1 To Total
        c(i) = c(i) & " = '" & v(i) & "'"
    Next
    
    Sql = "UPDATE " & Tabela & " SET " & VBA.Join(c, ", ")
    Sql = Sql & " WHERE " & Tabela & "_Id = " & Id
    
    SQL_UPDATE = Sql
    
End Function


Public Function CvData(Data As String) As String

    If VBA.Trim(Data) = Empty Then
        CvData = "null"
    Else
        CvData = "%$#@" & VBA.CLng(VBA.CDate(Data)) & "%$#@"
    End If

End Function


Public Function CvNum(valor As String) As String
    
    If VBA.Trim(valor) = Empty Then
        CvNum = "0"
    Else
        
        valor = VBA.Replace(valor, "R", Empty)
        valor = VBA.Replace(valor, "$", Empty)
        valor = VBA.Replace(valor, ".", Empty)
        
        valor = VBA.Trim(valor)
        
        CvNum = valor
        
    End If
    
End Function



Public Function BuscarSQL(Sql As String, Optional Guia As String = Empty, _
                        Optional Celula As String = Empty) As Long
    
    BuscarSQL = 0
    
    If Conectar = False Then Exit Function

On Error GoTo Erro

    Set Tabela = New ADODB.Recordset
    Tabela.CursorLocation = adUseClient
    Tabela.CursorType = adOpenStatic
    
    Tabela.Open Sql, Banco
    
    BuscarSQL = Tabela.RecordCount
    
    If Guia = Empty Then
        Resultado = Tabela.GetRows
    Else
        Sheets(Guia).Range(Celula).CopyFromRecordset Tabela
    End If

Erro:

    Call Desconectar
    
    If VBA.Err Then
        MsgBox "Erro Buscar SQL!" & VBA.Chr(10) & VBA.Chr(10) & VBA.Err.Description & _
        VBA.Chr(10) & VBA.Chr(10) & "SQL: " & Sql, vbCritical, "ATENÇÃO!"
    End If
    
End Function

Function AcSQL(valor As String) As String
    
    Dim n, t, v
    t = ""
    
    For n = 1 To VBA.Len(valor)
    
        v = VBA.Asc(VBA.Mid(valor, n, 1))
        
        Select Case v
            
            Case 39: t = t & "''"
            Case 65: t = t & "[ÁÀÂÄÃA]"
            Case 67: t = t & "[ÇC]"
            Case 69: t = t & "[ÉÈÊËE]"
            Case 73: t = t & "[ÍÌÎÏI]"
            Case 79: t = t & "[ÓÒÔÖÕO]"
            Case 85: t = t & "[ÚÙÛÜU]"
            Case 97: t = t & "[áàâäãa]"
            Case 99: t = t & "[çc]"
            Case 101: t = t & "[éèêëe]"
            Case 105: t = t & "[íìîïi]"
            Case 111: t = t & "[óòôöõo]"
            Case 117: t = t & "[úùûüu]"
            
            Case Else
            
                If v > 31 And v < 127 Then
                    t = t & VBA.Chr(v)
                Else
                    t = t & "_"
                End If
                
        End Select
        
    Next
    
    AcSQL = t

End Function
