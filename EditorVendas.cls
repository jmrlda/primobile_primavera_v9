VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EditorVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public encomenda_str As String
Public totalEncomenda As Integer
Public cliente As String
Public documento As String
Public vendedor As String
Public encomenda_id As String

Dim enc As String
Dim encomenda_lst() As String
Dim obj As New GcpBELinhaDocumentoVenda
Dim SQL As String
Dim objArtigo As GcpBEArtigo


'Instancia BaseDados
    Dim con As New ADODB.Connection
    Dim rec As New ADODB.Recordset
    Dim ins As New ADODB.Command
    





' Depois que identificar o cliente preencher a linha do editor de venda
' com os itens da sua encomenda
Private Sub EditorVendas_ClienteIdentificado(cliente As String, Cancel As Boolean)
    If Me.DocumentoVenda.TipoDoc = "VD" Then
        importar_encomenda_para_editor
        
    End If

End Sub




' Actualizar encomenda na basedados como encomenda ja processada .

Private Sub EditorVendas_DepoisDeGravar(Filial As String, Serie As String, Tipo As String, NumDoc As Long)
    actualizar_encomenda
End Sub

' Depois que o documento for identificado como encomenda remota
' abrir form para listagens e carregamento  das encomendas para editor
' e importar o respectivo cliente
Private Sub EditorVendas_TipoDocumentoIdentificado(Tipo As String, Cancel As Boolean)
    If Tipo = "VD" Then
        EncomendaForm.Show
        importar_entidade_para_editor
    End If

End Sub



' importar encomendas armazenados em cache
Private Sub importar_encomenda_para_editor()
    If Len(Me.encomenda_str) > 0 Then

        encomenda_lst = Split(encomenda_str, ";")
            Set obj = Nothing
            For i = 0 To totalEncomenda - 1
                enc = Trim(encomenda_lst(i))
                Set obj = Nothing
                Set objArtigo = Nothing
                   
                Set objArtigo = Aplicacao.BSO.Comercial.Artigos.Consulta(enc)
                obj.Artigo = objArtigo.Artigo
                    
                obj.quantidade = 3
                obj.TipoLinha = 10
                obj.Descricao = objArtigo.Descricao
                obj.PrecUnit = objArtigo.PCUltimo
                obj.Armazem = objArtigo.ArmazemSugestao
                obj.Localizacao = objArtigo.LocalizacaoSugestao
                obj.CodIva = objArtigo.Iva
                obj.vendedor = Me.vendedor
                     
                Me.DocumentoVenda.Linhas.Insere obj
         
                  Next i
        
    End If


End Sub


' importar entidade 'cliente' da encomenda a ser
' processado no editor de venda
Private Sub importar_entidade_para_editor()
  Me.DocumentoVenda.Entidade = Me.cliente
  
End Sub



' atualizar estado da encomenda de 'pendente' para 'processado'

Private Sub actualizar_encomenda()
    SQL = "update TDU_primobEncomenda set CDU_estado = 'processado' where CDU_encomenda = '" + encomenda_id + "'"
'    Aplicacao.BSO.Consulta (SQL)
        con.Open "Driver=SQL Server;Server=.\PRIMAVERA9;Database=PRIPRITERRA;uid=sa;pwd=jmr2013!"
    con.Execute SQL
    con.Close

End Sub
