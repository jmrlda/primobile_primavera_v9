VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EncomendaForm 
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14415
   OleObjectBlob   =   "EncomendaForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EncomendaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


   
  Dim obj As New GcpBELinhaDocumentoVenda
      Dim SQL As String
    Dim objLista As StdBELista
   
   Dim encomenda_lista(50) As New encomenda
   Dim encomenda_obj As New encomenda
    







    Dim objArtigo As GcpBEArtigo
      Dim editor As New GcpBELinhaDocumentoVenda
Private Sub btnCarregar_Click()
 If Me.lstEncomendaItem.ListCount > 0 Then
 export_encomenda_item_str
 Unload Me
 End If
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub lblEncomenda_Click()
    encomenda.Add
End Sub

Private Sub lblUsuario_Click()

End Sub

Private Sub lstEncomenda_Click()

    Dim enc As String
    'procurar pela linha selecionada
    For x = 0 To lstEncomenda.ListCount - 1
        limpar_cabecalho
        If lstEncomenda.Selected(x) = True Then
            Me.lstEncomendaItem.Clear
            enc = lstEncomenda.List(x, 0)
            preencher_encomendaItem (enc)
            preencher_cabecalho (x)
            Exit For
        End If
    Next x
    
End Sub


Private Sub UserForm_Activate()


          SQL = "select cli.nome as cliente, util.CDU_nome as vendedor, util.CDU_documento , enc.*   from TDU_primobEncomenda as enc, Clientes as cli, TDU_primobUtilizador as util where enc.CDU_cliente = cli.Cliente and enc.CDU_vendedor = util.CDU_utilizador and enc.CDU_estado='pendente'"
        Set objLista = Aplicacao.BSO.Consulta(SQL)
        Dim i As Integer
        Dim valor, estado, data, enc, nome As String
            
       
       i = 0
    
    If Not (objLista Is Nothing) Then
        While Not (objLista.NoInicio Or objLista.NoFim)    'existe registo
        
             encomenda_obj.cliente = objLista("cliente")
            
               encomenda_obj.id = objLista("CDU_encomenda")
               encomenda_obj.data = objLista("CDU_data_hora")
               encomenda_obj.valorTotal = objLista("CDU_valor")
               encomenda_obj.estado = objLista("CDU_estado")
                 encomenda_obj.documento = objLista("CDU_documento")
                 encomenda_obj.vendedor = objLista("vendedor")
            
           Me.lstEncomenda.AddItem (encomenda_obj.id)
           
           Me.lstEncomenda.List(i, 1) = encomenda_obj.cliente
           Me.lstEncomenda.List(i, 2) = encomenda_obj.valorTotal
           Me.lstEncomenda.List(i, 3) = encomenda_obj.data
           Me.lstEncomenda.List(i, 4) = encomenda_obj.estado
           Me.lstEncomenda.List(i, 5) = objLista("CDU_cliente")
           Me.lstEncomenda.List(i, 6) = objLista("CDU_vendedor")
           
            Set encomenda_lista(i) = encomenda_obj
            i = i + 1


             objLista.Seguinte
             
        Wend
    End If
End Sub

Private Sub preencher_encomendaItem(encomenda As String)
              SQL = "select a.Descricao, item.* from TDU_primobItemEncomenda as item, Artigo as a  where CDU_encomenda = '" + encomenda + "' and item.CDU_artigo = a.Artigo"
        Set objLista = Aplicacao.BSO.Consulta(SQL)
        Dim i As Integer
        Dim valor, estado, data, enc As String
            
       
       i = 0
    
    If Not (objLista Is Nothing) Then
        While Not (objLista.NoInicio Or objLista.NoFim)    'existe registo
        
            Artigo = objLista("CDU_artigo")
            Descricao = objLista("Descricao")
            quantidade = objLista("CDU_quantidade")
           
           Me.lstEncomendaItem.AddItem (Artigo)
           Me.lstEncomendaItem.List(i, 1) = Descricao
           Me.lstEncomendaItem.List(i, 2) = quantidade
           
            i = i + 1


             objLista.Seguinte
             
        Wend
    End If
End Sub

Private Sub preencher_cabecalho(x As Integer)
                lblValorTotal.Caption = lstEncomenda.List(x, 2)
            lblCliente.Caption = lstEncomenda.List(x, 1)
            lblEncomenda.Caption = lstEncomenda.List(x, 0)
            lblEstado.Caption = lstEncomenda.List(x, 3)
            lblUsuario.Caption = encomenda_lista(x).vendedor
            lblDocumento.Caption = encomenda_lista(x).documento
            
            
    lblCliente.Visible = True
    lblDocumento.Visible = True
    lblEncomenda.Visible = True
    lblEstado.Visible = True
    lblUsuario.Visible = True
    lblValorTotal.Visible = True
End Sub


Private Sub limpar_cabecalho()
            
    lblCliente.Visible = False
    lblDocumento.Visible = False
    lblEncomenda.Visible = False
    lblEstado.Visible = False
    lblUsuario.Visible = False
    lblValorTotal.Visible = False

End Sub

Private Function check_encomenda_selecionada() As Boolean

    Dim rv As Boolean
    rv = False
      For x = 0 To lstEncomenda.ListCount - 1
        limpar_cabecalho
        If lstEncomenda.Selected(x) = True Then
        rv = True
            Exit For
        End If
    Next x
    
    check_encomenda_selecionada = rv
    

End Function


' verifcar se a encomenda selecionada possui itens para processar
' resultado correto se selecionado uma encomenda antes da verificacao
Private Function has_encomenda_item() As Boolean

    Dim rv As Boolean
    rv = False
    
 If Me.lstEncomendaItem.ListCount > 0 Then
    rv = True
 End If
 
     check_encomenda_selecionada = rv
End Function


' Retornar string com  itens da encomenda
' separados por ponto e virgula
Private Function export_encomenda_item_str() As String
 Dim encomenda_item As String
 For x = 0 To lstEncomendaItem.ListCount - 1
    encomenda_item = encomenda_item + lstEncomendaItem.List(x, 0)
    encomenda_item = encomenda_item + ";"
    Next x
    
For x = 0 To lstEncomenda.ListCount - 1
    
Next x

    EditorVendas.encomenda_str = encomenda_item
    EditorVendas.totalEncomenda = lstEncomendaItem.ListCount
    Dim index As Integer
    index = get_index_encomenda_selecionada
    EditorVendas.cliente = lstEncomenda.List(index, 5)
    EditorVendas.vendedor = lstEncomenda.List(index, 6)
    EditorVendas.encomenda_id = lstEncomenda.List(index, 0)


End Function

' Retornar Lista  de itens da encomenda
'''Private Function get_encomenda_item_lst() As List
 '   get_encomenda_item_lst = lstEncomendaItem.List
'End Function
Private Function get_index_encomenda_selecionada() As Integer

    Dim rv As Integer
    rv = -1
    For x = 0 To lstEncomenda.ListCount - 1
        If lstEncomenda.Selected(x) = True Then
            rv = x
            Exit For
        End If
    Next x
    
    get_index_encomenda_selecionada = rv
    

End Function


