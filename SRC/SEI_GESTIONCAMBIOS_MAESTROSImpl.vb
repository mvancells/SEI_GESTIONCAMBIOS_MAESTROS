Imports System.Net
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports SEI_GESTIONCAMBIOS_MAESTROS.My
Public Class SEI_GESTIONCAMBIOS_MAESTROSImpl
    Inherits ModuleBase.AbstractModuleImplementation

#Region "Parent Module Property"
    Private _oModule As SEI_GESTIONCAMBIOS_MAESTROS

    Public Property oModule() As SEI_GESTIONCAMBIOS_MAESTROS
        Get
            Return _oModule
        End Get
        Set(ByVal value As SEI_GESTIONCAMBIOS_MAESTROS)
            _oModule = value
        End Set
    End Property
#End Region

#Region "Responses"

    Public Class SingleResponse
        Public Property ExecutionSuccess As Boolean
        Public Property FailureReason As String
    End Class

    Public Class SingleResponseCampos
        Inherits SingleResponse
        Public Property ListaCampos As New List(Of DetalleCampos)
    End Class

    Public Structure DetalleCampos
        Public Property Tabla As String
        Public Property CampoBD As String
        Public Property CampoObjeto As String
        Public Property EsEspecial As String
        Public Property TablaAsociada As String
        Public Property CampoAsociado As String
        Public Property Descripcion As String
    End Structure


#End Region

#Region "Estructuras"
    Public Structure enMenuUID
        Const MNU_Siguiente As String = "1288"
        Const MNU_Anterior As String = "1289"
        Const MNU_Primero As String = "1290"
        Const MNU_Ultimo As String = "1291"
        Const MNU_Buscar As String = "1281"
        Const MNU_Crear As String = "1282"
        Const MNU_AñadirLinea As String = "1292"
        Const MNU_EliminarLinea As String = "1293"
        Const MNU_DuplicarLinea As String = "1294"
    End Structure

    Public Structure Grid

        Const DT As String = "DT_Grid"
        Const Grid As String = "gDatos"
        Const Col_Code As String = "Code"
        Const Col_Campo As String = "Nombre del Campo"
        Const Col_Antiguo As String = "Valor Anterior"
        Const Col_Nuevo As String = "Valor Nuevo"
        Const Col_Sel As String = "Selección"
        Const Col_IdUsuario As String = "U_IdUsuario"
        Const Col_NombreUsuario As String = "Nombre de usuario"
        Const Col_Fecha As String = "Fecha del Cambio"
        Const Col_Estado As String = "Estado"
        Const Col_CampoBD As String = "U_CampoBBDD"
        Const Col_Tabla As String = "U_Tabla"
        Const Col_CampoClave As String = "Campo Clave"
        Const Col_ValorCampoClave As String = "Valor Campo Clave"
        Const Col_PropiedadObjeto As String = "U_Objeto"

    End Structure

    Public Structure Cab
        Const txtCardCode As String = "txtCCode"
        Const txtCardName As String = "txtCName"
        Const Boton_Link As String = "bLink"
        Const Boton_Filtrar As String = "bFiltrar"
        Const Boton_Grabar As String = "btnGrab"
        Const Boton_Descartar As String = "btnDes"
        Const Choose_CardCode As String = "cflBP"
        Const UDS_CardCode As String = "dsCCode"
        Const UDS_CardName As String = "dsCName"
    End Structure

#End Region
    Public Property UsuarioConPermisosIC As Boolean = False
    Public Property UsuarioConPermisosItems As Boolean = False
    Public Property YaEntre As Boolean = False


    'TODO: Write your implementation here...
#Region "Formulario IC"

    Public Sub HANDLE_OK_BUTTON_PRESSED_BEFORE_ACTION_BPFORM(ByVal pVal As BusinessObjectInfo)
        Dim oForm As SAPbouiCOM.Form
        oForm = oApplication.Forms.Item(pVal.FormUID)
        Try
            'Paso 1: revisar si el usuario tiene autorización para grabar directamente en el formulario
            Dim oResponseAutorizacion As New SingleResponse
            oResponseAutorizacion = getUserAuthorization(oCompany.UserName, pVal.FormTypeEx)
            If oResponseAutorizacion.ExecutionSuccess = False AndAlso oResponseAutorizacion.FailureReason.StartsWith("ERROR:") Then
                oApplication.MessageBox("Su usuario no tiene autorización para cambiar ciertos datos maestros de este formulario" & vbNewLine & "Sus cambios se guardarán para que otro usuario con permisos los pueda autorizar")
                UsuarioConPermisosIC = False
            Else
                UsuarioConPermisosIC = True
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub HANDLE_OK_BUTTON_PRESSED_AFTER_ACTION_BPFORM(ByVal pVal As BusinessObjectInfo)
        YaEntre = True
        Dim oForm As SAPbouiCOM.Form
        oForm = oApplication.Forms.Item(pVal.FormUID)
        Dim oBp As SAPbobsCOM.BusinessPartners = Nothing
        oForm.Freeze(True)
        Dim HoraCambio As Integer = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
        Dim DiaCambio As String = DNAUtils.Converters.ConvertDateTimeToSAPDate(DateTime.Now)
        Try
            'Sacamos los campos a monitorizar
            Dim CardType As String = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0).ToString.Trim
            Dim CardCode As String = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString.Trim
            Dim SubTipo As String = String.Empty
            Select Case CardType
                Case "S"
                    SubTipo = "Proveedores"
                Case "C"
                    SubTipo = "Clientes"
                Case Else
                    SubTipo = "Leads"
            End Select
            Dim ListaCampos_A_monitorizar As SingleResponseCampos = getFieldsToMonitorize(pVal.FormTypeEx, SubTipo)
            If ListaCampos_A_monitorizar.ExecutionSuccess = False Then
                Throw New Exception(ListaCampos_A_monitorizar.FailureReason)
            Else
                If Not IsNothing(ListaCampos_A_monitorizar.ListaCampos) AndAlso ListaCampos_A_monitorizar.ListaCampos.Count > 0 Then
                    oApplication.StatusBar.SetText("Analizando campos monitorizados...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    oBp = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    oBp.GetByKey(CardCode)
                    'Obtener la instancia de la modificación realizada
                    Dim LogInstanc As Integer = getLogInstanceFromBP(CardCode)
                    If LogInstanc = -1 Then Throw New Exception("Error al obtener la instancia de modificación del IC")
                    '////////////////////////////////////////////////
                    Dim oRs As SAPbobsCOM.Recordset
                    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
#Region "OCRD - Interlocutores/Datos Maestros"

                    oApplication.StatusBar.SetText("Revisando información de datos maestros...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    For Each Campo In ListaCampos_A_monitorizar.ListaCampos
                        Dim TablaMadre As String = String.Empty, TablaHistorico As String = String.Empty
                        TablaMadre = Campo.Tabla
                        Select Case TablaMadre
                            Case "OCRD" 'Maestro
                                TablaHistorico = "ACRD"
                                Dim QueryCambio As String = "SELECT COALESCE(T1.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_NewValue"", " & vbCrLf
                                QueryCambio += " COALESCE(T0.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_OldValue"" " & vbCrLf
                                QueryCambio += " FROM " & TablaHistorico & " T0 " & vbCrLf
                                QueryCambio += " LEFT JOIN " & TablaMadre & " T1 ON T0.""CardCode"" = T1.""CardCode""" & vbCrLf
                                QueryCambio += " WHERE COALESCE(T1.""" & Campo.CampoBD & """,'') NOT LIKE COALESCE(T0.""" & Campo.CampoBD & """,'')" & vbCrLf
                                QueryCambio += " AND T1.""CardCode""='" & CardCode & "'" & vbCrLf
                                QueryCambio += " AND T1.""CardType"" = '" & CardType & "'" & vbCrLf
                                QueryCambio += " AND T0.""LogInstanc"" = " & LogInstanc '--->ESTA LÍNEA ES PARA SACAR LA ÚLTIMA MODIFICACIÓN GRABADA
                                oRs.DoQuery(QueryCambio)
                                If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then 'Si esto devuelve algo, es porque hay diferencia entre lo que hay y lo que había.
                                    Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                    Dim table
                                    Dim Code
                                    Select Case SubTipo
                                        Case "Proveedores"
                                            table = oCompany.UserTables.Item("SEI_CAMBIOSP")
                                            Code = getNextNumberFromTable("[@SEI_CAMBIOSP]")
                                        Case "Clientes"
                                            table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                                            Code = getNextNumberFromTable("[@SEI_CAMBIOSC]")
                                        Case Else
                                            table = oCompany.UserTables.Item("SEI_CAMBIOSL")
                                            Code = getNextNumberFromTable("[@SEI_CAMBIOSL]")
                                    End Select
                                    table.Code = Code
                                    table.Name = Code
                                    table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                    'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                    table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                    table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                    table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                    table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                    table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                    table.UserFields.Fields.Item("U_ValorAnterior").Value = CStr(oRs.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                    table.UserFields.Fields.Item("U_ValorNuevo").Value = oRs.Fields.Item(Campo.CampoBD & "_NewValue").Value
                                    table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                    table.UserFields.Fields.Item("U_CardCode").Value = CardCode
                                    If table.Add() <> 0 Then
                                        Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                    End If
                                    SetPropertyValueDynamically(oBp, Campo.CampoObjeto, oRs.Fields.Item(Campo.CampoBD & "_OldValue").Value)


                                End If

                        End Select
                    Next
                    oApplication.StatusBar.SetText("Datos maestros revisados", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
#End Region

#Region "OCPR - Personas de Contacto"
                    oApplication.StatusBar.SetText("Revisando información de personas de contacto...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    For Each Campo In ListaCampos_A_monitorizar.ListaCampos
                        Dim TablaMadre As String = String.Empty, TablaHistorico As String = String.Empty
                        TablaMadre = Campo.Tabla
                        Select Case TablaMadre
                            Case "OCPR" 'Personas de contacto
                                TablaHistorico = "ACPR"
                                'Primero hay que sacar todas las personas de contacto que tiene el IC. 
                                Dim QueryPersonasContacto As String = "SELECT ""CntctCode"",""Name"" FROM OCPR WHERE ""CardCode"" = '" & CardCode & "'"
                                oRs.DoQuery(QueryPersonasContacto)
                                If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                                    'Por cada una de estas direcciones, debo revisar si hay algún cambio o han incluido alguna nueva.
                                    'Si no hay nada en la ACPR, es porque se está incluyendo la dirección.
                                    oRs.MoveFirst()
                                    While Not oRs.EoF
                                        Dim CntctCode As Integer = CInt(oRs.Fields.Item(0).Value.ToString.Trim)
                                        Dim CntctName As String = oRs.Fields.Item(1).Value.ToString.Trim
                                        Dim QueryContarCambiosPC As String = "SELECT COUNT(*) FROM ACPR WHERE ""CardCode"" = '" & CardCode & "' AND ""CntctCode"" = " & CntctCode
                                        Dim oRsPC As SAPbobsCOM.Recordset = Nothing
                                        Try
                                            oRsPC = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsPC.DoQuery(QueryContarCambiosPC)
                                            If Not IsNothing(oRsPC) AndAlso oRsPC.RecordCount > 0 Then
                                                Dim NumRegistros As Integer = CInt(oRsPC.Fields.Item(0).Value.ToString.Trim)
                                                If NumRegistros = 0 Then
#Region "La persona de contacto NO está en la tabla de cambios, por tanto, es un contacto nuevo"
                                                    Dim QueryNuevoPC As String = "SELECT COALESCE(""" & Campo.CampoBD & """,'') AS " & Campo.CampoBD & vbCrLf
                                                    QueryNuevoPC += " FROM OCPR "
                                                    QueryNuevoPC += " WHERE ""CardCode"" ='" & CardCode & "'" & vbCrLf
                                                    QueryNuevoPC += " AND ""CntctCode"" = " & CntctCode & vbCrLf
                                                    Dim oRSNuevoContacto As SAPbobsCOM.Recordset
                                                    oRSNuevoContacto = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRSNuevoContacto.DoQuery(QueryNuevoPC)
                                                    Dim NuevoValor As String = oRSNuevoContacto.Fields.Item(Campo.CampoBD).Value.ToString.Trim
                                                    If Not String.IsNullOrEmpty(NuevoValor) Then 'Si el campo tiene un valor distinto a vacío y está monitorizado, debemos volverlo a su valor previo
                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                                        Dim table
                                                        Dim Code
                                                        Select Case SubTipo
                                                            Case "Proveedores"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSP")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSP]")
                                                            Case "Clientes"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSC]")
                                                            Case Else
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSL")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSL]")
                                                        End Select
                                                        table.Code = Code
                                                        table.Name = Code
                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = ""
                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = NuevoValor
                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                                        table.UserFields.Fields.Item("U_CardCode").Value = CardCode
                                                        table.UserFields.Fields.Item("U_CampoClaveBBDD").Value = "CntctCode"
                                                        table.UserFields.Fields.Item("U_ValorCampoClaveBBD").Value = CntctCode
                                                        If table.Add() <> 0 Then
                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                                        End If
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                                                        For i As Integer = 0 To oBp.ContactEmployees.Count - 1
                                                            oBp.ContactEmployees.SetCurrentLine(i)
                                                            If oBp.ContactEmployees.InternalCode = CntctCode Then
                                                                SetPropertyValueDynamically(oBp.ContactEmployees, Campo.CampoObjeto, "")
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If
                                                    LiberarObjetoCOM(oRSNuevoContacto)
#End Region
                                                Else
#Region "La persona de contacto está en la tabla de cambios. Revisamos los cambios"
                                                    Dim QueryCambio As String = "SELECT COALESCE(T1.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_NewValue"", " & vbCrLf
                                                    QueryCambio += " COALESCE(T0.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_OldValue"" " & vbCrLf
                                                    QueryCambio += " FROM " & TablaMadre & " T0 " & vbCrLf
                                                    QueryCambio += " LEFT JOIN " & TablaHistorico & " T1 ON T0.""CardCode"" = T1.""CardCode"" AND T0.""CntctCode"" = T1.""CntctCode"" " & vbCrLf
                                                    QueryCambio += " WHERE T1.""CardCode"" = '" & CardCode & "' AND T1.CntctCode = " & CntctCode & vbCrLf
                                                    QueryCambio += " AND T1.""LogInstanc"" = " & LogInstanc
                                                    QueryCambio += " AND COALESCE(T1.""" & Campo.CampoBD & """,'') NOT LIKE COALESCE(T0.""" & Campo.CampoBD & """,'')" & vbCrLf
                                                    Dim oRsCambio As SAPbobsCOM.Recordset
                                                    oRsCambio = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRsCambio.DoQuery(QueryCambio)
                                                    If Not IsNothing(oRsCambio) AndAlso oRsCambio.RecordCount > 0 Then
                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                                        Dim table
                                                        Dim Code
                                                        Select Case SubTipo
                                                            Case "Proveedores"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSP")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSP]")
                                                            Case "Clientes"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSC]")
                                                            Case Else
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSL")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSL]")
                                                        End Select
                                                        table.Code = Code
                                                        table.Name = Code
                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = CStr(oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = oRsCambio.Fields.Item(Campo.CampoBD & "_NewValue").Value
                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                                        table.UserFields.Fields.Item("U_CardCode").Value = CardCode
                                                        table.UserFields.Fields.Item("U_CampoClaveBBDD").Value = "CntctCode"
                                                        table.UserFields.Fields.Item("U_ValorCampoClaveBBD").Value = CntctCode
                                                        If table.Add() <> 0 Then
                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                                        End If
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                                                        For i As Integer = 0 To oBp.ContactEmployees.Count - 1
                                                            oBp.ContactEmployees.SetCurrentLine(i)
                                                            If oBp.ContactEmployees.InternalCode = CntctCode Then
                                                                SetPropertyValueDynamically(oBp.ContactEmployees, Campo.CampoObjeto, oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If
#End Region
                                                End If
                                            Else
                                                Trace.WriteLineIf(ots.TraceError, "Error al determinar cuantos cambios hay de Personas de Contactos para la ID de contacto " & CntctCode & " del Interlocutor " & CardCode)
                                            End If
                                        Catch ex As Exception
                                            Trace.WriteLineIf(ots.TraceError, ex.Message)
                                        Finally
                                            LiberarObjetoCOM(oRsPC)
                                        End Try
                                        oRs.MoveNext()
                                    End While

                                Else
                                    Trace.WriteLineIf(ots.TraceInfo, "No hay personas de contacto para el IC " & CardCode)
                                End If
                        End Select
                    Next
                    oApplication.StatusBar.SetText("Personas de contacto revisadas", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)


#End Region

#Region "CRD1 - Direcciones"
                    oApplication.StatusBar.SetText("Revisando información de direcciones...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    For Each Campo In ListaCampos_A_monitorizar.ListaCampos
                        Dim TablaMadre As String = String.Empty, TablaHistorico As String = String.Empty
                        TablaMadre = Campo.Tabla
                        Select Case TablaMadre
                            Case "CRD1"
                                TablaHistorico = "ACR1"
                                'Primero hay que sacar todas las direcciones que tiene el IC. 
                                Dim QueryDirecciones As String = "SELECT ""Address"", ""AdresType"" FROM CRD1 WHERE ""CardCode"" = '" & CardCode & "'"
                                oRs.DoQuery(QueryDirecciones)
                                If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                                    oRs.MoveFirst()
                                    While Not oRs.EoF
                                        Dim Address As String = oRs.Fields.Item(0).Value.ToString.Trim
                                        Dim AdresType As String = oRs.Fields.Item(1).Value.ToString.Trim
                                        Dim QueryContarCambiosDir As String = "SELECT COUNT(*) FROM ACR1 WHERE ""CardCode"" = '" & CardCode & "' AND ""Address"" = '" & Address & "'"
                                        Dim oRsDir As SAPbobsCOM.Recordset = Nothing
                                        Try
                                            oRsDir = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsDir.DoQuery(QueryContarCambiosDir)
                                            If Not IsNothing(oRsDir) AndAlso oRsDir.RecordCount > 0 Then
                                                Dim NumRegistros As Integer = CInt(oRsDir.Fields.Item(0).Value.ToString.Trim)
                                                If NumRegistros = 0 Then
#Region "La dirección no está en la tabla de cambios. Es una dirección nueva"
                                                    Dim QueryNuevoPC As String = "SELECT COALESCE(""" & Campo.CampoBD & """,'') AS " & Campo.CampoBD & vbCrLf
                                                    QueryNuevoPC += " FROM CDR1 "
                                                    QueryNuevoPC += " WHERE ""CardCode"" ='" & CardCode & "'" & vbCrLf
                                                    QueryNuevoPC += " AND ""Address"" = " & Address & vbCrLf
                                                    Dim oRsNuevaDireccion As SAPbobsCOM.Recordset
                                                    oRsNuevaDireccion = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRsNuevaDireccion.DoQuery(QueryNuevoPC)
                                                    Dim NuevoValor As String = oRsNuevaDireccion.Fields.Item(Campo.CampoBD).Value.ToString.Trim
                                                    If Not String.IsNullOrEmpty(NuevoValor) Then 'Si el campo tiene un valor distinto a vacío y está monitorizado, debemos volverlo a su valor previo
                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                                        Dim table
                                                        Dim Code
                                                        Select Case SubTipo
                                                            Case "Proveedores"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSP")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSP]")
                                                            Case "Clientes"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSC]")
                                                            Case Else
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSL")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSL]")
                                                        End Select
                                                        table.Code = Code
                                                        table.Name = Code
                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = ""
                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = NuevoValor
                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                                        table.UserFields.Fields.Item("U_CardCode").Value = CardCode
                                                        table.UserFields.Fields.Item("U_CampoClaveBBDD").Value = "Address"
                                                        table.UserFields.Fields.Item("U_ValorCampoClaveBBD").Value = Address
                                                        table.UserFields.Fields.Item("U_SubTipoCampoClaveBBDD").Value = AdresType
                                                        If table.Add() <> 0 Then
                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                                        End If
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                                                        For i As Integer = 0 To oBp.Addresses.Count - 1
                                                            oBp.Addresses.SetCurrentLine(i)
                                                            If oBp.Addresses.AddressName = Address Then
                                                                SetPropertyValueDynamically(oBp.Addresses, Campo.CampoObjeto, "")
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If


#End Region
                                                Else
#Region "La dirección ya está en la tabla de cambios. Se revisan los cambios"
                                                    Dim QueryCambio As String = "SELECT COALESCE(T1.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_NewValue"", " & vbCrLf
                                                    QueryCambio += " COALESCE(T0.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_OldValue"" " & vbCrLf
                                                    QueryCambio += " FROM " & TablaMadre & " T0 " & vbCrLf
                                                    QueryCambio += " LEFT JOIN " & TablaHistorico & " T1 ON T0.""CardCode"" = T1.""CardCode"" AND T0.""Address"" = T1.""Address"" " & vbCrLf
                                                    QueryCambio += " WHERE T1.""CardCode"" = '" & CardCode & "' AND T1.""Address"" = '" & Address & "'" & vbCrLf
                                                    QueryCambio += " AND T1.""LogInstanc"" = " & LogInstanc
                                                    QueryCambio += " AND COALESCE(T1.""" & Campo.CampoBD & """,'') NOT LIKE COALESCE(T0.""" & Campo.CampoBD & """,'')" & vbCrLf
                                                    Dim oRsCambio As SAPbobsCOM.Recordset
                                                    oRsCambio = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRsCambio.DoQuery(QueryCambio)
                                                    If Not IsNothing(oRsCambio) AndAlso oRsCambio.RecordCount > 0 Then
                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                                        Dim table
                                                        Dim Code
                                                        Select Case SubTipo
                                                            Case "Proveedores"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSP")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSP]")
                                                            Case "Clientes"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSC]")
                                                            Case Else
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSL")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSL]")
                                                        End Select
                                                        table.Code = Code
                                                        table.Name = Code
                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = CStr(oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = oRsCambio.Fields.Item(Campo.CampoBD & "_NewValue").Value
                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                                        table.UserFields.Fields.Item("U_CardCode").Value = CardCode
                                                        table.UserFields.Fields.Item("U_CampoClaveBBDD").Value = "Address"
                                                        table.UserFields.Fields.Item("U_ValorCampoClaveBBD").Value = Address
                                                        table.UserFields.Fields.Item("U_SubTipoCampoClaveBBDD").Value = AdresType
                                                        If table.Add() <> 0 Then
                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                                        End If
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                                                        For i As Integer = 0 To oBp.Addresses.Count - 1
                                                            oBp.Addresses.SetCurrentLine(i)
                                                            If oBp.Addresses.AddressName = Address Then
                                                                SetPropertyValueDynamically(oBp.Addresses, Campo.CampoObjeto, oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If
#End Region
                                                End If
                                            Else
                                                Trace.WriteLineIf(ots.TraceError, "Error al determinar cuantos cambios hay de Direcciones el Interlocutor " & CardCode)
                                            End If

                                        Catch ex As Exception
                                            Trace.WriteLine(ots.TraceError, ex.Message)
                                        Finally
                                            LiberarObjetoCOM(oRsDir)
                                        End Try
                                        oRs.MoveNext()
                                    End While
                                Else
                                    Trace.WriteLineIf(ots.TraceInfo, "No hay direcciones para el IC " & CardCode)
                                End If

                        End Select
                    Next
                    oApplication.StatusBar.SetText("Direcciones revisadas", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
#End Region

#Region "CRD2 - Vías de Pago - Pendiente de Implementar"
                    oApplication.StatusBar.SetText("Revisando información de Vías de Pago...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)

                    '                    For Each Campo In ListaCampos_A_monitorizar.ListaCampos
                    '                        Dim TablaMadre As String = String.Empty, TablaHistorico As String = String.Empty
                    '                        TablaMadre = Campo.Tabla
                    '                        Select Case TablaMadre
                    '                            Case "CRD2"
                    '                                TablaHistorico = "ACR2"
                    '                                'Primero hay que sacar todas las direcciones que tiene el IC. 
                    '                                Dim QueryViasdePago As String = "SELECT ""Address"", ""AdresType"" FROM CRD1 WHERE ""CardCode"" = '" & CardCode & "'"
                    '                                oRs.DoQuery(QueryViasdePago)
                    '                                If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                    '                                    oRs.MoveFirst()
                    '                                    While Not oRs.EoF
                    '                                        Dim Address As String = oRs.Fields.Item(0).Value.ToString.Trim
                    '                                        Dim AdresType As String = oRs.Fields.Item(1).Value.ToString.Trim
                    '                                        Dim QueryContarCambiosDir As String = "SELECT COUNT(*) FROM ACR1 WHERE ""CardCode"" = '" & CardCode & "' AND ""Address"" = '" & Address & "'"
                    '                                        Dim oRsDir As SAPbobsCOM.Recordset = Nothing
                    '                                        Try
                    '                                            oRsDir = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                                            oRsDir.DoQuery(QueryContarCambiosDir)
                    '                                            If Not IsNothing(oRsDir) AndAlso oRsDir.RecordCount > 0 Then
                    '                                                Dim NumRegistros As Integer = CInt(oRsDir.Fields.Item(0).Value.ToString.Trim)
                    '                                                If NumRegistros = 0 Then
                    '#Region "La dirección no está en la tabla de cambios. Es una dirección nueva"
                    '                                                    Dim QueryNuevoPC As String = "SELECT COALESCE(""" & Campo.CampoBD & """,'') AS " & Campo.CampoBD & vbCrLf
                    '                                                    QueryNuevoPC += " FROM CDR1 "
                    '                                                    QueryNuevoPC += " WHERE ""CardCode"" ='" & CardCode & "'" & vbCrLf
                    '                                                    QueryNuevoPC += " AND ""Address"" = " & Address & vbCrLf
                    '                                                    Dim oRsNuevaDireccion As SAPbobsCOM.Recordset
                    '                                                    oRsNuevaDireccion = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                                                    oRsNuevaDireccion.DoQuery(QueryNuevoPC)
                    '                                                    Dim NuevoValor As String = oRsNuevaDireccion.Fields.Item(Campo.CampoBD).Value.ToString.Trim
                    '                                                    If Not String.IsNullOrEmpty(NuevoValor) Then 'Si el campo tiene un valor distinto a vacío y está monitorizado, debemos volverlo a su valor previo
                    '                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                    '                                                        Dim table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                    '                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                    '                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                    '                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                    '                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                    '                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                    '                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                    '                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                    '                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = ""
                    '                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = NuevoValor
                    '                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                    '                                                        If table.Add() <> 0 Then
                    '                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                    '                                                        End If
                    '                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                    '                                                        For i As Integer = 0 To oBp.Addresses.Count - 1
                    '                                                            oBp.Addresses.SetCurrentLine(i)
                    '                                                            If oBp.Addresses.AddressName = Address Then
                    '                                                                SetPropertyValueDynamically(oBp.Addresses, Campo.CampoObjeto, "")
                    '                                                                Exit For
                    '                                                            End If
                    '                                                        Next
                    '                                                    End If


                    '#End Region
                    '                                                Else
                    '#Region "La dirección ya está en la tabla de cambios. Se revisan los cambios"
                    '                                                    Dim QueryCambio As String = "SELECT COALESCE(T1.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_NewValue"", " & vbCrLf
                    '                                                    QueryCambio += " COALESCE(T0.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_OldValue"" " & vbCrLf
                    '                                                    QueryCambio += " FROM " & TablaMadre & " T0 " & vbCrLf
                    '                                                    QueryCambio += " LEFT JOIN " & TablaHistorico & " T1 ON T0.""CardCode"" = T1.""CardCode"" AND T0.""CntctCode"" = T1.""CntctCode"" " & vbCrLf
                    '                                                    QueryCambio += " WHERE T1.""CardCode"" = '" & CardCode & "' AND T1.""Address"" = '" & Address & "'" & vbCrLf
                    '                                                    QueryCambio += " AND T1.""LogInstanc"" = " & LogInstanc
                    '                                                    QueryCambio += " AND COALESCE(T1.""" & Campo.CampoBD & """,'') <> COALESCE(T0.""" & Campo.CampoBD & """,'')" & vbCrLf
                    '                                                    Dim oRsCambio As SAPbobsCOM.Recordset
                    '                                                    oRsCambio = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                                                    oRsCambio.DoQuery(QueryCambio)
                    '                                                    If Not IsNothing(oRsCambio) AndAlso oRsCambio.RecordCount > 0 Then
                    '                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                    '                                                        Dim table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                    '                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                    '                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                    '                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                    '                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                    '                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                    '                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                    '                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                    '                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value
                    '                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = oRsCambio.Fields.Item(Campo.CampoBD & "_NewValue").Value
                    '                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                    '                                                        If table.Add() <> 0 Then
                    '                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                    '                                                        End If
                    '                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                    '                                                        For i As Integer = 0 To oBp.Addresses.Count - 1
                    '                                                            oBp.Addresses.SetCurrentLine(i)
                    '                                                            If oBp.Addresses.AddressName = Address Then
                    '                                                                SetPropertyValueDynamically(oBp.Addresses, Campo.CampoObjeto, oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                    '                                                                Exit For
                    '                                                            End If
                    '                                                        Next
                    '                                                    End If
                    '#End Region
                    '                                                End If
                    '                                            Else
                    '                                                Trace.WriteLineIf(ots.TraceError, "Error al determinar cuantos cambios hay de Direcciones el Interlocutor " & CardCode)
                    '                                            End If

                    '                                        Catch ex As Exception
                    '                                            Trace.WriteLine(ots.TraceError, ex.Message)
                    '                                        Finally
                    '                                            LiberarObjetoCOM(oRsDir)
                    '                                        End Try


                    '                                        oRs.MoveNext()
                    '                                    End While
                    '                                Else
                    '                                    Trace.WriteLineIf(ots.TraceInfo, "No hay direcciones para el IC " & CardCode)
                    '                                End If

                    '                        End Select
                    '                    Next

                    oApplication.StatusBar.SetText("Vías de Paga Revisadas", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)


#End Region

#Region "CRD3 - Cuentas Asociadas"
                    oApplication.StatusBar.SetText("Revisando información de Cuentas Asociadas...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    For Each Campo In ListaCampos_A_monitorizar.ListaCampos
                        Dim TablaMadre As String = String.Empty, TablaHistorico As String = String.Empty
                        TablaMadre = Campo.Tabla
                        Select Case TablaMadre
                            Case "CRD3"
                                TablaHistorico = "ACR3"
                                'Primero hay que sacar todas las direcciones que tiene el IC. 
                                Dim QueryCuentasAsociadas As String = "SELECT ""AcctType"", ""AcctCode"" FROM CRD3 WHERE ""CardCode"" = '" & CardCode & "'"
                                oRs.DoQuery(QueryCuentasAsociadas)
                                If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                                    oRs.MoveFirst()
                                    While Not oRs.EoF
                                        Dim AcctType As String = oRs.Fields.Item(0).Value.ToString.Trim
                                        Dim AcctCode As String = oRs.Fields.Item(1).Value.ToString.Trim
                                        Dim QueryContarCambiosDir As String = "SELECT COUNT(*) FROM ACR3 WHERE ""CardCode"" = '" & CardCode & "' AND ""AcctType"" = '" & AcctType & "' AND ""AcctCode"" = '" & AcctCode & "'"
                                        Dim oRsCtas As SAPbobsCOM.Recordset = Nothing
                                        Try
                                            oRsCtas = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsCtas.DoQuery(QueryContarCambiosDir)
                                            If Not IsNothing(oRsCtas) AndAlso oRsCtas.RecordCount > 0 Then
                                                Dim NumRegistros As Integer = CInt(oRsCtas.Fields.Item(0).Value.ToString.Trim)
                                                If NumRegistros = 0 Then
#Region "La cuenta no está en la tabla de cambios. Es una cuenta nueva"
                                                    Dim QueryNuevaCuentaAsociada As String = "SELECT COALESCE(""" & Campo.CampoBD & """,'') AS " & Campo.CampoBD & vbCrLf
                                                    QueryNuevaCuentaAsociada += " FROM CDR3 "
                                                    QueryNuevaCuentaAsociada += " WHERE ""CardCode"" ='" & CardCode & "'" & vbCrLf
                                                    QueryNuevaCuentaAsociada += " AND ""AcctType"" = '" & AcctType & "'" & vbCrLf
                                                    QueryNuevaCuentaAsociada += " AND ""AcctCode"" = '" & AcctCode & "'" & vbCrLf
                                                    Dim oRsNuevaCuenta As SAPbobsCOM.Recordset
                                                    oRsNuevaCuenta = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRsNuevaCuenta.DoQuery(QueryNuevaCuentaAsociada)
                                                    Dim NuevoValor As String = oRsNuevaCuenta.Fields.Item(Campo.CampoBD).Value.ToString.Trim
                                                    If Not String.IsNullOrEmpty(NuevoValor) Then 'Si el campo tiene un valor distinto a vacío y está monitorizado, debemos volverlo a su valor previo
                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                                        Dim table
                                                        Dim Code
                                                        Select Case SubTipo
                                                            Case "Proveedores"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSP")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSP]")
                                                            Case "Clientes"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSC]")
                                                            Case Else
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSL")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSL]")
                                                        End Select
                                                        table.Code = Code
                                                        table.Name = Code
                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = ""
                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = NuevoValor
                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                                        table.UserFields.Fields.Item("U_CardCode").Value = CardCode
                                                        table.UserFields.Fields.Item("U_CampoClaveBBDD").Value = "AccountType"
                                                        table.UserFields.Fields.Item("U_ValorCampoClaveBBD").Value = AcctType
                                                        If table.Add() <> 0 Then
                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                                        End If
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                                                        For i As Integer = 0 To oBp.AccountRecivablePayables.Count - 1
                                                            oBp.AccountRecivablePayables.SetCurrentLine(i)
                                                            If oBp.AccountRecivablePayables.AccountType = AcctType Then
                                                                SetPropertyValueDynamically(oBp.AccountRecivablePayables, Campo.CampoObjeto, "")
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If


#End Region
                                                Else
#Region "La cuenta ya está en la tabla de cambios. Se revisan los cambios"
                                                    Dim QueryCambio As String = "SELECT COALESCE(T1.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_NewValue"", " & vbCrLf
                                                    QueryCambio += " COALESCE(T0.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_OldValue"" " & vbCrLf
                                                    QueryCambio += " FROM " & TablaMadre & " T0 " & vbCrLf
                                                    QueryCambio += " LEFT JOIN " & TablaHistorico & " T1 ON T0.""CardCode"" = T1.""CardCode"" AND T0.""AcctType"" = T1.""AcctType"" " & vbCrLf
                                                    QueryCambio += " WHERE T1.""CardCode"" = '" & CardCode & "' AND T1.""AcctType"" = '" & AcctType & "'" & vbCrLf
                                                    QueryCambio += " AND T1.""LogInstanc"" = " & LogInstanc
                                                    QueryCambio += " AND COALESCE(T1.""" & Campo.CampoBD & """,'') NOT LIKE COALESCE(T0.""" & Campo.CampoBD & """,'')" & vbCrLf
                                                    Dim oRsCambio As SAPbobsCOM.Recordset
                                                    oRsCambio = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRsCambio.DoQuery(QueryCambio)
                                                    If Not IsNothing(oRsCambio) AndAlso oRsCambio.RecordCount > 0 Then
                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                                        Dim table
                                                        Dim Code
                                                        Select Case SubTipo
                                                            Case "Proveedores"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSP")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSP]")
                                                            Case "Clientes"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSC]")
                                                            Case Else
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSL")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSL]")
                                                        End Select
                                                        table.Code = Code
                                                        table.Name = Code
                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = CStr(oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = oRsCambio.Fields.Item(Campo.CampoBD & "_NewValue").Value
                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                                        table.UserFields.Fields.Item("U_CardCode").Value = CardCode
                                                        table.UserFields.Fields.Item("U_CampoClaveBBDD").Value = "AccountType"
                                                        table.UserFields.Fields.Item("U_ValorCampoClaveBBD").Value = AcctType
                                                        If table.Add() <> 0 Then
                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                                        End If
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                                                        For i As Integer = 0 To oBp.AccountRecivablePayables.Count - 1
                                                            oBp.AccountRecivablePayables.SetCurrentLine(i)
                                                            If oBp.AccountRecivablePayables.AccountType = AcctType Then
                                                                SetPropertyValueDynamically(oBp.AccountRecivablePayables, Campo.CampoObjeto, oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                                                Exit For
                                                            End If
                                                        Next

                                                    End If
#End Region
                                                End If
                                            Else
                                                Trace.WriteLineIf(ots.TraceError, "Error al determinar cuantos cambios hay de cuentas asociadas el Interlocutor " & CardCode)
                                            End If

                                        Catch ex As Exception
                                            Trace.WriteLine(ots.TraceError, ex.Message)
                                        Finally
                                            LiberarObjetoCOM(oRsCtas)
                                        End Try


                                        oRs.MoveNext()
                                    End While
                                Else
                                    Trace.WriteLineIf(ots.TraceInfo, "No hay cuentas asociadas para el IC " & CardCode)
                                End If

                        End Select
                    Next
                    oApplication.StatusBar.SetText("Cuentas Asociadas revisadas", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)

#End Region

#Region "CRD5 - Fechas de Pago"

                    oApplication.StatusBar.SetText("Revisando información de Fechas de Pago...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    For Each Campo In ListaCampos_A_monitorizar.ListaCampos
                        Dim TablaMadre As String = String.Empty, TablaHistorico As String = String.Empty
                        TablaMadre = Campo.Tabla
                        Select Case TablaMadre
                            Case "CRD5"
                                TablaHistorico = "ACR5"
                                'Primero hay que sacar todas las direcciones que tiene el IC. 
                                Dim QueryFechasDePago As String = "SELECT ""CardCode"", ""PmntDate"" FROM CRD5 WHERE ""CardCode"" = '" & CardCode & "'"
                                oRs.DoQuery(QueryFechasDePago)
                                If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                                    oRs.MoveFirst()
                                    While Not oRs.EoF
                                        'Dim AcctType As String = oRs.Fields.Item(0).Value.ToString.Trim
                                        Dim PmntDate As String = oRs.Fields.Item(1).Value.ToString.Trim
                                        Dim QueryContarCambiosFecCamb As String = "SELECT COUNT(*) FROM ACR5 WHERE ""CardCode"" = '" & CardCode & "'"
                                        Dim oRsFechasPago As SAPbobsCOM.Recordset = Nothing
                                        Try
                                            oRsFechasPago = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsFechasPago.DoQuery(QueryContarCambiosFecCamb)
                                            If Not IsNothing(oRsFechasPago) AndAlso oRsFechasPago.RecordCount > 0 Then
                                                Dim NumRegistros As Integer = CInt(oRsFechasPago.Fields.Item(0).Value.ToString.Trim)
                                                If NumRegistros = 0 Then
#Region "La fecha de pago no está en la tabla de cambios. Es una fecha nueva"
                                                    Dim QueryNuevaFechaPago As String = "SELECT COALESCE(""" & Campo.CampoBD & """,'') AS " & Campo.CampoBD & vbCrLf
                                                    QueryNuevaFechaPago += " FROM CDR5 "
                                                    QueryNuevaFechaPago += " WHERE ""CardCode"" ='" & CardCode & "'" & vbCrLf
                                                    Dim oRsNuevaFechaPago As SAPbobsCOM.Recordset
                                                    oRsNuevaFechaPago = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRsNuevaFechaPago.DoQuery(QueryNuevaFechaPago)
                                                    Dim NuevoValor As String = oRsNuevaFechaPago.Fields.Item(Campo.CampoBD).Value.ToString.Trim
                                                    If Not String.IsNullOrEmpty(NuevoValor) Then 'Si el campo tiene un valor distinto a vacío y está monitorizado, debemos volverlo a su valor previo
                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                                        Dim table
                                                        Dim Code
                                                        Select Case SubTipo
                                                            Case "Proveedores"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSP")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSP]")
                                                            Case "Clientes"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSC]")
                                                            Case Else
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSL")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSL]")
                                                        End Select
                                                        table.Code = Code
                                                        table.Name = Code
                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = ""
                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = NuevoValor
                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                                        table.UserFields.Fields.Item("U_CardCode").Value = CardCode
                                                        table.UserFields.Fields.Item("U_CampoClaveBBDD").Value = "PaymentDate"
                                                        table.UserFields.Fields.Item("U_ValorCampoClaveBBD").Value = PmntDate
                                                        If table.Add() <> 0 Then
                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                                        End If
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                                                        For i As Integer = 0 To oBp.BPPaymentDates.Count - 1
                                                            oBp.BPPaymentDates.SetCurrentLine(i)
                                                            If oBp.BPPaymentDates.PaymentDate = PmntDate Then
                                                                SetPropertyValueDynamically(oBp.BPPaymentDates, Campo.CampoObjeto, "")
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If


#End Region
                                                Else
#Region "La Fecha de Pago ya está en la tabla de cambios. Se revisan los cambios"
                                                    Dim QueryCambio As String = "SELECT COALESCE(T1.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_NewValue"", " & vbCrLf
                                                    QueryCambio += " COALESCE(T0.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_OldValue"" " & vbCrLf
                                                    QueryCambio += " FROM " & TablaMadre & " T0 " & vbCrLf
                                                    QueryCambio += " LEFT JOIN " & TablaHistorico & " T1 ON T0.""CardCode"" = T1.""CardCode"" AND T0.""AcctType"" = T1.""AcctType"" " & vbCrLf
                                                    QueryCambio += " WHERE T1.""CardCode"" = '" & CardCode & "'" & vbCrLf
                                                    QueryCambio += " AND COALESCE(T1.""" & Campo.CampoBD & """,'') NOT LIKE COALESCE(T0.""" & Campo.CampoBD & """,'')" & vbCrLf
                                                    Dim oRsCambio As SAPbobsCOM.Recordset
                                                    oRsCambio = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRsCambio.DoQuery(QueryCambio)
                                                    If Not IsNothing(oRsCambio) AndAlso oRsCambio.RecordCount > 0 Then
                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                                        Dim table
                                                        Dim Code
                                                        Select Case SubTipo
                                                            Case "Proveedores"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSP")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSP]")
                                                            Case "Clientes"
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSC")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSC]")
                                                            Case Else
                                                                table = oCompany.UserTables.Item("SEI_CAMBIOSL")
                                                                Code = getNextNumberFromTable("[@SEI_CAMBIOSL]")
                                                        End Select
                                                        table.Code = Code
                                                        table.Name = Code
                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = CStr(oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = oRsCambio.Fields.Item(Campo.CampoBD & "_NewValue").Value
                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                                        table.UserFields.Fields.Item("U_CardCode").Value = CardCode
                                                        table.UserFields.Fields.Item("U_CampoClaveBBDD").Value = "PaymentDate"
                                                        table.UserFields.Fields.Item("U_ValorCampoClaveBBD").Value = PmntDate
                                                        If table.Add() <> 0 Then
                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                                        End If
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                                                        For i As Integer = 0 To oBp.BPPaymentDates.Count - 1
                                                            oBp.BPPaymentDates.SetCurrentLine(i)
                                                            If oBp.BPPaymentDates.PaymentDate = PmntDate Then
                                                                SetPropertyValueDynamically(oBp.BPPaymentDates, Campo.CampoObjeto, oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                                                Exit For
                                                            End If
                                                        Next

                                                    End If
#End Region
                                                End If
                                            Else
                                                Trace.WriteLineIf(ots.TraceError, "Error al determinar cuantos cambios hay de fechas de pago para el Interlocutor " & CardCode)
                                            End If

                                        Catch ex As Exception
                                            Trace.WriteLine(ots.TraceError, ex.Message)
                                        Finally
                                            LiberarObjetoCOM(oRsFechasPago)
                                        End Try


                                        oRs.MoveNext()
                                    End While
                                Else
                                    Trace.WriteLineIf(ots.TraceInfo, "No hay fechas de pago para el IC " & CardCode)
                                End If

                        End Select
                    Next
                    oApplication.StatusBar.SetText("Fechas de Pago revisadas", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)

#End Region

#Region "Actualización del IC"
                    Try
                        oBp.Update()
                    Catch ex As Exception
                    End Try
#End Region

                Else
                    oApplication.StatusBar.SetText("No hay campos a monitorizar en el maestro de " & SubTipo, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                End If
            End If
            oForm.Refresh()
            oForm.Freeze(False)
        Catch ex As Exception
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
            oForm.Freeze(False)
        End Try
        YaEntre = False
    End Sub

#End Region

#Region "Formulario Artículos"

    Public Sub HANDLE_OK_BUTTON_PRESSED_BEFORE_ACTION_ITEMSFORM(pVal As BusinessObjectInfo)
        Dim oForm As SAPbouiCOM.Form
        oForm = oApplication.Forms.Item(pVal.FormUID)
        Try
            'Paso 1: revisar si el usuario tiene autorización para grabar directamente en el formulario
            Dim oResponseAutorizacion As New SingleResponse
            oResponseAutorizacion = getUserAuthorization(oCompany.UserName, pVal.FormTypeEx)
            If oResponseAutorizacion.ExecutionSuccess = False AndAlso oResponseAutorizacion.FailureReason.StartsWith("ERROR:") Then
                oApplication.MessageBox("Su usuario no tiene autorización para cambiar ciertos datos maestros de este formulario" & vbNewLine & "Sus cambios se guardarán para que otro usuario con permisos los pueda autorizar")
                UsuarioConPermisosItems = False
            Else
                UsuarioConPermisosItems = True
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub HANDLE_OK_BUTTON_PRESSED_AFTER_ACTION_ITEMSFORM(pval As BusinessObjectInfo)
        Dim oForm As SAPbouiCOM.Form
        oForm = oApplication.Forms.Item(pval.FormUID)
        Dim oItem As SAPbobsCOM.Items = Nothing
        oForm.Freeze(True)
        Dim HoraCambio As Integer = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
        Dim DiaCambio As String = DNAUtils.Converters.ConvertDateTimeToSAPDate(DateTime.Now)
        Try
            Dim ItemCode As String = oForm.DataSources.DBDataSources.Item("OITM").GetValue("ItemCode", 0).ToString.Trim
            Dim ListaCampos_A_monitorizar As SingleResponseCampos = getFieldsToMonitorize(pval.FormTypeEx)
            If ListaCampos_A_monitorizar.ExecutionSuccess = False Then
                Throw New Exception(ListaCampos_A_monitorizar.FailureReason)
            Else
                If Not IsNothing(ListaCampos_A_monitorizar.ListaCampos) AndAlso ListaCampos_A_monitorizar.ListaCampos.Count > 0 Then
                    oApplication.StatusBar.SetText("Analizando campos monitorizados...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    oItem = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    oItem.GetByKey(ItemCode)
                    'Obtener la instancia de la modificación realizada
                    Dim LogInstanc As Integer = getLogInstanceFromItem(ItemCode)
                    If LogInstanc = -1 Then Throw New Exception("Error al obtener la instancia de modificación del Artículo")
                    '////////////////////////////////////////////////
                    Dim oRs As SAPbobsCOM.Recordset
                    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
#Region "OITM - Datos Maestros"
                    oApplication.StatusBar.SetText("Revisando información de datos maestros...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    For Each Campo In ListaCampos_A_monitorizar.ListaCampos
                        Dim TablaMadre As String = String.Empty, TablaHistorico As String = String.Empty
                        TablaMadre = Campo.Tabla
                        Dim TipoCampo As String = GetPropertyType(oItem, Campo.CampoObjeto)
                        Select Case TipoCampo
                            Case "String"
                            Case "Integer"
                            Case "DateTime"

                        End Select

                        Select Case TablaMadre
                            Case "OITM" 'Maestro
                                TablaHistorico = "AITM"
                                Dim QueryCambio As String = "SELECT COALESCE(T1.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_NewValue"", " & vbCrLf
                                QueryCambio += " COALESCE(T0.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_OldValue"" " & vbCrLf
                                QueryCambio += " FROM " & TablaHistorico & " T0 " & vbCrLf
                                QueryCambio += " LEFT JOIN " & TablaMadre & " T1 ON T0.""ItemCode"" = T1.""ItemCode""" & vbCrLf
                                QueryCambio += " WHERE COALESCE(T1.""" & Campo.CampoBD & """,'') <> COALESCE(T0.""" & Campo.CampoBD & """,'')" & vbCrLf
                                QueryCambio += " AND T1.""ItemCode""='" & ItemCode & "'" & vbCrLf
                                QueryCambio += " AND T0.""LogInstanc"" = " & LogInstanc '--->ESTA LÍNEA ES PARA SACAR LA ÚLTIMA MODIFICACIÓN GRABADA
                                oRs.DoQuery(QueryCambio)
                                If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then 'Si esto devuelve algo, es porque hay diferencia entre lo que hay y lo que había.
                                    Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                    Dim table = oCompany.UserTables.Item("SEI_CAMBIOSI")
                                    Dim Code As String = getNextNumberFromTable("[@SEI_CAMBIOSI]")
                                    table.Code = Code
                                    table.Name = Code
                                    table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                    'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                    table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                    table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                    table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                    table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                    table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                    table.UserFields.Fields.Item("U_ValorAnterior").Value = oRs.Fields.Item(Campo.CampoBD & "_OldValue").Value.ToString.Trim
                                    table.UserFields.Fields.Item("U_ValorNuevo").Value = oRs.Fields.Item(Campo.CampoBD & "_NewValue").Value.ToString.Trim
                                    table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                    table.UserFields.Fields.Item("U_ItemCode").Value = ItemCode
                                    If table.Add() <> 0 Then
                                        Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                    End If
                                    SetPropertyValueDynamically(oItem, Campo.CampoObjeto, oRs.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                End If

                        End Select
                    Next
                    oApplication.StatusBar.SetText("Datos maestros revisados", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)


#End Region

#Region "ITM2 - Vendedores/Proveedores Preferentes"
                    oApplication.StatusBar.SetText("Revisando información de proveedores preferentes...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    For Each Campo In ListaCampos_A_monitorizar.ListaCampos
                        Dim TablaMadre As String = String.Empty, TablaHistorico As String = String.Empty
                        TablaMadre = Campo.Tabla
                        Select Case TablaMadre
                            Case "ITM2" 'Proveedores Preferentes
                                TablaHistorico = "AIT2"
                                'Primero hay que sacar todas las personas de contacto que tiene el IC. 
                                Dim QueryPersonasContacto As String = "SELECT ""VendorCode"" FROM ITM2 WHERE ""ItemCode"" = '" & ItemCode & "'"
                                oRs.DoQuery(QueryPersonasContacto)
                                If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                                    'Por cada uno de estos proveedores, debo revisar si hay algún cambio o han incluido alguno nuevo.
                                    'Si no hay nada en la AIT2, es porque se está incluyendo el proveedor
                                    oRs.MoveFirst()
                                    While Not oRs.EoF
                                        Dim VendorCode As Integer = CInt(oRs.Fields.Item(0).Value.ToString.Trim)
                                        Dim QueryContarCambiosPP As String = "SELECT COUNT(*) FROM AIT2 WHERE ""ItemCode"" = '" & ItemCode & "' AND ""VendorCode"" = " & VendorCode
                                        Dim oRsPC As SAPbobsCOM.Recordset = Nothing
                                        Try
                                            oRsPC = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsPC.DoQuery(QueryContarCambiosPP)
                                            If Not IsNothing(oRsPC) AndAlso oRsPC.RecordCount > 0 Then
                                                Dim NumRegistros As Integer = CInt(oRsPC.Fields.Item(0).Value.ToString.Trim)
                                                If NumRegistros = 0 Then
#Region "El Proveedor NO está en la tabla de cambios, por tanto, es un proveedor nuevo"
                                                    Dim QueryNuevoPC As String = "SELECT COALESCE(""" & Campo.CampoBD & """,'') AS " & Campo.CampoBD & vbCrLf
                                                    QueryNuevoPC += " FROM ITM2 "
                                                    QueryNuevoPC += " WHERE ""ItemCode"" ='" & ItemCode & "'" & vbCrLf
                                                    QueryNuevoPC += " AND ""VendorCode"" = " & VendorCode & vbCrLf
                                                    Dim oRSNuevoProveedor As SAPbobsCOM.Recordset
                                                    oRSNuevoProveedor = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRSNuevoProveedor.DoQuery(QueryNuevoPC)
                                                    Dim NuevoValor As String = oRSNuevoProveedor.Fields.Item(Campo.CampoBD).Value.ToString.Trim
                                                    If Not String.IsNullOrEmpty(NuevoValor) Then 'Si el campo tiene un valor distinto a vacío y está monitorizado, debemos volverlo a su valor previo
                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                                        Dim table = oCompany.UserTables.Item("SEI_CAMBIOSI")
                                                        Dim Code As String = getNextNumberFromTable("[@SEI_CAMBIOSI]")
                                                        table.Code = Code
                                                        table.Name = Code
                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = ""
                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = NuevoValor
                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                                        table.UserFields.Fields.Item("U_ItemCode").Value = ItemCode
                                                        If table.Add() <> 0 Then
                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                                        End If
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el Artículo a su valor previo
                                                        For i As Integer = 0 To oItem.PreferredVendors.Count - 1
                                                            oItem.PreferredVendors.SetCurrentLine(i)
                                                            If oItem.PreferredVendors.BPCode = VendorCode Then
                                                                SetPropertyValueDynamically(oItem.PreferredVendors, Campo.CampoObjeto, "")
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If
                                                    LiberarObjetoCOM(oRSNuevoProveedor)
#End Region
                                                Else
#Region "La persona de contacto está en la tabla de cambios. Revisamos los cambios"
                                                    Dim QueryCambio As String = "SELECT COALESCE(T1.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_NewValue"", " & vbCrLf
                                                    QueryCambio += " COALESCE(T0.""" & Campo.CampoBD & """,'') AS """ & Campo.CampoBD & "_OldValue"" " & vbCrLf
                                                    QueryCambio += " FROM " & TablaMadre & " T0 " & vbCrLf
                                                    QueryCambio += " LEFT JOIN " & TablaHistorico & " T1 ON T0.""ItemCode"" = T1.""ItemCode"" AND T0.""VendorCode"" = T1.""VendorCode"" " & vbCrLf
                                                    QueryCambio += " WHERE T1.""ItemCode"" = '" & ItemCode & "' AND T1.VendorCode = " & VendorCode & vbCrLf
                                                    QueryCambio += " AND T1.""LogInstanc"" = " & LogInstanc
                                                    QueryCambio += " AND COALESCE(T1.""" & Campo.CampoBD & """,'') <> COALESCE(T0.""" & Campo.CampoBD & """,'')" & vbCrLf
                                                    Dim oRsCambio As SAPbobsCOM.Recordset
                                                    oRsCambio = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    oRsCambio.DoQuery(QueryCambio)
                                                    If Not IsNothing(oRsCambio) AndAlso oRsCambio.RecordCount > 0 Then
                                                        Dim Fecha As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")
                                                        Dim table = oCompany.UserTables.Item("SEI_CAMBIOSI")
                                                        Dim Code As String = getNextNumberFromTable("[@SEI_CAMBIOSI]")
                                                        table.Code = Code
                                                        table.Name = Code
                                                        table.UserFields.Fields.Item("U_FechaCambio").Value = DateTime.Now
                                                        'table.UserFields.Fields.Item("U_HoraCambio").Value = DNAUtils.Converters.ConvertDateTimeToSAPTime(DateTime.Now)
                                                        table.UserFields.Fields.Item("U_IdUsuario").Value = oCompany.UserSignature
                                                        table.UserFields.Fields.Item("U_NombreUsuario").Value = oCompany.UserName
                                                        table.UserFields.Fields.Item("U_CampoBBDD").Value = Campo.CampoBD
                                                        table.UserFields.Fields.Item("U_Tabla").Value = Campo.Tabla
                                                        table.UserFields.Fields.Item("U_Objeto").Value = Campo.CampoObjeto
                                                        table.UserFields.Fields.Item("U_ValorAnterior").Value = oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value
                                                        table.UserFields.Fields.Item("U_ValorNuevo").Value = oRsCambio.Fields.Item(Campo.CampoBD & "_NewValue").Value
                                                        table.UserFields.Fields.Item("U_Description").Value = Campo.Descripcion
                                                        table.UserFields.Fields.Item("U_ItemCode").Value = ItemCode
                                                        If table.Add() <> 0 Then
                                                            Trace.WriteLineIf(ots.TraceError, oCompany.GetLastErrorDescription)
                                                        End If
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el IC a su valor previo
                                                        'Una vez que he registrado el cambio en la tabla de cambios, debo devolver el Artículo a su valor previo
                                                        For i As Integer = 0 To oItem.PreferredVendors.Count - 1
                                                            oItem.PreferredVendors.SetCurrentLine(i)
                                                            If oItem.PreferredVendors.BPCode = VendorCode Then
                                                                SetPropertyValueDynamically(oItem.PreferredVendors, Campo.CampoObjeto, oRsCambio.Fields.Item(Campo.CampoBD & "_OldValue").Value)
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If
#End Region
                                                End If
                                            Else
                                                Trace.WriteLineIf(ots.TraceError, "Error al determinar cuantos cambios hay de Proveedores preferentes para el artículo " & ItemCode & " del Interlocutor " & VendorCode)
                                            End If
                                        Catch ex As Exception
                                            Trace.WriteLineIf(ots.TraceError, ex.Message)
                                        Finally
                                            LiberarObjetoCOM(oRsPC)
                                        End Try
                                        oRs.MoveNext()
                                    End While

                                Else
                                    Trace.WriteLineIf(ots.TraceInfo, "No hay proveedores preferentes para el artículo " & ItemCode)
                                End If
                        End Select
                    Next
                    oApplication.StatusBar.SetText("Proveedores preferentes revisados", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)


#End Region

                Else
                    oApplication.StatusBar.SetText("No hay campos a monitorizar en el maestro de Artículos", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                End If
            End If
            oForm.Refresh()
            oForm.Freeze(False)
        Catch ex As Exception
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

#Region "Generales"
    Public Sub HANDLE_MENU_CLICK(ByVal FormTypeEx As String, Optional ByVal SubType As String = "")
        Try
            Dim oResponseAutorizacion As New SingleResponse
            oResponseAutorizacion = getUserAuthorization(oCompany.UserName, FormTypeEx)
            If oResponseAutorizacion.ExecutionSuccess = False AndAlso oResponseAutorizacion.FailureReason.StartsWith("ERROR:") Then
                Select Case FormTypeEx
                    Case SEI_GEST.Default.BP_FORMTYPEEX
                        Select Case SubType
                            Case "S"
                                oApplication.MessageBox("Su usuario no tiene autorización para abrir el formulario de gestión de cambios de proveeedores")
                            Case "C"
                                oApplication.MessageBox("Su usuario no tiene autorización para abrir el formulario de gestión de cambios de clientes")
                        End Select
                    Case SEI_GEST.Default.ITEM_FORMTYPEEX
                        oApplication.MessageBox("Su usuario no tiene autorización para abrir el formulario de gestión de cambios de artículos")
                End Select
            Else
                HANDLE_OPENFORM_GESTION(FormTypeEx, SubType)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub HANDLE_OPENFORM_GESTION(ByVal FormTypeEx As String, Optional SubType As String = "")
        Try
            Select Case FormTypeEx
                Case SEI_GEST.Default.BP_FORMTYPEEX
                    If Not String.IsNullOrEmpty(SubType) Then
                        Dim uniqueID As String
                        Dim oForm As SAPbouiCOM.Form
                        uniqueID = oAddOnService.loadFormFromXMLString(oModule.getResourceContentsAsString("SEI_GestionIC.xml"))
                        oForm = oApplication.Forms.Item(uniqueID)
                        If String.IsNullOrEmpty(uniqueID) Then
                            oApplication.MessageBox(oApplication.GetLastBatchResults())
                        Else
                            oForm.Title &= IIf(SubType = "C", " - Ficha cliente", " - Ficha Proveedor")
                        End If
                        Formulario_AffectsFormMode(uniqueID)
                        Formulario_EnableMenu(uniqueID)
                        oForm.Visible = True
                        'oForm.State = BoFormStateEnum.fs_Maximized
                    End If
                Case SEI_GEST.Default.ITEM_FORMTYPEEX
                    'Dim uniqueID As String
                    'Dim oForm As SAPbouiCOM.Form
                    'uniqueID = oAddOnService.loadFormFromXMLString(oModule.getResourceContentsAsString("SEI_GestionIC.xml"))
                    'oForm = oApplication.Forms.Item(uniqueID)
                    'If String.IsNullOrEmpty(uniqueID) Then
                    '    oApplication.MessageBox(oApplication.GetLastBatchResults())
                    'Else
                    '    oForm.Title &= IIf(SubType = "C", " - Ficha cliente", " - Ficha Proveedor")
                    'End If
                    'Formulario_AffectsFormMode(uniqueID)
                    'Formulario_EnableMenu(uniqueID)
                    'oForm.Visible = True
                    'oForm.State = BoFormStateEnum.fs_Maximized
            End Select
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "Formulario Gestión IC"

    Public Sub HANDLE_FORM_GESTION_BOTON_DESCARTAR_ITEM_PRESSED(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, FormUID As String)
        Dim tot As Integer
        Dim oForm As SAPbouiCOM.Form
        Dim oGrid As SAPbouiCOM.Grid
        oForm = oApplication.Forms.Item(FormUID)
        oGrid = oForm.Items.Item(Grid.Grid).Specific
        Try
            If pVal.BeforeAction Then
                For i As Integer = 0 To oGrid.Rows.Count - 1
                    If oGrid.DataTable.GetValue(Grid.Col_Sel, i) = "Y" Then
                        tot += 1
                    End If
                Next
                If tot = 0 Then
                    oApplication.StatusBar.SetText("No se ha seleccionado ningún registro", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False
                End If
            Else
                For i As Integer = 0 To oGrid.Rows.Count - 1
                    If oGrid.DataTable.GetValue(Grid.Col_Sel, i) = "Y" Then
                        HANDLE_UPDATE_VALUE_CONFIGTABLE_BP(oGrid.DataTable.GetValue(Grid.Col_Code, i), "R", "@SEI_CAMBIOSC")
                    End If
                Next
                HANDLE_FORM_GESTION_FILTRAR_DATOS(FormUID)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub HANDLE_FORM_GESTION_BOTON_GRABAR_ITEM_PRESSED(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, FormUID As String)
        Dim bTemp_Empl As Boolean = False
        Dim bTemp_Proy As Boolean = False
        Dim tot As Integer

        Dim oForm As SAPbouiCOM.Form
        oForm = oApplication.Forms.Item(FormUID)
        Dim oItem As SAPbouiCOM.Item = oForm.Items.Item("txtCCode")
        Dim oSpecific As SAPbouiCOM.EditText = oItem.Specific
        Dim CardCode As String = oSpecific.Value.ToString.Trim
        'Dim CardCode As String = DirectCast(oForm.Items.Item("txtCCode"), SAPbouiCOM.EditText).Value.ToString.Trim
        Dim oGrid As SAPbouiCOM.Grid
        oGrid = DirectCast(oForm.Items.Item(Grid.Grid).Specific, SAPbouiCOM.Grid)
        oGrid = oForm.Items.Item(Grid.Grid).Specific
        Try
            If pVal.BeforeAction Then
                For i As Integer = 0 To oGrid.Rows.Count - 1
                    If oGrid.DataTable.GetValue(Grid.Col_Sel, i) = "Y" Then
                        tot += 1
                    End If
                Next
                If tot = 0 Then
                    oApplication.StatusBar.SetText("No se ha seleccionado ningún registro", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False
                End If
            Else
                Dim oBP As SAPbobsCOM.BusinessPartners = Nothing
                Try
                    oBP = oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners)
                    oBP.GetByKey(CardCode)
                    For i As Integer = 0 To oGrid.Rows.Count - 1
                        oApplication.StatusBar.SetText("Actualizando ficha del interlocutor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        If oGrid.DataTable.GetValue(Grid.Col_Sel, i) = "Y" Then
                            Dim Tabla As String = oGrid.DataTable.GetValue(Grid.Col_Tabla, i).ToString.Trim
                            Dim Propiedad As String = oGrid.DataTable.GetValue(Grid.Col_PropiedadObjeto, i).ToString.Trim
                            Dim CampoClave As String = oGrid.DataTable.GetValue(Grid.Col_CampoClave, i).ToString.Trim
                            Dim ValorCampoClave = oGrid.DataTable.GetValue(Grid.Col_ValorCampoClave, i)
                            Dim ValorNuevo = oGrid.DataTable.GetValue(Grid.Col_Nuevo, i)
                            Select Case Tabla
                                Case "OCRD" 'Maestros
                                    SetPropertyValueDynamically(oBP, Propiedad, ValorNuevo)
                                Case "OCPR" 'Personas de Contacto
                                    'El código de la persona de contacto lo sacaremos de ValorCampoClave.
                                    For x As Integer = 0 To oBP.ContactEmployees.Count - 1
                                        oBP.ContactEmployees.SetCurrentLine(x)
                                        If oBP.ContactEmployees.InternalCode = ValorCampoClave Then
                                            SetPropertyValueDynamically(oBP.ContactPerson, Propiedad, ValorNuevo)
                                            Exit For
                                        End If
                                    Next
                                Case "CRD1" 'Direcciones
                                    'El campo Address de la dirección lo sacaremos del ValorCampoClave
                                    For x As Integer = 0 To oBP.Addresses.Count - 1
                                        oBP.Addresses.SetCurrentLine(x)
                                        If oBP.Addresses.AddressName = ValorCampoClave Then
                                            SetPropertyValueDynamically(oBP.Addresses, Propiedad, ValorNuevo)
                                        End If
                                    Next
                                Case "CRD2" 'Vías de Pago - POr implementar

                                Case "CRD3" 'Cuentas Asociadas
                                    For x As Integer = 0 To oBP.AccountRecivablePayables.Count - 1
                                        oBP.AccountRecivablePayables.SetCurrentLine(i)
                                        If oBP.AccountRecivablePayables.AccountType = ValorCampoClave Then
                                            SetPropertyValueDynamically(oBP.AccountRecivablePayables, Propiedad, ValorNuevo)
                                            Exit For
                                        End If
                                    Next
                                Case "CRD5" 'Fehas de Pago
                                    For x As Integer = 0 To oBP.BPPaymentDates.Count - 1
                                        oBP.BPPaymentDates.SetCurrentLine(i)
                                        If oBP.BPPaymentDates.PaymentDate = ValorCampoClave Then
                                            SetPropertyValueDynamically(oBP.BPPaymentDates, Propiedad, ValorNuevo)
                                            Exit For
                                        End If
                                    Next
                            End Select
                            If oBP.Update <> 0 Then Throw New Exception("Error al actualizar el Interlocutor " & CardCode & ": " & oCompany.GetLastErrorDescription)
                            HANDLE_UPDATE_VALUE_CONFIGTABLE_BP(oGrid.DataTable.GetValue(Grid.Col_Code, i), "A", "@SEI_CAMBIOSC")
                        End If
                    Next
                Catch ex As Exception
                    Trace.WriteLineIf(ots.TraceError, ex.Message)
                    Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
                    oApplication.MessageBox(ex.Message)
                End Try
                HANDLE_FORM_GESTION_FILTRAR_DATOS(FormUID)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub HANDLE_FORM_GESTION_BOTON_FILTRAR_ITEM_PRESSED(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, FormUID As String)
        Dim bTemp_Empl As Boolean = False
        Dim bTemp_Proy As Boolean = False
        Dim i As Integer = 0
        Try
            If pVal.BeforeAction Then
                Dim oForm As SAPbouiCOM.Form = oApplication.Forms.Item(FormUID)
                If String.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item(Cab.UDS_CardCode).Value) Then
                    oApplication.StatusBar.SetText("Debe seleccionar un cliente para poder filtrar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False
                End If
            Else
                HANDLE_FORM_GESTION_FILTRAR_DATOS(FormUID)
            End If
        Catch ex As Exception
            oApplication.MessageBox("Error al filtrar la información de cambios: " & ex.Message)
            BubbleEvent = False
        End Try
    End Sub

    Private Sub HANDLE_FORM_GESTION_FILTRAR_DATOS(FormUID As String)
        Dim ls As String = ""

        Dim sFecha As String = ""
        Dim sFechaHasta As String = ""

        Dim oForm As SAPbouiCOM.Form
        Dim Tipo As String
        Dim oGrid As SAPbouiCOM.Grid
        Dim oDT_Grid As SAPbouiCOM.DataTable

        oForm = oApplication.Forms.Item(FormUID)
        If oForm.Title.Contains("Proveedor") Then Tipo = "S" Else Tipo = "C"
        Dim CardCode As String = oForm.DataSources.UserDataSources.Item(Cab.UDS_CardCode).Value.ToString.Trim
        oGrid = DirectCast(oForm.Items.Item(Grid.Grid).Specific, SAPbouiCOM.Grid)
        Try
            oDT_Grid = oForm.DataSources.DataTables.Item(Grid.DT)
            oGrid = oForm.Items.Item(Grid.Grid).Specific
            Dim Tabla As String
            Select Case oCompany.DbServerType
                Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
                    Select Case Tipo
                        Case "S"
                            Tabla = """@SEI_CAMBIOSP"""
                        Case Else
                            Tabla = """@SEI_CAMBIOSC"""
                    End Select

                Case Else
                    Select Case Tipo
                        Case "S"
                            Tabla = "[@SEI_CAMBIOSP]"
                        Case Else
                            Tabla = "[@SEI_CAMBIOSC]"
                    End Select
            End Select

            ls &= "SELECT T0.""Code"", T0.""U_FechaCambio"" AS ""Fecha del Cambio"", T0.""U_Description"" AS ""Nombre del Campo""," & vbCrLf
            ls &= "T0.""U_ValorAnterior"" AS ""Valor Anterior"", T0.""U_ValorNuevo"" AS ""Valor Nuevo"", T0.""U_NombreUsuario"" AS ""Nombre de usuario""," & vbCrLf
            ls &= "'N' AS ""Selección"", T0.""U_Estatus"" AS ""Estado"", T0.""U_IdUsuario"",T0.""U_CampoBBDD"",T0.""U_Tabla"",T0.""U_Objeto""," & vbCrLf
            ls &= " T0.""U_CampoClaveBBDD"" AS ""Campo Clave"", T0.""U_ValorCampoClaveBBD"" AS ""Valor Campo Clave"" " & vbCrLf
            ls &= " FROM " & Tabla & "T0 WHERE T0.""U_CardCode""  = '" & CardCode & "' AND  T0.""U_Estatus"" ='N'"
            oForm.Freeze(True)
            oDT_Grid.ExecuteQuery(ls)
            Configurar_ColumnsGrid(FormUID)
            If oDT_Grid.Rows.Count > 0 Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw New Exception("Error Mostrar_Datos: " & ex.Message)
        End Try
    End Sub

    Private Sub HANDLE_UPDATE_VALUE_CONFIGTABLE_BP(Code As String, Value As String, Table As String)
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Try
            Select Case oCompany.DbServerType
                Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
                    Table = """" & Table & """"
                Case Else
                    Table = "[" & Table & "]"
            End Select
            Dim ls As String = "UPDATE " & Table & " SET ""U_Estatus"" = '" & Value & "' WHERE ""Code"" ='" & Code & "'"
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(ls)
        Catch ex As Exception
            Trace.WriteLineIf(ots.TraceError, "Error al actualizar el registro de la tabla: " & Table & " y Code: " & Code & " --> " & ex.Message)
        Finally
            LiberarObjetoCOM(oRs)
        End Try
    End Sub

    Public Sub HANDLE_FORM_GESTION_CHOOSEFROMLIST(ByRef pVal As SAPbouiCOM.IItemEvent, ByRef BubbleEvent As Boolean, FormUID As String)
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim i As Integer
        Dim oForm As SAPbouiCOM.Form
        Dim oGrid As SAPbouiCOM.Grid
        oForm = oApplication.Forms.Item(FormUID)
        Dim Tipo As String = String.Empty
        If oForm.Title.Contains("Proveedor") Then
            Tipo = "S"
        Else
            Tipo = "C"
        End If
        oGrid = DirectCast(oForm.Items.Item(Grid.Grid).Specific, SAPbouiCOM.Grid)
        Dim Tabla As String
        Select Case oCompany.DbServerType
            Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
                Select Case Tipo
                    Case "S"
                        Tabla = """@SEI_CAMBIOSP"""
                    Case "C"
                        Tabla = """@SEI_CAMBIOSC"""
                End Select
            Case Else
                Select Case Tipo
                    Case "S"
                        Tabla = "[@SEI_CAMBIOSP]"
                    Case "C"
                        Tabla = "[@SEI_CAMBIOSC]"
                End Select

        End Select
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Try
            If pVal.BeforeAction Then
                Dim ls As String = "SELECT DISTINCT(T0.""U_CardCode"") AS ""CardCode"",(SELECT ""CardName"" FROM OCRD WHERE ""CardCode"" = T0.""U_CardCode"") AS ""CardName"" FROM " & Tabla & " T0 WHERE T0.""U_Estatus"" ='N'"
                oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRs.DoQuery(ls)
                If oRs.RecordCount <> 0 Then
                    oConds = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
                    i = 0
                    Do While Not oRs.EoF
                        If i <> 0 Then oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        oCond = oConds.Add
                        oCond.Alias = "CardCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = oRs.Fields.Item("CardCode").Value
                        oRs.MoveNext()
                        i += 1
                    Loop
                Else
                    BubbleEvent = False
                    Exit Sub
                End If
                oForm.ChooseFromLists.Item(Cab.Choose_CardCode).SetConditions(oConds)
            Else
                If Not pVal.SelectedObjects Is Nothing Then
                    oForm.DataSources.UserDataSources.Item(Cab.UDS_CardCode).Value = pVal.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                    oForm.DataSources.UserDataSources.Item(Cab.UDS_CardName).Value = pVal.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value
                Else
                    oForm.DataSources.UserDataSources.Item(Cab.UDS_CardCode).Value = ""
                    oForm.DataSources.UserDataSources.Item(Cab.UDS_CardName).Value = ""
                End If
            End If
        Catch ex As Exception
            Trace.WriteLineIf(ots.TraceError, "Error en el CHOOSEFROMLIST del formulario de Gestión: " & ex.Message)
        Finally
            LiberarObjetoCOM(oRs)
        End Try
    End Sub

    Public Sub Configurar_ColumnsGrid(FormUID As String)
        Dim oForm As SAPbouiCOM.Form
        oForm = oApplication.Forms.Item(FormUID)
        Dim oGrid As SAPbouiCOM.Grid
        oGrid = DirectCast(oForm.Items.Item(Grid.Grid).Specific, SAPbouiCOM.Grid)
        Try
            oGrid = oForm.Items.Item(Grid.Grid).Specific
            If oGrid.Rows.Count = 0 Then Exit Sub
            oForm.Freeze(True)
            oGrid.RowHeaders.Width = 20
            With oGrid.Columns
                .Item(Grid.Col_Code).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_Code).Editable = False
                .Item(Grid.Col_Code).Visible = False

                .Item(Grid.Col_Fecha).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_Fecha).Editable = False
                .Item(Grid.Col_Fecha).TitleObject.Sortable = True

                .Item(Grid.Col_Campo).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_Campo).Editable = False
                .Item(Grid.Col_Campo).TitleObject.Sortable = True


                .Item(Grid.Col_Antiguo).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_Antiguo).Editable = False
                .Item(Grid.Col_Antiguo).TitleObject.Sortable = True


                .Item(Grid.Col_Nuevo).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_Nuevo).Editable = False
                .Item(Grid.Col_Nuevo).TitleObject.Sortable = True

                .Item(Grid.Col_NombreUsuario).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_NombreUsuario).Editable = False
                .Item(Grid.Col_NombreUsuario).TitleObject.Sortable = True

                .Item(Grid.Col_Estado).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_Estado).Editable = False
                .Item(Grid.Col_Estado).Visible = False

                .Item(Grid.Col_IdUsuario).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_IdUsuario).Editable = False
                .Item(Grid.Col_IdUsuario).Visible = False

                .Item(Grid.Col_CampoBD).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_CampoBD).Editable = False
                .Item(Grid.Col_CampoBD).Visible = False

                .Item(Grid.Col_Tabla).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_Tabla).Editable = False
                .Item(Grid.Col_Tabla).Visible = False

                .Item(Grid.Col_PropiedadObjeto).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                .Item(Grid.Col_PropiedadObjeto).Editable = False
                .Item(Grid.Col_PropiedadObjeto).Visible = False

                .Item(Grid.Col_CampoClave).Type = BoGridColumnType.gct_EditText
                .Item(Grid.Col_CampoClave).Editable = False
                .Item(Grid.Col_CampoClave).Visible = False

                .Item(Grid.Col_ValorCampoClave).Type = BoGridColumnType.gct_EditText
                .Item(Grid.Col_ValorCampoClave).Editable = False
                .Item(Grid.Col_ValorCampoClave).Visible = False


                .Item(Grid.Col_Sel).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                .Item(Grid.Col_Sel).Editable = True
                .Item(Grid.Col_Sel).TitleObject.Sortable = False
            End With
            oGrid.AutoResizeColumns()
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try

    End Sub

    Public Sub HANDLE_FORM_GESTION_RESIZE(ByVal FormUID As String)
        Configurar_ColumnsGrid(FormUID)
    End Sub

#End Region

#Region "Funciones y Rutinas Generales"

    Private Function getNextNumberFromTable(ByVal TableName As String) As String
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim Result As String = String.Empty
        Dim ls As String = String.Empty
        Try
            oRs = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Select Case oCompany.DbServerType
                Case BoDataServerTypes.dst_HANADB
                    If TableName.Contains("[") Then TableName = TableName.Remove("[")
                    If TableName.Contains("]") Then TableName = TableName.Remove("]")
            End Select
            ls = "SELECT RIGHT('00000000'+ CAST(CAST(COALESCE(MAX(""Code""),0) AS INT)+1 AS VARCHAR(8)),8) FROM " & TableName
            oRs.DoQuery(ls)
            Result = oRs.Fields.Item(0).Value.ToString
        Catch ex As Exception
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
        Finally
            LiberarObjetoCOM(oRs)
        End Try
        Return Result
    End Function

    Private Function getUserAuthorization(ByVal sUserID As String, ByVal FormTypeEx As String) As SingleResponse
        Dim oResponse As New SingleResponse
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Try
            Dim sFilterField As String = String.Empty, sNombreFormulario As String = String.Empty

            Select Case FormTypeEx
                Case SEI_GEST.Default.BP_FORMTYPEEX 'Interlocutores comerciales
                    sFilterField = "U_Aut_IC"
                    sNombreFormulario = "Interlocutores Comerciales"
                Case SEI_GEST.Default.ITEM_FORMTYPEEX 'Artículos
                    sFilterField = "U_Aut_Item"
                    sNombreFormulario = "Artículos"
            End Select
            Dim oSQLQuery As New DNAUtils.DNASQLUtils.SQLQuery
            oSQLQuery.SelectFieldList.Add("ISNULL(" & sFilterField & ",'N')")
            oSQLQuery.FromTable = "OUSR"
            oSQLQuery.FilterList.Add(New DNAUtils.DNASQLUtils.SQLFilter("USER_CODE", sUserID, DNAUtils.DNASQLUtils.SQLFilter.SQLFilterOperandType.Value_, Nothing, Nothing, DNAUtils.DNASQLUtils.SQLFilter.SQLFilterOperation.Equals_, DNAUtils.DNASQLUtils.SQLFilter.SQLFilterConcatOp._END_CONCAT_))
            Dim laSelect As String = DNAUtils.DNASQLUtils.DNASQLUtils.GetInstance(oCompany).getSelectSqlSentence(oSQLQuery)
            'TEMPORAL!!. QUITAR LOS [ Y ] PORQUE CON EL ISNULL DA ERROR. PREGUNTAR A ANDER
            laSelect = laSelect.Replace("[", "").Replace("]", "")

            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(laSelect)
            If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                Dim sAutorizacion As String = oRs.Fields.Item(0).Value.ToString.Trim
                If String.IsNullOrEmpty(sAutorizacion) Then Throw New Exception("El usuario " & sUserID & " no tiene definido si dispone de autorización en el formulario de " & sNombreFormulario)
                Select Case sAutorizacion
                    Case "Y"
                        oResponse.ExecutionSuccess = True
                        oResponse.FailureReason = String.Empty
                    Case Else
                        Throw New Exception("ERROR: El usuario no tiene autorización en el formulario de " & sNombreFormulario)
                End Select
            Else
                Throw New Exception("ERROR: El usuario " & sUserID & " no tiene definido si dispone de autorización en el formulario de " & sNombreFormulario)
            End If
        Catch ex As Exception
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
            oResponse.ExecutionSuccess = False
            oResponse.FailureReason = ex.Message
        Finally
            LiberarObjetoCOM(oRs)
        End Try
        Return oResponse
    End Function

    Private Function getLogInstanceFromItem(ByVal ItemCode As String) As Integer
        'SELECT MAX(LogInstanc) FROM ACRD WHERE CardCode = '0100004'
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oResponse As Integer
        Try
            Dim ls As String = String.Empty
            ls = "SELECT MAX(""LogInstanc"") FROM AITM WHERE ""ItemCode"" = '" & ItemCode & "'"
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(ls)
            If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                oResponse = oRs.Fields.Item(0).Value
            Else
                Throw New Exception("No se detecta cambio en el Artículo " & ItemCode)
            End If
        Catch ex As Exception
            oResponse = -1
            Trace.WriteLineIf(ots.TraceError, "Error al identificar la última instancia de modificación del Artículo " & ItemCode & " : " & ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
        Finally
            LiberarObjetoCOM(oRs)
        End Try
        Return oResponse
    End Function

    Private Function getLogInstanceFromBP(ByVal CardCode As String) As Integer
        'SELECT MAX(LogInstanc) FROM ACRD WHERE CardCode = '0100004'
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oResponse As Integer
        Try
            Dim ls As String = String.Empty
            ls = "SELECT MAX(""LogInstanc"") FROM ACRD WHERE ""CardCode"" = '" & CardCode & "'"
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(ls)
            If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                oResponse = oRs.Fields.Item(0).Value
            Else
                Throw New Exception("No se detecta cambio en el IC " & CardCode)
            End If
        Catch ex As Exception
            oResponse = -1
            Trace.WriteLineIf(ots.TraceError, "Error al identificar la última instancia de modificación del IC " & CardCode & " : " & ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
        Finally
            LiberarObjetoCOM(oRs)
        End Try
        Return oResponse
    End Function

    Private Function getFieldsToMonitorize(ByVal FormTypeEx As String, Optional ByVal Subtipo As String = "") As SingleResponseCampos
        Dim oResponse As New SingleResponseCampos
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Try
            Dim ls As String = String.Empty
            ls = "SELECT * FROM ""SEI_MONITORIZEDFIELDS"" WHERE ""FormTypeEx""='" & FormTypeEx & "'"
            If Not String.IsNullOrEmpty(Subtipo) Then ls += " AND ""Code"" LIKE '%" & Subtipo & "'"
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(ls)
            If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                Dim ListaCampos As New List(Of DetalleCampos)
                While Not oRs.EoF
                    Dim Detalle As New DetalleCampos With {
                        .Tabla = oRs.Fields.Item("Tabla").Value.ToString.Trim,
                        .CampoBD = oRs.Fields.Item("CampoBD").Value.ToString.Trim,
                        .CampoObjeto = oRs.Fields.Item("CampoObjeto").Value.ToString.Trim,
                        .EsEspecial = oRs.Fields.Item("EsEspecial").Value.ToString.Trim,
                        .TablaAsociada = oRs.Fields.Item("TablaAsociada").Value.ToString.Trim,
                        .CampoAsociado = oRs.Fields.Item("CampoAsociado").Value.ToString.Trim,
                        .Descripcion = oRs.Fields.Item("Description").Value.ToString.Trim
                    }
                    ListaCampos.Add(Detalle)
                    oRs.MoveNext()
                End While
                oResponse.ExecutionSuccess = True
                oResponse.FailureReason = String.Empty
                oResponse.ListaCampos = ListaCampos
            Else
                oResponse.ExecutionSuccess = True
                oResponse.FailureReason = String.Empty
                oResponse.ListaCampos = Nothing
            End If
        Catch ex As Exception
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
            oResponse.ExecutionSuccess = False
            oResponse.FailureReason = ex.Message
            oResponse.ListaCampos = Nothing
        Finally
            If Not IsNothing(oRs) Then
                LiberarObjetoCOM(oRs)
            End If
        End Try
        Return oResponse
    End Function

    Private Function GetPropertyType(ByVal obj As Object, ByVal propertyName As String) As String
        Try
            If propertyName.StartsWith("U_") Then

            Else
                For Each oProperty As System.ComponentModel.PropertyDescriptor In System.ComponentModel.TypeDescriptor.GetProperties(obj)
                    If (oProperty.Name = propertyName) Then
                        Return oProperty.PropertyType.Name
                        Exit For
                    End If
                Next
            End If

        Catch ex As Exception
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
        End Try
    End Function

    Private Sub SetPropertyValueDynamically(ByVal obj As Object, ByVal propertyName As String, ByVal newValue As Object)
        Try
            If propertyName.StartsWith("U_") Then
                obj.UserFields.Fields.Item(propertyName).Value = newValue
            Else

                For Each oProperty As System.ComponentModel.PropertyDescriptor In System.ComponentModel.TypeDescriptor.GetProperties(obj)
                    If (oProperty.Name = propertyName) Then
                        If oProperty.PropertyType = Today.GetType() Then
                            Try
                                oProperty.SetValue(obj, newValue)
                            Catch ex As Exception
                                oProperty.SetValue(obj, DNAUtils.Converters.ConvertSAPDateToDateTime(newValue))
                            End Try
                        Else
                            oProperty.SetValue(obj, newValue)
                        End If

                        Exit For
                    End If
                Next

            End If
        Catch ex As Exception
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
        End Try
    End Sub

    Private Sub LiberarObjetoCOM(ByVal obj As Object)
        Try
            If Not IsNothing(obj) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Catch ex As Exception
        End Try
    End Sub


    Private Sub Formulario_EnableMenu(FormUID As String)

        Dim oForm As SAPbouiCOM.Form
        oForm = oApplication.Forms.Item(FormUID)

        oForm.EnableMenu(enMenuUID.MNU_Siguiente, False)
        oForm.EnableMenu(enMenuUID.MNU_Anterior, False)
        oForm.EnableMenu(enMenuUID.MNU_Primero, False)
        oForm.EnableMenu(enMenuUID.MNU_Ultimo, False)
        oForm.EnableMenu(enMenuUID.MNU_Buscar, False)
        oForm.EnableMenu(enMenuUID.MNU_Crear, False)

        oForm.EnableMenu(enMenuUID.MNU_AñadirLinea, False)
        oForm.EnableMenu(enMenuUID.MNU_EliminarLinea, False)
        oForm.EnableMenu(enMenuUID.MNU_DuplicarLinea, False)

    End Sub
    Private Sub Formulario_AffectsFormMode(FormUID As String)

        Dim oForm As SAPbouiCOM.Form
        oForm = oApplication.Forms.Item(FormUID)

        oForm.Items.Item("txtCCode").AffectsFormMode = True
        oForm.Items.Item("gDatos").AffectsFormMode = False

    End Sub

#End Region
End Class
