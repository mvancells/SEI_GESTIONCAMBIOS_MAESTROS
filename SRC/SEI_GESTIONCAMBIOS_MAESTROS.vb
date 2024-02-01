'----------------------------------------------------------------------------------------------------
'This is an automagically generated template for an Addonia AddOn Module.
'Please refer to INSTRUCTIONS.txt for further help.
'
'ModuleName: SEI_GESTIONCAMBIOS_MAESTROS
'Author: jsalgueiro.SEIDORBCN  @  PC03D36U
'Date: 10/16/2023 7:59:09 PM
'----------------------------------------------------------------------------------------------------
Imports Microsoft.VisualBasic.CompilerServices
Imports ModuleBase
Imports SAPbouiCOM
Imports SEI_GESTIONCAMBIOS_MAESTROS.My

Public Class SEI_GESTIONCAMBIOS_MAESTROS
    Inherits ModuleBase.AbstractModuleModel

#Region "Implementation Module Property"
    Private _implementation As SEI_GESTIONCAMBIOS_MAESTROSImpl
    Public Property oImplementation() As SEI_GESTIONCAMBIOS_MAESTROSImpl
        Get
            Return _implementation
        End Get
        Set(ByVal value As SEI_GESTIONCAMBIOS_MAESTROSImpl)
            _implementation = value
        End Set
    End Property

    Public Overrides Sub InitializeImplementation()
        _implementation = New SEI_GESTIONCAMBIOS_MAESTROSImpl
        _implementation.initializeModule(oApplication, oCompany, oAddOnService)
        _implementation.oModule = Me
    End Sub
#End Region


    Public Overrides Sub queryMetadataCreationConfig(ByRef oMetadataCreationSelector As ModuleBase.cMetadaCreationSelector)
        oMetadataCreationSelector.ShouldCreateMetadata = True
        oMetadataCreationSelector.ShouldCreateUserTables = True
        oMetadataCreationSelector.ShouldCreateUserFields = True
        oMetadataCreationSelector.ShouldCreateUDOs = True
        oMetadataCreationSelector.ShouldCreateSQLviews = True
    End Sub

    Public Overrides Function getFilterCollection() As List(Of cEventFilter)
        addFilter(SEI_GEST.Default.BP_FORMTYPEEX, BoEventTypes.et_FORM_LOAD)
        addFilter(SEI_GEST.Default.BP_FORMTYPEEX, BoEventTypes.et_FORM_DATA_UPDATE)
        addFilter(SEI_GEST.Default.BP_FORMTYPEEX, BoEventTypes.et_FORM_CLOSE)

        addFilter(SEI_GEST.Default.ITEM_FORMTYPEEX, BoEventTypes.et_FORM_LOAD)
        addFilter(SEI_GEST.Default.ITEM_FORMTYPEEX, BoEventTypes.et_FORM_DATA_UPDATE)
        addFilter(SEI_GEST.Default.ITEM_FORMTYPEEX, BoEventTypes.et_FORM_CLOSE)

        addFilter(SEI_GEST.Default.GESTIONIC_FORMID, BoEventTypes.et_FORM_RESIZE)
        addFilter(SEI_GEST.Default.GESTIONIC_FORMID, BoEventTypes.et_ITEM_PRESSED)
        addFilter(SEI_GEST.Default.GESTIONIC_FORMID, BoEventTypes.et_CHOOSE_FROM_LIST)

        Return MyBase.getFilterCollection()
    End Function


    Public Overrides Sub oApplication_FormDataEvent(ByRef BusinessObjectInfo As BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        MyBase.oApplication_FormDataEvent(BusinessObjectInfo, BubbleEvent)
#Region "Interlocutores Comerciales"
        If BusinessObjectInfo.FormTypeEx = SEI_GEST.Default.BP_FORMTYPEEX And BusinessObjectInfo.BeforeAction And BusinessObjectInfo.EventType = BoEventTypes.et_FORM_DATA_UPDATE Then
            oImplementation.HANDLE_OK_BUTTON_PRESSED_BEFORE_ACTION_BPFORM(BusinessObjectInfo)
        End If
        'Si el usuario no tiene permiso, entra
        If Not oImplementation.YaEntre And Not oImplementation.UsuarioConPermisosIC And BusinessObjectInfo.FormTypeEx = SEI_GEST.Default.BP_FORMTYPEEX And Not BusinessObjectInfo.BeforeAction And BusinessObjectInfo.ActionSuccess And BusinessObjectInfo.EventType = BoEventTypes.et_FORM_DATA_UPDATE Then
            oImplementation.HANDLE_OK_BUTTON_PRESSED_AFTER_ACTION_BPFORM(BusinessObjectInfo)
        End If
#End Region
#Region "Artículos"

        If BusinessObjectInfo.FormTypeEx = SEI_GEST.Default.ITEM_FORMTYPEEX And BusinessObjectInfo.BeforeAction And BusinessObjectInfo.EventType = BoEventTypes.et_FORM_DATA_UPDATE Then
            oImplementation.HANDLE_OK_BUTTON_PRESSED_BEFORE_ACTION_ITEMSFORM(BusinessObjectInfo)
        End If
        'Si el usuario no tiene permiso, entra
        If Not oImplementation.YaEntre And Not oImplementation.UsuarioConPermisosItems And BusinessObjectInfo.FormTypeEx = SEI_GEST.Default.ITEM_FORMTYPEEX And Not BusinessObjectInfo.BeforeAction And BusinessObjectInfo.ActionSuccess And BusinessObjectInfo.EventType = BoEventTypes.et_FORM_DATA_UPDATE Then
            oImplementation.HANDLE_OK_BUTTON_PRESSED_AFTER_ACTION_ITEMSFORM(BusinessObjectInfo)
        End If


#End Region
    End Sub

    'TODO: Write you module here...
    Public Overrides Sub oApplication_ItemEvent(FormUID As String, ByRef pVal As ItemEvent, ByRef BubbleEvent As Boolean)
        MyBase.oApplication_ItemEvent(FormUID, pVal, BubbleEvent)
        Select Case pVal.FormTypeEx

#Region "Maestros IC & Artículos"
            Case SEI_GEST.Default.BP_FORMTYPEEX
                If pVal.EventType = BoEventTypes.et_FORM_CLOSE And pVal.ActionSuccess Then
                    oImplementation.UsuarioConPermisosIC = False
                    oImplementation.YaEntre = False
                End If
            Case SEI_GEST.Default.ITEM_FORMTYPEEX
                If pVal.EventType = BoEventTypes.et_FORM_CLOSE And pVal.ActionSuccess Then
                    oImplementation.UsuarioConPermisosItems = False
                    oImplementation.YaEntre = False
                End If

#End Region

#Region "Gestión ICs"
            Case SEI_GEST.Default.GESTIONIC_FORMID
                If pVal.EventType = BoEventTypes.et_FORM_RESIZE And Not pVal.BeforeAction Then
                    oImplementation.HANDLE_FORM_GESTION_RESIZE(FormUID)
                End If
                If pVal.EventType = BoEventTypes.et_CHOOSE_FROM_LIST And pVal.ItemUID = oImplementation.Cab.txtCardCode Then
                    oImplementation.HANDLE_FORM_GESTION_CHOOSEFROMLIST(pVal, BubbleEvent, FormUID)
                End If
                If pVal.EventType = BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = oImplementation.Cab.Boton_Filtrar Then
                    oImplementation.HANDLE_FORM_GESTION_BOTON_FILTRAR_ITEM_PRESSED(pVal, BubbleEvent, FormUID)
                End If
                If pVal.EventType = BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = oImplementation.Cab.Boton_Grabar Then
                    oImplementation.HANDLE_FORM_GESTION_BOTON_GRABAR_ITEM_PRESSED(pVal, BubbleEvent, FormUID)
                End If
                If pVal.EventType = BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = oImplementation.Cab.Boton_Descartar Then
                    oImplementation.HANDLE_FORM_GESTION_BOTON_DESCARTAR_ITEM_PRESSED(pVal, BubbleEvent, FormUID)
                End If
#End Region

#Region "Gestión ICs"

#End Region

        End Select

    End Sub

    Public Overrides Sub oApplication_MenuEvent(ByRef pVal As MenuEvent, ByRef BubbleEvent As Boolean)
        MyBase.oApplication_MenuEvent(pVal, BubbleEvent)
        If pVal.MenuUID = SEI_GEST.Default.MENU_AutorizacionPr And Not pVal.BeforeAction Then
            oImplementation.HANDLE_MENU_CLICK(SEI_GEST.Default.BP_FORMTYPEEX, "S")
        End If

        If pVal.MenuUID = SEI_GEST.Default.MENU_AutorizacionCl And Not pVal.BeforeAction Then
            oImplementation.HANDLE_MENU_CLICK(SEI_GEST.Default.BP_FORMTYPEEX, "C")
        End If

        If pVal.MenuUID = SEI_GEST.Default.MENU_AutorizacionA And Not pVal.BeforeAction Then
            oImplementation.HANDLE_MENU_CLICK(SEI_GEST.Default.ITEM_FORMTYPEEX)
        End If
    End Sub


End Class
