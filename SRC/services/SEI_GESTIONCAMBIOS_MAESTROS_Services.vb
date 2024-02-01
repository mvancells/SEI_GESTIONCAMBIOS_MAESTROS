Public Class SEI_GESTIONCAMBIOS_MAESTROS_Services

#Region "Constructor/Singleton"

    Private Shared Instance As SEI_GESTIONCAMBIOS_MAESTROS_Services = Nothing

    Dim ots As New TraceSwitch("TraceSwitch", "Switch del Traceador")

    Private oCompany As SAPbobsCOM.Company = Nothing

    Public Sub New(ByRef _oCompany As SAPbobsCOM.Company)
        oCompany = _oCompany
    End Sub

    Public Shared Function GetInstance(ByRef _oCompany As SAPbobsCOM.Company) As SEI_GESTIONCAMBIOS_MAESTROS_Services

        If Instance Is Nothing Then
            Instance = New SEI_GESTIONCAMBIOS_MAESTROS_Services(_oCompany)
        End If

        Instance.oCompany = _oCompany

        Return Instance

    End Function

#End Region


End Class