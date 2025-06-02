

Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class DiasLaborales

#Region "Atributos"

    Private bCreado As Boolean

#End Region

#Region "Constructores"

    Private Sub InicializarVariables()
        Try
            Mes = 0
            DiasLaborales = 0
            bCreado = False
        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub

    Public Sub New(imesLaboral As Integer)
        Try

            Dim sSQl As String = "SELECT * " &
                              " FROM DiasLaborales with(nolock) " &
                              " WHERE dlMes=" & UTrim(imesLaboral)

            Dim DTDatos As New DataTable

            InicializarVariables()
            If Datos.CGPL.DameDatosDT(sSQl, DTDatos) Then
                Mes = CInt((NoNull(DTDatos.Rows(0).Item("dlMes"), "N")))
                DiasLaborales = CInt((NoNull(DTDatos.Rows(0).Item("dlDiasLaborales"), "N")))
                Me.bCreado = True
            End If


        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)

        End Try
    End Sub



#End Region

#Region "Propiedades"

    Public Property Mes As Integer
    Public Property DiasLaborales As Integer
    Public ReadOnly Property Creado As Boolean
        Get
            Creado = bCreado
        End Get
    End Property


#End Region

#Region "BBDD"
    Public Function Insertar() As Boolean
        Try
            Dim CodigoSQL As Integer
            Dim sSql As String = "INSERT INTO DiasLaborales (dlMes,dlDiasLaborales) " &
                                 " VALUES (" & Mes & "," &
                                                DiasLaborales & ") SELECT @@IDENTITY "

            CodigoSQL = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If CodigoSQL = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
                Datos.GuardarLog(TipoLogDescripcion.Alta & " DiasLaborales", CStr(Mes))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE DiasLaborales SET dlDiasLaborales=" & DiasLaborales.ToString & "" &
                                 " WHERE dlMes=" & Mes & ""

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & "DiasLaborales", DiasLaborales.ToString)
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM DiasLaborales " &
                                  " WHERE dlMes=" & Mes & ""

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " DiasLaborales", CStr(Mes))
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function
#End Region

End Class
