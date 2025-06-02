

Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class Formatos
#Region "Atributos"

    Private bCreado As Boolean

#End Region

#Region "Constructores"

    Private Sub InicializarVariables()
        Try
            Codigo = ""
            Nombre = ""
            bCreado = False
        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub

    Public Sub New(sCodigo As String)
        Try

            Dim sSQl As String = "SELECT * " &
                              " FROM formatos with(nolock) " &
                              " WHERE fmCod=" & UTrim(sCodigo)

            Dim DTDatos As New DataTable

            InicializarVariables()
            If Datos.CGPL.DameDatosDT(sSQl, DTDatos) Then
                Codigo = sCodigo
                Nombre = CStr((NoNull(DTDatos.Rows(0).Item("fmNombre"), "A")))
                Me.bCreado = True
            End If


        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)

        End Try
    End Sub



#End Region

#Region "Propiedades"

    Public Property Codigo As String
    Public Property Nombre As String
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
            Dim sSql As String = "INSERT INTO Formatos (fmcod,fmNombre) " &
                                 " VALUES (" & Codigo & "," &
                                                Nombre & ") SELECT @@IDENTITY "

            CodigoSQL = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If CodigoSQL = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
                Datos.GuardarLog(TipoLogDescripcion.Alta & " Formatos", CStr(Codigo))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE formatos SET fmNombre=" & Nombre & "" &
                                 " WHERE fmCod=" & Codigo & ""

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & "Formatos", Nombre)
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM formatos " &
                                  " WHERE fmCod=" & Codigo & ""

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Formatos", CStr(Codigo))
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function
#End Region
End Class
