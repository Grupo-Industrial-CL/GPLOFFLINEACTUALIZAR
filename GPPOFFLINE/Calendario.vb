
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP
Public Class Calendario
#Region "Atributos"

    Private iId As Integer
    Private iCodPuestotrabajo As Integer
    Private dFechaturno As Date
    Private cTurno As Char
    Private dInicio As DateTime
    Private dFin As DateTime
    Private iOperarios As Integer
    Private iMinutosTurno As Integer
    Private iTurnoRecalcular As Boolean
    Private iTurnoRecalcularHistorico As Boolean
    Private miPuestoTrabajo As PuestosTrabajo

    Private bCreado As Boolean

#End Region

#Region "Constructores"



    Private Sub InicializarVariables()
        Try
            iId = 0
            iCodPuestotrabajo = 0
            dFechaturno = ConstantesGPP.FechaGlobal
            dInicio = ConstantesGPP.FechaGlobal
            dFin = ConstantesGPP.FechaGlobal
            iOperarios = 0
            iMinutosTurno = 0
            cTurno = Nothing
            iTurnoRecalcular = False
            bCreado = False

        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub

    Public Sub New(ByVal Codigo As Integer,
                   ByVal Cod_PuestoTrabajo As Integer,
                   ByVal Fecha_turno As Date,
                   ByVal Turno As Char,
                   ByVal Fec_Inicio As DateTime,
                   ByVal Fec_Fin As DateTime,
                   ByVal Num_Operarios As Integer,
                   ByVal Minutos_Turno As Integer)
        Try
            InicializarVariables()

            iId = Codigo
            iCodPuestotrabajo = Cod_PuestoTrabajo
            dFechaturno = Fecha_turno
            dInicio = Fec_Inicio
            dFin = Fec_Fin
            iOperarios = Num_Operarios
            iMinutosTurno = Minutos_Turno
            cTurno = Turno

            bCreado = True

        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(ByVal Codigo As Integer,
                   ByVal TurnoRecalcular As Boolean,
                   ByVal Cod_PuestoTrabajo As Integer,
                   ByVal Fecha_turno As Date,
                   ByVal Turno As Char,
                   ByVal Fec_Inicio As DateTime,
                   ByVal Fec_Fin As DateTime,
                   ByVal Num_Operarios As Integer,
                   ByVal TurnoRecalcularHistorico As Boolean,
                   ByVal Minutos_Turno As Integer)
        Try
            InicializarVariables()

            iId = Codigo
            iTurnoRecalcular = TurnoRecalcular
            iCodPuestotrabajo = Cod_PuestoTrabajo
            dFechaturno = Fecha_turno
            dInicio = Fec_Inicio
            dFin = Fec_Fin
            iOperarios = Num_Operarios
            iTurnoRecalcularHistorico = TurnoRecalcularHistorico
            iMinutosTurno = Minutos_Turno
            cTurno = Turno

            bCreado = True

        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(Codigo As Integer)
        Try
            Dim sSQl As String = "SELECT * " &
                                 " FROM Calendario " &
                                 " WHERE clId=" & Codigo

            Dim DTDatos As New DataTable

            InicializarVariables()
            If Datos.CGPL.DameDatosDT(sSQl, DTDatos) Then
                With DTDatos.Rows(0)
                    iId = CInt(NoNull(.Item("clId"), "D"))
                    iCodPuestotrabajo = CInt(NoNull(.Item("clPuestoTrabajo"), "D"))
                    cTurno = CChar(NoNull(.Item("clTurno"), "A"))

                    dFechaturno = CDate(NoNull(.Item("clFechaTurno"), "DT"))
                    dInicio = CDate(NoNull(.Item("clInicioTurno"), "DT"))
                    dFin = CDate(NoNull(.Item("clFinTurno"), "DT"))

                    iOperarios = CInt(NoNull(.Item("clOperarios"), "D"))
                    iMinutosTurno = CInt(NoNull(.Item("clMinutosTurno"), "D"))
                End With

                Me.bCreado = True
            End If
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

#End Region

#Region "Propiedades"

    Public ReadOnly Property Creado As Boolean
        Get
            Creado = Me.bCreado
        End Get
    End Property

    Public ReadOnly Property PuestoTrabajo As PuestosTrabajo
        Get
            If miPuestoTrabajo Is Nothing Then
                miPuestoTrabajo = New PuestosTrabajo(iCodigoPuestoTrabajo:=Me.iCodPuestotrabajo)
            ElseIf miPuestoTrabajo.Creado = False Then
                miPuestoTrabajo = New PuestosTrabajo(iCodigoPuestoTrabajo:=Me.iCodPuestotrabajo)
            End If

            Return miPuestoTrabajo
        End Get
    End Property

    Public Property Id As Integer
        Get
            Return Me.iId
        End Get
        Set(value As Integer)
            Me.iId = value
        End Set
    End Property

    Public Property CodPuestotrabajo As Integer
        Get
            Return Me.iCodPuestotrabajo
        End Get
        Set(value As Integer)
            Me.iCodPuestotrabajo = value
        End Set
    End Property

    Public Property FechaTurno As Date
        Get
            Return Me.dFechaturno
        End Get
        Set(value As Date)
            Me.dFechaturno = value
        End Set
    End Property



    Public Property Turno As Char
        Get
            Return Me.cTurno
        End Get
        Set(value As Char)
            Me.cTurno = value
        End Set
    End Property

    Public Property InicioTurno As DateTime
        Get
            Return Me.dInicio
        End Get
        Set(value As DateTime)
            Me.dInicio = value
        End Set
    End Property

    Public Property FinTurno As DateTime
        Get
            Return Me.dFin
        End Get
        Set(value As DateTime)
            Me.dFin = value
        End Set
    End Property

    Public Property Operarios As Integer
        Get
            Return Me.iOperarios
        End Get
        Set(value As Integer)
            Me.iOperarios = value
        End Set
    End Property


    Public Property TurnoRecalcular As Boolean
        Get
            Return Me.iTurnoRecalcular
        End Get
        Set(value As Boolean)
            Me.iTurnoRecalcular = value
        End Set
    End Property

    Public ReadOnly Property MinutosTurno As Integer
        Get
            Return Me.iMinutosTurno
        End Get
    End Property

    Public Property TurnoRecalcularHistorico As Boolean
        Get
            Return Me.iTurnoRecalcularHistorico
        End Get
        Set(value As Boolean)
            Me.iTurnoRecalcularHistorico = value
        End Set
    End Property

#End Region

#Region "BBDD"
    Public Function Insertar() As Boolean
        Try
            Dim sSql As String = "INSERT INTO Calendario (clPuestoTrabajo,clFechaTurno,clTurno,clInicioTurno,clFinTurno,clOperarios) " &
                                 " VALUES (" & Me.iCodPuestotrabajo & ",'" &
                                                Me.dFechaturno & "','" &
                                                Me.cTurno & "','" &
                                                Me.dInicio & "','" &
                                                Me.dFin & "'," &
                                                Me.iOperarios & ") SELECT @@IDENTITY "

            Me.iId = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If Me.iId = -1 Then
                Insertar = False
            Else
                Insertar = True

                Datos.GuardarLog(TipoLogDescripcion.Alta & " Calendario", CStr(Me.iId))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE Calendario SET clPuestoTrabajo=" & Me.iCodPuestotrabajo & "," &
                                 "clFechaTurno='" & Me.dFechaturno & "'," &
                                 "clTurno='" & Me.cTurno & "'," &
                                 "clInicioTurno='" & Me.dInicio & "'," &
                                 "clFinTurno='" & Me.dFin & "'," &
                                 "clOperarios=" & Me.iOperarios &
                                 " WHERE clId=" & Me.iId

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)
            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " Calendario", CStr(Me.iId))
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM Calendario " &
                                 "WHERE clId=" & Me.iId

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Calendario", CStr(Me.iId))
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function EliminarPorFecha() As Boolean
        Try
            Dim sSql As String = "DELETE FROM Calendario " &
                                 "WHERE clPuestoTrabajo=" & Me.iCodPuestotrabajo & " and clFechaTurno='" & Me.dFechaturno & "'" & " and clTurno='" & Me.Turno & "'"

            Datos.CGPL.EjecutarConsulta(sSql)
            EliminarPorFecha = True
        Catch ex As Exception
            EliminarPorFecha = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

#End Region
End Class
