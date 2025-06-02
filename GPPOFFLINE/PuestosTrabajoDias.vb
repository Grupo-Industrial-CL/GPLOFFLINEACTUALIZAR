
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class PuestosTrabajoDias

#Region "Atributos"
    Private iCodPuestotrabajo As Integer
    Private sNombrePuestoTrabajo As String
    Private sHoraInicioM As TimeSpan
    Private sHoraInicioT As TimeSpan
    Private sHoraInicioN As TimeSpan
    Private sHoraFinM As TimeSpan
    Private sHoraFinT As TimeSpan
    Private sHoraFinN As TimeSpan
    Private sInicioM As DateTime
    Private sInicioT As DateTime
    Private sInicioN As DateTime
    Private sFinM As DateTime
    Private sFinT As DateTime
    Private sFinN As DateTime
    Private sDias As Integer
    Private bTurnoM As Boolean
    Private bTurnoT As Boolean
    Private bTurnoN As Boolean
    Private iOperarios As Integer
    Private iOperariosM As Integer
    Private iOperariosT As Integer
    Private iOperariosN As Integer
    Private bCreado As Boolean
    Private iId As Integer
    Private iClId As Integer
    Private iFecha As Date
    Private sObservaciones As String
    Private sMes As String

#End Region

#Region "Constructores"

    Private Sub InicializarVariables()
        Try
            'iId = 0
            'iCodPuestotrabajo = 0
            'dFechaturno = ConstantesGPP.FechaGlobal
            'dInicio = ConstantesGPP.FechaGlobal
            'dFin = ConstantesGPP.FechaGlobal
            'iOperarios = 0
            'iMinutosTurno = 0
            'cTurno = Nothing

            bCreado = False

        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub
    Public Sub New(CodPuestotrabajo As Integer)
        InicializarVariables()
    End Sub

    Public Sub New(ByVal CodPuestotrabajo As Integer,
                   ByVal Fecha As Date,
                   ByVal Dias As Integer,
                   ByVal InicioM As DateTime,
                   ByVal InicioT As DateTime,
                   ByVal InicioN As DateTime,
                   ByVal FinM As DateTime,
                   ByVal FinT As DateTime,
                   ByVal FinN As DateTime,
                   ByVal TurnoM As Boolean,
                   ByVal TurnoT As Boolean,
                   ByVal TurnoN As Boolean,
                   ByVal Operarios As Integer,
                   ByVal OperariosM As Integer,
                   ByVal OperariosT As Integer,
                   ByVal OperariosN As Integer,
                   ByVal Observaciones As String,
                    ByVal HoraInicioM As TimeSpan,
                    ByVal HoraInicioT As TimeSpan,
                    ByVal HoraInicioN As TimeSpan,
                    ByVal HoraFinM As TimeSpan,
                    ByVal HoraFinT As TimeSpan,
                    ByVal HoraFinN As TimeSpan
                   )
        Try
            InicializarVariables()

            iCodPuestotrabajo = CodPuestotrabajo
            iFecha = Fecha
            sDias = Dias ' If(Fecha.Date.Day.ToString().Length = 1, "0" & Fecha.Date.Day.ToString(), Fecha.Date.Day.ToString())
            sInicioM = InicioM
            sInicioT = InicioT
            sInicioN = InicioN
            sFinM = FinM
            sFinT = FinT
            sFinN = FinN
            bTurnoM = TurnoM
            bTurnoT = TurnoT
            bTurnoN = TurnoN
            iOperarios = Operarios
            iOperariosM = OperariosM
            iOperariosT = OperariosT
            iOperariosN = OperariosN
            sObservaciones = Observaciones
            sHoraInicioM = HoraInicioM
            sHoraInicioT = HoraInicioT
            sHoraInicioN = HoraInicioN
            sHoraFinM = HoraFinM
            sHoraFinT = HoraFinT
            sHoraFinN = HoraFinN
            bCreado = True

        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub
    Public Sub New(ByVal CodPuestotrabajo As Integer,
                   ByVal Fecha As Date,
                   ByVal Dias As Integer,
                   ByVal InicioM As DateTime,
                   ByVal InicioT As DateTime,
                   ByVal InicioN As DateTime,
                   ByVal FinM As DateTime,
                   ByVal FinT As DateTime,
                   ByVal FinN As DateTime,
                   ByVal TurnoM As Boolean,
                   ByVal TurnoT As Boolean,
                   ByVal TurnoN As Boolean,
                   ByVal Operarios As Integer,
                   ByVal OperariosM As Integer,
                   ByVal OperariosT As Integer,
                   ByVal OperariosN As Integer,
                   ByVal Observaciones As String,
                    ByVal HoraInicioM As TimeSpan,
                    ByVal HoraInicioT As TimeSpan,
                    ByVal HoraInicioN As TimeSpan,
                    ByVal HoraFinM As TimeSpan,
                    ByVal HoraFinT As TimeSpan,
                    ByVal HoraFinN As TimeSpan,
                   ByVal Mes As String
                   )
        Try
            InicializarVariables()

            iCodPuestotrabajo = CodPuestotrabajo
            iFecha = Fecha
            sDias = Dias ' If(Fecha.Date.Day.ToString().Length = 1, "0" & Fecha.Date.Day.ToString(), Fecha.Date.Day.ToString())
            sInicioM = InicioM
            sInicioT = InicioT
            sInicioN = InicioN
            sFinM = FinM
            sFinT = FinT
            sFinN = FinN
            bTurnoM = TurnoM
            bTurnoT = TurnoT
            bTurnoN = TurnoN
            iOperarios = Operarios
            iOperariosM = OperariosM
            iOperariosT = OperariosT
            iOperariosN = OperariosN
            sObservaciones = Observaciones
            sHoraInicioM = HoraInicioM
            sHoraInicioT = HoraInicioT
            sHoraInicioN = HoraInicioN
            sHoraFinM = HoraFinM
            sHoraFinT = HoraFinT
            sHoraFinN = HoraFinN
            bCreado = True
            sMes = Mes
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

#End Region


#Region "Propiedades"


    Public Property Fecha As Date
        Get
            Return Me.iFecha
        End Get
        Set(value As Date)
            Me.iFecha = value
        End Set
    End Property
    Public ReadOnly Property Creado As Boolean
        Get
            Creado = Me.bCreado
        End Get
    End Property


    Public Property ClId As Integer
        Get
            Return Me.iClId
        End Get
        Set(value As Integer)
            Me.iClId = value
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

    Public Property Dias As Integer
        Get
            Return Me.sDias
        End Get
        Set(value As Integer)
            Me.sDias = value
        End Set
    End Property

    Public Property NombrePuestotrabajo As String
        Get
            Return Me.sNombrePuestoTrabajo
        End Get
        Set(value As String)
            Me.sNombrePuestoTrabajo = value
        End Set
    End Property

    Public Property TurnoM As Boolean
        Get
            Return Me.bTurnoM
        End Get
        Set(value As Boolean)
            Me.bTurnoM = value
        End Set
    End Property

    Public Property TurnoT As Boolean
        Get
            Return Me.bTurnoT
        End Get
        Set(value As Boolean)
            Me.bTurnoT = value
        End Set
    End Property
    Public Property TurnoN As Boolean
        Get
            Return Me.bTurnoN
        End Get
        Set(value As Boolean)
            Me.bTurnoN = value
        End Set
    End Property
    Public Property OperariosM As Integer
        Get
            Return Me.iOperariosM
        End Get
        Set(value As Integer)
            Me.iOperariosM = value
        End Set
    End Property

    Public Property OperariosT As Integer
        Get
            Return Me.iOperariosT
        End Get
        Set(value As Integer)
            Me.iOperariosT = value
        End Set
    End Property

    Public Property OperariosN As Integer
        Get
            Return Me.iOperariosN
        End Get
        Set(value As Integer)
            Me.iOperariosN = value
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


    Public Property HoraInicioM As TimeSpan
        Get
            Return Me.sHoraInicioM
        End Get
        Set(value As TimeSpan)
            Me.sHoraInicioM = value
        End Set
    End Property

    Public Property HoraInicioT As TimeSpan
        Get
            Return Me.sHoraInicioT
        End Get
        Set(value As TimeSpan)
            Me.sHoraInicioT = value
        End Set
    End Property

    Public Property HoraInicioN As TimeSpan
        Get
            Return Me.sHoraInicioN
        End Get
        Set(value As TimeSpan)
            Me.sHoraInicioN = value
        End Set
    End Property


    Public Property HoraFinM As TimeSpan
        Get
            Return Me.sHoraFinM
        End Get
        Set(value As TimeSpan)
            Me.sHoraFinM = value
        End Set
    End Property

    Public Property HoraFinT As TimeSpan
        Get
            Return Me.sHoraFinT
        End Get
        Set(value As TimeSpan)
            Me.sHoraFinT = value
        End Set
    End Property

    Public Property HoraFinN As TimeSpan
        Get
            Return Me.sHoraFinN
        End Get
        Set(value As TimeSpan)
            Me.sHoraFinN = value
        End Set
    End Property

    Public Property InicioM As DateTime
        Get
            Return Me.sInicioM
        End Get
        Set(value As DateTime)
            Me.sInicioM = value
        End Set
    End Property

    Public Property InicioT As DateTime
        Get
            Return Me.sInicioT
        End Get
        Set(value As DateTime)
            Me.sInicioT = value
        End Set
    End Property

    Public Property InicioN As DateTime
        Get
            Return Me.sInicioN
        End Get
        Set(value As DateTime)
            Me.sInicioN = value
        End Set
    End Property

    Public Property FinM As DateTime
        Get
            Return Me.sFinM
        End Get
        Set(value As DateTime)
            Me.sFinM = value
        End Set
    End Property

    Public Property FinT As DateTime
        Get
            Return Me.sFinT
        End Get
        Set(value As DateTime)
            Me.sFinT = value
        End Set
    End Property

    Public Property FinN As DateTime
        Get
            Return Me.sFinN
        End Get
        Set(value As DateTime)
            Me.sFinN = value
        End Set
    End Property

    Public Property Observaciones As String
        Get
            Return Me.sObservaciones
        End Get
        Set(value As String)
            Me.sObservaciones = value
        End Set
    End Property

    Public Property Mes As String
        Get
            Return Me.sMes
        End Get
        Set(value As String)
            Me.sMes = value
        End Set
    End Property
    'Public Shared Narrowing Operator CType(v As Object) As CalendarioPeriodo
    '    Throw New NotImplementedException()
    'End Operator
#End Region


#Region "BBDD"
    Public Function Insertar() As Boolean
        Try
            ' Dim sSql As String = "INSERT INTO [CalendarioAnual]
            '([clPuestoTrabajo]
            ',[clFecha]
            ',[clMañana]
            ',[clInicioMañana]
            ',[clFinMañana]
            ',[clTarde]
            ',[clInicioTarde]
            ',[clFinTarde]
            ',[clNoche]
            ',[clInicioNoche]
            ',[clFinNoche]
            ',[clObservaciones]
            ',[clOperarios])
            ' VALUES(" &
            ' Me.iCodPuestotrabajo & ",'" &
            ' Me.iFecha & "'," &
            ' If(Me.bTurnoM, 1, 0) & ",'" &
            ' New DateTime(Me.sInicioM.Year, Me.sInicioM.Month, Me.sInicioM.Day, Me.sHoraInicioM.Hours, Me.sHoraInicioM.Minutes, Me.sHoraInicioM.Seconds) & "','" &
            ' New DateTime(Me.sFinM.Year, Me.sFinM.Month, Me.sFinM.Day, Me.sHoraFinM.Hours, Me.sHoraFinM.Minutes, Me.sHoraFinM.Seconds) & "'," &
            ' If(Me.bTurnoT, 1, 0) & ",'" &
            ' New DateTime(Me.sInicioT.Year, Me.sInicioT.Month, Me.sInicioT.Day, Me.sHoraInicioT.Hours, Me.sHoraInicioT.Minutes, Me.sHoraInicioT.Seconds) & "','" &
            ' New DateTime(Me.sFinT.Year, Me.sFinT.Month, Me.sFinT.Day, Me.sHoraFinT.Hours, Me.sHoraFinT.Minutes, Me.sHoraFinT.Seconds) & "'," &
            ' If(Me.bTurnoN, 1, 0) & ",'" &
            ' New DateTime(Me.sInicioN.Year, Me.sInicioN.Month, Me.sInicioN.Day, Me.sHoraInicioN.Hours, Me.sHoraInicioN.Minutes, Me.sHoraInicioN.Seconds) & "','" &
            ' New DateTime(Me.sFinN.Year, Me.sFinN.Month, Me.sFinN.Day, Me.sHoraFinN.Hours, Me.sHoraFinN.Minutes, Me.sHoraFinN.Seconds) & "','" &
            ' Me.Observaciones & "'," &
            ' Me.iOperarios & ") SELECT @@IDENTITY "

            Dim sSql As String = "INSERT INTO [CalendarioProduccion]
           ([caPuestoTrabajo]
           ,[caFecha]
           ,[caTurnoM]
           ,[caInicioTurnoM]
           ,[caFinTurnoM]
           ,[caTurnoT]
           ,[caInicioTurnoT]
           ,[caFinTurnoT]
           ,[caTurnoN]
           ,[caInicioTurnoN]
           ,[caFinTurnoN]
           ,[caComentarios]
           ,[caNumOperarios]
            ,[caNumOperariosM]
            ,[caNumOperariosT]
            ,[caNumOperariosN])
            VALUES(" &
            Me.iCodPuestotrabajo & ",'" &
            Me.iFecha & "'," &
            If(Me.bTurnoM, 1, 0) & ",'" &
            New DateTime(Me.sInicioM.Year, Me.sInicioM.Month, Me.sInicioM.Day, Me.sHoraInicioM.Hours, Me.sHoraInicioM.Minutes, Me.sHoraInicioM.Seconds) & "','" &
            New DateTime(Me.sFinM.Year, Me.sFinM.Month, Me.sFinM.Day, Me.sHoraFinM.Hours, Me.sHoraFinM.Minutes, Me.sHoraFinM.Seconds) & "'," &
            If(Me.bTurnoT, 1, 0) & ",'" &
            New DateTime(Me.sInicioT.Year, Me.sInicioT.Month, Me.sInicioT.Day, Me.sHoraInicioT.Hours, Me.sHoraInicioT.Minutes, Me.sHoraInicioT.Seconds) & "','" &
            New DateTime(Me.sFinT.Year, Me.sFinT.Month, Me.sFinT.Day, Me.sHoraFinT.Hours, Me.sHoraFinT.Minutes, Me.sHoraFinT.Seconds) & "'," &
            If(Me.bTurnoN, 1, 0) & ",'" &
            New DateTime(Me.sInicioN.Year, Me.sInicioN.Month, Me.sInicioN.Day, Me.sHoraInicioN.Hours, Me.sHoraInicioN.Minutes, Me.sHoraInicioN.Seconds) & "','" &
            New DateTime(Me.sFinN.Year, Me.sFinN.Month, Me.sFinN.Day, Me.sHoraFinN.Hours, Me.sHoraFinN.Minutes, Me.sHoraFinN.Seconds) & "','" &
            Me.Observaciones & "'," &
            Me.iOperarios & "," & Me.iOperariosM & "," & Me.iOperariosT & "," & Me.iOperariosN & ") SELECT @@IDENTITY "

            Me.iId = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If Me.iId = -1 Then
                Insertar = False
            Else
                Insertar = True

                'Datos.GuardarLog(TipoLogDescripcion.Alta & " Calendario", CStr(Me.iId))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    'Public Function Modificar() As Boolean
    '    Try
    '        Dim sSql As String = "UPDATE Calendario SET clPuestoTrabajo=" & Me.iCodPuestotrabajo & "," &
    '                             "clFechaTurno='" & Me.dFechaturno & "'," &
    '                             "clTurno='" & Me.cTurno & "'," &
    '                             "clInicioTurno='" & Me.dInicio & "'," &
    '                             "clFinTurno='" & Me.dFin & "'," &
    '                             "clOperarios=" & Me.iOperarios &
    '                             " WHERE clId=" & Me.iId

    '        Modificar = Datos.CGPL.EjecutarConsulta(sSql)
    '        If Modificar Then
    '            Datos.GuardarLog(TipoLogDescripcion.Modificar & " Calendario", CStr(Me.iId))
    '        End If
    '    Catch ex As Exception
    '        Modificar = False
    '        Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Function

    'Public Function Eliminar() As Boolean
    '    Try
    '        Dim sSql As String = "DELETE FROM Calendario " &
    '                             "WHERE clId=" & Me.iId

    '        Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
    '        If Eliminar Then
    '            Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Calendario", CStr(Me.iId))
    '        End If
    '    Catch ex As Exception
    '        Eliminar = False
    '        Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Function

    Public Function EliminarPorFecha() As Boolean
        Try

            'Dim sSql As String = "DELETE FROM CalendarioAnual " &
            '                     "WHERE clPuestoTrabajo = " & Me.iCodPuestotrabajo & " and clFecha = '" & iFecha.Year & "-" &
            '                     If(iFecha.Month.ToString().Length = 1, "0" & iFecha.Month.ToString(), iFecha.Month.ToString()) &
            '                     "-" & If(iFecha.Day.ToString().Length = 1, "0" & iFecha.Day.ToString(), iFecha.Day.ToString()) & "'"
            Dim sSql As String = "DELETE FROM CalendarioProduccion " &
                                 "WHERE caPuestoTrabajo = " & Me.iCodPuestotrabajo & " and caFecha = '" & iFecha.Year & "-" &
                                 If(iFecha.Month.ToString().Length = 1, "0" & iFecha.Month.ToString(), iFecha.Month.ToString()) &
                                 "-" & If(iFecha.Day.ToString().Length = 1, "0" & iFecha.Day.ToString(), iFecha.Day.ToString()) & "'"

            Datos.CGPL.EjecutarConsulta(sSql)
            EliminarPorFecha = True
        Catch ex As Exception
            EliminarPorFecha = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

#End Region

End Class
