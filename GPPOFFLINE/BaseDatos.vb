Imports System.Data.Common
Imports System.Data
Imports System.Configuration
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Data.OleDb
'prueba - exclusion mutua
Imports System.Threading

''' <summary>
''' Representa la base de datos en el sistema.
''' Ofrece los métodos de acceso a misma.
''' </summary>7
Public Class BaseDatos

    Private Conexion As IDbConnection = Nothing
    Private Commando As IDbCommand = Nothing
    Private Transaccion As IDbTransaction = Nothing
    Private sCadenaConexion As String
    'Private Shared _Factory As DbProviderFactory = Nothing
    Private _Factory As DbProviderFactory = Nothing
    Private tipoBBDD As BBDD
    Private pParametros() As IDbDataParameter

    'prueba - exclusion mutua
    Private Shared mutex As New Mutex
    Public Enum BBDD
        SQL = 1
        ODBC = 2
        OLEDB = 3
        FIREBIRD = 4
    End Enum

    '************************************ FUNCIONES PRIVADAS *********************************************
    Private Sub CrearComando(ByVal Sql As String,
                              ByVal TipoConsulta As String)

        Try
            Select Case tipoBBDD
                Case BBDD.SQL : Me.Commando = New SqlCommand
                Case BBDD.ODBC : Me.Commando = New OdbcCommand
                Case BBDD.OLEDB : Me.Commando = New OleDbCommand
                    'Case BBDD.FIREBIRD : Me.Commando = New FbCommand
            End Select

            Me.Commando = _Factory.CreateCommand()
            Me.Commando.Connection = Me.Conexion
            Me.Commando.CommandTimeout = 99999999
            Select Case TipoConsulta
                Case "P" 'Procedimiento almacenado
                    Me.Commando.CommandType = CommandType.StoredProcedure

                Case "T"
                    Me.Commando.CommandType = CommandType.TableDirect

                Case Else 'Sentencia SQL
                    Me.Commando.CommandType = CommandType.Text
            End Select

            Me.Commando.CommandText = Sql

            If Not Me.Transaccion Is Nothing Then
                Me.Commando.Transaction = Me.Transaccion
            End If



        Catch ex As Exception
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.CrearComando ", ex)
        End Try

    End Sub


    ''' <summary>
    ''' Crea un commando para procedimiento almacenado
    ''' </summary>
    ''' <param name="procedimiento"></param>
    ''' <remarks></remarks>
    Private Sub CrearComandoProcedure(ByVal procedimiento As String)
        Try
            Select Case tipoBBDD
                Case BBDD.SQL : Me.Commando = New SqlCommand
                Case BBDD.ODBC : Me.Commando = New OdbcCommand
                Case BBDD.OLEDB : Me.Commando = New OleDbCommand
                    'Case BBDD.FIREBIRD : Me.Commando = New FbCommand
            End Select

            Me.Commando = _Factory.CreateCommand()
            Me.Commando.Connection = Me.Conexion
            Me.Commando.CommandTimeout = 99999999
            Me.Commando.CommandType = CommandType.StoredProcedure
            Me.Commando.CommandText = procedimiento

            If Not Me.Transaccion Is Nothing Then
                Me.Commando.Transaction = Me.Transaccion
            End If
        Catch ex As Exception
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.CrearComandoProcedure ", ex)
        End Try

    End Sub

    ''' <summary>
    ''' Devuelve un IDataAdapter que segun tipoBBDD puede ser SQlDataAdapter, OdbcDataAdapter o OleDbDataAdapter
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function EjecutarConsultaDataAdapter() As IDataAdapter
        Try
            Select Case tipoBBDD
                Case BBDD.SQL : EjecutarConsultaDataAdapter = New SqlDataAdapter(CType(Me.Commando, SqlCommand))
                Case BBDD.ODBC : EjecutarConsultaDataAdapter = New OdbcDataAdapter(CType(Me.Commando, OdbcCommand))
                Case BBDD.OLEDB : EjecutarConsultaDataAdapter = New OleDbDataAdapter(CType(Me.Commando, OleDbCommand))
                    'Case BBDD.FIREBIRD : EjecutarConsultaDataAdapter = New FbDataAdapter(CType(Me.Commando, FbCommand))
                Case Else : Return Nothing
            End Select
        Catch ex As Exception
            Return Nothing
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.EjecutarConsultaDataAdapter ", ex)
        End Try

    End Function

    ''' <summary>
    ''' Ejecuta el comando creado y retorna un escalar.
    ''' </summary>
    ''' <returns>El escalar que es el resultado del comando.</returns>
    ''' <exception cref="BaseDatosException">Si ocurre un error al ejecutar el comando.</exception>
    Private Function EjecutarEscalar() As Integer
        Dim escalar As Integer = 0
        Try
            escalar = Integer.Parse(Me.Commando.ExecuteScalar().ToString())
        Catch ex As InvalidCastException
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.EjecutarEscalar ", ex)
        End Try
        Return escalar
    End Function

    ''' <summary>
    ''' Ejecuta un commando
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub EjecutarComando()
        Try
            Me.Commando.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.EjecutarComando ", ex)
        End Try

    End Sub


    '************************************ FUNCIONES PUBLICAS *********************************************

    Public Sub New()
        Try

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Crea una nueva conexion a la base de datos mediante una Key del app.config
    ''' </summary>
    ''' <param name="KeyConfig"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal KeyConfig As String)

        Try
            Dim proveedor As String = ConfigurationManager.AppSettings.Get("PROVEEDOR_ADONET")

            Select Case Trim(proveedor)
                Case "System.Data.SqlClient" : tipoBBDD = BBDD.SQL
                Case "System.Data.Odbc" : tipoBBDD = BBDD.ODBC
                Case "System.Data.OleDb" : tipoBBDD = BBDD.OLEDB
                    'Case "FirebirdSql.Data.FirebirdClient" : tipoBBDD = BBDD.FIREBIRD
            End Select

            Me.sCadenaConexion = ConfigurationManager.AppSettings.Get(KeyConfig)
            'BaseDatos._Factory = DbProviderFactories.GetFactory(proveedor)
            Me._Factory = DbProviderFactories.GetFactory(proveedor)
            pParametros = Nothing

        Catch ex As Exception
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.New(sKeyConfig) ", ex)
        End Try
    End Sub

    Public Sub New(ByVal ProveedorBBDD As String, ByVal KeyConfig As String)

        Try
            Dim proveedor As String = ConfigurationManager.AppSettings.Get(ProveedorBBDD)

            Select Case Trim(proveedor)
                Case "System.Data.SqlClient" : tipoBBDD = BBDD.SQL
                Case "System.Data.Odbc" : tipoBBDD = BBDD.ODBC
                Case "System.Data.OleDb" : tipoBBDD = BBDD.OLEDB
                    'Case "FirebirdSql.Data.FirebirdClient" : tipoBBDD = BBDD.FIREBIRD
            End Select

            Me.sCadenaConexion = ConfigurationManager.AppSettings.Get(KeyConfig)
            Me._Factory = DbProviderFactories.GetFactory(proveedor)
            pParametros = Nothing

        Catch ex As Exception
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.New(sKeyConfig) ", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Crea una nueva conexion a la base de datos mediante DSN, Usuario y Contraseña
    ''' </summary>
    ''' <param name="BaseDatos"></param>
    ''' <param name="Usuario"></param>
    ''' <param name="Pass"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal TipoConexion As BaseDatos.BBDD,
                   ByVal BaseDatos As String,
                   ByVal Usuario As String,
                   ByVal Pass As String)
        Try
            Dim proveedor As String = ""
            tipoBBDD = TipoConexion
            Select Case tipoBBDD
                Case BBDD.SQL : proveedor = "System.Data.SqlClient"
                Case BBDD.ODBC : proveedor = "System.Data.Odbc"
                Case BBDD.OLEDB : proveedor = "System.Data.OleDb"
                    'Case BBDD.FIREBIRD : proveedor = "FirebirdSql.Data.FirebirdClient"
            End Select

            If tipoBBDD = BBDD.FIREBIRD Then
                Me.sCadenaConexion = "User=" & Usuario & ";Password=" & Pass & ";Database=" & BaseDatos &
                                     ";Port=3050;Dialect=3;Charset=NONE;Role=;Connection lifetime=15;Pooling=true;" &
                                     "MinPoolSize=0;MaxPoolSize=50;Packet Size=8192;ServerType=0;"
            Else
                Me.sCadenaConexion = "DSN=" & BaseDatos & ";UID=" & Usuario & ";PWD=" & Pass
            End If

            'BaseDatos._Factory = DbProviderFactories.GetFactory(proveedor)
            Me._Factory = DbProviderFactories.GetFactory(proveedor)
            pParametros = Nothing

        Catch ex As Exception
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.New(Parametros) ", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Crea una nueva conexion a la base de datos mediante TipoConexion(SQL,ODBC,OLEDB) y la cadena de conexion
    ''' </summary>
    ''' <param name="tipoConexion"></param>
    ''' <param name="CadenaConexion"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal TipoConexion As BaseDatos.BBDD,
                   ByVal CadenaConexion As String)
        Try
            Dim proveedor As String = ""
            tipoBBDD = TipoConexion
            Select Case tipoBBDD
                Case BBDD.SQL : proveedor = "System.Data.SqlClient"
                Case BBDD.ODBC : proveedor = "System.Data.Odbc"
                Case BBDD.OLEDB : proveedor = "System.Data.OleDb"
                    ' Case BBDD.FIREBIRD : proveedor = "FirebirdSql.Data.FirebirdClient"
            End Select

            Me.sCadenaConexion = CadenaConexion
            Me._Factory = DbProviderFactories.GetFactory(proveedor)

            pParametros = Nothing
        Catch ex As Exception
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.New(tipoConexion,CadenaConexion) ", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Se concecta con la base de datos.
    ''' </summary>
    ''' <exception cref="BaseDatosException">Si existe un error al conectarse.</exception> 
    Public Sub Conectar()
        Try
            If Not Me.Conexion Is Nothing Then
                If Me.Conexion.State.Equals(ConnectionState.Closed) Then
                    Throw New Exception("La conexion ya se encuentra abierta")
                    Exit Sub
                End If
            End If

            If Me.Conexion Is Nothing Then
                Select Case tipoBBDD
                    Case BBDD.SQL : Me.Conexion = New SqlConnection
                    Case BBDD.ODBC : Me.Conexion = New OdbcConnection
                    Case BBDD.OLEDB : Me.Conexion = New OleDbConnection
                        'Case BBDD.FIREBIRD : Me.Conexion = New FbConnection
                End Select

                Me.Conexion = _Factory.CreateConnection()
                Me.Conexion.ConnectionString = Me.sCadenaConexion
            End If
            Me.Conexion.Open()

            If Me.Conexion.State <> ConnectionState.Open Then
                Throw New Exception("ERROR AL CONECTAR CON LA BASE DE DATOS " & Me.Conexion.Database)
            End If

        Catch ex As Exception
            Throw New Exception("ERROR :" & ex.Message & " - Cadena conexion : " & Me.sCadenaConexion & " BASEDATOS.Conectar ", ex)
        End Try

    End Sub

    ''' <summary>
    ''' Desconecta de la base de datos
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Desconectar()

        Try
            If Me.Conexion.State.Equals(ConnectionState.Open) Then
                Me.Conexion.Close()
                Me.Conexion = Nothing
            End If
        Catch ex As DataException
            Me.Conexion = Nothing
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.DESCONECTAR ", ex)
        Catch ex As InvalidOperationException
            Me.Conexion = Nothing
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.DESCONECTAR ", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Devuelve un datatable con los datos retornado de la consulta o procedimiento almacenado.
    ''' </summary>
    ''' <param name="SQL"></param>
    ''' <param name="Tabla"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DameDatosDT(ByVal SQL As String,
                                ByRef Tabla As DataTable) As Boolean
        Try
            mutex.WaitOne()
            Dim arsDatos As IDataAdapter
            Dim datos As New DataSet
            Dim i As Integer = 0
            Select Case tipoBBDD
                Case BBDD.SQL : arsDatos = New SqlDataAdapter
                Case BBDD.ODBC : arsDatos = New OdbcDataAdapter
                Case BBDD.OLEDB : arsDatos = New OleDbDataAdapter
                    'Case BBDD.FIREBIRD : arsDatos = New FbDataAdapter
            End Select

            CrearComando(SQL, "C")

            arsDatos = Me.EjecutarConsultaDataAdapter()
            arsDatos.Fill(datos)

            If datos.Tables.Count > 0 Then
                Tabla = datos.Tables(0)
            End If

            If Tabla.Rows.Count > 0 Then
                DameDatosDT = True
            Else
                DameDatosDT = False
            End If
        Catch ex As Exception
            DameDatosDT = False
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.DameDatosDT ", ex)
        Finally
            mutex.ReleaseMutex()
        End Try
    End Function

    ''' <summary>
    ''' Devuelve un datatable con los datos retornado de la consulta o procedimiento almacenado.
    ''' sTipoConsulta= P->Procedimiento Almacenado, C->Consulta SQL , T->TableDirect
    ''' </summary>
    ''' <param name="NombreProcedure"></param>
    ''' <param name="HayParametros"></param>
    ''' <param name="Tabla"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DameDatosDTProcedimiento(ByVal NombreProcedure As String,
                                             ByVal HayParametros As Boolean,
                                             ByRef Tabla As DataTable) As Boolean

        Dim arsDatos As IDataAdapter
        Dim datos As New DataSet
        Dim i As Integer = 0

        Try

            Select Case tipoBBDD
                Case BBDD.SQL : arsDatos = New SqlDataAdapter
                Case BBDD.ODBC : arsDatos = New OdbcDataAdapter
                Case BBDD.OLEDB : arsDatos = New OleDbDataAdapter
                    ' Case BBDD.FIREBIRD : arsDatos = New FbDataAdapter
            End Select


            CrearComando(NombreProcedure, "P")

            If HayParametros Then
                While i <= UBound(pParametros)
                    Me.Commando.Parameters.Add(pParametros(i))
                    i = i + 1
                End While
            End If

            arsDatos = Me.EjecutarConsultaDataAdapter()
            arsDatos.Fill(datos)
            If datos.Tables.Count > 0 Then
                Tabla = datos.Tables(0)
            End If
            ' Tabla = datos.Tables(0)

            If Tabla.Rows.Count > 0 Then
                DameDatosDTProcedimiento = True
            Else
                DameDatosDTProcedimiento = False
            End If

        Catch ex As Exception
            DameDatosDTProcedimiento = False
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.DameDatosDTProcedimiento ", ex)
        End Try
    End Function

    ''' <summary>
    ''' Devuelve un dataadapter de la consulta y los parametros pasados.
    ''' sTipoConsulta= P->Procedimiento Almacenado, C->Consulta SQL , T->TableDirect
    ''' </summary>
    ''' <param name="sProc_SQL"></param>
    ''' <param name="bHayParametros"></param>
    ''' <param name="sTipoConsulta"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Dame_Datos_DA(ByVal sProc_SQL As String,
                                  ByVal bHayParametros As Boolean,
                                  ByVal sTipoConsulta As String) As IDataAdapter
        Try
            Dim i As Integer = 0

            CrearComando(sProc_SQL, sTipoConsulta)

            If bHayParametros Then
                While i <= UBound(pParametros)
                    Me.Commando.Parameters.Add(pParametros(i))
                    i = i + 1
                End While
            End If

            Return Me.EjecutarConsultaDataAdapter()

        Catch ex As Exception
            Dame_Datos_DA = Nothing
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.DameDatos_DA ", ex)
        End Try
    End Function

    ''' <summary>
    ''' Devuelve los datos generados por un proceso almacenado con parametros en un dataReader.
    ''' sTipoConsulta= P->Procedimiento Almacenado, C->Consulta SQL , T->TableDirect
    ''' </summary>
    ''' <param name="sNombreProc"></param>
    ''' <param name="bHayParametros"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Dame_Datos_DR(ByVal sNombreProc As String,
                                  ByVal bHayParametros As Boolean,
                                  ByVal sTipoConsulta As String) As IDataReader
        Try
            Dim i As Integer = 0
            CrearComando(sNombreProc, sTipoConsulta)
            If bHayParametros Then
                While i <= UBound(pParametros)
                    Me.Commando.Parameters.Add(pParametros(i))
                    i = i + 1
                End While
            End If

            Return Me.Commando.ExecuteReader()

        Catch ex As Exception
            Return Nothing
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.DameDatosDR ", ex)
        End Try
    End Function

    ''' <summary>
    ''' Función que ejecuta una consulta de datos
    ''' </summary>
    ''' <param name="Sql"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EjecutarConsulta(ByVal Sql As String) As Boolean
        Try
            mutex.WaitOne()
            Me.CrearComando(Sql, "S")
            Me.EjecutarComando()
            EjecutarConsulta = True
        Catch ex As Exception
            EjecutarConsulta = False
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.EjecutarConsulta ", ex)
        Finally
            mutex.ReleaseMutex()
        End Try
    End Function

    Public Function EjecutarConsultaEscalar(ByVal Sql As String) As Long
        Try
            mutex.WaitOne()
            Me.CrearComando(Sql, "S")
            EjecutarConsultaEscalar = Me.EjecutarEscalar()

        Catch ex As Exception
            EjecutarConsultaEscalar = -1
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.EjecutarConsultaEscalar ", ex)
        Finally
            mutex.ReleaseMutex()
        End Try
    End Function


    ''' <summary>
    ''' Ejecuta un procedimiento almacenado, con o sin parametros
    ''' </summary>
    ''' <param name="NombreProcedure"></param>
    ''' <param name="HayParametros"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EjecutarPA(ByVal NombreProcedure As String,
                                 ByVal HayParametros As Boolean) As Boolean
        'Esta instancia ejecuta un procedimiento almacenado INSERT, DELETE o UPDATE
        Try
            Dim i As Integer = 0

            CrearComando(NombreProcedure, "P")
            If HayParametros Then
                While i <= UBound(pParametros)
                    Me.Commando.Parameters.Add(pParametros(i))
                    i = i + 1
                End While
            End If

            If Me.Commando.ExecuteNonQuery = 0 Then
                EjecutarPA = False
            Else
                EjecutarPA = True
            End If

            Me.Commando.Parameters.Clear()

        Catch ex As Exception
            EjecutarPA = False
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.EjecutarPA ", ex)
        End Try

    End Function


    ''' <summary>
    ''' Comienza una transacion 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ComenzarTransaccion()
        If Me.Transaccion Is Nothing Then
            Me.Transaccion = Me.Conexion.BeginTransaction()
        End If
    End Sub


    ''' <summary>
    ''' Cancela la transaccion lanzada
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CancelarTransaccion()
        If Not Me.Transaccion Is Nothing Then
            Me.Transaccion.Rollback()
            Me.Transaccion = Nothing
        End If
    End Sub

    ''' <summary>
    ''' Confirma la transaccion lanzada
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ConfirmarTransaccion()
        If Not Me.Transaccion Is Nothing Then
            Me.Transaccion.Commit()
            Me.Transaccion = Nothing
        End If
    End Sub


    ''' <summary>
    ''' Nos dice cual es el estado de la conexion
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EstadoConexion() As ConnectionState
        Try
            EstadoConexion = Conexion.State
        Catch ex As Exception
            EstadoConexion = ConnectionState.Closed
        End Try
    End Function


    ''' <summary>
    ''' Asigna un parametro nulo
    ''' </summary>
    ''' <param name="nombre"></param>
    ''' <remarks></remarks>
    Public Sub AsignarParametroNulo(ByVal nombre As String)
        AsignarParametro(nombre, "", "NULL")
    End Sub

    ''' <summary>
    ''' Asigna un parametro de tipo Cadena
    ''' </summary>
    ''' <param name="nombre"></param>
    ''' <param name="valor"></param>
    ''' <remarks></remarks>
    Public Sub AsignarParametroCadena(ByVal nombre As String, ByVal valor As String)
        AsignarParametro(nombre, "'", valor)
    End Sub


    ''' <summary>
    ''' Asigna un parametro de tipo Entero
    ''' </summary>
    ''' <param name="nombre"></param>
    ''' <param name="valor"></param>
    ''' <remarks></remarks>
    Public Sub AsignarParametroEntero(ByVal nombre As String, ByVal valor As Integer)
        AsignarParametro(nombre, "", valor.ToString())
    End Sub

    ''' <summary>
    ''' Asigna el parametro a la consulta
    ''' </summary>
    ''' <param name="nombre"></param>
    ''' <param name="separador"></param>
    ''' <param name="valor"></param>
    ''' <remarks></remarks>
    Private Sub AsignarParametro(ByVal nombre As String, ByVal separador As String, ByVal valor As String)
        Try
            Dim indice As Integer = Me.Commando.CommandText.IndexOf(nombre)
            Dim prefijo As String = Me.Commando.CommandText.Substring(0, indice)
            Dim sufijo As String = Me.Commando.CommandText.Substring(indice + nombre.Length)
            Me.Commando.CommandText = prefijo + separador + valor + separador + sufijo
        Catch ex As Exception
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.AsignarParametro ", ex)
        End Try

    End Sub


    ''' <summary>
    ''' Asigna un parametro de tipo Fecha
    ''' </summary>
    ''' <param name="nombre"></param>
    ''' <param name="valor"></param>
    ''' <remarks></remarks>
    Public Sub AsignarParametroFecha(ByVal nombre As String, ByVal valor As DateTime)
        AsignarParametro(nombre, "'", valor.ToString())
    End Sub

    ''' <summary>
    ''' Añade un parametro a la consulta o procemiento almacenado. 
    ''' sTipo = L->long, B->boolean, D->double, DC->Decimal, DT->DateTime, S->String, I->Integer, SM->SmallInt
    ''' </summary>
    ''' <param name="iIndice"></param>
    ''' <param name="sNombre"></param>
    ''' <param name="sTipo"></param>
    ''' <param name="sValor"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Añadir_Parametro(ByVal iIndice As Integer,
                                      ByVal sNombre As String,
                                      ByVal sTipo As String,
                                      ByVal sValor As String) As Boolean
        Try

            ReDim Preserve pParametros(iIndice)
            Select Case tipoBBDD
                Case BBDD.SQL
                    pParametros(iIndice) = New SqlParameter
                Case BBDD.ODBC
                    pParametros(iIndice) = New OdbcParameter
                Case BBDD.OLEDB
                    pParametros(iIndice) = New OleDbParameter
                Case BBDD.FIREBIRD
                    ' pParametros(iIndice) = New FbParameter
            End Select

            pParametros(iIndice).ParameterName = sNombre

            Select Case sTipo
                Case "L" 'Long
                    pParametros(iIndice).DbType = DbType.Int64
                Case "D" 'Double
                    pParametros(iIndice).DbType = DbType.Double
                Case "B" 'Bit- Booleano
                    pParametros(iIndice).DbType = DbType.Boolean
                Case "DC" 'Decimal
                    pParametros(iIndice).DbType = DbType.Decimal
                Case "DT" 'Datetime
                    pParametros(iIndice).DbType = DbType.DateTime
                Case "S" 'String
                    pParametros(iIndice).DbType = DbType.String
                Case "I" 'Integer
                    pParametros(iIndice).DbType = DbType.Int32
                Case "SM" 'Smallint
                    pParametros(iIndice).DbType = DbType.Int16
            End Select

            pParametros(iIndice).Value = sValor
            Añadir_Parametro = True
        Catch ex As Exception
            Añadir_Parametro = False
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.Añadir_Parametro ", ex)
        End Try

    End Function
    ''' <summary>
    ''' Añade un parametro Binario a la consulta o procemiento almacenado.  Mapping: .net byte() - sqlDbType VarBynary - DbType Binary
    ''' </summary>
    ''' <param name="iIndice"></param>
    ''' <param name="sNombre"></param>
    ''' <param name="sValor"></param>
    ''' <returns></returns>
    Public Function Añadir_ParametroBinario(ByVal iIndice As Integer,
                                      ByVal sNombre As String,
                                      ByVal sValor() As Byte) As Boolean
        Try

            ReDim Preserve pParametros(iIndice)
            Select Case tipoBBDD
                Case BBDD.SQL
                    pParametros(iIndice) = New SqlParameter
                Case BBDD.ODBC
                    pParametros(iIndice) = New OdbcParameter
                Case BBDD.OLEDB
                    pParametros(iIndice) = New OleDbParameter
                Case BBDD.FIREBIRD
                    ' pParametros(iIndice) = New FbParameter
            End Select

            pParametros(iIndice).ParameterName = sNombre

            pParametros(iIndice).DbType = DbType.Binary

            pParametros(iIndice).Value = sValor

            Añadir_ParametroBinario = True
        Catch ex As Exception
            Añadir_ParametroBinario = False
            Throw New Exception("ERROR :" & ex.Message & " BASEDATOS.Añadir_ParametroBinario ", ex)
        End Try

    End Function


    ''' <summary>
    ''' Cuenta cuantos registros tiene el datareader
    ''' </summary>
    ''' <param name="arsDatos"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RecordCount(ByVal arsDatos As IDataReader) As Long
        Try
            RecordCount = 0
            With arsDatos
                While .Read
                    RecordCount = RecordCount + 1
                End While
            End With
        Catch ex As Exception
            RecordCount = -1
        End Try

    End Function

End Class
