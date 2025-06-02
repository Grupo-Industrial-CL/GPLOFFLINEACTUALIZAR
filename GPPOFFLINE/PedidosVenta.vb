Imports GPLOFFLINEACTUALIZAR.ConstantesGPP
Imports GPLOFFLINEACTUALIZAR.DatosProduccion


Public Class PedidosVenta

#Region "Atributos"
    Private miFabricacion As List(Of Fabricaciones)
    Private sNombresPuestoTrabajo As List(Of PuestosTrabajo)
    Private miMaterial As Material
    Private iStockActual As Integer
    Private miStockCritico As Boolean
    Private sNombrePuestoTrabajo As String
    Private sMaterial As String
    Private sPuestoTrabajo As String
    Private sFechaPlan As Date
#End Region

#Region "contructor"
    Public Sub New()
        iStockActual = -999999999
        sNombrePuestoTrabajo = ""
    End Sub
#End Region



    Public Property Fecha As Date
    Public Property FechaPrevista As Date
    Public Property FechaReal As String
    Public Property CodClienteSolic As Integer
    Public Property CodClienteDest As Integer
    'Public Property Material As String
    Public Property CodigoMaterialFab As String
    Public Property Material As String
        Get
            'If sMaterial <> "" Then

            '    sNombresPuestoTrabajo = DatosProduccion.DameListaPuestosTrabajoMaterial(Material:=sMaterial)
            'End If
            Return sMaterial
        End Get
        Set(value As String)
            sMaterial = value
        End Set
    End Property
    Public Property Grupo As String
    Public Property Kilos As Integer
    Public Property Unidad As String
    Public Property KilosPtes As Integer

    Public Property UnidadesPtes As Integer

    Public Property KgPtes As Double
    Public Property KilosEnt As Integer

    ' Public Property NumEntrega As String
    Public Property LineaPedido As String

    Public Property ClaseEntrega As String
    Public Property TipoPosicion As String
    Public Property TipoEnvio As String
    Public Property Centro As String
    Public Property Almacen As String
    Public Property Pedido As String
    Public Property OrdenTransporte As String
    Public Property EstadoOrdenTpte As String
    Public Property EntregaPendiente As Boolean
    Public Property NombreMaterial As String
    Public Property NombreCliente As String
    Public Property StockActual As Integer

    Public Property NuevoStockActual As Integer

    Public Property NuevoStockAPedidoVenta As Integer
    Public Property StatusGLobal As String
    Public Property StatusEntrega As String

    Public Property NombrePuestoTrabajo As String
        Get
            Return sNombrePuestoTrabajo.Trim()
        End Get
        Set(value As String)
            sNombrePuestoTrabajo = value
        End Set
    End Property
    Public ReadOnly Property StockCritico As Boolean
        Get
            If KilosPtes > StockPedidosVenta Then
                miStockCritico = True
            Else
                miStockCritico = False
            End If

            Return miStockCritico
        End Get
    End Property
    Public ReadOnly Property GrupoCompras As String
        Get
            Return MaterialDetalle.GrupoComprasDetalle.Nombre
        End Get
    End Property
    Public Property FechaPlan As Date
        Get
            'If MaterialDetalle.Codigo = "70907287" Then
            '    Dim h = 9

            'End If
            'Dim FechaPlanificacion As Date = New Date
            'Dim misTurnos As New List(Of Calendario)
            'If NombresPuestoTrabajo.Count = 1 Then
            '    misTurnos = DameTurnosMaquina(FechaPrevista.AddDays(-(MaterialDetalle.DiasFabPropia + MaterialDetalle.DiasPP) * 5), NombresPuestoTrabajo(0).CodigoPuestoTrabajo)
            'Else
            '    Dim Cod = MaterialDetalle.HojasRutaDefecto.PuestosTrabajoHojaRuta(0).CodigoPuestoTrabajo
            '    misTurnos = DameTurnosMaquina(FechaPrevista.AddDays(-(MaterialDetalle.DiasFabPropia + MaterialDetalle.DiasPP) * 5), NombresPuestoTrabajo(0).CodigoPuestoTrabajo)
            'End If

            'Dim turnosARestar = (MaterialDetalle.DiasFabPropia + MaterialDetalle.DiasPP) * 3
            ''Obtenemos en que turno esta la la Fecha Prevista
            'Dim IdTurno = misTurnos.Where(Function(w) FechaPrevista >= w.InicioTurno).ToList().Select(Function(s) s.Id).Max() - turnosARestar
            'FechaPlanificacion = misTurnos.Where(Function(w) w.Id = IdTurno).FirstOrDefault().InicioTurno


            'FechaPlanificacion = FechaPrevista.AddDays(-(MaterialDetalle.DiasFabPropia + MaterialDetalle.DiasPP))


            Return sFechaPlan
        End Get
        Set(value As Date)
            sFechaPlan = value
        End Set
    End Property
    Public ReadOnly Property Formato As String
        Get
            Dim sFormsato As String = ""
            If MaterialDetalle.HojasRutaDefecto IsNot Nothing Then
                sFormsato = MaterialDetalle.HojasRutaDefecto.FormatoDetalle.Nombre
            End If
            Return sFormsato
        End Get
    End Property


    Public ReadOnly Property NumPedido As String
        Get
            Return Pedido.Trim & "-" & LineaPedido.Trim
        End Get
    End Property

    Public ReadOnly Property NombresPuestoTrabajo As List(Of PuestosTrabajo)
        'Get
        '    Return sNombresPuestoTrabajo
        'End Get
        Get
            If Material <> "" Then
                sNombresPuestoTrabajo = DatosProduccion.DameListaPuestosTrabajoMaterial(Material:=Material)
            Else
                sNombresPuestoTrabajo = New List(Of PuestosTrabajo)
            End If
            Return sNombresPuestoTrabajo
        End Get

    End Property

    Public Property PuestoTrabajo As String
        Get
            Return sPuestoTrabajo
        End Get
        Set(value As String)
            sPuestoTrabajo = value
        End Set
    End Property
    Public ReadOnly Property FabricacionExistente As List(Of Fabricaciones)
        Get
            If Pedido <> "" Then
                miFabricacion = DatosProduccion.DameFabricacionExistente(NumPedSap:=Pedido,
                                                                                     PosPedSap:=LineaPedido,
                                                                                     Estado_Fabricacion:=EstadoFabricacion.Ninguna,
                                                                                     CodPuestoTrabajo:=0,
                                                                                     IdEnvio:=0)
            Else
                miFabricacion = New List(Of Fabricaciones)
            End If

            Return miFabricacion
        End Get

    End Property
    Public ReadOnly Property ValidacionPullsystem As Boolean
        Get
            Return DatosProduccion.ValidacionPullsystem(codMaterial:=Material.Trim(), FechaPrevistaFin:=FechaPrevista)
        End Get
    End Property

    Public ReadOnly Property StockPedidosVenta As Integer
        Get
            If iStockActual = -999999999 Then
                iStockActual = DatosSAPConexion.DatosSAP.DameStockMARD(CENTRO:="12",
                                                                       ALMACEN:="",
                                                                       MATERIAL:=Material,
                                                                       GTIN:=0,
                                                                       Tipo_material:="")
            End If

            Return Me.iStockActual

        End Get
    End Property

    Public ReadOnly Property MaterialDetalle As Material
        Get
            If Not String.IsNullOrEmpty(Material) Then
                miMaterial = New Material(Material.Trim())
            End If
            Return miMaterial
        End Get
    End Property
    Public ReadOnly Property PuestosTrabajo_Maquina As List(Of PuestosTrabajo)
        Get

            Dim ptMaquinas As New List(Of PuestosTrabajo)

            If MaterialDetalle.HojasRutaDefecto Is Nothing Then
                Return ptMaquinas
            End If
            If MaterialDetalle.HojasRutaDefecto.PuestosTrabajoHojaRuta Is Nothing Then
                Return ptMaquinas
            End If
            ptMaquinas = (From dato In MaterialDetalle.HojasRutaDefecto.PuestosTrabajoHojaRuta Where dato.Tipo = TipoPuestoTrabajo.Maquina Select dato).ToList

            Return ptMaquinas

        End Get
    End Property

    Public ReadOnly Property Minutos_Tiempo_Fabricacion(bIncluirPreparacion As Boolean) As Integer
        Get
            Dim iMinutos As Integer = 0
            Dim iMinutosOperacion As Decimal = 0

            If miMaterial Is Nothing And Not String.IsNullOrEmpty(Material) Then
                miMaterial = New Material(Material.Trim())
            End If


            If miMaterial.HojasRutaDefecto Is Nothing Then
                Return 0
            End If

            For Each miOper In miMaterial.HojasRutaDefecto.OperacHojaRutaLista
                If miOper.CantidadBase <> 0 Then
                    If bIncluirPreparacion = False Then
                        iMinutosOperacion = miOper.MinutosMaquina +
                                            miOper.MinutosLimpieza
                    Else
                        iMinutosOperacion = miOper.MinutosMaquina +
                                            miOper.MinutosPreparacion +
                                            miOper.MinutosLimpieza
                    End If

                    iMinutos += CInt(KilosEnt * iMinutosOperacion / miOper.CantidadBase)
                End If
            Next

            Return iMinutos
        End Get
    End Property


    Public ReadOnly Property ValorRdoTanque As Integer
        Get
            ' esto es temporal se tiene que rediseñar con la alta de Tanques y ahi debe tener como
            ' propiedad la cacidad del mismo y al centro producutivo al que pertenece
            Dim TanqueRedondeo As Integer = 0
            'If sHojaRuraDefault = "" Then
            '    Return TanqueRedondeo
            'End If

            If KgPtes <= 0 Then
                Return TanqueRedondeo
            End If

            Dim iCodCentroProd = PuestosTrabajo_Maquina(0).Centro_Prod.Codigo

            Select Case iCodCentroProd
                Case 1 ' Cosmetica
                    'Cosmética: 200 Kg, 500 Kg, 3.000 Kg y 6.000 Kg.
                    Select Case KgPtes
                        Case 0
                            TanqueRedondeo = 0
                        Case 1 To 200
                            TanqueRedondeo = TanqueCosmetica.TC200
                        Case 201 To 500
                            TanqueRedondeo = TanqueCosmetica.TC500
                        Case 501 To 3000
                            TanqueRedondeo = TanqueCosmetica.TC3000
                        Case 3001 To 6000
                            TanqueRedondeo = TanqueCosmetica.TC6000
                        Case Else
                            TanqueRedondeo = TanqueCosmetica.TC6000
                    End Select

                Case 2 ' Fragancias
                    ' No hay tanques para fragancias
                    TanqueRedondeo = 0
                Case 3 ' Higiene 
                    'Higiene: 3.000 Kg, 9.000 Kg y 20.000 Kg
                    '-El redondeo que se hace de los kilos en los reactores de higiene se realiza actualmente a 3000kgs, 9000kgs y 20000kgs, y 
                    'lo debería realizar a 700kgs, 3000kgs, 9000kgs y 19000kgs.
                    '2023-09-19 
                    'Se regreso como estaba
                    Select Case KgPtes
                        Case 0
                            TanqueRedondeo = 0
                        Case 1 To 700
                            TanqueRedondeo = TanqueHigiene.TH3000
                        Case 701 To 3000
                            TanqueRedondeo = TanqueHigiene.TH3000
                        Case 3001 To 9000
                            TanqueRedondeo = TanqueHigiene.TH9000
                        Case 9001 To 20000
                            TanqueRedondeo = TanqueHigiene.TH20000
                        Case Else
                            TanqueRedondeo = TanqueHigiene.TH20000
                    End Select
            End Select

            Return TanqueRedondeo
        End Get

    End Property

End Class
