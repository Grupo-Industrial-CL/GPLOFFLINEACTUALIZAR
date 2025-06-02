Public Class ConstantesGPP

    ' esto es temporal se tiene que rediseñar con la alta de Tanques y ahi debe tener como
    ' propiedad la cacidad del mismo y al centro producutivo al que pertenece
    'Higiene: 3.000 Kg, 9.000 Kg y 20.000 Kg

    'Cosmética: 200 Kg, 500 Kg, 3.000 Kg y 6.000 Kg.

    Public Enum TanqueCosmetica As Integer
        'Cosmética: 200 Kg, 500 Kg, 3.000 Kg y 6.000 Kg.
        TC200 = 200
        TC500 = 500
        TC3000 = 3000
        TC6000 = 6000
    End Enum

    Public Enum TanqueHigiene As Integer
        '-El redondeo que se hace de los kilos en los reactores de higiene se realiza actualmente a 3000kgs, 9000kgs y 20000kgs, y 
        'lo debería realizar a 700kgs, 3000kgs, 9000kgs y 19000kgs.
        '2023-09-19 
        'Higiene: 3.000 Kg, 9.000 Kg y 20.000 Kg
        TH700 = 700
        TH3000 = 3000
        TH9000 = 9000
        '11-04-2023 Cambiar redondeo de 20000 a 19000 a peticion de Juan Ramon
        TH20000 = 19000
    End Enum
    ' esto es temporal se tiene que rediseñar con la alta de Tanques y ahi debe tener como
    ' propiedad la cacidad del mismo y al centro producutivo al que pertenece


    Public Const AlmacenPerseidaBelleza = "100"
    Public Const AlmanceMMdefault = "1202"
    Public Const AlmacenMMdefaultMP = "1203"
    Public Const AlmacenMMgeneral = "1201"

    Public Enum EstadosDisponibilidad As Integer
        DisponibleConStockActual = 5
        DisponibleEnFechaPrev = 10
        DisponibleFueraFechaPrev = 20
        EnRiesgoFaltaStock = 30
        DependePedidoCompra = 40
        EnRiesgoNoExistenPedidosCompra = 45
        EnRiesgoPedidoCompraInsuficiente = 50
    End Enum

    'Public Const MesesAntiguedad_SalidaPaletFIFO = 5
    Public Enum Tipo_Almacen As Integer
        Producto_Terminado = 400
        PT_Almendralejo = 401
        PT_Zafra = 402
        PT_Friovaldi = 403
        PT_Farmacia = 404
        PT_Devoluciones = 405
        CL_Corte = 407
    End Enum


    Public Structure Tipo_Almacen_MM
        Const General = "1201"
        Const Fabricacion = "1202"
        Const CL_Corte = "1207"
        Const Almendralejo = "12AL"
        Const Almacen_Beauty = "12BE"
        Const Canal_Farmacia = "12CF"
        Const Devolucion = "12DE"
        Const Friovaldi = "12FV"
        Const Zafra = "12ZA"
    End Structure

    Public Structure OperacionFormulario
        Public Const Nuevo As Integer = 1
        Public Const Editar As Integer = 2
        Public Const Lectura As Integer = 0
        '1=nuevo 2=Editar 0=Lectura
    End Structure

    Public Structure TiposTrasporte
        Public Const A = "A"
        Public Const M As String = "M"
        Public Const T As String = "T"
    End Structure
    Public Structure TiposTrasporte_Descripcion
        Public Const A = "Aereo"
        Public Const M As String = "Marítimo"
        Public Const T As String = "Terrestre"
    End Structure


    Public Enum PerfilDeCarga
        PorUbicacion = 0
        PorAntiguedad = 1
    End Enum
    Public Structure PerfilDeCarga_Descripcion
        Const PorUbicacion = "Por ubicacion Optima"
        Const PorAntiguedad = "Por Antiguedad de Lote"
    End Structure
    Public Structure Tipo_Almacen_Descripcion
        Const Producto_Terminado As String = "Producto Terminado"
        Const PT_Almendralejo As String = "PT Almendralejo"
        Const PT_Zafra As String = "PT Zafra"
        Const PT_Friovaldi As String = "PT Friovaldi"
        Const PT_Farmacia As String = "PT Farmancia"
        Const PT_Devoluciones As String = "PT Devoluciones"
        Const CL_Corte As String = "PT CL"
    End Structure

    Public Structure Estatus_Pedido_Venta
        Public Const NoRelevante = ""
        Public Const NoTratado As String = "A"
        Public Const TratadoParcialmente As String = "B"
        Public Const Concluido As String = "C"
    End Structure

    Public Structure Estatus_Pedido_VentaDescripcion
        Public Const NoRelevante = "No Relevante"
        Public Const NoTratado As String = "No Tratado"
        Public Const TratadoParcialmente As String = "Tratado Parcialmente"
        Public Const Concluido As String = "Concluido"
    End Structure

    Public Structure SectorProveedor
        Const TODOS As String = ""
        Const OC As String = "OC"
        Const OP As String = "OP"
        Const OT As String = "OT"
        Const ZO As String = "ZO"
    End Structure



    Public Structure ConfiguracionSAP
        Public Const Produccion As String = "PRODUCCION"
        Public Const Desarrollo As String = "DESARROLLO"
        Public Const ReintentosConexionSap As Integer = 120
    End Structure

    Public Structure TipoPuestoTrabajo
        Public Const Maquina As String = "0001"
        Public Const Persona As String = "0003"
        Public Const Externo As String = "10000000"
        Public Const Otros As String = ""
        Public Const CooperativasExterna As String = "EXTERNO"
        Public Const esCentroExterno As Boolean = True
    End Structure

    ' parametros para la creacion de entregas WM
    Public Const PuestoExpedicion As String = "1201"
    Public Const ClaseEntraga As String = "ZLWM"
    Public Const OrganizacionVentas As String = "12"
    Public Const CanalDistribucion As String = "12"
    Public Const Divicion As String = "12"
    Public Const Unidades As String = "ST"



    Public Const ReintentosConexionSap As Integer = 60

    Public Const ClaveControl_PtoTrabajo_Externo As String = "ZPE4"

    'Public Const CentroSAP As String = "0300"

    Public Const UbicProvisional As String = "999"


    Public Const UbicacionPlaya As String = "999999"
    Public Const UbicPrep1 As String = "PREPCAR1"
    Public Const UbicCamion As String = "CAMION"

    Public Const IntervaloGraficos As Integer = 2000

    Public Const MinutosRefrescoGrupoMaquinas As Integer = 1
    Public Const SegundosRefrescoGrupoMaquinas As Integer = 10

    Public Const MandanteSAP As String = "200"
    Public Const TipoAlmacenCalle As String = "C"
    'Public Const CentroAlmacen As String = "0300"
    'Public Const CentroAlmacenesExt As String = "0310"
    'Public Const CentroConsignas As String = "0320"
    Public Const AlmacenBigBag As String = "201"

    Public Const TurnoManana As String = "M"
    Public Const TurnoTarde As String = "T"
    Public Const TurnoNoche As String = "N"

    Public Const MascaraCodigosDeBarras As String = "00000000000000000000"
    Public Const MascaraCodigoProducto As String = "000000000000000000"
    Public Const MascaraUbicaciones As String = "000"
    Public Const MascaraLote As String = "0000000000"
    Public Const MascaraPedido As String = "0000000000"
    Public Const MascaraEntrega As String = "0000000000"
    Public Const MascaraPosicion As String = "000000"
    Public Const MascaraCliente As String = "0000000000"
    Public Const EstadoTpteRealizado As String = "6"

    Public Const FechaGlobal As Date = #1/1/1900#

    Public Structure TipoLogDescripcion
        Const Alta = "Alta de"
        Const Modificar = "Modificar"
        Const Eliminar = "Eliminar"
        Const Contabilizar = "Contabilizar"
    End Structure

    Public Structure ListaMeses
        Const Enero = 1
        Const Febrero = 2
        Const Marzo = 3
        Const Abril = 4
        Const Mayo = 5
        Const Junio = 6
        Const Julio = 7
        Const Agosto = 8
        Const Septiembre = 9
        Const Octubre = 10
        Const Noviembre = 11
        Const Diciembre = 12
    End Structure

    Public Structure StructureLote
        Public Lote As Int64
        Public MascaraLote As String
        Public Calidad As String
        Public Produccion As String
        Public Instalacion As String
        Public NombreInstalacion As String
        Public Motivo As String
        Public Causa As String
        Public Creado As Boolean
    End Structure

    Public Enum TipoAgenda As Integer
        Ninguno = 0
        Personal = 1
        Grupo = 2
        SinAsignar = 3
    End Enum

    Public Enum OpcionesMenu_MDI As Integer
        MenuDatos = 1
        MenuOficina = 2
        MenuAlmacen = 3
    End Enum

    Public Structure TipoMensajeSAP
        Const Success As String = "S"
        Const Error_ As String = "E"
        Const Warning As String = "W"
        Const Info As String = "I"
        Const Abort As String = "A"
    End Structure

    Public Structure ClaseOrdenEnvFab
        Const Envasado As String = "OENV"
        Const Fabricacion As String = "OFAB"
        Const CopExtern As String = "OESP"
    End Structure

    Public Enum OperacionesMDI As Integer
        GestiónUsuarios = 1
        GestionPerfiles = 2
        GestionOperaciones = 3
        Importar = 4
        ForeCastVentas = 5
        ConsultaMateriales = 6
        EntradaPullSystem = 7
        DiasLaborales = 8
        ControlDeProduccion = 9
        Calendario = 10
        Forecast2 = 11
        PedidosVenta = 12
        Recursos = 13
        ConsultaStock = 14
        AnalisisCapacidad = 15
        GestionCarteraPedidos = 16
        GrupoMaquinas = 18
        CambiarFechaPedidoVenta = 19
        CentrosProd = 25
        ConsultaParadas = 26
        Estadisticas = 29
        EnvioProveedores = 30
        ImputarCantPentPedidosVenta = 31
        Trazabilidad = 32
        ImputarRecibidosEnviosProveedor = 33
        EnviosExternosProveedor = 34
        Ubicar = 35
        CargasPendientes = 36
        NotificarSAP = 37
        MapaNaves = 38
        GestionCaducidad = 39
        Buscar = 40
        GestionCamiones = 41
        ParteDeCarga = 42
        AsignarMuelle = 43
        CargarCamion = 44
        ModificarHoraCarga = 45
        DatosEntregaSAP = 46
        NotificarOrdenEnvasado = 47
        Sociedades = 48
        Clientes = 49
        Picking = 50
        MoverPaletsSeleccion = 51

        EnviosPendientesCE = 55
        AsignarMuelleCE = 56
        CargarCamionCE = 57
        AnalisisDisponibilidad = 58
        DevolucionSobrante = 59
        CalendarioBeta = 60
        GrupoCompra = 61
        AsignacionRecursos = 62
        FormatoPedidos = 63
    End Enum

    Public Enum RatiosGenerales
        Consumo_Materia_Prima = 1
        Consumo_Producto_Terminado_Reciclado = 2
        Consumo_Materia_Prima_Reciclado = 3
        Produccion_Semielaborado = 4
        Envasado = 5
        Movimientos_Recalificacion = 6
        Movimientos_InterModal = 7
        Expediciones_Cisterna = 8
        Expediciones_BigBag = 9
        Ventas = 10
        Devoluciones = 11
    End Enum

    Public Enum RatiosSecundarios
        Consumo_PTA = 12
        Consumo_MEG = 13
        Consumo_DEG = 14
        Consumo_IPA = 15
        Produccion_Amorphous = 16
        Produccion_PET = 17
        Ventas_Planta = 18
        Ventas_Planta_BigBag = 19
        Ventas_Planta_Cisterna = 20
        Devoluciones_Planta = 21
        Ventas_Slogan = 22
        Devoluciones_Slogan = 23
    End Enum

    Public Enum TipoImportacion
        Clientes = 0
        Materiales = 1
        ListaMateriales = 2
        CabeceraListaMaterial = 3
        ListaMaterialxMaterial = 4
        HojaDeRuta = 5
        CabeceraHojaDeRuta = 6
        HojaDeRutaxMaterial = 7
        OperacionXHojaDeRuta = 8
        PuestoDeTrabajo = 9
        VersionFabricacionxMaterial = 10
        Formatos = 11
        GruposCompra = 12
        Proveedores = 13
    End Enum

    Public Enum TipoPrecio
        Normal = 0
        Merchant = 1
    End Enum

    Public Enum TipoMovimiento
        Mov101 = 101
        Mov102 = 102
        Mov131 = 131
        Mov132 = 132
        Mov309 = 309
        Mov261 = 261
        Mov262 = 262
        Mov601 = 601
        Mov602 = 602
        Mov641 = 641
        Mov642 = 642
        Mov653 = 653
        Mov654 = 654
        Mov657 = 657
        Mov658 = 658
    End Enum

    Public Const AlmacenesVirtuales = "VS"
    Public Const AlmacenesRecovery = "RW"
    Public Const AlmacenesTransito = "205"
    Public Const AlmacenesMuestra = "295"
    Public Const AlmacenesDevoluciones = "298"
    Public Const LoteOFFSpec09 = "09"
    Public Const LoteOFFSpec01 = "01"
    Public Const LoteProductoConforme = "00"

    Public Const SegundosHoras = 3600
    Public Const segundosDia = 86400
    Public Const segundosMinuto = 60

    'Public Structure Centros
    '    'Public Const Principal As String = "12"
    '    'Public Const Consigna As String = "0320"
    '    'Public Const AlmExterno As String = "0310"
    'End Structure

    'Public Enum TipoSector
    '    Sector22 = 22
    '    Sector23 = 23
    'End Enum

    'Public Structure TipoGrupoRatio
    '    Public Const Grupo05 = "05"
    'End Structure

    'Public Structure IndicadorMovimientoRatio
    '    Public Const B As String = "B"
    'End Structure

    'Public Structure Almacenes
    '    Public Const RW01 As String = "RW01"
    '    Public Const SinAlmacen As String = ""
    'End Structure

    'Public Structure Area_Produccion
    '    Public Const ENVASADO = "12_ENVAS"
    '    Public Const SEMIELABORADOS = "12_SEMIEL"
    'End Structure



    'Public Structure TiposConsumption
    '    Public Const PTA As String = "PTA"
    '    Public Const DEG As String = "DEG"
    '    Public Const MEG As String = "MEG"
    '    Public Const IPA As String = "IPA"
    'End Structure

    Public Enum DecimalesGrid
        Moneda = 99
        Numero = 98
    End Enum

    Public Structure DatosVarios
        Public Const CodigoEspaña As Integer = 108
        Public Const ReintentosConexionSap As Integer = 120
    End Structure

    Public Enum TipoEmbalaje As Integer
        Desconocido = 0
        BigBag = 1
        Cisterna = 2
    End Enum

    Public Structure TipoEmbalajeRatio
        Public Const BigBag As String = "021"
        Public Const Bulk As String = "027"
    End Structure

    Public Structure TipoUnidadBase
        Public Const KG As String = "KG"
        Public Const valorKG As Integer = 1
        Public Const G As String = "G"
        Public Const valorG As Integer = 1000
        Public Const MG As String = "MG"
        Public Const valorMG As Integer = 1000000
    End Structure

    Public Enum TipoGrupo As Short
        Commodity = 0
        PolimeroTecnico = 1
        Reciclado = 2
    End Enum

    Public Enum NumeroSemana As Integer
        Todas = 0
        Primera = 1
        Segunda = 2
        Tercera = 3
        Cuarta = 4
        Quinta = 5
    End Enum

    Public Enum EstadoUbicacion
        ABIERTA
        BLOQUEADA_ENTRADA
        BLOQUEADA_SALIDA
        BLOQUEADA
    End Enum

    Public Structure EstadoUbicacionDescripcion
        Const ABIERTA As String = "ABIERTA"
        Const BLOQUEADA_ENTRADA As String = "BLOQUEADA ENTRADA"
        Const BLOQUEADA_SALIDA As String = "BLOQUEADA SALIDA"
        Const BLOQUEADA As String = "BLOQUEADA"
    End Structure

    Public Structure TipoTamanioLote
        Public Const LoteExactoPerseida As String = "Y0"
        Public Const LoteSemanalPerseida As String = "Y1"
        Public Const LoteDosSemanasPerseida As String = "Y2"
        Public Const LoteCuatroSemanasPerseida As String = "Y4"
        Public Const Otros As String = ""
    End Structure

    Public Structure TipoTamanioLoteDescripcion
        Public Const LoteExactoPerseida As String = "Lote Exacto - Perseida"
        Public Const LoteSemanalPerseida As String = "Lote Semanal - Perseida"
        Public Const LoteDosSemanasPerseida As String = "Lote dos semanas - Perseida"
        Public Const LoteCuatroSemanasPerseida As String = "Lote cuatro semanas - Perseida"
        Public Const Otros As String = "Otros"
    End Structure

    Public Enum EstadoOrdenTpte As Short
        'EnPlanifNec = 0
        Ninguno = 0
        PlanifNecFin = 1
        Registro = 2
        InicioCarga = 3
        FinCarga = 4
        DespachoExpedicion = 5
        InicioTpte = 6
        FinTpte = 7
        'Ninguno = -1
    End Enum

    Public Structure EstadoOrdenTpteDescripcion
        'Public Const EnPlanifNec As String = "Planif. Nec."
        Public Const PlanifNecFin = "Fin Planif. Nec."
        Public Const Registro As String = "Registro"
        Public Const InicioCarga As String = "Inicio Carga"
        Public Const FinCarga As String = "Fin Carga"
        Public Const DespachoExpedicion As String = "Despacho Exp"
        Public Const InicioTpte As String = "Inicio Tpte"
        Public Const FinTpte As String = "Fin Tpte"
        Public Const Ninguno As String = "Ninguno"
    End Structure

    Public Enum TipoFabricacion As Short
        Produccion = 0
        Envasado = 1
        Ninguno = -1
    End Enum
    'Public Structure OperacionFormulario
    '    Public Const Nuevo As Integer = 1
    '    Public Const Editar As Integer = 2
    '    Public Const Lectura As Integer = 0
    '    '1=nuevo 2=Editar 0=Lectura
    'End Structure
    Public Structure TipoFabricacionDescripcion
        Public Const Fabricacion As String = "FABRICACIÓN"
        Public Const Envasado As String = "ENVASADO"
    End Structure

    Public Enum EstadoFabricacion As Short
        PteFabricar = 0
        EnMarcha = 1
        Finalizada = 2
        PlanFuturo = 3
        Ninguna = -1
        Anulada = 99
    End Enum

    Public Structure EstadoFabricacionDescripcion
        Public Const PteFabricar As String = "PTE. FABRICAR"
        Public Const EnMarcha As String = "EN MARCHA"
        Public Const Finalizada As String = "FINALIZADA"
        Public Const Anulada As String = "ANULADA"
        Public Const Ninguna As String = "NINGUNO"
    End Structure

    Public Enum EstadoEntregaPrevista As Short
        Activa = 1
        Inactiva = 0
        Desconocido = -1
    End Enum

    Public Structure EstadoEntregaPrevistaDescripcion
        Public Const Activa As String = "Activa"
        Public Const Inactiva As String = "Inactiva"
        Public Const Desconocido As String = "Desconocido"
    End Structure

    Public Structure EstadosEnvioProveedor
        Public Const PteOrdeSC As Integer = 0
        Public Const PteEnvio As Integer = 1
        Public Const Enviado As Integer = 2
        Public Const Recepcionado As Integer = 3
        Public Const EnCarga As Integer = 4
        Public Const EnvioParcial As Integer = 5
        Public Const PteCarretilla As Integer = 6
        Public Const Anulado As Integer = 99
    End Structure

    Public Structure EstadosEnvioProveedorDescripcion
        Public Const PteOrdeSC As String = "Pte. Orden Sub."
        Public Const PteEnvio As String = "Pte. Envio."
        Public Const Enviado As String = "Enviado."
        Public Const Recepcionado As String = "Recepcionado"
        Public Const EnCarga As String = "En carga"
        Public Const EnvioParcial As String = "Envio Parcial."
        Public Const PteCarretilla As String = "Pte. Carretilla"
        Public Const Anulado As String = "Anulado"
    End Structure

    Public Enum EstadoEnvasadoBB As Short
        PteLiberar = 0
        Liberada = 1
        EnCurso = 2
        Finalizada = 3
        Anulada = 4
        Ninguno = 99
    End Enum

    Public Structure EstadoEnvasadoBBDescripcion
        Public Const PteLiberar As String = "PTE. LIBERAR"
        Public Const Liberada As String = "LIBERADA"
        Public Const EnCurso As String = "EN CURSO"
        Public Const Finalizada As String = "FINALIZADA"
        Public Const Anulada As String = "ANULADA"
        Public Const Ninguno As String = "NINGUNO"
    End Structure

    Public Enum OperacionBigBagLeido As Short
        Ubicar = 1
        Cargar_Camion = 2
        DesCargar_Camion = 3
        Cambio_Ubicacion = 4
        Otros = 5
    End Enum

    Public Structure OperacionBigBagLeidoBDescripcion
        Public Const Ubicar As String = "UBICAR"
        Public Const Cargar_Camion As String = "CARGAR EN CAMIÓN"
        Public Const DesCargar_Camion As String = "DESCARGAR DE CAMIÓN"
        Public Const Cambio_Ubicacion As String = "CAMBIAR DE UBICACIÓN"
        Public Const Otros As String = "OTROS"
    End Structure

    Public Enum EstadoCamion
        Todos = 0
        SinCargar = 1
        EnCarga = 2
        Cargado = 3
        Albaranado = 4
        Anulado = -1
    End Enum

    Public Const claveMaterialInactivo = "PO"

    Public Structure EstadoCamionDescripcion
        Const Todos As String = "Todos"
        Const SinCargar As String = "PTE CARGAR"
        Const EnCarga As String = "EN CARGA"
        Const Cargado As String = "CARGADO"
        Const Albaranado As String = "ALBARANADO"
        Const Anulado = "ANULADO"
    End Structure

    Public Structure TipoEnvio_Entregas
        Const CAMION As String = "01"
        Const FERROCARRIL As String = "03"
        Const CORREO As String = "02"
        Const CONSIGNA As String = "CO"
        Const BARCO As String = "04"
        Const AVION As String = "05"
        Const NORELEVANTE As String = "99"
    End Structure

    Public Structure TipoEnvio_EntregasDescripcion
        Const CAMION As String = "CAMION"
        Const FERROCARRIL As String = "FERROCARRIL"
        Const CORREO As String = "CORREO"
        Const CONSIGNA As String = "CONSIGNA"
        Const BARCO As String = "BARCO"
        Const AVION As String = "AVION"
        Const NORELEVANTE As String = "NO RELEVANTE"
    End Structure

    Public Structure ClaseEntrega_SAP
        Const Normal As String = "LF"
        Const Traslado As String = "NL"
        Const Otros As String = ""
    End Structure

    Public Structure ClaseEntrega_SAP_Descripcion
        Const Normal As String = "Normal"
        Const Traslado As String = "Traslado"
        Const Otros As String = "Otros"
    End Structure

    Public Structure Centros_SAP
        Const Plastiverd As String = "0300"
        Const Externo As String = "0310"
        Const Consignas As String = "0320"
        Const Otros As String = ""
    End Structure

    Public Structure Centros_SAP_Descripcion
        Const Plastiverd As String = "Plastiverd"
        Const Externo As String = "Alm. Externos"
        Const Consignas As String = "Consignas"
        Const Otros As String = "Otros"
    End Structure

    Public Structure EstadoBigBags
        Const Ninguno As String = "-999"
        Const SinLeer As String = "0"
        Const Ubicado As String = ""
        Const Cargado As String = "2"
        Const EnSalidaStock As String = "3"
        Const Perdido As String = "4"
    End Structure

    Public Structure EstadoBigBagsDescripcion
        Const Ninguno As String = ""
        Const SinLeer As String = "SIN LEER"
        Const Ubicado As String = "UBICADO"
        Const Cargado As String = "CARGADO"
        Const EnSalidaStock As String = "EN CAMION"
        Const Perdido As String = "PERDIDO"
    End Structure

    Public Const CodigoMatDesconocido As Integer = -1

    Public Enum EstadoParada As Short
        Inactiva = 0
        Activa = 1
        PteIndicarMotivo = 2
    End Enum

    Public Structure EstadoParadaDescripcion
        Public Const Inactiva As String = "INACTIVA"
        Public Const Activa As String = "ACTIVA"
        Public Const PteIndicarMotivo = "PTE. INDICAR MOTIVO"
    End Structure

    Public Structure TipoMaterial
        Public Const NINGUNO As String = ""
        Public Const ProdTerminado As String = "1206"
        Public Const MatPrima As String = "1201"
        Public Const Packaging As String = "1202"
        Public Const Semielaborado As String = "1215"
        Public Const Fabricaciones As String = "1205"
        Public Const MatCliente As String = "1203"
        Public Const Repuestos As String = "1209"
    End Structure

    Public Structure TipoMaterialDescripcion
        Public Const NINGUNO As String = ""
        Public Const ProdTerminado As String = "Producto Terminado"
        Public Const MatPrima As String = "Materia Prima"
        Public Const Packaging As String = "Packaging"
        Public Const Semielaborado As String = "Semielaborado"
        Public Const Fabricaciones As String = "Fabricaciones"
        Public Const MatCliente As String = "Mat.Cliente"
        Public Const Repuestos As String = "Repuestos"
    End Structure

    Public Structure ListaMesesDescripcion
        Const Enero = "ENERO"
        Const Febrero = "FEBRERO"
        Const Marzo = "MARZO"
        Const Abril = "ABRIL"
        Const Mayo = "MAYO"
        Const Junio = "JUNIO"
        Const Julio = "JULIO"
        Const Agosto = "AGOSTO"
        Const Septiembre = "SEPTIMBRE"
        Const Octubre = "OCTUBRE"
        Const Noviembre = "NOVIEMBRE"
        Const Diciembre = "DICIEMBRE"

    End Structure

    Public Structure LiteralesSAP
        Public Const Unidad As String = "UN"
    End Structure

    Public Shared Function enumToString(ByVal valor As Object, ByVal tipo As Type) As String
        Try
            Select Case tipo
                Case GetType(Estatus_Pedido_Venta)
                    Dim vTipo As String = CType(valor, String)
                    Select Case vTipo
                        Case Estatus_Pedido_Venta.NoRelevante
                            Return Estatus_Pedido_VentaDescripcion.NoRelevante
                        Case Estatus_Pedido_Venta.NoTratado
                            Return Estatus_Pedido_VentaDescripcion.NoTratado
                        Case Estatus_Pedido_Venta.TratadoParcialmente
                            Return Estatus_Pedido_VentaDescripcion.TratadoParcialmente
                        Case Estatus_Pedido_Venta.Concluido
                            Return Estatus_Pedido_VentaDescripcion.Concluido
                        Case Else
                            Return ""
                    End Select


                Case GetType(PerfilDeCarga)
                    Dim vTipo As Integer = CType(valor, Integer)
                    Select Case vTipo
                        Case PerfilDeCarga.PorAntiguedad
                            Return PerfilDeCarga_Descripcion.PorAntiguedad
                        Case PerfilDeCarga.PorUbicacion
                            Return PerfilDeCarga_Descripcion.PorUbicacion
                    End Select

                Case GetType(TipoTamanioLote)
                    Dim vTipo As String = CType(valor, String)
                    Select Case vTipo
                        Case TipoTamanioLote.LoteExactoPerseida
                            Return TipoTamanioLoteDescripcion.LoteExactoPerseida
                        Case TipoTamanioLote.LoteSemanalPerseida
                            Return TipoTamanioLoteDescripcion.LoteSemanalPerseida
                        Case TipoTamanioLote.LoteDosSemanasPerseida
                            Return TipoTamanioLoteDescripcion.LoteDosSemanasPerseida
                        Case TipoTamanioLote.LoteCuatroSemanasPerseida
                            Return TipoTamanioLoteDescripcion.LoteCuatroSemanasPerseida
                        Case TipoTamanioLote.Otros
                            Return TipoTamanioLoteDescripcion.Otros
                        Case Else
                            Return ""
                    End Select
                Case GetType(EstadosEnvioProveedor)
                    Dim vTipo As String = CType(valor, String)
                    Select Case vTipo
                        Case EstadosEnvioProveedor.PteOrdeSC.ToString
                            Return EstadosEnvioProveedorDescripcion.PteOrdeSC
                        Case EstadosEnvioProveedor.PteEnvio.ToString
                            Return EstadosEnvioProveedorDescripcion.PteEnvio
                        Case EstadosEnvioProveedor.Enviado.ToString
                            Return EstadosEnvioProveedorDescripcion.Enviado
                        Case EstadosEnvioProveedor.Recepcionado.ToString
                            Return EstadosEnvioProveedorDescripcion.Recepcionado
                        Case EstadosEnvioProveedor.Anulado.ToString
                            Return EstadosEnvioProveedorDescripcion.Anulado
                        Case EstadosEnvioProveedor.EnCarga.ToString
                            Return EstadosEnvioProveedorDescripcion.EnCarga
                        Case EstadosEnvioProveedor.EnvioParcial.ToString
                            Return EstadosEnvioProveedorDescripcion.EnvioParcial
                        Case EstadosEnvioProveedor.PteCarretilla.ToString
                            Return EstadosEnvioProveedorDescripcion.PteCarretilla
                        Case Else
                            Return ""
                    End Select

                Case GetType(Tipo_Almacen)
                    Dim vTipo As Tipo_Almacen = CType(valor, Tipo_Almacen)
                    Select Case vTipo
                        Case Tipo_Almacen.Producto_Terminado
                            Return Tipo_Almacen_Descripcion.Producto_Terminado
                        Case Tipo_Almacen.PT_Almendralejo
                            Return Tipo_Almacen_Descripcion.PT_Almendralejo
                        Case Tipo_Almacen.PT_Devoluciones
                            Return Tipo_Almacen_Descripcion.PT_Devoluciones
                        Case Tipo_Almacen.PT_Farmacia
                            Return Tipo_Almacen_Descripcion.PT_Farmacia
                        Case Tipo_Almacen.PT_Friovaldi
                            Return Tipo_Almacen_Descripcion.PT_Friovaldi
                        Case Tipo_Almacen.PT_Zafra
                            Return Tipo_Almacen_Descripcion.PT_Zafra
                        Case Tipo_Almacen.CL_Corte
                            Return Tipo_Almacen_Descripcion.CL_Corte
                        Case Else
                            Return ""
                    End Select

                Case GetType(TiposTrasporte)
                    Dim vTipo As String = CType(valor, String)
                    Select Case vTipo
                        Case TiposTrasporte.T
                            Return TiposTrasporte_Descripcion.T
                        Case TiposTrasporte.M
                            Return TiposTrasporte_Descripcion.M
                        Case TiposTrasporte.A
                            Return TiposTrasporte_Descripcion.A
                        Case Else
                            Return ""
                    End Select
                Case GetType(TipoMaterial)
                    Dim vTipo As String = CType(valor, String)
                    Select Case vTipo
                        Case TipoMaterial.NINGUNO
                            Return TipoMaterialDescripcion.NINGUNO
                        Case TipoMaterial.ProdTerminado
                            Return TipoMaterialDescripcion.ProdTerminado
                        Case TipoMaterial.MatPrima
                            Return TipoMaterialDescripcion.MatPrima
                        Case TipoMaterial.Packaging
                            Return TipoMaterialDescripcion.Packaging
                        Case TipoMaterial.Semielaborado
                            Return TipoMaterialDescripcion.Semielaborado
                        Case TipoMaterial.Fabricaciones
                            Return TipoMaterialDescripcion.Fabricaciones
                        Case TipoMaterial.MatCliente
                            Return TipoMaterialDescripcion.MatCliente
                        Case TipoMaterial.Repuestos
                            Return TipoMaterialDescripcion.Repuestos
                        Case TipoMaterial.NINGUNO
                            Return TipoMaterialDescripcion.NINGUNO
                        Case Else
                            Return ""
                    End Select


                Case GetType(ListaMeses)
                    Dim vTipo As String = CType(valor, String)
                    Select Case vTipo
                        Case ListaMeses.Enero
                            Return ListaMesesDescripcion.Enero
                        Case ListaMeses.Febrero
                            Return ListaMesesDescripcion.Febrero
                        Case ListaMeses.Marzo
                            Return ListaMesesDescripcion.Marzo
                        Case ListaMeses.Abril
                            Return ListaMesesDescripcion.Abril
                        Case ListaMeses.Mayo
                            Return ListaMesesDescripcion.Mayo
                        Case ListaMeses.Junio
                            Return ListaMesesDescripcion.Julio
                        Case ListaMeses.Julio
                            Return ListaMesesDescripcion.Julio
                        Case ListaMeses.Agosto
                            Return ListaMesesDescripcion.Agosto
                        Case ListaMeses.Septiembre
                            Return ListaMesesDescripcion.Septiembre
                        Case ListaMeses.Octubre
                            Return ListaMesesDescripcion.Octubre
                        Case ListaMeses.Noviembre
                            Return ListaMesesDescripcion.Noviembre
                        Case ListaMeses.Diciembre
                            Return ListaMesesDescripcion.Diciembre
                        Case Else
                            Return ""
                    End Select

                Case GetType(ClaseEntrega_SAP)
                    Dim vTipo As String = CType(valor, String)
                    Select Case vTipo
                        Case ClaseEntrega_SAP.Normal
                            Return ClaseEntrega_SAP_Descripcion.Normal
                        Case ClaseEntrega_SAP.Traslado
                            Return ClaseEntrega_SAP_Descripcion.Traslado
                        Case ClaseEntrega_SAP.Otros
                            Return ClaseEntrega_SAP_Descripcion.Otros
                        Case Else
                            Return ""
                    End Select

                Case GetType(Centros_SAP)
                    Dim vTipo As String = CType(valor, String)
                    Select Case vTipo
                        Case Centros_SAP.Plastiverd
                            Return Centros_SAP_Descripcion.Plastiverd
                        Case Centros_SAP.Consignas
                            Return Centros_SAP_Descripcion.Consignas
                        Case Centros_SAP.Externo
                            Return Centros_SAP_Descripcion.Externo
                        Case Centros_SAP.Otros
                            Return Centros_SAP_Descripcion.Otros
                        Case Else
                            Return ""
                    End Select

                Case GetType(TipoEnvio_Entregas)
                    Dim vTipo As String = CType(valor, String)
                    Select Case vTipo
                        Case TipoEnvio_Entregas.CAMION
                            Return TipoEnvio_EntregasDescripcion.CAMION
                        Case TipoEnvio_Entregas.AVION
                            Return TipoEnvio_EntregasDescripcion.AVION
                        Case TipoEnvio_Entregas.BARCO
                            Return TipoEnvio_EntregasDescripcion.BARCO
                        Case TipoEnvio_Entregas.CONSIGNA
                            Return TipoEnvio_EntregasDescripcion.CONSIGNA
                        Case TipoEnvio_Entregas.CORREO
                            Return TipoEnvio_EntregasDescripcion.CORREO
                        Case TipoEnvio_Entregas.FERROCARRIL
                            Return TipoEnvio_EntregasDescripcion.FERROCARRIL
                        Case TipoEnvio_Entregas.NORELEVANTE
                            Return TipoEnvio_EntregasDescripcion.NORELEVANTE
                        Case Else
                            Return ""
                    End Select

                Case GetType(TipoFabricacion)
                    Dim vValor As TipoFabricacion = CType(valor, TipoFabricacion)
                    Select Case vValor
                        Case TipoFabricacion.Produccion
                            Return TipoFabricacionDescripcion.Fabricacion
                        Case TipoFabricacion.Envasado
                            Return TipoFabricacionDescripcion.Envasado
                        Case Else
                            Return ""
                    End Select

                Case GetType(EstadoParada)
                    Dim vValor As EstadoParada = CType(valor, EstadoParada)
                    Select Case vValor
                        Case EstadoParada.Activa
                            Return EstadoParadaDescripcion.Activa
                        Case EstadoParada.Inactiva
                            Return EstadoParadaDescripcion.Inactiva
                        Case EstadoParada.PteIndicarMotivo
                            Return EstadoParadaDescripcion.PteIndicarMotivo
                    End Select

                Case GetType(EstadoOrdenTpte)
                    Dim vValor As EstadoOrdenTpte = CType(valor, EstadoOrdenTpte)
                    Select Case vValor
                        'Case EstadoOrdenTpte.EnPlanifNec
                        '    Return EstadoOrdenTpteDescripcion.EnPlanifNec
                        Case EstadoOrdenTpte.PlanifNecFin
                            Return EstadoOrdenTpteDescripcion.PlanifNecFin
                        Case EstadoOrdenTpte.Ninguno
                            Return EstadoOrdenTpteDescripcion.Ninguno
                        Case EstadoOrdenTpte.Registro
                            Return EstadoOrdenTpteDescripcion.Registro
                        Case EstadoOrdenTpte.InicioCarga
                            Return EstadoOrdenTpteDescripcion.InicioCarga
                        Case EstadoOrdenTpte.FinCarga
                            Return EstadoOrdenTpteDescripcion.FinCarga
                        Case EstadoOrdenTpte.InicioTpte
                            Return EstadoOrdenTpteDescripcion.InicioTpte
                        Case EstadoOrdenTpte.FinTpte
                            Return EstadoOrdenTpteDescripcion.FinTpte
                        Case EstadoOrdenTpte.DespachoExpedicion
                            Return EstadoOrdenTpteDescripcion.DespachoExpedicion
                        Case Else
                            Return ""
                    End Select

                Case GetType(EstadoEntregaPrevista)
                    Dim vValor As EstadoEntregaPrevista = CType(valor, EstadoEntregaPrevista)
                    Select Case vValor
                        Case EstadoEntregaPrevista.Activa
                            Return EstadoEntregaPrevistaDescripcion.Activa
                        Case EstadoEntregaPrevista.Inactiva
                            Return EstadoEntregaPrevistaDescripcion.Inactiva
                        Case EstadoEntregaPrevista.Desconocido
                            Return EstadoEntregaPrevistaDescripcion.Desconocido
                        Case Else
                            Return ""
                    End Select

                Case GetType(EstadoFabricacion)
                    Dim vValor As EstadoFabricacion = CType(valor, EstadoFabricacion)
                    Select Case vValor
                        Case EstadoFabricacion.PteFabricar
                            Return EstadoFabricacionDescripcion.PteFabricar
                        Case EstadoFabricacion.EnMarcha
                            Return EstadoFabricacionDescripcion.EnMarcha
                        Case EstadoFabricacion.Finalizada
                            Return EstadoFabricacionDescripcion.Finalizada
                        Case EstadoFabricacion.Anulada
                            Return EstadoFabricacionDescripcion.Anulada
                        Case EstadoFabricacion.Ninguna
                            Return EstadoFabricacionDescripcion.Ninguna
                        Case Else
                            Return ""

                    End Select

                Case GetType(OperacionBigBagLeido)
                    Dim vValor As OperacionBigBagLeido = CType(valor, OperacionBigBagLeido)
                    Select Case vValor
                        Case OperacionBigBagLeido.Ubicar
                            Return OperacionBigBagLeidoBDescripcion.Ubicar
                        Case OperacionBigBagLeido.Cargar_Camion
                            Return OperacionBigBagLeidoBDescripcion.Cargar_Camion
                        Case OperacionBigBagLeido.DesCargar_Camion
                            Return OperacionBigBagLeidoBDescripcion.DesCargar_Camion
                        Case OperacionBigBagLeido.Cambio_Ubicacion
                            Return OperacionBigBagLeidoBDescripcion.Cambio_Ubicacion
                        Case OperacionBigBagLeido.Otros
                            Return OperacionBigBagLeidoBDescripcion.Otros
                        Case Else
                            Return ""
                    End Select

                Case GetType(EstadoEnvasadoBB)
                    Dim vValor As EstadoEnvasadoBB = CType(valor, EstadoEnvasadoBB)
                    Select Case vValor
                        Case EstadoEnvasadoBB.PteLiberar
                            Return EstadoEnvasadoBBDescripcion.PteLiberar
                        Case EstadoEnvasadoBB.Liberada
                            Return EstadoEnvasadoBBDescripcion.Liberada
                        Case EstadoEnvasadoBB.EnCurso
                            Return EstadoEnvasadoBBDescripcion.EnCurso
                        Case EstadoEnvasadoBB.Finalizada
                            Return EstadoEnvasadoBBDescripcion.Finalizada
                        Case EstadoEnvasadoBB.Anulada
                            Return EstadoEnvasadoBBDescripcion.Anulada
                        Case EstadoEnvasadoBB.Ninguno
                            Return EstadoEnvasadoBBDescripcion.Ninguno
                        Case Else
                            Return ""
                    End Select

                Case GetType(EstadoCamion)
                    Dim vTipo As EstadoCamion = CType(valor, EstadoCamion)
                    Select Case vTipo
                        Case EstadoCamion.Todos
                            Return EstadoCamionDescripcion.Todos
                        Case EstadoCamion.SinCargar
                            Return EstadoCamionDescripcion.SinCargar
                        Case EstadoCamion.EnCarga
                            Return EstadoCamionDescripcion.EnCarga
                        Case EstadoCamion.Cargado
                            Return EstadoCamionDescripcion.Cargado
                        Case EstadoCamion.Albaranado
                            Return EstadoCamionDescripcion.Albaranado
                        Case EstadoCamion.Anulado
                            Return EstadoCamionDescripcion.Anulado
                        Case Else
                            Return ""
                    End Select




                Case GetType(EstadoBigBags)
                    Dim vTipo As String = CType(valor, String)
                    Select Case vTipo
                        Case EstadoBigBags.Ninguno
                            Return EstadoBigBagsDescripcion.Ninguno
                        Case EstadoBigBags.SinLeer
                            Return EstadoBigBagsDescripcion.SinLeer
                        Case EstadoBigBags.Ubicado
                            Return EstadoBigBagsDescripcion.Ubicado
                        Case EstadoBigBags.Cargado
                            Return EstadoBigBagsDescripcion.Cargado
                        Case EstadoBigBags.Perdido
                            Return EstadoBigBagsDescripcion.Perdido
                        Case EstadoBigBags.EnSalidaStock
                            Return EstadoBigBagsDescripcion.EnSalidaStock
                        Case Else
                            Return ""
                    End Select
                Case Else
                    Return ""
            End Select
        Catch ex As Exception
            enumToString = ""
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & MethodBase.GetCurrentMethod().DeclaringType.Name & "." & MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

End Class
