Public Class stockAltualSAP

#Region "Atributos"

    Private bCreado As Boolean
    Private miMaterial As New Material
#End Region


#Region "Constructores"

    Private Sub InicializarVariables()
        Try
            Material = ""
            Centro = ""
            GTIN = ""
            LoteSAP = ""
            Almacen = ""
            Stock_LU = 0
            Stock_BL = 0
            bCreado = False

        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub

    Public Sub New(sMaterial As String,
                   sCentro As String,
                   sGTIN As String,
                   sLoteSAP As String,
                   sAlmacen As String,
                   iStockLu As Integer,
                   iStockBl As Integer)

        InicializarVariables()
        Material = sMaterial
        Centro = sCentro
        GTIN = sGTIN
        LoteSAP = sLoteSAP
        Almacen = sAlmacen
        Stock_LU = iStockLu
        Stock_BL = Stock_BL

        Me.bCreado = True

    End Sub


#End Region

#Region "Propiedades"
    Public Property Material As String
    Public Property Centro As String
    Public Property GTIN As String
    Public Property LoteSAP As String
    Public Property Almacen As String
    Public Property Stock_LU As Integer
    Public Property Stock_BL As Integer

    Public ReadOnly Property MaterialDetalle As Material
        Get
            If miMaterial.Creado = False Then
                miMaterial = New Material(Material)
            End If

            Return miMaterial
        End Get
    End Property

#End Region

End Class
