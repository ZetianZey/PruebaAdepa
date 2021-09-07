' NOTE: You can use the "Rename" command on the context menu to change the interface name "IService1" in both code and config file together.
<ServiceContract()>
Public Interface IServicioDALClientes

    <OperationContract()>
    Function RescataDatosParcialesCliente(SecInd As String) As Respuesta
    <OperationContract()>
    Function RescataDatosBusquedaCliente(Busqueda As String, Campo As String) As Respuesta

End Interface

<DataContract>
Public Class Respuesta
    <DataMember()>
    Public Property Mensaje As String
    <DataMember()>
    Public Property Codigo As String
    <DataMember()>
    Public Property Respuesta As Boolean
    <DataMember()>
    Public Property Objeto1 As DataSet

    <DataMember()>
    Public Property Objeto2 As DataSet
    <DataMember()>
    Public Property Objeto3 As Object
    <DataMember()>
    Public Property Objeto4 As Object
    <DataMember()>
    Public Property String1 As String
    <DataMember()>
    Public Property String2 As String
    <DataMember()>
    Public Property Int1 As Integer
    <DataMember()>
    Public Property Int2 As Integer

End Class