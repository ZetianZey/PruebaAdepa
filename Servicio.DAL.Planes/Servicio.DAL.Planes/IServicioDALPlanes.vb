' NOTE: You can use the "Rename" command on the context menu to change the interface name "IServicioDALPlanes" in both code and config file together.
<ServiceContract()>
Public Interface IServicioDALPlanes

    <OperationContract()>
    Function TablasGenerales(sistema As String, tabla As String, condicion As String) As Respuesta

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
<DataContract>
Public Class TablaGral
    <DataMember()>
    Property TABLA As String
    <DataMember()>
    Property KEY_ALFA As String
    <DataMember()>
    Property KEY_NUMERO As Integer
    <DataMember()>
    Property DESCRIPCIO As String
    <DataMember()>
    Property ABREVIACIO As String
    <DataMember()>
    Property ALFA1 As String
    <DataMember()>
    Property ALFA2 As String
    <DataMember()>
    Property ALFA3 As String
    <DataMember()>
    Property ALFA4 As String
    <DataMember()>
    Property ALFA5 As String
    <DataMember()>
    Property ALFA6 As String
    <DataMember()>
    Property ALFA7 As String
    <DataMember()>
    Property ALFA8 As String
    <DataMember()>
    Property ALFA9 As String
    <DataMember()>
    Property ALFA10 As String
    <DataMember()>
    Property ALFA11 As String
    <DataMember()>
    Property ALFA12 As String
    <DataMember()>
    Property ALFA13 As String
    <DataMember()>
    Property ALFA14 As String
    <DataMember()>
    Property ALFA15 As String
    <DataMember()>
    Property VALOR1 As Integer
    <DataMember()>
    Property VALOR2 As Integer
    <DataMember()>
    Property VALOR3 As Integer
    <DataMember()>
    Property VALOR4 As Integer
    <DataMember()>
    Property VALOR5 As Integer
    <DataMember()>
    Property VALOR6 As Integer
    <DataMember()>
    Property VALOR7 As Integer
    <DataMember()>
    Property VALOR8 As Integer
    <DataMember()>
    Property VALOR9 As Integer
    <DataMember()>
    Property VALOR10 As Integer
    <DataMember()>
    Property VALOR11 As Integer
    <DataMember()>
    Property VALOR12 As Integer
    <DataMember()>
    Property VALOR13 As Integer
    <DataMember()>
    Property VALOR14 As Integer
    <DataMember()>
    Property VALOR15 As Integer
    <DataMember()>
    Property CORR_SUSC As Integer
    <DataMember()>
    Property MEMO1 As String
    <DataMember()>
    Property MEMO2 As String
    <DataMember()>
    Property SISTEMA As String

End Class