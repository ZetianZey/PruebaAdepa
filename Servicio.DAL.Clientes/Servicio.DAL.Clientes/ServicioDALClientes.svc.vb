Public Class ServicioDALClientes
    Implements IServicioDALClientes

    Public Sub New()
    End Sub

    Public Function RescataDatosParcialesCliente(SecInd As String) As Respuesta Implements IServicioDALClientes.RescataDatosParcialesCliente

        Dim Respuesta As Respuesta
        Dim ds As New DataSet
        Dim ObjLibClientes As New LibClientesNet.ClsClientes
        Try

            ds = ObjLibClientes.RescataDatosParcialesCliente(SecInd)

            If ObjLibClientes.MensajeRetorno = "" Then
                Respuesta = New Respuesta()
                Respuesta.Mensaje = ""
                Respuesta.Codigo = "0"
                Respuesta.Respuesta = True
                Respuesta.Objeto1 = ds
            Else
                Respuesta = New Respuesta()
                Respuesta.Mensaje = ObjLibClientes.MensajeRetorno
                Respuesta.Codigo = ObjLibClientes.CodigoRetorno.ToString()
                Respuesta.Respuesta = False
            End If

        Catch ex As Exception
            Respuesta = New Respuesta()
            Respuesta.Mensaje = ex.Message
            Respuesta.Codigo = "-1"
            Respuesta.Respuesta = False
        Finally
            ObjLibClientes = Nothing
        End Try
        Return Respuesta
    End Function

    Public Function RescataDatosBusquedaCliente(Campo As String, Busqueda As String) As Respuesta Implements IServicioDALClientes.RescataDatosBusquedaCliente

        Dim Respuesta As Respuesta
        Dim ds As DataSet
        Dim ObjLibClientes As New LibClientesNet.ClsClientes
        Try



            ds = ObjLibClientes.RescataDatosBusquedaCliente(Campo, Busqueda)

            '----------------------------------
            If ObjLibClientes.MensajeRetorno = "" Then
                Respuesta = New Respuesta()
                Respuesta.Mensaje = ""
                Respuesta.Codigo = "0"
                Respuesta.Respuesta = True
                Respuesta.Objeto1 = ds
            Else
                Respuesta = New Respuesta()
                Respuesta.Mensaje = ObjLibClientes.MensajeRetorno
                Respuesta.Codigo = ObjLibClientes.CodigoRetorno.ToString()
                Respuesta.Respuesta = False
            End If


        Catch ex As Exception
            Respuesta = New Respuesta()
            Respuesta.Mensaje = ex.Message
            Respuesta.Codigo = "-1"
            Respuesta.Respuesta = False

        Finally
            ObjLibClientes = Nothing
        End Try
        Return Respuesta
    End Function



End Class
