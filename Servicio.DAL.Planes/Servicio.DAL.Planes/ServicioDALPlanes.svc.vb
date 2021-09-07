Option Strict Off
Imports System
Imports System.IO
Imports System.Text
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Data.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data.Sql
Public Class ServicioDALPlanes
    Implements IServicioDALPlanes

    Protected DB As Database
    Protected a As New Object
    Public Sub New()
        Dim connectionStrinName As String = "cnPlanes"
        Dim ObjDatosLocal As New ClaseAccesoDatos
        Dim NombreDSN = ""
        Dim StrDSN = ""
        Dim docPath As String = AppDomain.CurrentDomain.BaseDirectory
        Try

            'Using writer As StreamWriter =
            'New StreamWriter(Path.Combine(docPath, "LOG_CONNECTION.txt"), True)
            '    writer.WriteLine("Planes.vb")
            'End Using
            ObjDatosLocal.BuscaRutasIni("Planes.ini", "Planes")
            ObjDatosLocal.CargarConexionDatos(ObjDatosLocal.TipoDSN, ObjDatosLocal.NameDSN1, StrDSN, "ADO.NET")
            NombreDSN = ObjDatosLocal.NameDSN1

            'Using writer As StreamWriter =
            'New StreamWriter(Path.Combine(docPath, "LOG_CONNECTION.txt"), True)
            '    writer.WriteLine("str_connection_" & StrDSN)
            'End Using

            Dim dbEngine As SqlDatabase = New SqlDatabase(StrDSN)
            Dim factory As New DatabaseProviderFactory
            DB = dbEngine
        Catch ex As Exception
            Using writer As StreamWriter =
            New StreamWriter(Path.Combine(docPath, "LOG_ERROR.txt"), True)
                writer.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss => ") & "ERROR_" & ex.Message.Replace("'", ""))
            End Using
        End Try
    End Sub

    Private Shared Function CreateDbConnection(ByVal providerName As String, ByVal connectionString As String) As DbConnection
        Dim connection As DbConnection = Nothing

        If connectionString IsNot Nothing Then

            Try
                Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(providerName)
                connection = factory.CreateConnection()
                connection.ConnectionString = connectionString
            Catch ex As Exception

                If connection IsNot Nothing Then
                    connection = Nothing
                End If

                Console.WriteLine(ex.Message)
            End Try
        End If

        Return connection
    End Function


    Public Function TablasGenerales(Sistema As String,
                                                      Tabla As String,
                                                      Condicion As String) As Respuesta Implements IServicioDALPlanes.TablasGenerales
        Dim resultSetMapper = MapBuilder(Of TablaGral).MapAllProperties().Build()
        Dim result As IQueryable(Of TablaGral) = Nothing
        Dim Respuesta As Respuesta
        Try
            result = DB.ExecuteSprocAccessor("Tab_ConsultarRegistro", resultSetMapper, Sistema, " * ", Tabla, Condicion, 0).AsQueryable()

            Dim ds As New DataSet()

            Dim Tab As DataTable = New DataTable("Datos")

            Tab.Columns.Add("TABLA", System.Type.GetType("System.String"))
            Tab.Columns.Add("KEY_ALFA", System.Type.GetType("System.String"))
            Tab.Columns.Add("KEY_NUMERO", System.Type.GetType("System.Double"))
            Tab.Columns.Add("DESCRIPCIO", System.Type.GetType("System.String"))
            Tab.Columns.Add("ABREVIACIO", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA1", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA2", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA3", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA4", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA5", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA6", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA7", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA8", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA9", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA10", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA11", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA12", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA13", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA14", System.Type.GetType("System.String"))
            Tab.Columns.Add("ALFA15", System.Type.GetType("System.String"))
            Tab.Columns.Add("VALOR1", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR2", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR3", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR4", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR5", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR6", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR7", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR8", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR9", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR10", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR11", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR12", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR13", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR14", System.Type.GetType("System.Double"))
            Tab.Columns.Add("VALOR15", System.Type.GetType("System.Double"))
            Tab.Columns.Add("CORR_SUSC", System.Type.GetType("System.Double"))
            Tab.Columns.Add("MEMO1", System.Type.GetType("System.String"))
            Tab.Columns.Add("MEMO2", System.Type.GetType("System.String"))
            Tab.Columns.Add("SISTEMA", System.Type.GetType("System.String"))

            For Each gral As TablaGral In result
                Dim row As DataRow = Tab.NewRow()
                row("TABLA") = gral.TABLA
                row("KEY_ALFA") = gral.KEY_ALFA
                row("KEY_NUMERO") = gral.KEY_NUMERO
                row("DESCRIPCIO") = gral.DESCRIPCIO
                row("ABREVIACIO") = gral.ABREVIACIO
                row("ALFA1") = gral.ALFA1
                row("ALFA2") = gral.ALFA2
                row("ALFA3") = gral.ALFA3
                row("ALFA4") = gral.ALFA4
                row("ALFA5") = gral.ALFA5
                row("ALFA6") = gral.ALFA6
                row("ALFA7") = gral.ALFA7
                row("ALFA8") = gral.ALFA8
                row("ALFA9") = gral.ALFA9
                row("ALFA10") = gral.ALFA10
                row("ALFA11") = gral.ALFA11
                row("ALFA12") = gral.ALFA12
                row("ALFA13") = gral.ALFA13
                row("ALFA14") = gral.ALFA14
                row("ALFA15") = gral.ALFA15
                row("VALOR1") = gral.VALOR1
                row("VALOR2") = gral.VALOR2
                row("VALOR3") = gral.VALOR3
                row("VALOR4") = gral.VALOR4
                row("VALOR5") = gral.VALOR5
                row("VALOR6") = gral.VALOR6
                row("VALOR7") = gral.VALOR7
                row("VALOR8") = gral.VALOR8
                row("VALOR9") = gral.VALOR9
                row("VALOR10") = gral.VALOR10
                row("VALOR11") = gral.VALOR11
                row("VALOR12") = gral.VALOR12
                row("VALOR13") = gral.VALOR13
                row("VALOR14") = gral.VALOR14
                row("VALOR15") = gral.VALOR15
                row("CORR_SUSC") = gral.CORR_SUSC
                row("MEMO1") = gral.MEMO1
                row("MEMO2") = gral.MEMO2
                row("SISTEMA") = gral.SISTEMA
                Tab.Rows.Add(row)
            Next
            ds.Tables.Add(Tab)
            Respuesta = New Respuesta()
            Respuesta.Mensaje = ""
            Respuesta.Codigo = "0"
            Respuesta.Respuesta = True
            Respuesta.Objeto1 = ds

        Catch ex As Exception
            Respuesta = New Respuesta()
            Respuesta.Mensaje = ex.Message
            Respuesta.Codigo = "-1"
            Respuesta.Respuesta = False
        End Try
        Return Respuesta
    End Function
End Class

Public Class ClaseAccesoDatos
    'Clase Acceso Datos .NET
    Public PathFilesIni As String, PathDatabase As String,
            NameDatabase1 As String,
            UserIdSistema As String, NameUsuario As String,
            TipoDSN As String, FisicalServerName,
            ServerName As String, UserId As String, PassWord As String
    'Variables para definir tipo de acceso externo o interno para
    'consultar seguridad de planes o externa
    Public TipoAcceso As String, Puerta As String, Sistema As String


    'Información  de Modulo de Seguridad
    Public PathdatabaseSeg As String, NamedatabaseSeg As String

    'Información de Tablas Externas Vinculadas a Bases de Datos Access
    Public FoxVersion As String, PathFoxBase As String, PathFoxBase2 As String

    'Información de Conexión Utilizando ODBC
    Public NameDSN1 As String, NameDSN2 As String,
           DataVersion1 As String

    'Información para Conexión TCP/IP
    Public HostRemote As String, _PortRemote As String

    'Información de Rutas de Reports
    Public PathFilesRpt As String

    'Información de Utilitarios Anexos al Sistema
    Public PathCargaBCP As String

    'Información para envio de Mail a Operador
    Public NombrePerfil As String, NombreDestinatario As String, NombreDestinatarioCC As String
    Public Url1 As String, Url2 As String

    'Tipo de Validación de Seguridad: DOMAIN/PLANES
    Public TipoSeguridad As String

    'Ruta para carpeta de clientes
    Public PathDatabaseDoc As String

    'Código de Aplicación
    Public _CodAPL As String
    'Otras Variables 
    Public ServerName_Rutina As String, UserId_Rutina As String, PassId_Rutina As String, NameDatabase1_Rutina As String,
           DataVersion1_Rutina As String

    Public Function CrearCarro(ByVal pComputador As String,
                               ByVal pUsuario As String,
                               ByVal pModalidad As String,
                               ByVal pSistema As String,
                               ByVal pPrograma As String,
                      Optional ByVal pServicio As String = "",
                      Optional ByVal pClave As String = "") As String

        CrearCarro = IIf(Len(Trim(pComputador)) > 0, pComputador, " ") & "@" &
                     IIf(Len(Trim(pUsuario)) > 0, pUsuario, " ") & "@" &
                     IIf(Len(Trim(pModalidad)) > 0, pModalidad, " ") & "@" &
                     IIf(Len(Trim(pSistema)) > 0, pSistema, " ") & "@" &
                     IIf(Len(Trim(pPrograma)) > 0, pPrograma, " ") & "@" &
                     IIf(Len(Trim(pServicio)) > 0, pServicio, " ") & "@" &
                     IIf(Len(Trim(pClave)) > 0, pClave, " ") & "@"
    End Function

    Public Shared Sub Lee()
        Dim path As String = "c:  \temp\MyTest.txt"
        Dim fs As FileStream

        ' Delete the file if it exists.
        If File.Exists(path) = False Then
            ' Create the file.
            fs = File.Create(path)
            Dim info As Byte() = New UTF8Encoding(True).GetBytes("This is some text in the file.")

            ' Add some information to the file.
            fs.Write(info, 0, info.Length)
            fs.Close()
        End If

        ' Open the stream and read it back.
        fs = File.OpenRead(path)
        Dim b(1024) As Byte
        Dim temp As UTF8Encoding = New UTF8Encoding(True)

        Do While fs.Read(b, 0, b.Length) > 0
            Console.WriteLine(temp.GetString(b))
        Loop
        fs.Close()
    End Sub

    Public Function App_Path() As String
        Dim Ruta As String
        Ruta = System.AppDomain.CurrentDomain.BaseDirectory()
        Ruta = Ruta.Replace("/", "\")
        Return Ruta
    End Function

    Public Sub BuscaRutasIni(ByVal NombreIni As String, ByVal ModuloIni As String)
        'Lee desde "Invpath.Ini" la ruta en donde
        'se debe encontrar ini buscado
        Dim MyPath As String, Pos As Integer, Largo As Integer, LineaLeida As String


        MyPath = App_Path() & "Invpaths.ini"

        Dim docPath As String = AppDomain.CurrentDomain.BaseDirectory
        'Using writer As StreamWriter =
        '    New StreamWriter(Path.Combine(docPath, "LOG_CONNECTION.txt"), True)
        '    writer.WriteLine("MyPath_" & MyPath)
        'End Using

        PathFilesIni = ""
        If Len(Trim(Dir(MyPath))) > 0 Then
            Dim sr As StreamReader = File.OpenText(MyPath)
            LineaLeida = sr.ReadLine()
            While Not LineaLeida Is Nothing
                Pos = InStr(1, LineaLeida, "=")
                Largo = Len(Trim(LineaLeida))
                If Pos > 0 Then
                    Select Case UCase(Mid(LineaLeida, 1, Pos - 1))
                        Case "PATHFILESINI"
                            PathFilesIni = Mid(LineaLeida, Pos + 1, Largo) & "\"
                    End Select
                End If
                LineaLeida = sr.ReadLine()
            End While
            sr.Close()
        End If
        If Len(Trim(PathFilesIni)) > 0 Then
            LecturaIni(NombreIni, ModuloIni, PathFilesIni)
        Else
            LecturaIni(NombreIni, ModuloIni)
        End If
    End Sub


    Public Sub LecturaIni(ByVal NombreIni As String,
                          ByVal ModuloIni As String,
                 Optional ByVal Ruta As String = "")
        'On Error GoTo ControlError
        'CodigoRutina = ""
        'MensajeRutina = ""
        'TipoErrorRutina = 0

        'Lee Ini y almacena las variables de entorno
        Dim MyPath As String, Pos As Integer, Largo As Integer,
            LineaLeida As String, LeerModulo As Boolean
        If Len(Trim(Ruta)) = 0 Then
            MyPath = App_Path()
        Else
            MyPath = Ruta
        End If
        LeerModulo = False
        PathDatabase = "" : NameDatabase1 = "" : PathFoxBase = "" : PathFoxBase2 = "" : PathFilesRpt = ""
        PathdatabaseSeg = "" : NamedatabaseSeg = "" : HostRemote = "" : _PortRemote = ""

        Dim sr As StreamReader = File.OpenText(MyPath & NombreIni)
        LineaLeida = sr.ReadLine()
        While Not LineaLeida Is Nothing
            If UCase(Trim(LineaLeida)) = "[" & UCase(ModuloIni) & "]" Then
                LeerModulo = True
                LineaLeida = sr.ReadLine()
            End If
            If LeerModulo Then
                If InStr(1, LineaLeida, "[") > 0 Then
                    Exit While
                Else
                    Pos = InStr(1, LineaLeida, "=")
                    Largo = Len(Trim(LineaLeida))
                    If Pos > 0 Then
                        Select Case UCase(Mid(LineaLeida, 1, Pos - 1))
                            Case "PATHDATABASE"
                                PathDatabase = Mid(LineaLeida, Pos + 1, Largo) & "\"
                            Case "DATABASENAME1"
                                NameDatabase1 = Mid(LineaLeida, Pos + 1, Largo)
                                NameDatabase1_Rutina = NameDatabase1
                            Case "PATHFOXBASE"
                                PathFoxBase = Mid(LineaLeida, Pos + 1, Largo)
                            Case "PATHFOXBASE2"
                                PathFoxBase2 = Mid(LineaLeida, Pos + 1, Largo)
                            Case "FOXVERSION"
                                FoxVersion = Mid(LineaLeida, Pos + 1, Largo)
                            Case "DATAVERSION1"
                                DataVersion1 = Mid(LineaLeida, Pos + 1, Largo)
                                DataVersion1_Rutina = DataVersion1
                            Case "NAMEDSN1"
                                NameDSN1 = Mid(LineaLeida, Pos + 1, Largo)
                            Case "NAMEDSN2"
                                NameDSN2 = Mid(LineaLeida, Pos + 1, Largo)
                            Case "PATHFILESRPT"
                                PathFilesRpt = Mid(LineaLeida, Pos + 1, Largo) & "\"
                            Case "PATHDATABASESEG"
                                PathdatabaseSeg = Mid(LineaLeida, Pos + 1, Largo) & "\"
                            Case "NAMEDATABASESEG"
                                NamedatabaseSeg = Mid(LineaLeida, Pos + 1, Largo)
                            Case "HOSTNAME"
                                HostRemote = Mid(LineaLeida, Pos + 1, Largo)
                            Case "HOSTPORT"
                                _PortRemote = Val(Mid(LineaLeida, Pos + 1, Largo))
                            Case "PATHCARGABCP"
                                PathCargaBCP = Mid(LineaLeida, Pos + 1, Largo) & "\"
                            Case "TIPODSN"
                                TipoDSN = Mid(LineaLeida, Pos + 1, Largo)
                            Case "FISICALSERVERNAME"
                                FisicalServerName = Mid(LineaLeida, Pos + 1, Largo)
                            Case "SERVERNAME"
                                ServerName = Mid(LineaLeida, Pos + 1, Largo)
                                ServerName_Rutina = ServerName
                            Case "USERID"
                                UserId = Mid(LineaLeida, Pos + 1, Largo)
                                UserId_Rutina = UserId
                            Case "PASSWORD"
                                PassWord = Mid(LineaLeida, Pos + 1, Largo)
                                PassId_Rutina = PassWord
                            Case "NOMBREPERFIL"
                                NombrePerfil = Mid(LineaLeida, Pos + 1, Largo)
                            Case "NOMBREDESTINATARIO"
                                NombreDestinatario = Mid(LineaLeida, Pos + 1, Largo)
                            Case "NOMBREDESTINATARIOCC"
                                NombreDestinatarioCC = Mid(LineaLeida, Pos + 1, Largo)
                            Case "URL1"
                                Url1 = Mid(LineaLeida, Pos + 1, Largo)
                            Case "URL2"
                                Url2 = Mid(LineaLeida, Pos + 1, Largo)
                            Case "TIPOACCESO"
                                TipoAcceso = Mid(LineaLeida, Pos + 1, Largo)
                            Case "PUERTA"
                                Puerta = Mid(LineaLeida, Pos + 1, Largo)
                            Case "SISTEMA"
                                Sistema = Mid(LineaLeida, Pos + 1, Largo)
                            Case "TIPOSEGURIDAD"
                                TipoSeguridad = Mid(LineaLeida, Pos + 1, Largo)
                            Case "PATHDATABASEDOC"
                                PathDatabaseDoc = Mid(LineaLeida, Pos + 1, Largo)
                            Case "CODIGOSISTEMA"
                                _CodAPL = Mid(LineaLeida, Pos + 1, Largo)
                        End Select
                    End If
                End If
            End If
            LineaLeida = sr.ReadLine()
        End While
        sr.Close()
        If Puerta <> "" Then
            Dim ObjSeg As Object
            ObjSeg = CreateObject("LibreriaSegNet.clsEncriptado")
            If PathDatabase <> "" Then PathDatabase = ObjSeg.DecryptString(Mid(PathDatabase, 1, Len(PathDatabase) - 1), TipoAcceso, Puerta) & "\"
            If NameDatabase1 <> "" Then NameDatabase1 = ObjSeg.DecryptString(NameDatabase1, TipoAcceso, Puerta)
            NameDatabase1_Rutina = NameDatabase1
            If PathFoxBase <> "" Then PathFoxBase = ObjSeg.DecryptString(PathFoxBase, TipoAcceso, Puerta)
            If PathFoxBase2 <> "" Then PathFoxBase2 = ObjSeg.DecryptString(PathFoxBase2, TipoAcceso, Puerta)
            If FoxVersion <> "" Then FoxVersion = ObjSeg.DecryptString(FoxVersion, TipoAcceso, Puerta)
            If DataVersion1 <> "" Then DataVersion1 = ObjSeg.DecryptString(DataVersion1, TipoAcceso, Puerta)
            DataVersion1_Rutina = DataVersion1
            If NameDSN1 <> "" Then NameDSN1 = ObjSeg.DecryptString(NameDSN1, TipoAcceso, Puerta)
            If NameDSN2 <> "" Then NameDSN2 = ObjSeg.DecryptString(NameDSN2, TipoAcceso, Puerta)
            If PathFilesRpt <> "" Then PathFilesRpt = ObjSeg.DecryptString(Mid(PathFilesRpt, 1, Len(PathFilesRpt) - 1), TipoAcceso, Puerta) & "\"
            If PathdatabaseSeg <> "" Then PathdatabaseSeg = ObjSeg.DecryptString(Mid(PathdatabaseSeg, 1, Len(PathdatabaseSeg) - 1), TipoAcceso, Puerta) & "\"
            If NamedatabaseSeg <> "" Then NamedatabaseSeg = ObjSeg.DecryptString(NamedatabaseSeg, TipoAcceso, Puerta)
            If HostRemote <> "" Then HostRemote = ObjSeg.DecryptString(HostRemote, TipoAcceso, Puerta)
            If _PortRemote <> "" Then _PortRemote = ObjSeg.DecryptString(_PortRemote, TipoAcceso, Puerta)
            If PathCargaBCP <> "" Then PathCargaBCP = ObjSeg.DecryptString(Mid(PathCargaBCP, 1, Len(PathCargaBCP) - 1), TipoAcceso, Puerta) & "\"
            If TipoDSN <> "" Then TipoDSN = ObjSeg.DecryptString(TipoDSN, TipoAcceso, Puerta)
            If FisicalServerName <> "" Then FisicalServerName = ObjSeg.DecryptString(FisicalServerName, TipoAcceso, Puerta)
            If ServerName <> "" Then ServerName = ObjSeg.DecryptString(ServerName, TipoAcceso, Puerta)
            ServerName_Rutina = ServerName
            If UserId <> "" Then UserId = ObjSeg.DecryptString(UserId, TipoAcceso, Puerta)
            UserId_Rutina = UserId
            If PassWord <> "" Then PassWord = ObjSeg.DecryptString(PassWord, TipoAcceso, Puerta)
            PassId_Rutina = PassWord
            If NombrePerfil <> "" Then NombrePerfil = ObjSeg.DecryptString(NombrePerfil, TipoAcceso, Puerta)
            If NombreDestinatario <> "" Then NombreDestinatario = ObjSeg.DecryptString(NombreDestinatario, TipoAcceso, Puerta)
            If NombreDestinatarioCC <> "" Then NombreDestinatarioCC = ObjSeg.DecryptString(NombreDestinatarioCC, TipoAcceso, Puerta)
            If Url1 <> "" Then Url1 = ObjSeg.DecryptString(Url1, TipoAcceso, Puerta)
            If Url2 <> "" Then Url2 = ObjSeg.DecryptString(Url2, TipoAcceso, Puerta)
            If Sistema <> "" Then Sistema = ObjSeg.DecryptString(Sistema, TipoAcceso, Puerta)
            If TipoSeguridad <> "" Then TipoSeguridad = ObjSeg.DecryptString(TipoSeguridad, TipoAcceso, Puerta)
            If PathDatabaseDoc <> "" Then PathDatabaseDoc = ObjSeg.DecryptString(PathDatabaseDoc, TipoAcceso, Puerta)
            If _CodAPL <> "" Then _CodAPL = ObjSeg.DecryptString(_CodAPL, TipoAcceso, Puerta)
        End If
        'If Not LeerModulo Then
        '   CodigoRutina = "CADC000004"
        '   MensajeRutina = "No fue encontrada sección '" & "[" & ModuloIni & "]' en archivo '" & NombreIni & "'."
        '   TipoErrorRutina = vbExclamation
        '   SugerenciaRutina = "Debe ser agregada esta sección al archivo '" & NombreIni & "'"
        'End If
    End Sub
    Public Sub CargarConexionDatos(ByVal Tipo As String,
                                   ByRef vNameDSN As String,
                                   ByRef vStrConexion As String,
                          Optional ByVal Motor As String = "")

        Select Case Trim(Motor)
            Case "" 'ADO
                Select Case UCase(Trim(Tipo))
                    Case "ACCESS"
                        vNameDSN = NameDSN1
                        vStrConexion = "UID=;PWD=;Driver={" & DataVersion1 & "};" &
                                    "Server=;Database=" & NameDatabase1 & ";"
                    Case "FOXPRO"
                        vNameDSN = NameDSN2
                        vStrConexion = "UID=;PWD=;Driver={" & FoxVersion & "};" &
                                    "Server=;Database=" & PathFoxBase & ";"
                    Case "ADO"
                        vNameDSN = NameDSN1
                        vStrConexion = "UID=" & UserId & ";PWD=" & PassWord & ";Driver=" & DataVersion1 & ";" &
                                    "Server=" & ServerName & ";Database=" & NameDatabase1 & ";App=Planes;"

                    Case "SQL SERVER"
                        vNameDSN = NameDSN1
                        vStrConexion = "UID=" & UserId & ";PWD=" & PassWord & ";Driver={" & DataVersion1 & "};" &
                                    "Server=" & ServerName & ";Database=" & NameDatabase1 & ";App=Planes;"

                    Case "ADO_WUSER"
                        vNameDSN = NameDSN1
                        vStrConexion = "Server=" & ServerName & ";Provider=SQLOLEDB.1; INTEGRATED SECURITY=SSPI; PERSIST SECURITY INFO=FALSE;" &
                                    "DATABASE=" & NameDatabase1 & "; Connection Timeout=360; App=Planes;"
                End Select
            Case "DAO"
                Select Case UCase(Trim(Tipo))
                    Case "ACCESS"
                        vNameDSN = PathDatabase & NameDatabase1
                        vStrConexion = ""
                    Case "FOXPRO"
                        vNameDSN = PathFoxBase
                        vStrConexion = FoxVersion
                End Select
            Case "ADO.NET"
                Select Case UCase(Trim(Tipo))
               'Case "ACCESS"
               '   vNameDSN = NameDSN1
               '   vStrConexion = "UID=;PWD=;Driver={" & DataVersion1 & "};" & _
               '               "Server=;Database=" & NameDatabase1 & ";"
               'Case "FOXPRO"
               '   vNameDSN = NameDSN2
               '   vStrConexion = "UID=;PWD=;Driver={" & FoxVersion & "};" & _
               '               "Server=;Database=" & PathFoxBase & ";"
                    Case "ADO"
                        vNameDSN = NameDSN1
                        vStrConexion = "workstation id=" & Environment.MachineName & ";packet size=4096;" &
                                       "user id=" & UserId & ";" & "data source=" & ServerName & ";" &
                                       "persist security info=True;initial catalog=" & NameDatabase1 &
                                       ";password=" & PassWord & ";Connect Timeout=360"

                  'Case "SQL SERVER"   'AHI QUE ARMAR ESTE STRING
                  '   vNameDSN = NameDSN1
                  '   vStrConexion = "UID=" & UserId & ";PWD=" & PassWord & ";Driver={" & DataVersion1 & "};" & _
                  '               "Server=" & ServerName & ";Database=" & NameDatabase1 & ";App=Planes;"

                    Case "ADO_WUSER"
                        vNameDSN = NameDSN1
                        vStrConexion = "workstation id=" & Environment.MachineName & ";packet size=4096;" &
                                       "integrated security=SSPI;" & "data source=" & ServerName &
                                       ";persist security info=False;initial catalog=" & NameDatabase1 & ";" &
                                       "Connect Timeout=360"
                End Select
        End Select
    End Sub
    Public Property CodAPL() As Integer
        Get
            Return CInt(_CodAPL)
        End Get
        Set(ByVal value As Integer)
            _CodAPL = value
        End Set
    End Property
    Public Property PortRemote() As Integer
        Get
            Return CInt(_PortRemote)
        End Get
        Set(ByVal value As Integer)
            _PortRemote = value
        End Set
    End Property
End Class

