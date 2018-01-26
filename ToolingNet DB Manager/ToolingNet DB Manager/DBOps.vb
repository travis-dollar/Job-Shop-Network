Imports System.IO
Imports System.Xml
Imports LogManager

''' <summary>
''' Performs XML Database Operations. Files must be accesible.
''' </summary>
Public Class DBOps






#Region "Utilities"


    ''' <summary>
    ''' Checks to see if the database directory is usable.
    ''' </summary>
    ''' <returns>True if the directory exists. False if it is missing or an error occurs.</returns>
    Private Function CheckDirectory() As Boolean

        Try

            If Directory.Exists(My.Settings.DBDirectory.ToString) = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Dim sLog As New LogManager.Log

            sLog.SendErrorLog(ex, True, "Could not find the database directory.")

            Return False


        End Try


    End Function

#End Region


    ''' <summary>
    ''' The client configuration for the database operations
    ''' </summary>
    Private Class ClientConfig

        Public d3ListActive As String
        Public personnel As String
        Public workstations As String
        Public activeJobs As String
        Public archivedJobs As String
        Public pollDelay As Integer
        Public pollDelayUnits As String
        Public serverIP As String


        Public Sub New()

        End Sub



        ''' <summary>
        ''' Loads the client configuration into the class parameters
        ''' </summary>
        Private Sub LoadClientConfig()
            Dim xParams As XElement


            If CheckClientConfig(True) = True Then
                Try
                    xParams = XElement.Load(My.Settings.DBDirectory & "ClientConfig.xml")

                    For Each xItem As XElement In xParams.Elements

                        'Load the XML file location structure
                        If xItem.Name = "XML_Locations" Then
                            For Each xSubItem As XElement In xItem.Elements
                                Select Case xSubItem.Name
                                    Case "D3_List_Active"
                                        d3ListActive = xSubItem.Value
                                    Case "Personnel"
                                        personnel = xSubItem.Value
                                    Case "Workstations"
                                        workstations = xSubItem.Value
                                    Case "ActiveD3Directory"
                                        activeJobs = xSubItem.Value
                                    Case "ArchivedD3Directory"
                                        archivedJobs = xSubItem.Value

                                End Select
                            Next
                        End If

                        'Get the server PCs IP address
                        If xItem.Name = "ServerIP" Then
                            serverIP = xItem.Value
                        End If

                        'Get the db polling delay
                        If xItem.Name = "PollDelay" Then
                            pollDelay = CInt(xItem.Value)
                            pollDelayUnits = xItem.Attribute("units").Value
                        End If
                    Next

                Catch ex As Exception

                End Try





            End If



        End Sub





        ''' <summary>
        ''' Checks to see if client configuration file is present
        ''' </summary>
        ''' <param name="createIfNull">Set to true to create file if it is missing</param>
        ''' <returns></returns>
        Private Function CheckClientConfig(Optional ByVal createIfNull As Boolean = False) As Boolean
            Try
                If Directory.Exists(My.Settings.DBDirectory) = True Then
                    If File.Exists(My.Settings.DBDirectory & "ClientConfig.xml") = True Then
                        Return True
                    ElseIf createIfNull = True Then
                        Dim xClientConf As XElement = CreateNewClientConfig()
                        xClientConf.Save(My.Settings.DBDirectory & "ClientConfig.xml")
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Creates a default client configuration
        ''' </summary>
        ''' <returns>XElement containing the default client configuraiton</returns>
        Private Function CreateNewClientConfig() As XElement
            Dim xRoot As New XElement("Root")
            Dim xmlLocations As New XElement("XML_Locations")

            'Add the config version to the configuration document. Used to check the validity of the configuration file.
            xRoot.Add(New XElement("ProgramVersion", Application.ProductVersion.ToString))

            'Add the default xml locations. This also happens to be the database structure
            xmlLocations.Add(New XElement("RootDBDirectory", My.Settings.DBDirectory))          'Root directory of the XML filestorage
            xmlLocations.Add(New XElement("D3_List_Active", "D3_List_Active.xml"))              'Active D3s list
            xmlLocations.Add(New XElement("Personnel", "Personnel.xml"))                        'Personnel Records for the database
            xmlLocations.Add(New XElement("Workstations", "Workstations.xml"))                  'The workstation information
            xmlLocations.Add(New XElement("ActiveD3Directory", "ActiveD3s\"))                   'The repository of currently active D3s
            xmlLocations.Add(New XElement("ArchivedD3Directory", "ArchivedD3s\"))               'The repository of currently inactive D3s
            xmlLocations.Add(New XElement("PeerList", "PeerList.xml"))                          'List of all connected computers on the program's network

            'Add the locations to the root settings
            xRoot.Add(xmlLocations)

            'Add the container for the ServerIP
            xRoot.Add(New XElement("ServerIP"))

            'Add the poll delay
            Dim xPoll As New XElement("PollDelay", "15")
            xPoll.SetAttributeValue("units", "sec")
            xRoot.Add(xPoll)


            Return xRoot


        End Function

    End Class

End Class
