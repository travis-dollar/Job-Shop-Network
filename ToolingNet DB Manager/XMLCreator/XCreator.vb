Imports System.Xml
Imports System.IO
Imports LogManager


Public Class XCreator

    ''' <summary>
    ''' Creates a new xelement from a specified template
    ''' </summary>
    ''' <param name="XMLDocType">The type of document to be created</param>
    ''' <returns>Xelement containing the specified template</returns>
    Public Function CreateFromTemplate(ByVal XMLDocType As XMLDocumentTypes.DocTypes) As XElement

        Dim xNull As New XElement("Root")   'An empty element to be returned in the event of an invalid constant selection

        Select Case XMLDocType
            Case XMLDocumentTypes.DocTypes.clientConfig
                Return CreateNewClientConfig()
            Case XMLDocumentTypes.DocTypes.activeD3List
                MsgBox("This Functionality not yet implemented")
            Case XMLDocumentTypes.DocTypes.PersonnelList
                MsgBox("This Functionality not yet implemented")
            Case XMLDocumentTypes.DocTypes.WorkstationList
                MsgBox("This Functionality not yet implemented")
        End Select

        Return xNull

    End Function





    ''' <summary>
    ''' The named types of xml documents that can be created
    ''' </summary>
    Public Class XMLDocumentTypes

        Enum DocTypes
            clientConfig = 1
            activeD3List = 2
            PersonnelList = 3
            WorkstationList = 4
        End Enum
    End Class





    ''' <summary>
    ''' Creates a default client configuration
    ''' </summary>
    ''' <returns>XElement containing the default client configuraiton</returns>
    Private Function CreateNewClientConfig() As XElement
        Dim xRoot As New XElement("Root")
        Dim xmlLocations As New XElement("XML_Locations")

        'Add the config version to the configuration document. Used to check the validity of the configuration file.
        xRoot.Add(New XElement("ProgramVersion", My.Application.Info.Version.ToString))

        'Add the default xml locations. This also happens to be the database structure
        xmlLocations.Add(New XElement("RootDBDirectory"))          'Root directory of the XML filestorage
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
