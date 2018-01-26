Imports System.IO
Imports System.Xml

''' <summary>
''' Performs all public and private logging of issues
''' </summary>
Public Class Log
#Region "Usage Logging"





#End Region







#Region "Error Logging"

    ''' <summary>
    ''' Sends an error log to the output locations noted in the settings
    ''' </summary>
    ''' <param name="ex">The exception that triggered this subroutine.</param>
    ''' <param name="returnMsg">Optional error message box to be returned to user interface.</param>
    ''' <param name="programComments">Additional comments to be included in the log.</param>
    Public Sub SendErrorLog(ByRef ex As Exception, ByVal privateDir As String, Optional ByVal returnMsg As Boolean = False, Optional ByVal programComments As String = "nothing", ByVal Optional publicDir As String = "nothing")

        'First, create the error log item
        Dim xLog As XElement = CreateErrorLogEntry(ex, programComments)

        'Return the error msg if asked to do so
        If returnMsg = True Then
            MsgBox(xLog.Value.ToString)
        End If

        'Now export the log to the desired locations
        Try
            If CheckPrivateDir(True) = True Then
                Dim xInternalLog As XElement
                If File.Exists(privateDir & "ErrorLogs.xml") = True Then
                    xInternalLog = XElement.Load(privateDir & "ErrorLogs.xml")
                Else
                    xInternalLog = New XElement("Root")
                End If

                xInternalLog.Add(xLog)

                xInternalLog.Save(publicDir & "ErrorLogs.xml")

            End If

            'Now try to add the external log
            If publicDir = "nothing" Then Exit Sub      'Exit the subroutine early if the public dir hasn't been provided

            If CheckPubDir(publicDir) = True Then

                Dim xExternalLog As XElement
                If File.Exists(publicDir & "ErrorLogs.xml") = True Then
                    xExternalLog = XElement.Load(publicDir & "ErrorLogs.xml")
                Else
                    xExternalLog = New XElement("Root")
                End If

                xExternalLog.Add(xLog)

                xExternalLog.Save(publicDir & "ErrorLogs.xml")
            End If

        Catch ex2 As Exception
            MsgBox(ex.Message.ToString)
        End Try



    End Sub


    ''' <summary>
    ''' Checks to see if public logging is activated
    ''' </summary>
    ''' <returns>True if public logging is activated and directory exists, false otherwise.</returns>
    ''' <param name="logDir">The directory path in which the public log should reside</param>
    Private Function CheckPubDir(ByVal logDir As String) As Boolean
        Try

            If Directory.Exists(logDir) = True Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Checks for the existence of the private directory. Creates it if allowed
    ''' </summary>
    ''' <returns>True if directory exists or was created. False otherwise</returns>
    ''' <param name="createIfNull">Creates the directory if it doesn't exist and this parameter is set to true.</param>
    ''' <param name="logDir">The directory in which the private logs should reside</param>
    Private Function CheckPrivateDir(ByVal logDir As String, Optional ByVal createIfNull As Boolean = False) As Boolean
        Try
            If Directory.Exists(logDir) = True Then
                Return True
            ElseIf createIfNull = True Then
                Directory.CreateDirectory(logDir)

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try


    End Function

    ''' <summary>
    ''' Creates an entry for an error log.
    ''' </summary>
    ''' <param name="ex">The exception that caused the error log.</param>
    ''' <param name="programComments">Any additional comments to include in the program.</param>
    ''' <returns>An XElement containing the log information (does not contain root element)</returns>
    Private Function CreateErrorLogEntry(ByRef ex As Exception, Optional ByVal programComments As String = "nothing") As XElement
        'Create the Xelement
        Dim tmp As New XElement("Entry")

        tmp.SetAttributeValue("Hresult", ex.HResult.ToString)
        tmp.SetAttributeValue("TimeStamp", Date.Now.ToString)
        tmp.SetAttributeValue("TargetSite", ex.TargetSite.ToString)

        If programComments = "nothing" Then
            tmp.Value = ex.Message.ToString
        Else
            tmp.Value = programComments & " : " & ex.Message.ToString
        End If

        Return tmp
    End Function
#End Region

End Class
