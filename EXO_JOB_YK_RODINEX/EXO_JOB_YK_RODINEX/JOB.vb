Imports System.Data.SqlClient
Module JOB
    Public Sub Main()
        Dim iCountExeJOB As Integer = 0
        Dim sIDMax As String = "0"
        Dim oLog As EXO_Log.EXO_Log = Nothing
        Dim sError As String = ""
        Dim oFiles() As String = Nothing
        Dim sPath As String = ""

        Try
            sPath = My.Application.Info.DirectoryPath.ToString
            If System.IO.Directory.Exists(sPath & "\Logs") Then
                oFiles = System.IO.Directory.GetFileSystemEntries(sPath & "\Logs")
                For Each sFile As String In oFiles
                    System.IO.File.Delete(System.IO.Path.GetFullPath(sFile))
                Next
            Else
                System.IO.Directory.CreateDirectory(sPath & "\Logs")
            End If
            oLog = New EXO_Log.EXO_Log(sPath & "\Logs\logWS_", 50, EXO_Log.EXO_Log.Nivel.todos, 4, "", EXO_Log.EXO_Log.GestionFichero.dia)
            If Procesos.ActivaProceso("PROC1") = True Then
                    oLog.escribeMensaje("-- Procedimiento Artículos --", EXO_Log.EXO_Log.Tipo.informacion)
                    Procesos.Articulos(sPath)
                    oLog.escribeMensaje("-- Fin Procedimiento Artículos --", EXO_Log.EXO_Log.Tipo.informacion)
                End If

                oLog.escribeMensaje("-- Procedimiento Stock --", EXO_Log.EXO_Log.Tipo.informacion)
                'hago una parada de 10 segundos
                System.Threading.Thread.Sleep(10000)
                Procesos.Stock(sPath)
            oLog.escribeMensaje("-- Fin Procedimiento Stock --", EXO_Log.EXO_Log.Tipo.informacion)
        Catch ex As Exception
            If ex.InnerException IsNot Nothing AndAlso ex.InnerException.Message <> "" Then
                sError = ex.InnerException.Message
            Else
                sError = ex.Message
            End If
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        End Try
    End Sub
End Module
