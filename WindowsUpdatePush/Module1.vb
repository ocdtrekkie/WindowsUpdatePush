﻿Imports WUApiLib
Imports System.Security.Principal

Module Module1
    Function InstallationResultToText(result)
        Select Case result
            Case 2
                InstallationResultToText = "Succeeded"
            Case 3
                InstallationResultToText = "Succeeded with errors"
            Case 4
                InstallationResultToText = "Failed"
            Case 5
                InstallationResultToText = "Cancelled"
            Case Else
                InstallationResultToText = "Unexpected (" & result & ")"
        End Select
    End Function

    Sub Main()
        Dim serverSelection As Integer = 0
        Dim serviceId As String = ""

        Dim arguments As String() = Environment.GetCommandLineArgs()
        If arguments.Length > 1 Then
            Select Case arguments(1)
                Case "-microsoft"
                    Console.WriteLine("Microsoft Update source selected.")
                    serverSelection = 3
                    serviceId = "7971f918-a847-4430-9279-4a52d1efe18d"
            End Select
        End If

        'https://docs.microsoft.com/en-us/windows/win32/wua_sdk/searching--downloading--and-installing-updates
        Dim winUpdateSession As New UpdateSession()
        winUpdateSession.ClientApplicationID = "WindowsUpdatePush"
        Dim winUpdateSearcher As IUpdateSearcher = winUpdateSession.CreateUpdateSearcher()
        winUpdateSearcher.ServerSelection = serverSelection
        If serviceId <> "" Then
            winUpdateSearcher.ServiceID = serviceId
        End If

        Try
            Console.WriteLine("Checking for updates...")
            Dim winUpdateResult As ISearchResult = winUpdateSearcher.Search("IsInstalled=0 and Type='Software' And IsHidden=0")
            Dim winUpdatesToDownload As New UpdateCollection
            Console.WriteLine("Found " & winUpdateResult.Updates.Count & " updates" & Environment.NewLine)
            For Each winUpdate As IUpdate In winUpdateResult.Updates
                If winUpdate.InstallationBehavior.CanRequestUserInput = True Then
                    Console.WriteLine("Skipping " & winUpdate.Title & " because it requires user input")
                Else
                    If winUpdate.EulaAccepted = False Then
                        Console.WriteLine("Accepting EULA for " & winUpdate.Title)
                        winUpdate.AcceptEula()
                    Else
                        Console.WriteLine(winUpdate.Title)
                    End If
                    winUpdatesToDownload.Add(winUpdate)
                End If
            Next

            'https://stackoverflow.com/a/49115285
            Dim principal = New WindowsPrincipal(WindowsIdentity.GetCurrent())
            If principal.IsInRole(WindowsBuiltInRole.Administrator) Then
                Console.WriteLine("Script is running with elevation")
            Else
                Console.WriteLine("Script not running with elevation")
                Environment.ExitCode = 5
                Exit Sub
            End If

            If winUpdatesToDownload.Count > 0 Then
                Console.WriteLine("Downloading updates...")
                Dim winUpdateDownloader As UpdateDownloader = winUpdateSession.CreateUpdateDownloader()
                winUpdateDownloader.Updates = winUpdatesToDownload
                winUpdateDownloader.Download()

                Dim winRebootMayBeRequired = False
                Dim winUpdatesToInstall As New UpdateCollection
                For Each winUpdate As IUpdate In winUpdatesToDownload
                    If winUpdate.IsDownloaded = True Then
                        winUpdatesToInstall.Add(winUpdate)
                        If winUpdate.InstallationBehavior.RebootBehavior > 0 Then
                            winRebootMayBeRequired = True
                        End If
                    End If
                Next

                If winUpdatesToInstall.Count > 0 Then
                    Console.WriteLine("Installing updates...")
                    Dim winUpdateInstaller As UpdateInstaller = winUpdateSession.CreateUpdateInstaller()
                    winUpdateInstaller.Updates = winUpdatesToInstall
                    Dim winUpdateInstallResult As IInstallationResult = winUpdateInstaller.Install()
                    Console.WriteLine("Installation result: " & InstallationResultToText(winUpdateInstallResult.ResultCode))
                    Console.WriteLine("Reboot required: " & winUpdateInstallResult.RebootRequired)
                    If winUpdateInstallResult.RebootRequired = True Then
                        Environment.ExitCode = 3010
                        Exit Sub
                    End If
                End If
            End If
            Environment.ExitCode = 0
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Environment.ExitCode = 31
        End Try
    End Sub

End Module
