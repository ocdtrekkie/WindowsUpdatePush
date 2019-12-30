Imports WUApiLib
Imports System.Security.Principal

Module Module1

    Sub Main()
        'https://docs.microsoft.com/en-us/windows/win32/wua_sdk/searching--downloading--and-installing-updates
        Dim winUpdateSession As New UpdateSession()
        winUpdateSession.ClientApplicationID = "WindowsUpdatePush"
        Dim winUpdateSearcher As IUpdateSearcher = winUpdateSession.CreateUpdateSearcher()

        Try
            Console.WriteLine("Checking for updates...")
            Dim winUpdateResult As ISearchResult = winUpdateSearcher.Search("IsInstalled=0 and Type='Software' And IsHidden=0")
            Dim winUpdatesToDownload As New UpdateCollection
            Console.WriteLine("Found " & winUpdateResult.Updates.Count & " updates" & Environment.NewLine)
            For Each winUpdate As IUpdate In winUpdateResult.Updates
                Console.WriteLine(winUpdate.Title)
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
                End
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
                    Console.WriteLine("Installation result: " & winUpdateInstallResult.ResultCode)
                    Console.WriteLine("Reboot required: " & winUpdateInstallResult.RebootRequired)
                    If winUpdateInstallResult.RebootRequired = True Then
                        Environment.ExitCode = 3010
                        End
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
