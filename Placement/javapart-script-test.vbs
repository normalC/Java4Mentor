'Acceptable Usage Policy
'
'  This software is NOT officially supported by Mentor Graphics.
'
'  ####################################################################
'  ####################################################################
'  ## The following  software  is  "freeware" which  Mentor Graphics ##
'  ## Corporation  provides as a courtesy  to our users.  "freeware" ##
'  ## is provided  "as is" and  Mentor  Graphics makes no warranties ##
'  ## with  respect  to "freeware",  either  expressed  or  implied, ##
'  ## including any implied warranties of merchantability or fitness ##
'  ## for a particular purpose.                                      ##
'  ####################################################################
'  ####################################################################
'

Option Explicit
Dim App 'As MGCPCB.Application
Dim doc 'As MGCPCB.Document
Dim fsoo 'As FileSystemObject
Dim fa 'As TextStream
Dim FileName
Dim WshShell 'As Workbench.ShellCmd

Main

Public Sub Main()

    ' Compute the placement density for top and bottom of the PCB based on fromtos .
    '

    ' get expedition document
    'Set App = GetObject(, "MGCPCB.Application")
    Set App = Application
    Set doc = App.activeDocument

    If ValidateDocument(doc) = 0 Then
        MsgBox "Document Validation failed."
    End If

    doc.currentUnit = epcbUnitMils

    '################################################################
    'Start out with all the Netlines unselected
    doc.FromTos.Selected = False

    'Lock the server for blazing speed
    App.LockServer (True)
    
    '#################### Start Getting Data ########################

    Dim size
    Dim nt 'As Net
    Dim nts 'As Nets
    Dim objnet, objnetold
    Dim jobn, jobm, aatk_env, bx, by, i, xa, ya


    Dim todayDate, todayTime
    
    Set fsoo = CreateObject("Scripting.FileSystemObject")
    jobn = doc.Path

    jobm = Replace(jobn, "\", "/")

    aatk_env = Scripting.GetEnvVariable("AATK")
    fsoo.CopyFile aatk_env & "\Java\Placement\Placement.jar", jobn , True

    size = 25
    
    todayDate = Date
    todayTime = Time

    If fsoo.FileExists(jobn  & "Placement.ini") Then
        fsoo.DeleteFile jobn  & "Placement.ini"
    End If
        
    Set fa = fsoo.CreateTextFile(jobn  & "Placement.ini", True)
    'Set fa = fsoo.CreateTextFile("c:\WDIR\OTAKI-AATK\Java\Placement\Placement.ini", True)
    
    bx = CInt((doc.BoardOutline.Extrema.MaxX - doc.BoardOutline.Extrema.MinX) / size)
    by = CInt((doc.BoardOutline.Extrema.MaxY - doc.BoardOutline.Extrema.MinY) / size)
    
    fa.WriteLine ("#Autogenerated by javapart-script.vbs application")
    fa.WriteLine ("#" & todayDate & " " & todayTime)
    ' fa.WriteLine ("appOutdir=" & """" & jobn & "Config/xyplace.dat" & """")
    fa.WriteLine ("appOutdir=" & jobm & "Config/xyplace.dat")
    
    fa.WriteLine ("appWidth=1000") ' & bx * 2)
    fa.WriteLine ("appHeight=800") ' & by * 2)
    fa.Write ("edges=")
    
    Set nts = doc.Nets(epcbSelectAll)

    'For Each nt In nts

    Dim netsColl 'As Nets
    Dim netobj 'As Net
    Dim obj, obja
    Dim x, y, cnt, z
    Dim connectivity
    
    ' get all the nets in the design
    'Set tracesColl = doc.Traces
    'tracesColl.Sort
    Set netsColl = doc.Nets
    netsColl.Sort

    ' Update status bar
    Call App.Gui.StatusBarText("Writing " & netsColl.Count & " nets", epcbStatusField1)
    x = 0

    For i = netsColl.Count To 1 Step -1

        ' Update status bar
        Call App.Gui.StatusBarText("Writing " & i & " nets", epcbStatusField1)
    
        ' get the first trace
        Set netobj = netsColl.Item(i)
        objnet = netobj.Name
    
        objnetold = netobj.Name
       
        'Set connectivity = tracesColl(i).ConnectedObjects
        Set connectivity = netsColl(i).Pins
        ' for each thing connected to this trace
        cnt = connectivity.Count
        
        If cnt > 1 And connectivity.Count < 25 Then

            For z = 1 To cnt - 1
                ' get connected object
                Set obj = connectivity.Item(z)
                Set obja = connectivity.Item(z + 1)

                ' if it's a padstack object (pin/via/fiducial/mountinghole)
                If obj.ObjectClass = epcbObjectClassPadstackObject Then

                    ' if it's specifically a pin then select it
                    If obj.Type = epcbPadstackObjectPin Then

                        '## If it is a virtual pin don't list it.
                        'If (obj.Component.Name <> obj.Name) Then
                        
                        x  = CInt((obj.Component.Extrema.MaxX  - obj.Component.Extrema.MinX)  / size)
                        y  = CInt((obj.Component.Extrema.MaxY  - obj.Component.Extrema.MinY)  / size)
                        xa = CInt((obja.Component.Extrema.MaxX - obja.Component.Extrema.MinX) / size)
                        ya = CInt((obja.Component.Extrema.MaxY - obja.Component.Extrema.MinY) / size)
                            
                        If (obj.Component.Name < obja.Component.Name) Then
                            fa.Write (obj.Component.Name & "^" & x & "^" & y & "-" & obja.Component.Name & "^" & xa & "^" & ya & "/40,")
                        Else
                            fa.Write (obja.Component.Name & "^" & xa & "^" & ya & "-" & obj.Component.Name & "^" & x & "^" & y & "/40,")
                        End If
                            
                        'End If

                    End If
                End If

            Next

        End If

        '        If i = 180 Then
        '            fa.WriteLine ("")
        '        End If

    Next

    fa.WriteLine ("appImplicitINI=Yes")
    fa.WriteLine ("appTitle=Placement")
    fa.WriteLine ("appS=" & size)
    'fa.WriteLine ("<param name=center value=""" & obj.Component.Name & """>")
    fa.WriteLine ("appStatusbar=Yes")

    '*************************** End network ***********************
            
    '
    fa.Close
  
    'Unlock the server before it screws you up
    App.UnlockServer
        
    FileName = jobn & "Placement.jar"
    'MsgBox FileName
    Set WshShell = CreateObject("WScript.Shell")
    'WshShell.WorkingDirectory jobn & "output"
    
    WshShell.Run FileName

End Sub

'------------------------------------------------------------------------------

Private Function ValidateDocument(doc)
    Dim key, licenseServer, licenseToken

    ' Ask Expedition's document for the key
    key = doc.Validate(0)

    ' Get license server
    Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application")

    ' Ask the license server for the license token
    licenseToken = licenseServer.GetToken(key)

    ' Release license server
    Set licenseServer = Nothing

    ' Turn off error messages.  Validate may fail if the token is incorrect
    On Error Resume Next
    Err.Clear

    ' Ask the document to validate the license token
    doc.Validate (licenseToken)

    If Err.Number <> 0 Then
        ValidateDocument = 0
    Else
        ValidateDocument = 1
    End If

End Function

