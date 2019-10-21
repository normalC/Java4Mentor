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
Attribute VB_Name = "Module1"

Dim App As MGCPCB.Application
Dim doc As MGCPCB.Document

'Main

Public Sub Main()

    ' Compute the placement density for top and bottom of the PCB based on fromtos .
    '

    ' get expedition document
    Set App = GetObject(, "MGCPCB.Application")
    'Set App = Application
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
    Dim nt As Net
    Dim nts As Nets
    Dim jobn
    Dim bx, by
    Dim fsoo As FileSystemObject
    Dim fa As TextStream
    Set fsoo = CreateObject("Scripting.FileSystemObject")
    jobn = doc.Path
    size = 25

    'If fsoo.FileExists(jobn & "config\" & "place.cvs") Then
    '    fsoo.DeleteFile jobn & "config\" & "place.csv"
    'End If
        
    'Set fa = fsoo.OpenTextFile(jobn & "output\" & "place.csv", 8, True)
    
    If fsoo.FileExists("C:\Documents and Settings\kendall\My Documents\My Website\Placement\bin\place.html") Then
        fsoo.DeleteFile "C:\Documents and Settings\kendall\My Documents\My Website\Placement\bin\place.html"
    End If
        
    Set fa = fsoo.CreateTextFile("C:\Users\kendall.MGC\Documents\My Website\Placement\bin\place.html", True)
    'fa.WriteLine ("From,TO,count")
    
    bx = CInt((doc.BoardOutline.Extrema.MaxX - doc.BoardOutline.Extrema.MinX) / size)
    by = CInt((doc.BoardOutline.Extrema.MaxY - doc.BoardOutline.Extrema.MinY) / size)
    
    fa.WriteLine ("<title>Graph Layout</title>")
    fa.WriteLine ("<hr>")
    fa.WriteLine ("<applet code=""Placement.class"" width=" & bx * 2 & " height=" & by * 2 & ">")
    fa.WriteLine ("<param name=edges value=""")
    
    Set nts = doc.Nets(epcbSelectAll)

    'For Each nt In nts

    Dim netsColl As Nets
    Dim netobj As Net
    Dim obj, obja
    Dim x, y, ax, ay, c, cnt, z
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

        If (objnet = "CLKC") Then
            MsgBox objnet
        End If

        'If (objnet <> objnetold) Then
        'If (objnetold <> "firstobject") Then
                    
        '    fa.Writeline ("")
        'End If

        'If (objnet <> "(Net0)") Then
        'fa.WriteLine ("net =, " & objnet)
        'fa.Write (" pins")
        'End If
        'End If
    
        objnetold = netobj.Name
    
        ' and highlight it
        'outlinet.Highlighted = True
    
        ' get all things directly connected to this first trace
        'tracesColl = netsColl(i).traces
        'For z = tracesColl.Count To 1 Step -1
        
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
                        
                        x = CInt((obj.Component.Extrema.MaxX - obj.Component.Extrema.MinX) / size)
                        y = CInt((obj.Component.Extrema.MaxY - obj.Component.Extrema.MinY) / size)
                        xa = CInt((obja.Component.Extrema.MaxX - obja.Component.Extrema.MinX) / size)
                        ya = CInt((obja.Component.Extrema.MaxY - obja.Component.Extrema.MinY) / size)
                            
                        If (obj.Component.Name < obja.Component.Name) Then
                            fa.WriteLine (obj.Component.Name & "^" & x & "^" & y & "-" & obja.Component.Name & "^" & xa & "^" & ya & "/50,")
                        Else
                            fa.WriteLine (obja.Component.Name & "^" & xa & "^" & ya & "-" & obj.Component.Name & "^" & x & "^" & y & "/50,")
                        End If
                            
                        'End If

                        'If (c = 20) Then
                        '    fa.WriteLine ("")
                        '    fa.Write ("   ")
                        '    c = 0
                        'End If

                        'c = c + 1

                    End If
                End If

            Next

        End If

        '        If i = 180 Then
        '            fa.WriteLine ("")
        '        End If

    Next

    fa.WriteLine ("")
    fa.WriteLine (""">")
    fa.WriteLine ("<param name=s=""" & size & """>")
    'fa.WriteLine ("<param name=center value=""" & obj.Component.Name & """>")
    fa.WriteLine ("</applet>")
    fa.WriteLine ("<hr>")

    '*************************** End network ***********************
            
    'fa.WriteLine (fobj & "," & tobj & ",1")
    'End If

    'Debug.Print ft
    'Next
  
    'Unlock the server before it screws you up
    App.UnlockServer

End Sub

Sub boa()
    ' get the area of placement on this board layer in square inches
    'areaCu = maskCu.Area(emeUnitInch)
    
    ' get board outline and find its area
    Set bo = doc.BoardOutline
   
    pnts = bo.Geometry.PointsArray
    Set maskBO = mskeng.Masks.Add
    Call maskBO.shapes.AddByPointsArray(1 + UBound(pnts, 2), pnts)
    areaBO = maskBO.Area(emeUnitInch)
    
    ' compute the copper density on this layer as a percentage
    'MsgBox "Placement Density for layer " & i & " = " & FormatNumber(areaCu * 100 / areaBO, 1) & "%"
    'MsgBox ("Board Placement Area " & areaBO)
    'Next
    sel = areaBO
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

Sub extra_crap()
    ' repeat for all the other copper data on this layer and add to maskA
    '########'Set clgs = doc.conductorLayerGfxs(epcbCondGfxConductiveShape + epcbCondGfxPlaneShape + epcbCondGfxTeardrop + epcbCondGfxCopperBalancingData, , lyr)
    '########'For Each clg In clgs
    '########'        pnts = clg.Geometry.PointsArray
    '########'        If 1 + UBound(pnts, 2) > 2 Then
    '########'                Call shapesCu.AddByPointsArray(1 + UBound(pnts, 2), pnts)
    '########'        End If
    '########'Next
        
    ' get all the pads on this layer and add to maskA
    '########'Set pos = doc.padstackobjects
    '########'For Each po In pos
    '########'        Set pds = po.pads(lyr)
    '########'        For Each pd In pds
    '########'                Set gms = pd.geometries
    '########'                For Each gm In gms
    '########'                        pnts = gm.PointsArray
    '########'                        If 1 + UBound(pnts, 2) > 2 Then
    '########'                                Call shapesCu.AddByPointsArray(1 + UBound(pnts, 2), pnts)
    '########'                        End If
    '########'                Next
    '########'        Next
    '########'Next

End Sub
