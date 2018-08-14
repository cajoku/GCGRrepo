Attribute VB_Name = "AggregateData"
Private Const xmlnamespace As String = "TBD"
Private Const xmlversion As String = "TBD"


Private xmlDOC As DOMDocument60
Private xmlROOTelement As IXMLDOMElement
Private xmlPARENTelement As IXMLDOMElement
Private xmlGRANDPARENTelement As IXMLDOMElement
Private objXMLelement As IXMLDOMElement
Private objXMLattr As IXMLDOMAttribute
    
Private paCOLL As Collection
Private testCOLL As Collection
Private gcCOLL As Collection
Private grCOLL As Collection

Private projectNAME As String
Private projectLOCATION As String
Private opsjobNUM As String
Private pretaskNUM As String
Private docDATE As String
Private projectCOST As String
Private vacation As String
Private preconSTART As String
Private preconEND As String
Private conStart As String
Private conEnd As String
Private preconLIT As String
Private conLIT As String
Private preconRAISE As String
Private conRAISE As String
Private IT As String
Private gas As String
Private phone As String
Private units As String
Private area As String
Private preTOTAL As String
Private conTOTAL As String
Private gcTOTAL As String
Private grTOTAL As String
Private gcARR() As Variant
Private grARR() As Variant

Sub CostLabCSV()


    Dim csvWB As Workbook
    Dim beginRAN As Range
    Dim i As Integer, j As Double
    Dim pfrm As progressFRM

    j = 0.3
    Set pfrm = New progressFRM
    pfrm.progressON "CostLab Item Import", "Aggregating Line Items"
    
    grARR = detailITEMS(masterOBJ.grdWS)
    preTOTAL = masterOBJ.WS.[\prelabor]
    conTOTAL = masterOBJ.WS.[\conlabor]
    gcTOTAL = masterOBJ.WS.[\gctotal]
    
    pfrm.progressUPDATE "Creating CSV", 0
    
    Set csvWB = Application.Workbooks.Open(Filename:=itemImportFile, ReadOnly:=True)
    csvWB.Windows(1).Visible = False
    Set beginRAN = csvWB.Worksheets(1).[A1]
    
    Do Until beginRAN.Offset(0, 1).Value = ""
        Set beginRAN = beginRAN.Offset(1, 0)
    Loop
    
    For i = 0 To UBound(grARR, 1)
        beginRAN.Value = grARR(i, 0)
        beginRAN.Offset(0, 2).Value = grARR(i, 1)
        beginRAN.Offset(0, 3).Value = grARR(i, 2)
        beginRAN.Offset(0, 4).Value = grARR(i, 3)
        beginRAN.Offset(0, 12).Value = grARR(i, 4)
        Set beginRAN = beginRAN.Offset(1, 0)
        pfrm.progressUPDATE "Creating CSV", ((i + 1) / UBound(grARR, 1))
    Next
    
    With beginRAN
        .Value = "General Conditions"
        .Offset(0, 2).Value = 1
        .Offset(0, 3).Value = "lsum"
        .Offset(0, 4).Value = gcTOTAL
        .Offset(0, 12).Value = "98 00 00"
    End With
    Set beginRAN = beginRAN.Offset(1, 0)
    
    With beginRAN
        .Value = "Preonstruction Staffing"
        .Offset(0, 2).Value = 1
        .Offset(0, 3).Value = "lsum"
        .Offset(0, 4).Value = preTOTAL
        .Offset(0, 12).Value = "98 11 00"
    End With
    Set beginRAN = beginRAN.Offset(1, 0)
    
    With beginRAN
        .Value = "Construction Staffing"
        .Offset(0, 2).Value = 1
        .Offset(0, 3).Value = "lsum"
        .Offset(0, 4).Value = conTOTAL
        .Offset(0, 12).Value = "98 21 00"
    End With

    pfrm.turnOFF
    
    csvWB.Windows(1).Visible = True
    
End Sub
Sub createXML()
On Error GoTo e1

    Dim wb As Workbook
    Dim WS As Worksheet, sht As Worksheet
    Dim i As Integer, FileNo As Integer
    
    Set wb = masterWB
    Set WS = masterOBJ.WS
    
    Set paCOLL = New Collection
    Set testCOLL = New Collection
    Set gcCOLL = New Collection
    Set grCOLL = New Collection
    
    Set xmlDOC = New DOMDocument60
    
    Set xmlROOTelement = xmlDOC.createElement("TBD")
    xmlDOC.appendChild xmlROOTelement
    
    Set xmlGRANDPARENTelement = xmlDOC.createElement("ProjectMetrics")
    xmlROOTelement.appendChild xmlGRANDPARENTelement
    
    projectNAME = projectATTS(WS.[\proj])
    projectLOCATION = projectATTS(WS.[\loc])
    opsjobNUM = projectATTS(WS.[\ops])
    pretaskNUM = projectATTS(WS.[\task])
    docDATE = projectATTS(WS.[\date])
    projectCOST = projectATTS(WS.[\pcost])
    vacation = projectATTS(WS.[\vaca])
    preconSTART = projectATTS(WS.[\pstart])
    preconEND = projectATTS(WS.[\pend])
    conStart = projectATTS(WS.[\cstart])
    conEnd = projectATTS(WS.[\cend])
    preconLIT = projectATTS(WS.[\preLIT])
    conLIT = projectATTS(WS.[\conLIT])
    preconRAISE = projectATTS(WS.[\preRaise])
    conRAISE = projectATTS(WS.[\conRaise])
    IT = projectATTS(WS.[\IT])
    gas = projectATTS(WS.[\gas])
    phone = projectATTS(WS.[\iphone])
    units = projectATTS(WS.[\unit])
    area = projectATTS(WS.[\area])
    preTOTAL = projectATTS(WS.[\pretotal])
    conTOTAL = projectATTS(WS.[\contotal])
    gcTOTAL = projectATTS(WS.[\gcontotal])
    grTOTAL = projectATTS(WS.[\greqtotal])
    
    For i = 1 To paCOLL.Count
        Set objXMLelement = xmlDOC.createElement(paCOLL(i))
        objXMLelement.Text = testCOLL(i)
        xmlGRANDPARENTelement.appendChild objXMLelement
    Next
    
    gcARR = detailITEMS(masterOBJ.gcdWS)
    lineATTS gcARR, True
    
    grARR = detailITEMS(masterOBJ.grdWS)
    lineATTS grARR, False
    
    xmlDOC.Save DataLogFile
    
Exit Sub
e1:
    LogError "AggregateData", "createXML", Err.Description, Err
End Sub

Function detailITEMS(WS As Worksheet) As Variant

    
    Dim RAN As Range, cell As Range
    Dim groupRAN As Range, valRAN As Range, qtyRAN As Range, uomRAN As Range
    Dim itemCOLL As Collection, valCOLL As Collection, qtyCOLL As Collection, uomCOLL As Collection, ccCOLL As Collection
    Dim arr() As Variant
    Dim iCount As Integer, i As Integer
    
    Set RAN = boxRANGE(WS, "\r_start", "\r_end", "\c_desc")
    Set groupRAN = WS.[\c_group]
    Set valRAN = WS.[\c_val].Offset(0, -1)
    Set qtyRAN = WS.[\c_qt]
    Set uomRAN = WS.[\c_qt].Offset(0, 1)
    Set itemCOLL = New Collection
    Set valCOLL = New Collection
    Set qtyCOLL = New Collection
    Set uomCOLL = New Collection
    Set ccCOLL = New Collection
    
    For Each cell In RAN
        If Intersect(cell.EntireRow, groupRAN.EntireColumn).Value = 1 Then
            itemCOLL.Add cell.Value
            valCOLL.Add Intersect(cell.EntireRow, valRAN.EntireColumn).Value
            qtyCOLL.Add Intersect(cell.EntireRow, qtyRAN.EntireColumn).Value
            uomCOLL.Add Intersect(cell.EntireRow, uomRAN.EntireColumn).Value
            ccCOLL.Add Intersect(cell.EntireRow, groupRAN.Offset(0, 4).EntireColumn).Value
        End If
    Next

    iCount = itemCOLL.Count
    
    ReDim arr(iCount - 1, 4)
    
    For i = 0 To iCount - 1
        arr(i, 0) = itemCOLL(i + 1)
        arr(i, 1) = qtyCOLL(i + 1)
        arr(i, 2) = uomCOLL(i + 1)
        arr(i, 3) = valCOLL(i + 1)
        arr(i, 4) = ccCOLL(i + 1)
    Next
    
    detailITEMS = arr
    
End Function

Function projectATTS(attRAN As Range) As String

    
    Dim pOPEN As String, pCLOSE As String
    Dim attNAME As String, attVAL As String
    
    attNAME = Replace(attRAN.Offset(0, -1).Value, " ", "")
    attNAME = Replace(attNAME, "#", "")
    attNAME = Replace(attNAME, "/", "")
    attVAL = attRAN.Value
    
    pOPEN = "<" & attNAME & ">": pCLOSE = "</" & attNAME & ">"
    
    projectATTS = attVAL
    
    paCOLL.Add attNAME
    testCOLL.Add attVAL
    
    
    

End Function


Sub lineATTS(lineARR As Variant, GCITEMS As Boolean)

    Dim i As Integer
    
    If GCITEMS = True Then
        Set xmlGRANDPARENTelement = xmlDOC.createElement("GCItems")
    Else
        Set xmlGRANDPARENTelement = xmlDOC.createElement("GRItems")
    End If
    xmlROOTelement.appendChild xmlGRANDPARENTelement
    
    
    For i = 0 To UBound(lineARR, 1)
        Set xmlPARENTelement = xmlDOC.createElement("LineItem")
        xmlGRANDPARENTelement.appendChild xmlPARENTelement
    
        Set objXMLelement = xmlDOC.createElement("Name")
        objXMLelement.Text = lineARR(i, 0)
        xmlPARENTelement.appendChild objXMLelement
        
        Set objXMLelement = xmlDOC.createElement("Quantity")
        objXMLelement.Text = lineARR(i, 1)
        xmlPARENTelement.appendChild objXMLelement

        Set objXMLelement = xmlDOC.createElement("UnitOfMeasure")
        objXMLelement.Text = lineARR(i, 2)
        xmlPARENTelement.appendChild objXMLelement
        
        Set objXMLelement = xmlDOC.createElement("Value")
        objXMLelement.Text = lineARR(i, 3)
        xmlPARENTelement.appendChild objXMLelement
        
        Set objXMLelement = xmlDOC.createElement("CostCode")
        objXMLelement.Text = lineARR(i, 4)
        xmlPARENTelement.appendChild objXMLelement
        
    Next
    
    'Debug.Print lineATTS

End Sub


Sub testXML()

    Dim yarrr As MXXMLWriter60
    Dim xmlDOC As DOMDocument
    Dim xmlROOTelement As IXMLDOMElement
    Dim objXMLelement As IXMLDOMElement
    Dim objXMLattr As IXMLDOMAttribute
   
    Set xmlDOC = New DOMDocument60
   
    Set xmlROOTelement = xmlDOC.createElement("LineItem")
    xmlDOC.appendChild xmlROOTelement
    
    'set objxmlelement = obj.createelement(
    


End Sub

Private Sub testing()
'   Dim xmlDOC As DOMDocument
'   Dim xmlROOTelement As IXMLDOMElement
'   Dim objXMLelement As IXMLDOMElement
'   Dim objXMLattr As IXMLDOMAttribute

   Set xmlDOC = New DOMDocument60
   
   '~~> Creates root element
   Set xmlROOTelement = xmlDOC.createElement("Entry")
   xmlDOC.appendChild xmlROOTelement
   
   '~~> Create Date element
   Set objXMLelement = xmlDOC.createElement("Date")
   objXMLelement.Text = Now
   xmlROOTelement.appendChild objXMLelement
   
   '~~> Creates Attribute to the Date Element and set value
'   Set objXMLattr = xmlDOC.createAttribute("Value")
'   objXMLattr.NodeValue = "3/2/2012"
'   objXMLelement.setAttributeNode objXMLattr

   '~~> Create Time element
   Set objXMLelement = xmlDOC.createElement("Time")
   objXMLelement.Text = "colby o clock'"
   xmlROOTelement.appendChild objXMLelement
   
   '~~> Creates Attribute to the Time Element and set value
'   Set objXMLattr = xmlDOC.createAttribute("Value")
'   objXMLattr.NodeValue = "12 PM"
'   objXMLelement.setAttributeNode objXMLattr
   
   '~~> Creates Name element
   Set objXMLelement = xmlDOC.createElement("Name")
   objXMLelement.Text = "The Truth"
   xmlROOTelement.appendChild objXMLelement
   

   xmlDOC.Save ("C:\Users\CAjoku\Desktop\trash.xml")
End Sub

