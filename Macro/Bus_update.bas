Attribute VB_Name = "Bus_update"
Sub Bus_update1()

    Dim cheminFichier As String
    Dim nomFichierCATIA As String

    cheminFichier = ThisWorkbook.Path
    
    'BUS ROOF PLATE
    nomFichierCATIA = "Avionics_Unit\Bus_bottom_plate.CATPart"

    openfolder = cheminFichier & "\" & nomFichierCATIA
    Debug.Print openfolder
    Set CATIA = GetObject(, "CATIA.Application")
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    Set partDocumentbus1 = CATIA.ActiveDocument
    Set Partbus1 = partDocumentbus1.part
    Set Parametersbus1 = Partbus1.Parameters
    
    'Attachment
    
    Parametersbus1.Item("Bus_length").Value = Worksheets("Parameters").Range("D29").Value
    Parametersbus1.Item("Bus_width").Value = Worksheets("Parameters").Range("D30").Value
    Parametersbus1.Item("Bus_depth").Value = Worksheets("Parameters").Range("D31").Value
    Parametersbus1.Item("Bus_thickness").Value = Worksheets("Parameters").Range("D32").Value
    Parametersbus1.Item("Bus_fixing_screw_hole_dia").Value = Worksheets("Parameters").Range("D34").Value
    Parametersbus1.Item("Bus_screw_length").Value = Worksheets("Parameters").Range("D35").Value
    Parametersbus1.Item("Bus_screw_dia").Value = Worksheets("Parameters").Range("D33").Value
    
    Partbus1.Update
    CATDoc.Save
    
    'partDocumentbus1.Close
    
    'BUS CONNECTOR WALL
    
    nomFichierCATIA = "Avionics_Unit\Bus_connector_wall.CATPart"

    openfolder = cheminFichier & "\" & nomFichierCATIA
    
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    Set partDocumentbus2 = CATIA.ActiveDocument
    Set Partbus2 = partDocumentbus2.part
    Set Parametersbus2 = Partbus2.Parameters
    
    
    'Attachment
    
    Parametersbus2.Item("Bus_length").Value = Worksheets("Parameters").Range("D29").Value
    Parametersbus2.Item("Bus_width").Value = Worksheets("Parameters").Range("D30").Value
    
    Parametersbus2.Item("Bus_thickness").Value = Worksheets("Parameters").Range("D32").Value
    Parametersbus2.Item("Bus_fixing_screw_hole_dia").Value = Worksheets("Parameters").Range("D34").Value
    Parametersbus2.Item("Bus_screw_length").Value = Worksheets("Parameters").Range("D35").Value
      
    
    Partbus2.Update
    CATDoc.Save
    'partDocumentbus2.Close
    
    'BUS FRONT PLATE
    
    nomFichierCATIA = "Avionics_Unit\Bus_front_plate.CATPart"

    openfolder = cheminFichier & "\" & nomFichierCATIA
    
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    Set partDocumentbus3 = CATIA.ActiveDocument
    Set Partbus3 = partDocumentbus3.part
    Set Parametersbus3 = Partbus3.Parameters
    
    
    'Attachment
    
    Parametersbus3.Item("Bus_width").Value = Worksheets("Parameters").Range("D30").Value
    Parametersbus3.Item("Bus_depth").Value = Worksheets("Parameters").Range("D31").Value
    Parametersbus3.Item("Bus_thickness").Value = Worksheets("Parameters").Range("D32").Value
    Parametersbus3.Item("Bus_fixing_screw_hole_dia").Value = Worksheets("Parameters").Range("D34").Value
  
    Partbus3.Update
    CATDoc.Save
    
    
    
    'BUS INTERNAL PAYLOAD
    
    nomFichierCATIA = "Avionics_Unit\Bus_internal_payload.CATPart"

    openfolder = cheminFichier & "\" & nomFichierCATIA
    
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    Set partDocumentbus4 = CATIA.ActiveDocument
    Set Partbus4 = partDocumentbus4.part
    Set Parametersbus4 = Partbus4.Parameters
    
    'Attachment
    
    Parametersbus4.Item("Bus_pay_length").Value = Worksheets("Parameters").Range("K29").Value
    Parametersbus4.Item("Bus_pay_width").Value = Worksheets("Parameters").Range("K30").Value
    Parametersbus4.Item("Bus_pay_depth").Value = Worksheets("Parameters").Range("K31").Value
    
    Partbus4.Update
    CATDoc.Save
    
    


    'BUS Assembly
    
    nomFichierCATIA = "Avionics_Unit\Avionics_unit.CATProduct"

    openfolder = cheminFichier & "\" & nomFichierCATIA
    
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    Set ProductDocument = CATIA.ActiveDocument
    Set Product = ProductDocument.Product
    Set Parametersproduct = Product.Parameters


    Parametersproduct.Item("Bus_payload_X").Value = Worksheets("Parameters").Range("K35").Value
    Parametersproduct.Item("Bus_payload_Y").Value = Worksheets("Parameters").Range("K36").Value
    Parametersproduct.Item("Bus_payload_Z").Value = Worksheets("Parameters").Range("K37").Value
    
    Product.Update
    CATDoc.Save

End Sub
    
