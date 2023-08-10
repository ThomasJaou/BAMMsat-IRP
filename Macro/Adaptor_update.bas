Attribute VB_Name = "Adaptor_update"
Sub Adaptor_updtate1()

    Dim cheminFichier As String
    Dim nomFichierCATIA As String



    cheminFichier = ThisWorkbook.Path



    'PLATE
    nomFichierCATIA = "Adaptor_plate\Adaptor_plate.CATPart"

    openfolder = cheminFichier & "\" & nomFichierCATIA
    Set CATIA = GetObject(, "CATIA.Application")
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    Set partDocumentplate1 = CATIA.ActiveDocument
    Set Partplate1 = partDocumentplate1.part
    Set Parametersplate1 = Partplate1.Parameters
    
    Set bodies1 = Partplate1.Bodies
    Set body1 = bodies1.Item("PartBody")
    Set shapes1 = body1.Shapes
    
    Set rectPattern1 = shapes1.Item("RectPattern.2")
    Set rectPattern2 = shapes1.Item("RectPattern.3")
    Set rectPattern3 = shapes1.Item("RectPattern.5")
    If Worksheets("Parameters").Range("D65").Value = 2 Then
        Partplate1.Inactivate rectPattern2
    Else
        Partplate1.Activate rectPattern2
    End If
            
    If Worksheets("Parameters").Range("D64").Value = 2 Then
        Partplate1.Inactivate rectPattern1
        Partplate1.Inactivate rectPattern3
    Else
        Partplate1.Activate rectPattern1
        Partplate1.Activate rectPattern3
    End If
    
    
    
    'Plate dimmension updtate
    
    Parametersplate1.Item("Plate_length").Value = Worksheets("Parameters").Range("D56").Value
    Parametersplate1.Item("Plate_width").Value = Worksheets("Parameters").Range("D57").Value
    Parametersplate1.Item("Plate_thickness").Value = Worksheets("Parameters").Range("D58").Value
    Parametersplate1.Item("Plate_fixing_position_length").Value = Worksheets("Parameters").Range("D59").Value
    Parametersplate1.Item("Plate_fixing_position_width").Value = Worksheets("Parameters").Range("D60").Value
        
    Parametersplate1.Item("Bus_length").Value = Worksheets("Parameters").Range("D29").Value
    Parametersplate1.Item("Bus_depth").Value = Worksheets("Parameters").Range("D31").Value
    
    Parametersplate1.Item("PV_length").Value = Worksheets("Parameters").Range("D8").Value
    Parametersplate1.Item("PV_depth").Value = Worksheets("Parameters").Range("D10").Value
    Parametersplate1.Item("PV_contact_surface").Value = Worksheets("Parameters").Range("D12").Value
    Parametersplate1.Item("PV_lid_thickness").Value = Worksheets("Parameters").Range("D13").Value
    
    
    Parametersplate1.Item("Stud_positioning").Value = Worksheets("Parameters").Range("D63").Value
    Parametersplate1.Item("distance_stud_length").Value = Worksheets("Parameters").Range("D66").Value
    Parametersplate1.Item("distance_stud_depth").Value = Worksheets("Parameters").Range("D67").Value
     
    Parametersplate1.Item("Nb_of_stud_length").Value = Worksheets("Parameters").Range("D64").Value
    Parametersplate1.Item("Nb_of_stud_depth").Value = Worksheets("Parameters").Range("D65").Value
    Parametersplate1.Item("Stud_screwplate_dia").Value = Worksheets("Parameters").Range("D50").Value
    
    
    Parametersplate1.Item("PV_stif").Value = Worksheets("Parameters").Range("D22").Value
    
    
    Partplate1.Update
    CATDoc.Save
    
    'partDocumentplate1.Close
    
    
    
    'STUD
    nomFichierCATIA = "Adaptor_plate\Stud.CATPart"

    openfolder = cheminFichier & "\" & nomFichierCATIA
    
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    Set partDocumentplate2 = CATIA.ActiveDocument
    Set Partplate2 = partDocumentplate2.part
    Set Parametersplate2 = Partplate2.Parameters
    
    Parametersplate2.Item("Stud_length").Value = Worksheets("Parameters").Range("D43").Value
    Parametersplate2.Item("Stud_dia").Value = Worksheets("Parameters").Range("D44").Value
    
    Parametersplate2.Item("Stud_screwpay_length").Value = Worksheets("Parameters").Range("D45").Value
    Parametersplate2.Item("Stud_screwpay_dia").Value = Worksheets("Parameters").Range("D47").Value
    
    Parametersplate2.Item("Stud_screwplate_length").Value = Worksheets("Parameters").Range("D48").Value
    Parametersplate2.Item("Stud_screwplate_dia").Value = Worksheets("Parameters").Range("D50").Value
    
    Partplate2.Update
    CATDoc.Save
    
    
    
End Sub
