Attribute VB_Name = "PV_update"
Sub ModifierParametresCATIA()
     
    
    'Chemin d'accès du fichier
    Dim cheminFichier As String
    Dim nomFichierCATIA As String
    
    cheminFichier = ThisWorkbook.Path
    nomFichierCATIA = "External_shell.CATPart"
    
    openfolder = cheminFichier & "\" & nomFichierCATIA
    
    'Creation de l'instance
    
    Set CATIA = GetObject(, "CATIA.Application")
    
    'Ouvrir le fichier CATIA
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    If Not CATDoc Is Nothing Then
        If TypeName(CATDoc) = "PartDocument" Then
                       
                   
            Set partDocument1 = CATIA.ActiveDocument
            Set part1 = partDocument1.part
            Set bodies1 = part1.Bodies
            Set body1 = bodies1.Item("PartBody")
            Set shapes1 = body1.Shapes
            Set rectPattern1 = shapes1.Item("RectPattern.16")
            Set rectPattern2 = shapes1.Item("RectPattern.18")
            Set Parameters1 = part1.Parameters
            
            If Worksheets("Parameters").Range("D65").Value = 2 Then
                part1.Inactivate rectPattern1
            Else
                part1.Activate rectPattern1
            End If
            
            If Worksheets("Parameters").Range("D64").Value = 2 Then
                part1.Inactivate rectPattern2
            Else
                part1.Activate rectPattern2
            End If
            
            
            'Volume parameters
            Parameters1.Item("PV_length").Value = Worksheets("Parameters").Range("D8").Value
            Parameters1.Item("PV_width").Value = Worksheets("Parameters").Range("D9").Value
            Parameters1.Item("PV_depth").Value = Worksheets("Parameters").Range("D10").Value
            Parameters1.Item("PV_wall").Value = Worksheets("Parameters").Range("D11").Value
            Parameters1.Item("PV_contact_surface").Value = Worksheets("Parameters").Range("D12").Value
            
            Parameters1.Item("PV_lid_thickness").Value = Worksheets("Parameters").Range("D13").Value
            Parameters1.Item("Nb_of_stud_length").Value = Worksheets("Parameters").Range("D64").Value
            Parameters1.Item("Nb_of_stud_depth").Value = Worksheets("Parameters").Range("D65").Value
            Parameters1.Item("Stud_positioning").Value = Worksheets("Parameters").Range("D63").Value
            Parameters1.Item("distance_stud_length").Value = Worksheets("Parameters").Range("D66").Value
            Parameters1.Item("distance_stud_depth").Value = Worksheets("Parameters").Range("D67").Value
            
            Parameters1.Item("Stud_screwpay_length").Value = Worksheets("Parameters").Range("D45").Value
            Parameters1.Item("Stud_screwpay_dia").Value = Worksheets("Parameters").Range("D47").Value
            Parameters1.Item("Stud_screwpay_dia_hole").Value = Worksheets("Parameters").Range("D46").Value
            
            'Screw parameters
            Parameters1.Item("PV_Screw_positioning").Value = Worksheets("Parameters").Range("D14").Value
            Parameters1.Item("PV_nb_of_Screw_depth").Value = Worksheets("Parameters").Range("D15").Value
            Parameters1.Item("PV_distance_depth").Value = Worksheets("Parameters").Range("D16").Value
            Parameters1.Item("PV_nb_of_Screw_width").Value = Worksheets("Parameters").Range("D17").Value
            Parameters1.Item("PV_distance_width").Value = Worksheets("Parameters").Range("D18").Value
            Parameters1.Item("PV_Screw_dia").Value = Worksheets("Parameters").Range("D19").Value
            Parameters1.Item("PV_Screw_length").Value = Worksheets("Parameters").Range("D21").Value
            
            Parameters1.Item("PV_stif").Value = Worksheets("Parameters").Range("D22").Value
            
            part1.Update
            CATDoc.Save
        End If
        
       
    End If
    
    
    
    'LID UPDATE
    nomFichierCATIA = "lid.CATPart"
    openfolder = cheminFichier & "\" & nomFichierCATIA
    
    
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    
    Set partDocument2 = CATIA.ActiveDocument
    Set Part2 = partDocument2.part
    Set Parameters2 = Part2.Parameters
    
            Parameters2.Item("PV_width").Value = Worksheets("Parameters").Range("D9").Value
            Parameters2.Item("PV_depth").Value = Worksheets("Parameters").Range("D10").Value
            Parameters2.Item("PV_wall").Value = Worksheets("Parameters").Range("D11").Value
            Parameters2.Item("PV_contact_surface").Value = Worksheets("Parameters").Range("D12").Value
            Parameters2.Item("PV_lid_thickness").Value = Worksheets("Parameters").Range("D13").Value
            
            
    
            Parameters2.Item("PV_Screw_positioning").Value = Worksheets("Parameters").Range("D14").Value
            Parameters2.Item("PV_nb_of_Screw_depth").Value = Worksheets("Parameters").Range("D15").Value
            Parameters2.Item("PV_distance_depth").Value = Worksheets("Parameters").Range("D16").Value
            Parameters2.Item("PV_nb_of_Screw_width").Value = Worksheets("Parameters").Range("D17").Value
            Parameters2.Item("PV_distance_width").Value = Worksheets("Parameters").Range("D18").Value
            Parameters2.Item("PV_Screw_hole_dia").Value = Worksheets("Parameters").Range("D20").Value
    
    
            Part2.Update
            CATDoc.Save
            
            
            
    'LID UPDATE
    nomFichierCATIA = "lid_EI.CATPart"
    openfolder = cheminFichier & "\" & nomFichierCATIA
    
    
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    
    Set partDocument2 = CATIA.ActiveDocument
    Set Part2 = partDocument2.part
    Set Parameters2 = Part2.Parameters
    
            Parameters2.Item("PV_width").Value = Worksheets("Parameters").Range("D9").Value
            Parameters2.Item("PV_depth").Value = Worksheets("Parameters").Range("D10").Value
            Parameters2.Item("PV_wall").Value = Worksheets("Parameters").Range("D11").Value
            Parameters2.Item("PV_contact_surface").Value = Worksheets("Parameters").Range("D12").Value
            Parameters2.Item("PV_lid_thickness").Value = Worksheets("Parameters").Range("D13").Value
            
            
    
            Parameters2.Item("PV_Screw_positioning").Value = Worksheets("Parameters").Range("D14").Value
            Parameters2.Item("PV_nb_of_Screw_depth").Value = Worksheets("Parameters").Range("D15").Value
            Parameters2.Item("PV_distance_depth").Value = Worksheets("Parameters").Range("D16").Value
            Parameters2.Item("PV_nb_of_Screw_width").Value = Worksheets("Parameters").Range("D17").Value
            Parameters2.Item("PV_distance_width").Value = Worksheets("Parameters").Range("D18").Value
            Parameters2.Item("PV_Screw_hole_dia").Value = Worksheets("Parameters").Range("D20").Value
            
            Parameters2.Item("PV_EI_X").Value = Worksheets("Parameters").Range("K43").Value
            Parameters2.Item("PV_EI_Y").Value = Worksheets("Parameters").Range("K44").Value
            Parameters2.Item("PV_EI_r").Value = Worksheets("Parameters").Range("K45").Value
            
    
            Part2.Update
            CATDoc.Save
            
    'Opening the Pyaload CATPart
    nomFichierCATIA = "Internal_payload.CATPart"
    openfolder = cheminFichier & "\" & nomFichierCATIA
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    'Accesing to the parameters
    Set partDocument3 = CATIA.ActiveDocument
    Set Part3 = partDocument3.part
    Set Parameters3 = Part3.Parameters
        
    'Updating the value of the parameters
        Parameters3.Item("pay_length").Value = Worksheets("Parameters").Range("K8").Value
        Parameters3.Item("pay_width").Value = Worksheets("Parameters").Range("K9").Value
        Parameters3.Item("pay_depth").Value = Worksheets("Parameters").Range("K10").Value
            
    'Updating the Mass volumic
    
        Part3.Update
        CATDoc.Save
        
        
    'Update the Bus unit
    
    
    Call Bus_update1
    Call Adaptor_updtate1
        
    'Modification if the Assembly
    nomFichierCATIA = "BAMMSat_assembly.CATProduct"
    openfolder = cheminFichier & "\" & nomFichierCATIA
    
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
    Set ProductDocument1 = CATIA.ActiveDocument
    Set Product1 = ProductDocument1.Product
    Set Parameters_product = Product1.Parameters
    
        Parameters_product.Item("pay_X").Value = Worksheets("Parameters").Range("K16").Value
        Parameters_product.Item("pay_Y").Value = Worksheets("Parameters").Range("K17").Value
        Parameters_product.Item("pay_Z").Value = Worksheets("Parameters").Range("K18").Value
    
        Product1.Update
        CATDoc.Save
        
    
    'MsgBox "All documents have been succefully updated"
    
    Set CATIA = GetObject(, "CATIA.Application")
    
    For Each doc In CATIA.Documents
        doc.Close
        
    Next doc
    
    
    nomFichierCATIA = "BAMMSat_assembly.CATProduct"
    openfolder = cheminFichier & "\" & nomFichierCATIA
    
    Set CATDoc = CATIA.Documents.Open(openfolder)
    
End Sub
