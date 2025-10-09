' Title: Change ModelState in assembly/subassembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/change-modelstate-in-assembly-subassembly/td-p/13775448#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:04:22.801049

public string CreateSimplifyModel(AssemblyDocument assemblyDocument, string folder)
{
    try
    {
        var modelName = System.IO.Path.GetFileNameWithoutExtension(assemblyDocument.FullFileName);
        var simplifyPartPath = System.IO.Path.Combine("C:\\Vault\\Projekty\\URZĄDZENIA STAL POWIETRZE\\OGRZEWACZE POM. STAL KOZY\\BJORN", "Uproszczenia", modelName + ".ipt");

        PartDocument oPartDoc = (PartDocument)_mInv.Documents.Add(DocumentTypeEnum.kPartDocumentObject, "", false);
        PartComponentDefinition oPartDef = oPartDoc.ComponentDefinition;

        DerivedAssemblyDefinition oDerivedAssemblyDef =
            oPartDef.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(assemblyDocument.FullDocumentName);

   
        oDerivedAssemblyDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams;

        oPartDef.ReferenceComponents.DerivedAssemblyComponents.Add(oDerivedAssemblyDef);

        object itemErpId = GetErpID(System.IO.Path.GetFileName(assemblyDocument.FullFileName));

        oPartDef.SurfaceBodies[1].Name = $"{itemErpId}";
        oPartDoc.SaveAs(simplifyPartPath, false);
        oPartDoc.Close(true);

        return simplifyPartPath;

    }
    catch (Exception ex)
    {
        throw new Exception("Błąd podczas generowania uproszczonego modelu.", ex);
    }
}

public void TraverseAssembly(ComponentOccurrences occurrences, int level)
{
    ComponentOccurrence occ = null;
    string oOccName = null;
    string splittedoOccName = null;
    try
    {
        foreach (ComponentOccurrence oOcc in occurrences)
        {
           
            oOccName = oOcc.Name;
            splittedoOccName = oOccName.Split(':')[0];
            if (nameCountDict.ContainsKey(splittedoOccName))
            {
                nameCountDict[splittedoOccName]++;
            }
            else
            {
                nameCountDict.Add(splittedoOccName, 1);
            }

            if (oOcc.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
            {

                string partName = System.IO.Path.GetFileName(oOcc.ReferencedDocumentDescriptor.FullDocumentName);

                object itemErpId = GetErpID(partName);

                oOcc.SurfaceBodies[1].Name = (string)itemErpId;
            }

            else if (oOcc.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject && !string.IsNullOrEmpty(oOcc.Name) && !oOcc.Name.Contains("Weldbead"))
            {
                AssemblyDocument assemblyDocumentOcc = oOcc.ReferencedFileDescriptor.ReferencedDocument as AssemblyDocument;

                string assemblyName = System.IO.Path.GetFileName(assemblyDocumentOcc.FullDocumentName);
                string folder = System.IO.Path.GetDirectoryName(assemblyDocumentOcc.FullDocumentName);
                var modelName = System.IO.Path.GetFileNameWithoutExtension(assemblyDocumentOcc.FullFileName);
                var simplyfyModelPath = System.IO.Path.Combine("C:\\Vault\\Projekty\\URZĄDZENIA STAL POWIETRZE\\OGRZEWACZE POM. STAL KOZY\\BJORN", "Uproszczenia", modelName + ".ipt");

                if (assemblyDocumentOcc.SubType == DocumentSubTypeIds.Weldment)
                {

                    if (System.IO.File.Exists(simplyfyModelPath))
                    {
                        continue;
                    }

                    simplyfyModelPath = CreateSimplifyModel(assemblyDocumentOcc, folder);
                    object itemErpId = GetErpID(assemblyName);

                    ComponentDefinition oOccCompDef = oOcc.Definition;
                    LevelOfDetailRepresentation levelOfDetail = null;

                    if (oOccCompDef.Type == ObjectTypeEnum.kWeldmentComponentDefinitionObject)
                    {
                        WeldmentComponentDefinition oOccWeldCompDef = (WeldmentComponentDefinition)oOccCompDef;
                        levelOfDetail = oOccWeldCompDef.RepresentationsManager.LevelOfDetailRepresentations.AddSubstitute(simplyfyModelPath, $"{itemErpId.ToString()}", false);
                    }
                    else
                    {
                        AssemblyComponentDefinition oOccAssemblyCompDef = (AssemblyComponentDefinition)oOccCompDef;
                        levelOfDetail = oOccAssemblyCompDef.RepresentationsManager.LevelOfDetailRepresentations.AddSubstitute(simplyfyModelPath, $"{itemErpId.ToString()}", false);
                    }


                    oOcc.ActiveModelState = $"{itemErpId.ToString()}";

                    Debug.WriteLine(oOcc.ActiveModelState);
                }
                else
                {
                    object itemErpId = GetErpID(System.IO.Path.GetFileName(oOcc.ReferencedFileDescriptor.FullFileName));
                    oOcc.Name = $"{itemErpId}:{nameCountDict[splittedoOccName]}";

                    if (oOcc.BOMStructure != Inventor.BOMStructureEnum.kPurchasedBOMStructure)
                    {
                        TraverseAssembly(oOcc.Definition.Occurrences, level + 1);
                    }
                }
            }
        }
    }
    catch(Exception ex)
    {
        throw new Exception("Wystapił błąd podczas przeszukiwania modelu.");
    }
}