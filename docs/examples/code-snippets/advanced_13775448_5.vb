' Title: Change ModelState in assembly/subassembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/change-modelstate-in-assembly-subassembly/td-p/13775448#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:40:19.940101

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
          itemErpId = GetErpID(assemblyName);

          ComponentDefinition oOccCompDef = oOcc.Definition;
          ModelState modelState = null;

          if (oOccCompDef.Type == ObjectTypeEnum.kWeldmentComponentDefinitionObject)
          {
              WeldmentComponentDefinition oOccWeldCompDef = (WeldmentComponentDefinition)oOccCompDef;
              modelState = oOccWeldCompDef.ModelStates.AddSubstitute(simplyfyModelPath, $"{itemErpId.ToString()}", true);
          }
          else
          {
              AssemblyComponentDefinition oOccAssemblyCompDef = (AssemblyComponentDefinition)oOccCompDef;
              modelState = oOccAssemblyCompDef.ModelStates.AddSubstitute(simplyfyModelPath, $"{itemErpId.ToString()}", true);
          }

          oOcc.ActiveModelState = modelState.Name; // exception apear
          modelState.Activate(); 

          Debug.WriteLine(oOcc.ActiveModelState);
      }
      else
      {
          itemErpId = GetErpID(System.IO.Path.GetFileName(oOcc.ReferencedFileDescriptor.FullFileName));
          oOcc.Name = $"{itemErpId}:{nameCountDict[splittedoOccName]}";

          if (oOcc.BOMStructure != Inventor.BOMStructureEnum.kPurchasedBOMStructure)
          {
              TraverseAssembly(oOcc.Definition.Occurrences, level + 1);
          }
      }
  }