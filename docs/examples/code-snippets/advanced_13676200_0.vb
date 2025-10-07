' Title: Create Finish in Assembly with multiple Occurrence of one part - ERROR
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-finish-in-assembly-with-multiple-occurrence-of-one-part/td-p/13676200
' Category: advanced
' Scraped: 2025-10-07T13:14:51.629081

private void CreateFinish()
 {
     Asset oAppearance = null;
     try
     {
         string appearanceName = "Orange"; 
         AssemblyDocument assembly = iApp.ActiveDocument as AssemblyDocument;

         AssetLibrary assetLib = iApp.AssetLibraries["314DE259-5443-4621-BFBD-1730C6CC9AE9"];
         try
         {
             oAppearance = assembly.AppearanceAssets[appearanceName];
         }
         catch
         {
             oAppearance = assetLib.AppearanceAssets[appearanceName].CopyTo(assembly, true);
         }


         ObjectCollection oSurfaceBodyCollection = iApp.TransientObjects.CreateObjectCollection();
         foreach (ComponentOccurrence occ in assembly.ComponentDefinition.Occurrences)
         {
             foreach (SurfaceBodyProxy body in occ.SurfaceBodies)
             {
                 oSurfaceBodyCollection.Add(body);
             }
             //break; // For testing only the first occurrence
         }

         FinishFeatures oFinishFeatures = assembly.ComponentDefinition.Features.FinishFeatures;
         ObjectCollection oExcludedEntities = iApp.TransientObjects.CreateObjectCollection();

         FinishDefinition oFinishDef = oFinishFeatures.CreateDefinition(
             oSurfaceBodyCollection,
             oExcludedEntities,
             FinishTypeEnum.kAppearanceFinishType,
             "Prime/Paint RAL 2008",
             oAppearance);

         // Create finish feature.
         FinishFeature oFinish = oFinishFeatures.Add(oFinishDef);
         oFinish.Name = "Prime/Paint RAL 2008";
         oFinish.Suppressed = true;         
     }
     catch (Exception e)
     {
         MessageBox.Show("Create Finish faild" + System.Environment.NewLine + e.Message,
             "Error",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error);
         return;
     }
 }