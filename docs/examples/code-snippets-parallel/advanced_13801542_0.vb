' Title: GLB export - Inventor server
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/glb-export-inventor-server/td-p/13801542
' Category: advanced
' Scraped: 2025-10-07T14:07:38.687959

public void TraverseAssembly(ComponentOccurrences occurrences)
        {
            List<ComponentOccurrences> nextLevel = new List<ComponentOccurrences>();

            foreach (ComponentOccurrence oOcc in occurrences)
            {
                string oOccName = oOcc.Name;
                oOcc.Visible = true;
                if (oOcc.Suppressed) { continue; }
                oOcc.IsAssociativeToDesignViewRepresentation = false;

                Document oOccDoc = (Document)oOcc.ReferencedDocumentDescriptor.ReferencedDocument;

                string oOccDocName = oOccDoc.DisplayName.Split('.').FirstOrDefault();
                string oOccDocSubType = oOccDoc.SubType;

                string modelName = System.IO.Path.GetFileNameWithoutExtension(oOcc.ReferencedFileDescriptor.FullFileName);
                string simplifyPath = System.IO.Path.Combine(_simplyfyFolder, modelName + ".ipt");
                string asmName = System.IO.Path.GetFileName(oOcc.ReferencedFileDescriptor.FullFileName);

                object erpId = VaultMethods.GetErpIdByPartName(asmName, _connection);
                if (erpId == null) continue;

                if (!nameCountDict.ContainsKey(erpId.ToString()))
                    nameCountDict[erpId.ToString()] = 0;

                nameCountDict[erpId.ToString()]++;
                int partCount = nameCountDict[erpId.ToString()];

                string material = GetComponentMaterial(oOccDoc);

                if (!string.IsNullOrEmpty(oOcc.Name) &&
                    !oOcc.Name.Contains("Weldbead"))
                {
                    if (oOccDocSubType == DocumentSubTypeIds.Assembly)
                    {
                        var oOccAseCompDef = oOcc.Definition as AssemblyComponentDefinition;
                        oOcc.Name = $"{erpId}:{partCount}";

                        if (oOcc.BOMStructure != Inventor.BOMStructureEnum.kPurchasedBOMStructure)
                        {
                            if (oOccAseCompDef.IsModelStateMember)
                            {
                                var factoryDocument = (AssemblyDocument)oOccAseCompDef.FactoryDocument;
                                factoryDocument.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditAllMembers;
                                nextLevel.Add(factoryDocument.ComponentDefinition.Occurrences);
                            }
                            else
                            {
                                nextLevel.Add(oOcc.Definition.Occurrences);
                            }
                        }
                        else if (oOcc.BOMStructure == Inventor.BOMStructureEnum.kPurchasedBOMStructure)
                        {
                            AssemblyDocument asmDoc = null;
                            AssemblyComponentDefinition oOccAsmCompDef = null;
                            if (oOccAseCompDef.IsModelStateMember)
                            {
                                asmDoc = (AssemblyDocument)oOccAseCompDef.FactoryDocument;
                            }
                            else
                            {
                                asmDoc = (AssemblyDocument)oOcc.ReferencedDocumentDescriptor.ReferencedDocument;
                            }
                            oOccAsmCompDef = asmDoc.ComponentDefinition;
                            string folder = System.IO.Path.GetDirectoryName(asmDoc.FullFileName);

                            if (!System.IO.File.Exists(simplifyPath))
                            {
                                simplifyPath = CreateSimplifyModel(asmDoc, folder);
                                ModelState newState = oOccAsmCompDef.ModelStates.AddSubstitute(simplifyPath, $"{erpId}", false);
                            }
                            oOcc.Name = $"{erpId}:{partCount}";
                            oOcc.ActiveModelState = $"{erpId}";
                        }
                    }
                    else if (oOccDocSubType == DocumentSubTypeIds.Weldment)
                    {
                        var oOccAseCompDef = oOcc.Definition as AssemblyComponentDefinition;
                        AssemblyDocument asmDoc = null;
                        AssemblyComponentDefinition oOccAsmCompDef = null;
                        if (oOccAseCompDef.IsModelStateMember)
                        {
                            asmDoc = (AssemblyDocument)oOccAseCompDef.FactoryDocument;
                        }
                        else
                        {
                            asmDoc = (AssemblyDocument)oOcc.ReferencedDocumentDescriptor.ReferencedDocument;
                        }
                        oOccAsmCompDef = asmDoc.ComponentDefinition;
                        string folder = System.IO.Path.GetDirectoryName(asmDoc.FullFileName);


                        if (!System.IO.File.Exists(simplifyPath))
                        {
                            simplifyPath = CreateSimplifyModel(asmDoc, folder);
                            ModelState newState = oOccAsmCompDef.ModelStates.AddSubstitute(simplifyPath, $"{erpId}", false);
                        }
                        oOcc.Name = $"{erpId}:{partCount}";
                        oOcc.ActiveModelState = $"{erpId}";
                    }
                    else
                    {
                        oOcc.Name = $"{erpId}:{partCount}";
                       
                    }
                }
            }

            foreach (var occs in nextLevel)
            {
                TraverseAssembly(occs);
            }

        }