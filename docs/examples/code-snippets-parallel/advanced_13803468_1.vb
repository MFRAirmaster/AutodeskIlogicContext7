' Title: SurfaceBody Appearance Override ignored?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/surfacebody-appearance-override-ignored/td-p/13803468#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:02:12.589921

foreach ((string Name, SurfaceBody SrfBody, Matrix Transformation) item in srfBodies)
                {
                    SurfaceBody srfBdy = item.SrfBody;
                    if (srfBdy.AppearanceSourceType == AppearanceSourceTypeEnum.kPartAppearance)
                    {
                        PartComponentDefinition partDef = (PartComponentDefinition)srfBdy.Parent;
                        PartDocument partDoc = (PartDocument)partDef.Document;                        
                        srfBdy.Appearance = partDoc.ActiveAppearance;
                    }
                    GraphicsNode node = cg.AddNode(nodeID++);
                    SurfaceGraphics srfG = node.AddSurfaceGraphics(srfBdy);                                     
                    node.Transformation = item.Transformation;
                }                
                invApp.ActiveView.Update();