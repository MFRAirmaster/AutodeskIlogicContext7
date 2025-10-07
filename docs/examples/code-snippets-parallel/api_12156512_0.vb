' Title: Delete ImportedComponent from part environment
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/delete-importedcomponent-from-part-environment/td-p/12156512#messageview_0
' Category: api
' Scraped: 2025-10-07T14:09:28.191732

if (_document != null && _document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument _oPart = _document as PartDocument;
                var oPartCompdef = _oPart.ComponentDefinition;                
                ImportedComponents impComps = oPartCompdef.ReferenceComponents.ImportedComponents;
                foreach (ImportedComponent CompOcc in impComps)
                {                   
                  string fileUrn = CompOcc.Name.Replace(".stp","");
                   if (exchangeItem.ExchangeID.Contains(fileUrn))
                   {
                       CompOcc.Delete();
                   }
                }

            }