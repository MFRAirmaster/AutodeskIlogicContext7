' Title: Delete ImportedComponent from part environment
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/delete-importedcomponent-from-part-environment/td-p/12156512
' Category: api
' Scraped: 2025-10-09T08:52:05.117601

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