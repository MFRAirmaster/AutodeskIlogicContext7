' Title: Enable Add-in in Inventor Nesting
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/enable-add-in-in-inventor-nesting/td-p/13759956
' Category: advanced
' Scraped: 2025-10-07T14:02:32.270071

// Example: enable button only in Nesting environment
bool inNesting = environment.DisplayName.Equals("Nesting", StringComparison.OrdinalIgnoreCase);
bool isAssemblyDocument = m_inventorApplication.ActiveDocument?.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject;