' Title: Issues with SurfaceBody.CalculateFacets in Inventor 2025 Add-in (.NET 8 / x64)
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/issues-with-surfacebody-calculatefacets-in-inventor-2025-add-in/td-p/13736696
' Category: advanced
' Scraped: 2025-10-07T12:31:42.484949

Dim vertexCount As Long
Dim facetCount As Long
Dim vertexCoords() As Double
Dim normalVectors() As Double
Dim vertexIndices() As Long

Call oBody.CalculateFacets(tolerance, vertexCount, facetCount, vertexCoords, normalVectors, vertexIndices)