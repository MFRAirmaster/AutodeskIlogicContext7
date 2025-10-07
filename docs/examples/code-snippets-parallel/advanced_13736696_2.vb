' Title: Issues with SurfaceBody.CalculateFacets in Inventor 2025 Add-in (.NET 8 / x64)
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/issues-with-surfacebody-calculatefacets-in-inventor-2025-add-in/td-p/13736696
' Category: advanced
' Scraped: 2025-10-07T13:55:13.820960

int vertexCount = 0;
int facetCount = 0;
double[] vertexCoords = null;
double[] normalVectors = null;
int[] vertexIndices = null;

body.CalculateFacets(tolerance, out vertexCount, out facetCount, out vertexCoords, out normalVectors, out vertexIndices);