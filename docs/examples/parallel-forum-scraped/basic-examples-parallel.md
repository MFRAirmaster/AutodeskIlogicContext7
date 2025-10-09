# Parallel Basic iLogic Examples

Generated from 6 forum posts with advanced parallel scraping.

**Generated:** 2025-10-09T09:10:07.546310

---

## View &quot;Dim&quot; Variable Values

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/view-quot-dim-quot-variable-values/td-p/13332453#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/view-quot-dim-quot-variable-values/td-p/13332453#messageview_0)

**Author:** Joseph.Lasser

**Date:** ‎02-21-2025
	
		
		07:40 AM

**Description:** This is likely a poor question, but is there a a way to view a list of all declared variables in a rule and their associated values?  For example, please reference the script below:Dim Test1 = 1
Dim Test2 = 2
Dim Test3 = 3
Dim Sum = Test1+Test2+Test3Within the rule editor, would there be a way to see that the variable Sum is equal to 6? 
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Dim Test1 = 1
Dim Test2 = 2
Dim Test3 = 3
Dim Sum = Test1+Test2+Test3
```

```vb
Dim Val1 As Double = 1.5
Dim Val2 As Double = 3
Dim Val3 As Double = 5.5
Dim values As New List(Of Double) From {Val1, Val2, Val3 }
Dim sum As Double = values.Sum
Logger.Info("Sum = " & sum.ToString())
```

```vb
Dim Val1 As Double = 1.5
Dim Val2 As Double = 3
Dim Val3 As Double = 5.5
Dim values As New List(Of Double) From {Val1, Val2, Val3 }
Dim sum As Double = values.Sum
Logger.Info("Sum = " & sum.ToString())
```

---

## 3D-Pdf export to same location as 3D-file

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/3d-pdf-export-to-same-location-as-3d-file/td-p/13757662#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/3d-pdf-export-to-same-location-as-3d-file/td-p/13757662#messageview_0)

**Author:** Kupiszewski_A_

**Date:** ‎08-07-2025
	
		
		02:17 AM

**Description:** Hello everyone, I'm trying to write an ilogic rule that creates a 3D PDF and saves it in the same folder (or subfolder) as the original 3D part. I've tried many things, but it just doesn't work.  Can someone please help me? Thank you very much!
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Dim ExportFileName = ThisDoc.ChangeExtension(".pdf")
```

```vb
Dim ExportFileName = oFolder & ThisDoc.FileName & ".pdf"
```

---

## Greater than Less Than

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/greater-than-less-than/td-p/13754306](https://forums.autodesk.com/t5/inventor-programming-forum/greater-than-less-than/td-p/13754306)

**Author:** layochim

**Date:** ‎08-05-2025
	
		
		04:34 AM

**Description:** Hello, I'm trying to create a statement for my conveyor heights.I have two heights Infeed_End and Discharge_End. I'm trying to write "If Discharge_End > 115 And Infeed_End < 25 Then Discharge_End = 114 And Infeed_End = 26" I'm having no luck on this.  Can anyone help me on this?
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
If (Discharge_End > 115) And (Infeed_End < 25) Then 
	Discharge_End = 114 
	Infeed_End  = 26
End If
```

```vb
If ((Discharge_End > 115) And (Infeed_End < 25) ) Or (Belt_Length = 200) Then 
	Discharge_End = 114 
	Infeed_End  = 26
End If
```

---

## Greater than Less Than

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/greater-than-less-than/td-p/13754306#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/greater-than-less-than/td-p/13754306#messageview_0)

**Author:** layochim

**Date:** ‎08-05-2025
	
		
		04:34 AM

**Description:** Hello, I'm trying to create a statement for my conveyor heights.I have two heights Infeed_End and Discharge_End. I'm trying to write "If Discharge_End > 115 And Infeed_End < 25 Then Discharge_End = 114 And Infeed_End = 26" I'm having no luck on this.  Can anyone help me on this?
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
If (Discharge_End > 115) And (Infeed_End < 25) Then 
	Discharge_End = 114 
	Infeed_End  = 26
End If
```

```vb
If ((Discharge_End > 115) And (Infeed_End < 25) ) Or (Belt_Length = 200) Then 
	Discharge_End = 114 
	Infeed_End  = 26
End If
```

---

## View &quot;Dim&quot; Variable Values

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/view-quot-dim-quot-variable-values/td-p/13332453](https://forums.autodesk.com/t5/inventor-programming-forum/view-quot-dim-quot-variable-values/td-p/13332453)

**Author:** Joseph.Lasser

**Date:** ‎02-21-2025
	
		
		07:40 AM

**Description:** This is likely a poor question, but is there a a way to view a list of all declared variables in a rule and their associated values?  For example, please reference the script below:Dim Test1 = 1
Dim Test2 = 2
Dim Test3 = 3
Dim Sum = Test1+Test2+Test3Within the rule editor, would there be a way to see that the variable Sum is equal to 6? 
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Dim Test1 = 1
Dim Test2 = 2
Dim Test3 = 3
Dim Sum = Test1+Test2+Test3
```

```vb
Dim Val1 As Double = 1.5
Dim Val2 As Double = 3
Dim Val3 As Double = 5.5
Dim values As New List(Of Double) From {Val1, Val2, Val3 }
Dim sum As Double = values.Sum
Logger.Info("Sum = " & sum.ToString())
```

```vb
Dim Val1 As Double = 1.5
Dim Val2 As Double = 3
Dim Val3 As Double = 5.5
Dim values As New List(Of Double) From {Val1, Val2, Val3 }
Dim sum As Double = values.Sum
Logger.Info("Sum = " & sum.ToString())
```

---

## 3D-Pdf export to same location as 3D-file

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/3d-pdf-export-to-same-location-as-3d-file/td-p/13757662](https://forums.autodesk.com/t5/inventor-programming-forum/3d-pdf-export-to-same-location-as-3d-file/td-p/13757662)

**Author:** Kupiszewski_A_

**Date:** ‎08-07-2025
	
		
		02:17 AM

**Description:** Hello everyone, I'm trying to write an ilogic rule that creates a 3D PDF and saves it in the same folder (or subfolder) as the original 3D part. I've tried many things, but it just doesn't work.  Can someone please help me? Thank you very much!
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Dim ExportFileName = ThisDoc.ChangeExtension(".pdf")
```

```vb
Dim ExportFileName = oFolder & ThisDoc.FileName & ".pdf"
```

---

