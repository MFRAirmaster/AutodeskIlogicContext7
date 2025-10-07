' Title: Drawing HoleThreadNote Formatted Text Bug Precision and Tolerance when not set.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/drawing-holethreadnote-formatted-text-bug-precision-and/td-p/13222899
' Category: advanced
' Scraped: 2025-10-07T13:28:21.001538

Sub Main

	Dim sel As SelectSet = Thisdoc.Document.SelectSet

	If sel.Count = 0 OrElse typeof sel.Item(1) IsNot HoleThreadNote
		MsgBox("Please select a HoleThreadNote and only a HoleThreadNote!")
		Exit Sub
	End If

	Dim note As HoleThreadNote = ThisDoc.Document.SelectSet.Item(1)
	Dim noteText = note.FormattedHoleThreadNote

	System.Windows.Forms.Clipboard.SetText(note.FormattedHoleThreadNote)

	If note.FormattedHoleThreadNote.Contains("CBoreDiameter")

		noteText = "<StyleOverride Font='RomanD' FontSize='0.254'><HoleProperty HolePropertyID='kCBoreDiameterHoleProperty' Precision='3' AlternatePrecision='3' SetTolerances='False'></HoleProperty></StyleOverride>"

	Else

		noteText = "<StyleOverride Font='RomanD' FontSize='0.254'><HoleProperty HolePropertyID='kHoleDiameterHoleProperty' Precision='3' AlternatePrecision='3' SetTolerances='False'></HoleProperty></StyleOverride>"

	End If

	note.FormattedHoleThreadNote = noteText

End Sub