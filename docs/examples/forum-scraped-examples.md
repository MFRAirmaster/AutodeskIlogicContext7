# iLogic Forum Examples

This document contains code examples scraped from the Autodesk Inventor Programming Forum. These are real-world examples from community members solving practical problems.

## Table of Contents
- [Measuring Face Area](#measuring-face-area)
- [SelectSet with MakePath](#selectset-with-makepath)
- [ControlDefinition Execute](#controldefinition-execute)
- [3D Sketch Between Work Points](#3d-sketch-between-work-points)

---

## Measuring Face Area

**Source:** [Forum Post](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-to-measure-area-of-two-faces/td-p/13831748)

**Problem:** Measuring the area of named faces and tracking changes as sketches are modified.

**Code Example:**

```vb
Dim iLogicAuto = iLogicVb.Automation

Dim namedEntities = iLogicAuto.GetNamedEntities(ThisDoc.Document)

Dim f As Face
f = namedEntities.FindEntity(COLD_OPENING), (HOT_OPENING)
Dim Eval As SurfaceEvaluator
Dim area As Double
Eval = f.Evaluator
area = Eval.Area
MsgBox("Face_Area : " & area)
```

**Concepts Used:**
- `iLogicVb.Automation` - Access to iLogic automation features
- `GetNamedEntities()` - Retrieve named entities from the document
- `FindEntity()` - Find a specific named entity
- `SurfaceEvaluator` - Evaluate surface properties
- `Evaluator.Area` - Get the surface area of a face

**Use Cases:**
- Tracking surface area changes during design iterations
- Calculating coating or finishing requirements
- Quality control and validation

---

## SelectSet with MakePath

**Source:** [Forum Post](https://forums.autodesk.com/t5/inventor-programming-forum/utilizing-makepath-in-selectset-select/td-p/13835725)

**Problem:** Selecting nested components in an assembly using MakePath.

**Code Examples:**

```vb
' Selecting a nested component
oDoc.SelectSet.Select(Component.InventorComponent(MakePath("75430:2", "75424:1")))

' Selecting a first-level component
oDoc.SelectSet.Select(Component.InventorComponent("75430:2"))
```

**Concepts Used:**
- `SelectSet.Select()` - Add objects to the selection set
- `Component.InventorComponent()` - Reference a component
- `MakePath()` - Create a path to nested components using occurrence names

**Use Cases:**
- Automating selection of deeply nested components
- Batch operations on specific assembly components
- Programmatic component manipulation

**Notes:**
- MakePath requires occurrence names separated by commas
- First-level components can be referenced directly without MakePath

---

## ControlDefinition Execute

**Source:** [Forum Post](https://forums.autodesk.com/t5/inventor-programming-forum/controldefinition-execute2-does-not-run-unless-it-is-the-last/td-p/13838361)

**Problem:** Executing Vault Revision Table update command programmatically.

**Code Example:**

```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(True)
MessageBox.Show("Updated Revision Table")
```

**Alternative (working in triggers):**

```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(False)
```

**Concepts Used:**
- `ThisApplication.CommandManager` - Access to command manager
- `ControlDefinitions.Item()` - Get a specific command definition
- `Execute2()` - Execute a command programmatically

**Use Cases:**
- Automating Vault operations
- Triggering built-in Inventor commands from iLogic
- Integration with PLM/PDM systems

**Important Notes:**
- `Execute2(True)` runs the command and waits for completion
- `Execute2(False)` runs the command asynchronously
- When used in triggers, code after Execute2(True) may not execute as expected
- Consider using Execute2(False) and placing it as the last command

---

## 3D Sketch Between Work Points

**Source:** [Forum Post](https://forums.autodesk.com/t5/inventor-programming-forum/3d-sketch-line-between-two-ucs-center-points/td-p/13839610)

**Problem:** Creating a 3D sketch line between two UCS center points automatically, without manual selection.

**Code Example (with manual selection):**

```vb
Dim doc As PartDocument = ThisDoc.Document
Dim def As PartComponentDefinition = doc.ComponentDefinition

Dim wpA1 = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kWorkPointFilter, "Select work point A1")
Dim wpB1 = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kWorkPointFilter, "Select work point B1")

Dim sketch3D As Sketch3D = def.Sketches3D.Add()
sketch3D.SketchLines3D.AddByTwoPoints(wpA1, wpB1)
sketch3D.Name = ("Strut 01")
```

**Concepts Used:**
- `PartDocument` and `PartComponentDefinition` - Document access
- `CommandManager.Pick()` - Interactive selection with filtering
- `SelectionFilterEnum.kWorkPointFilter` - Filter for work points
- `Sketches3D.Add()` - Create a new 3D sketch
- `SketchLines3D.AddByTwoPoints()` - Create a line between two points
- `Sketch.Name` - Assign a name to the sketch

**Use Cases:**
- Creating structural elements between reference points
- Automated frame generation
- Parametric strut/beam creation

**Improvement Needed:**
- Replace manual Pick() with automatic retrieval of named UCS points
- Use WorkPoint collection with named entity retrieval

---

## Additional Resources

For more iLogic examples and documentation, see:
- [Core Concepts](../core-concepts/01-basic-syntax.md)
- [Common Patterns](../common-patterns/01-parameter-manipulation.md)
- [Complete Parametric Part Example](complete-parametric-part.md)

---

*Last Updated: 2025-01-07*
*Forum examples may require adaptation for your specific use case*
