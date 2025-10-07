# Autodesk Inventor iLogic Programming Documentation

Welcome to the comprehensive iLogic programming reference for Autodesk Inventor. This documentation is designed to help developers, engineers, and automation specialists master iLogic programming for design automation and parametric modeling.

## What is iLogic?

iLogic is a rule-based programming environment built into Autodesk Inventor that allows you to automate design tasks, create intelligent parametric models, and build custom workflows using Visual Basic .NET syntax.

## Documentation Structure

### 📚 [Core Concepts](./core-concepts/)
Fundamental iLogic concepts including:
- Basic syntax and structure
- Parameters and properties
- Rules and external rules
- Event triggers
- Variable scope and data types

### 🔧 [Common Patterns](./common-patterns/)
Frequently used programming patterns:
- Parameter manipulation
- Property access and modification
- Geometry operations
- File and assembly operations
- BOM (Bill of Materials) manipulation
- User interactions and dialogs

### 📖 [API Reference](./api-reference/)
Detailed API documentation:
- Document objects
- Component operations
- Parameter objects
- Property sets
- Geometry API
- File I/O operations

### 🔍 [Troubleshooting](./troubleshooting/)
Common issues and solutions:
- Error messages and fixes
- Performance optimization
- Debugging techniques
- Common pitfalls

### ✅ [Best Practices](./best-practices/)
Professional coding standards:
- Code organization
- Error handling
- Performance optimization
- Maintainability
- Documentation standards

### 💡 [Examples](./examples/)
Real-world code examples:
- Simple automation tasks
- Complex parametric designs
- Custom user forms
- Integration with external data

## Quick Start

```vb
' Simple iLogic rule example
Dim partLength As Double = 100.0
Parameter("Length") = partLength

' Update the model
iLogicVb.UpdateWhenDone = True

MessageBox.Show("Part length set to " & partLength & " mm", "iLogic")
```

## Key Features

- **VB.NET Syntax**: Familiar programming language
- **Direct Model Access**: Direct manipulation of Inventor objects
- **Event-Driven**: Respond to document events
- **Integration**: Connect with databases, Excel, and external systems
- **Extensible**: Create add-ins and custom functions

## Getting Started

1. Open Autodesk Inventor
2. Navigate to the **Manage** tab
3. Click **iLogic** > **Add Rule**
4. Start coding!

## Resources

- Official Autodesk Inventor API Documentation
- Autodesk Community Forums
- This comprehensive guide

## Contributing

This documentation is maintained for the benefit of the iLogic community. Contributions and improvements are welcome.

## Version Compatibility

This documentation covers iLogic features compatible with Autodesk Inventor 2020 and newer versions. Some features may vary by version.

---

**Note**: iLogic uses VB.NET syntax. While similar to VBA, there are important differences in object model and syntax.
