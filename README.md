# docx

[![tests](https://github.com/izetmolla/docx/workflows/test/badge.svg)](https://github.com/izetmolla/docx/actions?query=workflow%3Atest)
[![goreport](https://goreportcard.com/badge/github.com/izetmolla/docx)](https://goreportcard.com/report/github.com/izetmolla/docx)
[![GoDoc reference ](https://img.shields.io/badge/godoc-reference-blue.svg)](https://pkg.go.dev/github.com/izetmolla/docx)

**Process MS Word documents using Go's powerful text/template package.**

This library provides template-based processing for docx documents using Go's standard `text/template` package. It handles WordprocessingML fragmentation issues while providing advanced templating capabilities including conditional logic, loops, and custom functions.

## Features

- **Template-Based**: Uses Go's `text/template` package for advanced document processing
- **Fast**: Operates directly on byte contents for optimal performance
- **Zero Dependencies**: Built with Go standard library only
- **Advanced Features**: Conditional logic, loops, custom functions, and complex data structures
- **Fragmentation Handling**: Automatically handles WordprocessingML fragmentation issues

## Installation

```bash
go get github.com/izetmolla/docx
```

## Quick Start

```go
package main

import (
    "log"
    "strings"
    "text/template"
    "time"
    
    "github.com/izetmolla/docx"
)

func main() {
    // Open document
    doc, err := docx.Open("template.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    // Define custom functions
    funcMap := template.FuncMap{
        "upper": strings.ToUpper,
        "formatDate": func(t time.Time) string {
            return t.Format("2006-01-02")
        },
    }

    // Set data
    data := map[string]interface{}{
        "user": map[string]interface{}{
            "name":     "John Doe",
            "email":    "john@example.com",
            "isActive": true,
        },
        "company": map[string]interface{}{
            "name":    "Tech Corp",
            "revenue": 1500000.50,
        },
        "currentDate": time.Now(),
    }

    // Execute template
    err = doc.ExecuteTemplateWithFuncs(data, funcMap)
    if err != nil {
        log.Fatal(err)
    }

    // Save document
    err = doc.WriteToFile("output.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

## Template Syntax

Use Go template syntax in your Word documents:

```
Hello {{.user.name | upper}},

Welcome to {{.company.name}}!

{{if .user.isActive}}
Your account is active.
{{else}}
Your account is inactive.
{{end}}

Employee List:
{{range .employees}}
- {{.name}}: {{.position}}
{{end}}

Revenue: {{.company.revenue | formatCurrency}}
Generated on: {{.currentDate | formatDate}}
```

## Template Features

### Variable Access
- **Simple**: `{{.name}}` - Access top-level fields
- **Nested**: `{{.user.profile.name}}` - Access nested fields
- **Indexed**: `{{.items.0.name}}` - Access array elements

### Conditional Logic
```go
{{if .isActive}}
User is active
{{else}}
User is inactive
{{end}}

{{if .user.isAdmin}}
Admin privileges
{{else if .user.isModerator}}
Moderator privileges
{{else}}
Regular user
{{end}}
```

### Loops
```go
{{range .employees}}
- {{.name}}: {{.position}}
{{end}}

{{range $index, $employee := .employees}}
{{$index}}. {{$employee.name}} ({{$employee.position}})
{{end}}
```

### Function Pipelines
```go
{{.name | upper | trim}}
{{.amount | formatCurrency}}
{{.description | truncate 50}}
```

### Custom Functions
```go
funcMap := template.FuncMap{
    "upper": strings.ToUpper,
    "lower": strings.ToLower,
    "formatCurrency": func(amount float64) string {
        return fmt.Sprintf("$%.2f", amount)
    },
    "formatDate": func(t time.Time) string {
        return t.Format("2006-01-02")
    },
    "join": func(items []string, sep string) string {
        return strings.Join(items, sep)
    },
    "add": func(a, b int) int {
        return a + b
    },
    "isEven": func(num int) bool {
        return num%2 == 0
    },
}
```

## API Reference

### Document Methods

#### Opening Documents
```go
// Open from file
doc, err := docx.Open("template.docx")

// Open from bytes
doc, err := docx.OpenBytes(documentBytes)
```

#### Template Processing
```go
// Simple template execution
err = doc.ExecuteTemplate(data)

// Template execution with custom functions
err = doc.ExecuteTemplateWithFuncs(data, funcMap)

// Add custom functions
doc.AddTemplateFuncs(funcMap)

// Set template data
doc.SetTemplateData(data)
```

#### File Operations
```go
// Write to file
err = doc.WriteToFile("output.docx")

// Write to writer
err = doc.Write(writer)

// Replace images
err = doc.SetFile("word/media/image1.jpg", imageBytes)

// Get file content
content := doc.GetFile("word/document.xml")
```

#### Cleanup
```go
// Close document
doc.Close()
```

### Data Types

```go
// TemplateData can be any Go type
type TemplateData interface{}

// Example data structures
type Person struct {
    Name   string
    Age    int
    Email  string
    Active bool
}

type Company struct {
    Name      string
    Employees []Person
    Revenue   float64
}
```

## Examples

### Example 1: Simple Template Processing

```go
package main

import (
    "log"
    "github.com/izetmolla/docx"
)

func main() {
    doc, err := docx.Open("template.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    data := map[string]interface{}{
        "name":  "John Doe",
        "age":   30,
        "email": "john@example.com",
    }

    err = doc.ExecuteTemplate(data)
    if err != nil {
        log.Fatal(err)
    }

    err = doc.WriteToFile("output.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Example 2: Advanced Template with Custom Functions

```go
package main

import (
    "fmt"
    "log"
    "strings"
    "text/template"
    "time"
    
    "github.com/izetmolla/docx"
)

func main() {
    doc, err := docx.Open("template.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    // Define custom functions
    funcMap := template.FuncMap{
        "upper":    strings.ToUpper,
        "lower":    strings.ToLower,
        "formatCurrency": func(amount float64) string {
            return fmt.Sprintf("$%.2f", amount)
        },
        "formatDate": func(t time.Time) string {
            return t.Format("January 2, 2006")
        },
        "join": func(items []string, sep string) string {
            return strings.Join(items, sep)
        },
    }

    // Complex data structure
    data := map[string]interface{}{
        "company": map[string]interface{}{
            "name":    "Tech Solutions Inc",
            "revenue": 2500000.50,
            "founded": 2010,
        },
        "employees": []map[string]interface{}{
            {"name": "Alice Johnson", "position": "Developer", "skills": []string{"Go", "Python"}},
            {"name": "Bob Smith", "position": "Designer", "skills": []string{"UI", "UX"}},
            {"name": "Carol Davis", "position": "Manager", "skills": []string{"Leadership", "Strategy"}},
        },
        "currentDate": time.Now(),
    }

    err = doc.ExecuteTemplateWithFuncs(data, funcMap)
    if err != nil {
        log.Fatal(err)
    }

    err = doc.WriteToFile("output.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Example 3: Using Structs

```go
package main

import (
    "log"
    "github.com/izetmolla/docx"
)

type Person struct {
    Name   string
    Age    int
    Email  string
    Active bool
}

type Company struct {
    Name      string
    Employees []Person
    Revenue   float64
}

func main() {
    doc, err := docx.Open("template.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    company := Company{
        Name: "Tech Corp",
        Employees: []Person{
            {Name: "Alice", Age: 28, Email: "alice@techcorp.com", Active: true},
            {Name: "Bob", Age: 35, Email: "bob@techcorp.com", Active: false},
        },
        Revenue: 1500000.50,
    }

    err = doc.ExecuteTemplate(company)
    if err != nil {
        log.Fatal(err)
    }

    err = doc.WriteToFile("output.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Example 4: Image Replacement

```go
package main

import (
    "io/ioutil"
    "log"
    "github.com/izetmolla/docx"
)

func main() {
    doc, err := docx.Open("template.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    // Process templates
    data := map[string]interface{}{
        "title": "Company Report",
        "date":  "2024-01-15",
    }
    err = doc.ExecuteTemplate(data)
    if err != nil {
        log.Fatal(err)
    }

    // Replace image
    imageBytes, err := ioutil.ReadFile("new-logo.png")
    if err != nil {
        log.Fatal(err)
    }
    
    err = doc.SetFile("word/media/image1.png", imageBytes)
    if err != nil {
        log.Fatal(err)
    }

    err = doc.WriteToFile("output.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

## Template Document Content Examples

### Basic Template
```
Company Report
==============

Company: {{.company.name | upper}}
Founded: {{.company.founded}}
Revenue: {{.company.revenue | formatCurrency}}

Generated on: {{.currentDate | formatDate}}
```

### Advanced Template with Conditionals and Loops
```
Employee Directory
=================

{{if .company.employees}}
Total Employees: {{len .company.employees}}

{{range .company.employees}}
Employee: {{.name}}
Age: {{.age}} years old
Email: {{.email}}
{{if .active}}
Status: Active
{{else}}
Status: Inactive
{{end}}
Skills: {{.skills | join ", "}}

{{end}}
{{else}}
No employees found.
{{end}}

Company Statistics:
- Total Revenue: {{.company.revenue | formatCurrency}}
- Average Age: {{.averageAge}}
```

## Error Handling

```go
err = doc.ExecuteTemplateWithFuncs(data, funcMap)
if err != nil {
    log.Printf("Template error: %v", err)
    // Handle error appropriately
}
```

Common errors include:
- Missing template data fields
- Invalid template syntax
- Function execution errors
- Document parsing issues

## Performance Considerations

- Template parsing is done once per placeholder
- Large documents with many template placeholders may take longer to process
- Consider caching template functions for repeated use
- Use simple data structures when possible for better performance

## Best Practices

1. **Use descriptive field names** in your data structures
2. **Validate template syntax** before processing large documents
3. **Handle errors gracefully** with proper error checking
4. **Use custom functions** for common formatting needs
5. **Test templates** with sample data before production use
6. **Keep data structures simple** when possible
7. **Use conditional logic** sparingly to maintain document readability

## License

This software is licensed under the [MIT license](LICENSE).