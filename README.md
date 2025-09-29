# docx

[![tests](https://github.com/izetmolla/docx/workflows/test/badge.svg)](https://github.com/izetmolla/docx/actions?query=workflow%3Atest)
[![goreport](https://goreportcard.com/badge/github.com/izetmolla/docx)](https://goreportcard.com/report/github.com/izetmolla/docx)
[![GoDoc reference ](https://img.shields.io/badge/godoc-reference-blue.svg)](https://pkg.go.dev/github.com/izetmolla/docx)

**Process MS Word documents using Go's powerful text/template package.**

This library provides template-based processing for docx documents using Go's standard `text/template` package. It handles WordprocessingML fragmentation issues while providing advanced templating capabilities including conditional logic, loops, and custom functions.

## Features

- **Template-Based**: Uses Go's `text/template` package for advanced document processing
- **String-Based Replacement**: Simple `{placeholder}` replacement using PlaceholderMap
- **Fast**: Operates directly on byte contents for optimal performance
- **Zero Dependencies**: Built with Go standard library only
- **Advanced Features**: Conditional logic, loops, custom functions, and complex data structures
- **Fragmentation Handling**: Automatically handles WordprocessingML fragmentation issues
- **Debug Mode**: Comprehensive debug logging for troubleshooting template issues
- **Missing Field Handling**: Gracefully skips missing fields without corrupting documents
- **Convenience Functions**: One-line template processing with automatic output file generation
- **Cloud Storage Ready**: Return processed documents as bytes for direct upload to MinIO, S3, etc.
- **Serverless Processing**: Process templates entirely in memory - perfect for AWS Lambda, Cloud Functions

## Installation

```bash
go get github.com/izetmolla/docx
```

## Quick Start

### String-Based Placeholder Replacement

```go
package main

import (
    "log"
    "github.com/izetmolla/docx"
)

func main() {
    // Define placeholder replacements
    replaceMap := docx.PlaceholderMap{
        "key":                         "REPLACE some more",
        "key-with-dash":               "REPLACE",
        "key-with-dashes":             "REPLACE",
        "key with space":              "REPLACE",
        "key_with_underscore":         "REPLACE",
        "multiline":                   "REPLACE",
        "key.with.dots":               "REPLACE",
        "mixed-key.separator_styles#": "REPLACE",
        "yet-another_placeholder":     "REPLACE",
        "foo":                         "bar",
    }

    // Process document in one line - automatically creates "template_output.docx"
    err := docx.CompleteReplaceAll("template.docx", replaceMap)
    if err != nil {
        log.Fatal(err)
    }
    
    log.Println("Document processed successfully!")
}
```

### Simple One-Line Template Processing

```go
package main

import (
    "log"
    "github.com/izetmolla/docx"
)

func main() {
    // Data structure
    data := map[string]interface{}{
        "name":    "John Doe",
        "email":   "john@example.com",
        "company": "Tech Corp",
    }

    // Process template in one line - automatically creates "template_output.docx"
    err := docx.CompleteTemplate("template.docx", data)
    if err != nil {
        log.Fatal(err)
    }
    
    log.Println("Document processed successfully!")
}
```

### Advanced Processing with Custom Functions

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

    // Process template with custom functions - creates "template_output.docx"
    err := docx.CompleteTemplateWithFuncs("template.docx", data, funcMap)
    if err != nil {
        log.Fatal(err)
    }
    
    log.Println("Document processed successfully!")
}
```

### Traditional Step-by-Step Processing

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

## Placeholder Syntax

### String-Based Placeholders

Use simple `{placeholder}` syntax in your Word documents:

```
Hello {name},

Welcome to {company}!

Your account status: {status}
Revenue: {revenue}
Generated on: {date}
```

### Template Syntax

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

### Convenience Functions

#### One-Line String-Based Replacement
```go
// Process document and auto-generate output file
err := docx.CompleteReplaceAll("template.docx", replaceMap)
// Creates: template_output.docx

// Process document with custom output path
err := docx.CompleteReplaceAllToFile("template.docx", replaceMap, "custom_output.docx")

// Process document and return as bytes for cloud upload
docBytes, err := docx.CompleteReplaceAllToBytes("template.docx", replaceMap)

// Process template bytes with placeholders (serverless)
processedBytes, err := docx.CompleteReplaceAllFromBytesToBytes(templateBytes, replaceMap)
```

#### One-Line Template Processing
```go
// Process template and auto-generate output file
err := docx.CompleteTemplate("template.docx", data)
// Creates: template_output.docx

// Process template with custom output path
err := docx.CompleteTemplateToFile("template.docx", data, "custom_output.docx")

// Process template with custom functions
err := docx.CompleteTemplateWithFuncs("template.docx", data, funcMap)
// Creates: template_output.docx

// Process template with custom functions and output path
err := docx.CompleteTemplateWithFuncsToFile("template.docx", data, funcMap, "output.docx")
```

#### Cloud Storage Upload (MinIO, S3, etc.)
```go
// Process template and return as bytes for cloud upload
docBytes, err := docx.CompleteTemplateToBytes("template.docx", data)
if err != nil {
    log.Fatal(err)
}

// Upload to MinIO
minioClient.PutObject(bucketName, objectName, bytes.NewReader(docBytes), int64(len(docBytes)), minio.PutObjectOptions{})

// Process template with custom functions and return as bytes
docBytes, err := docx.CompleteTemplateWithFuncsToBytes("template.docx", data, funcMap)
if err != nil {
    log.Fatal(err)
}

// Upload to any cloud storage
uploadToCloud(docBytes, "report.docx")
```

#### Serverless Processing (MinIO-to-MinIO, No File System)
```go
// Download template bytes from MinIO
templateBytes, err := minioClient.GetObject(bucketName, "templates/report-template.docx", minio.GetObjectOptions{})
if err != nil {
    log.Fatal(err)
}
defer templateBytes.Close()

templateData, err := ioutil.ReadAll(templateBytes)
if err != nil {
    log.Fatal(err)
}

// Process template bytes with data (no file system involved)
processedBytes, err := docx.CompleteTemplateFromBytesToBytes(templateData, data)
if err != nil {
    log.Fatal(err)
}

// Upload processed bytes back to MinIO
minioClient.PutObject(
    bucketName, 
    "processed/report.docx", 
    bytes.NewReader(processedBytes), 
    int64(len(processedBytes)), 
    minio.PutObjectOptions{},
)

// With custom functions
processedBytes, err := docx.CompleteTemplateFromBytesToBytesWithFuncs(templateData, data, funcMap)
```

### Document Methods

#### Opening Documents
```go
// Open from file
doc, err := docx.Open("template.docx")

// Open from bytes
doc, err := docx.OpenBytes(documentBytes)
```

#### String-Based Replacement
```go
// Replace all placeholders
err = doc.ReplaceAll(replaceMap)

// Enable debug logging for replacement
doc.stringReplacer.SetDebug(true)
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

// Debug configuration
doc.SetDebug(true)  // Enable debug logging
doc.SetDebug(false) // Disable debug logging (default)
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
// PlaceholderMap for string-based replacement
type PlaceholderMap map[string]string

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

## Debug Mode

The library provides comprehensive debug logging to help troubleshoot template processing issues, especially when dealing with missing fields or complex data structures.

### Enabling Debug Mode

```go
doc, err := docx.Open("template.docx")
if err != nil {
    log.Fatal(err)
}
defer doc.Close()

// Enable debug logging
doc.SetDebug(true)

// Your template processing...
err = doc.ExecuteTemplate(data)
```

### Debug Output

When debug mode is enabled, you'll see detailed information about:

- **Template execution start/completion**
- **Number of placeholders found**
- **Each placeholder being processed**
- **Field existence checking** (map vs struct)
- **Missing field detection and skipping**
- **Template execution results**
- **Error handling for missing fields**

#### Example Debug Output

```
[DEBUG] Starting template execution...
[DEBUG] Found 6 template placeholders
[DEBUG] Processing placeholder: {{.ppp}}
[DEBUG] Field ppp: exists in map = false
[DEBUG] Skipping placeholder {{.ppp}} - missing fields detected
[DEBUG] Processing placeholder: {{.company}}
[DEBUG] Field company: exists in map = true
[DEBUG] Replacing placeholder {{.company}} with result: Tech Corp
[DEBUG] Processing placeholder: {{.title}}
[DEBUG] Field title: exists in map = true
[DEBUG] Replacing placeholder {{.title}} with result: Software Engineer
[DEBUG] Template execution completed successfully
```

### Debug Use Cases

#### 1. Troubleshooting Missing Fields

```go
doc.SetDebug(true)

data := map[string]interface{}{
    "name": "John Doe",
    "age":  30,
    // Missing "email" field
}

err = doc.ExecuteTemplate(data)
// Debug will show: "Field email: exists in map = false"
// Debug will show: "Skipping placeholder {{.email}} - missing fields detected"
```

#### 2. Understanding Template Processing Order

```go
doc.SetDebug(true)

// Debug shows the order of placeholder processing (right-to-left)
// This helps understand why certain replacements work or fail
```

#### 3. Verifying Field Existence

```go
doc.SetDebug(true)

// Debug shows whether fields exist in maps vs structs
// Helps identify data structure issues
```

### Silent Mode (Default)

Debug mode is **disabled by default** for silent operation:

```go
doc, err := docx.Open("template.docx")
if err != nil {
    log.Fatal(err)
}
defer doc.Close()

// Debug is disabled by default - completely silent
// doc.SetDebug(false) // Optional - this is the default

err = doc.ExecuteTemplate(data)
// No debug output - silent operation
```

### Debug API Reference

```go
// Enable debug logging
doc.SetDebug(true)

// Disable debug logging (default)
doc.SetDebug(false)

// Legacy method (still supported)
doc.SetTemplateDebug(true)  // Deprecated: Use SetDebug instead
```

### Debug Best Practices

1. **Enable debug during development** to understand template behavior
2. **Disable debug in production** for optimal performance
3. **Use debug to identify missing fields** before they cause issues
4. **Check debug output** when templates don't behave as expected
5. **Debug helps understand** the right-to-left processing order

## Examples

### Example 1: String-Based Placeholder Replacement

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

    replaceMap := docx.PlaceholderMap{
        "name":    "John Doe",
        "company": "Tech Corp",
        "status":  "Active",
        "revenue": "$1,500,000",
        "date":    "2024-01-15",
    }

    err = doc.ReplaceAll(replaceMap)
    if err != nil {
        log.Fatal(err)
    }

    err = doc.WriteToFile("output.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Example 2: Advanced String-Based Replacement with Convenience Functions

```go
package main

import (
    "log"
    "github.com/izetmolla/docx"
)

func main() {
    replaceMap := docx.PlaceholderMap{
        "key":                         "REPLACE some more",
        "key-with-dash":               "REPLACE",
        "key-with-dashes":             "REPLACE",
        "key with space":              "REPLACE",
        "key_with_underscore":         "REPLACE",
        "multiline":                   "REPLACE",
        "key.with.dots":               "REPLACE",
        "mixed-key.separator_styles#": "REPLACE",
        "yet-another_placeholder":     "REPLACE",
        "foo":                         "bar",
    }

    log.Println("=== Using String-Based Replacement Convenience Functions ===")

    // Method 1: Simple one-line processing
    err := docx.CompleteReplaceAll("template.docx", replaceMap)
    if err != nil {
        log.Fatal("CompleteReplaceAll failed:", err)
    }
    log.Println("✅ CompleteReplaceAll: template_output.docx created")

    // Method 2: Custom output path
    err = docx.CompleteReplaceAllToFile("template.docx", replaceMap, "report.docx")
    if err != nil {
        log.Fatal("CompleteReplaceAllToFile failed:", err)
    }
    log.Println("✅ CompleteReplaceAllToFile: report.docx created")

    // Method 3: Return as bytes for cloud upload
    docBytes, err := docx.CompleteReplaceAllToBytes("template.docx", replaceMap)
    if err != nil {
        log.Fatal("CompleteReplaceAllToBytes failed:", err)
    }
    log.Printf("✅ CompleteReplaceAllToBytes: %d bytes generated", len(docBytes))

    log.Println("\nAll string-based replacement functions work perfectly!")
}
```

### Example 3: Simple Template Processing

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

### Example 4: Advanced Template with Custom Functions

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

### Example 5: Using Structs

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

### Example 6: Image Replacement

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

### Example 7: Convenience Functions

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
    // Data structure
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

    // Custom functions
    funcMap := template.FuncMap{
        "upper": strings.ToUpper,
        "formatCurrency": func(amount float64) string {
            return fmt.Sprintf("$%.2f", amount)
        },
    }

    log.Println("=== Using Convenience Functions ===")

    // Method 1: Simple one-line processing
    err := docx.CompleteTemplate("template.docx", data)
    if err != nil {
        log.Fatal("CompleteTemplate failed:", err)
    }
    log.Println("✅ CompleteTemplate: template_output.docx created")

    // Method 2: Custom output path
    err = docx.CompleteTemplateToFile("template.docx", data, "report.docx")
    if err != nil {
        log.Fatal("CompleteTemplateToFile failed:", err)
    }
    log.Println("✅ CompleteTemplateToFile: report.docx created")

    // Method 3: With custom functions
    err = docx.CompleteTemplateWithFuncs("template.docx", data, funcMap)
    if err != nil {
        log.Fatal("CompleteTemplateWithFuncs failed:", err)
    }
    log.Println("✅ CompleteTemplateWithFuncs: template_output.docx created")

    // Method 4: With custom functions and output path
    err = docx.CompleteTemplateWithFuncsToFile("template.docx", data, funcMap, "final_report.docx")
    if err != nil {
        log.Fatal("CompleteTemplateWithFuncsToFile failed:", err)
    }
    log.Println("✅ CompleteTemplateWithFuncsToFile: final_report.docx created")

    log.Println("\nAll convenience functions work perfectly!")
}
```

### Example 8: Cloud Storage Upload (MinIO)

```go
package main

import (
    "bytes"
    "log"
    "strings"
    "text/template"
    
    "github.com/minio/minio-go/v7"
    "github.com/minio/minio-go/v7/pkg/credentials"
    "github.com/izetmolla/docx"
)

func main() {
    // Data structure
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
    }

    // Custom functions
    funcMap := template.FuncMap{
        "upper": strings.ToUpper,
        "formatCurrency": func(amount float64) string {
            return fmt.Sprintf("$%.2f", amount)
        },
    }

    log.Println("=== Processing template for cloud upload ===")

    // Process template and get bytes
    docBytes, err := docx.CompleteTemplateWithFuncsToBytes("template.docx", data, funcMap)
    if err != nil {
        log.Fatal("Template processing failed:", err)
    }
    
    log.Printf("✅ Generated document: %d bytes", len(docBytes))

    // Initialize MinIO client
    minioClient, err := minio.New("localhost:9000", &minio.Options{
        Creds:  credentials.NewStaticV4("minioadmin", "minioadmin", ""),
        Secure: false,
    })
    if err != nil {
        log.Fatal("MinIO client creation failed:", err)
    }

    // Upload to MinIO
    bucketName := "documents"
    objectName := "reports/user-report.docx"
    
    _, err = minioClient.PutObject(
        bucketName,
        objectName,
        bytes.NewReader(docBytes),
        int64(len(docBytes)),
        minio.PutObjectOptions{
            ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        },
    )
    if err != nil {
        log.Fatal("MinIO upload failed:", err)
    }

    log.Printf("✅ Document uploaded successfully to MinIO: %s/%s", bucketName, objectName)
    log.Println("Document is now available in cloud storage!")
}
```

### Example 9: Serverless MinIO-to-MinIO Processing

```go
package main

import (
    "bytes"
    "io/ioutil"
    "log"
    "strings"
    "text/template"
    
    "github.com/minio/minio-go/v7"
    "github.com/minio/minio-go/v7/pkg/credentials"
    "github.com/izetmolla/docx"
)

func main() {
    // Data structure
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
    }

    // Custom functions
    funcMap := template.FuncMap{
        "upper": strings.ToUpper,
        "formatCurrency": func(amount float64) string {
            return fmt.Sprintf("$%.2f", amount)
        },
    }

    log.Println("=== Serverless MinIO-to-MinIO Processing ===")

    // Initialize MinIO client
    minioClient, err := minio.New("localhost:9000", &minio.Options{
        Creds:  credentials.NewStaticV4("minioadmin", "minioadmin", ""),
        Secure: false,
    })
    if err != nil {
        log.Fatal("MinIO client creation failed:", err)
    }

    bucketName := "documents"

    // Step 1: Download template bytes from MinIO
    log.Println("📥 Downloading template from MinIO...")
    templateObject, err := minioClient.GetObject(bucketName, "templates/report-template.docx", minio.GetObjectOptions{})
    if err != nil {
        log.Fatal("Failed to download template from MinIO:", err)
    }
    defer templateObject.Close()

    templateBytes, err := ioutil.ReadAll(templateObject)
    if err != nil {
        log.Fatal("Failed to read template bytes:", err)
    }
    log.Printf("✅ Template downloaded: %d bytes", len(templateBytes))

    // Step 2: Process template bytes with data (NO FILE SYSTEM INVOLVED!)
    log.Println("⚙️ Processing template in memory...")
    processedBytes, err := docx.CompleteTemplateFromBytesToBytesWithFuncs(templateBytes, data, funcMap)
    if err != nil {
        log.Fatal("Template processing failed:", err)
    }
    log.Printf("✅ Template processed: %d bytes", len(processedBytes))

    // Step 3: Upload processed bytes back to MinIO
    log.Println("📤 Uploading processed document to MinIO...")
    outputObjectName := "processed/reports/user-report.docx"
    
    _, err = minioClient.PutObject(
        bucketName,
        outputObjectName,
        bytes.NewReader(processedBytes),
        int64(len(processedBytes)),
        minio.PutObjectOptions{
            ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        },
    )
    if err != nil {
        log.Fatal("Failed to upload to MinIO:", err)
    }

    log.Printf("✅ Document uploaded successfully to MinIO: %s/%s", bucketName, outputObjectName)
    log.Println("🎉 Complete serverless workflow: MinIO → Process → MinIO")
    log.Println("💡 Perfect for AWS Lambda, Google Cloud Functions, Azure Functions!")
}
```

### Example 10: Debug Mode Usage

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

    // Enable debug logging to troubleshoot issues
    doc.SetDebug(true)

    // Data with some missing fields
    data := map[string]interface{}{
        "name":    "John Doe",
        "age":     30,
        "email":   "john@example.com",
        // Missing: "title", "company" fields
    }

    log.Println("Processing template with debug enabled...")
    
    err = doc.ExecuteTemplate(data)
    if err != nil {
        log.Fatal(err)
    }

    err = doc.WriteToFile("output_debug.docx")
    if err != nil {
        log.Fatal(err)
    }

    log.Println("Check the debug output above to see:")
    log.Println("- Which fields were found/missing")
    log.Println("- How placeholders were processed")
    log.Println("- Why certain replacements were skipped")
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