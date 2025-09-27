package docx

import (
	"bytes"
	"strings"
	"testing"
	"text/template"
	"time"
)

func TestTemplateReplacer_ExecuteTemplate(t *testing.T) {
	// Open test document
	doc, err := Open("./test/template.docx")
	if err != nil {
		t.Error(err)
		return
	}

	// Test data
	data := map[string]interface{}{
		"name":        "John Doe",
		"age":         30,
		"email":       "john.doe@example.com",
		"isActive":    true,
		"currentDate": time.Now(),
	}

	// Add custom functions
	funcMap := template.FuncMap{
		"upper": func(s string) string {
			return strings.ToUpper(s)
		},
	}

	// Execute template
	err = doc.ExecuteTemplateWithFuncs(data, funcMap)
	if err != nil {
		t.Error("template execution failed", err)
		return
	}

	// Write output
	err = doc.WriteToFile("./test/template_output.docx")
	if err != nil {
		t.Error("unable to write", err)
		return
	}

	// Cleanup
	// _ = os.Remove("./test/template_output.docx")
}

func TestParseTemplatePlaceholders(t *testing.T) {
	// Mock document content with template placeholders
	docBytes := []byte(`
		<w:r>
			<w:t>Hello {{.name}}, you are {{.age}} years old.</w:t>
		</w:r>
		<w:r>
			<w:t>{{if .isActive}}Active{{else}}Inactive{{end}}</w:t>
		</w:r>
		<w:r>
			<w:t>{{range .items}}{{.name}}{{end}}</w:t>
		</w:r>
	`)

	// Mock runs
	runs := DocumentRuns{
		&Run{
			ID: 1,
			Text: TagPair{
				OpenTag:  Position{Start: 10, End: 20},
				CloseTag: Position{Start: 50, End: 60},
			},
			HasText: true,
		},
	}

	// Parse template placeholders
	placeholders, err := ParseTemplatePlaceholders(runs, docBytes, "test.xml")
	if err != nil {
		t.Error("failed to parse template placeholders", err)
		return
	}

	// Verify placeholders were found
	if len(placeholders) == 0 {
		t.Error("no template placeholders found")
		return
	}

	// Check first placeholder
	firstPlaceholder := placeholders[0]
	if firstPlaceholder.Key != ".name" {
		t.Errorf("expected key '.name', got '%s'", firstPlaceholder.Key)
	}
}

func TestTemplatePlaceholderProcessing(t *testing.T) {
	// Test template content processing
	templateContent := "{{.name | upper}}"

	// Create a simple template with functions
	funcMap := template.FuncMap{
		"upper": strings.ToUpper,
	}
	tmpl, err := template.New("test").Funcs(funcMap).Parse(templateContent)
	if err != nil {
		t.Error("failed to parse template", err)
		return
	}

	// Test data
	data := map[string]interface{}{
		"name": "john doe",
	}

	// Execute template
	var buf bytes.Buffer
	err = tmpl.Execute(&buf, data)
	if err != nil {
		t.Error("failed to execute template", err)
		return
	}

	result := buf.String()
	expected := "JOHN DOE"
	if result != expected {
		t.Errorf("expected '%s', got '%s'", expected, result)
	}
}

func TestDocumentTemplateAPI(t *testing.T) {
	// Test the simplified Document API
	doc, err := Open("./test/template.docx")
	if err != nil {
		t.Error(err)
		return
	}

	// Test setting template data
	data := map[string]interface{}{
		"test": "value",
	}
	doc.SetTemplateData(data)

	// Test adding template functions
	funcMap := template.FuncMap{
		"upper": strings.ToUpper,
	}
	doc.AddTemplateFuncs(funcMap)

	// Test template execution
	err = doc.ExecuteTemplate(data)
	if err != nil {
		t.Error("template execution failed", err)
		return
	}
}

func TestTemplateReplacer_MissingFields(t *testing.T) {
	// Open test document
	doc, err := Open("./test/template.docx")
	if err != nil {
		t.Error(err)
		return
	}

	// Test data with missing fields (only some fields present)
	data := map[string]interface{}{
		"name": "John Doe",
		// Missing: age, email, isActive, currentDate
	}

	// Execute template - this should not fail even with missing fields
	err = doc.ExecuteTemplate(data)
	if err != nil {
		t.Error("template execution should not fail with missing fields", err)
		return
	}

	// Write output to verify document is not corrupted
	err = doc.WriteToFile("./test/missing_fields_output.docx")
	if err != nil {
		t.Error("unable to write output with missing fields", err)
		return
	}

	t.Log("Successfully handled missing fields without corruption")
}
