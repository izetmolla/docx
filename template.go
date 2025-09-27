package docx

import (
	"bytes"
	"fmt"
	"regexp"
	"strings"
	"text/template"
)

// TemplateData represents the data structure that can be used in templates
type TemplateData interface{}

// TemplateReplacer provides template-based replacement functionality
type TemplateReplacer struct {
	document *Document
	tmpl     *template.Template
	data     TemplateData
}

// NewTemplateReplacer creates a new template replacer for the given document
func NewTemplateReplacer(doc *Document) *TemplateReplacer {
	return &TemplateReplacer{
		document: doc,
		tmpl:     template.New("docx-template"),
	}
}

// SetData sets the data to be used for template execution
func (tr *TemplateReplacer) SetData(data TemplateData) {
	tr.data = data
}

// AddFuncs adds custom functions to the template
func (tr *TemplateReplacer) AddFuncs(funcMap template.FuncMap) {
	tr.tmpl = tr.tmpl.Funcs(funcMap)
}

// ExecuteTemplate replaces all template placeholders in the document
// Template placeholders use Go template syntax: {{.field}}, {{if .condition}}...{{end}}, etc.
func (tr *TemplateReplacer) ExecuteTemplate() error {
	if tr.data == nil {
		return fmt.Errorf("template data not set, call SetData() first")
	}

	// Extract all template placeholders from the document
	templatePlaceholders, err := tr.extractTemplatePlaceholders()
	if err != nil {
		return fmt.Errorf("failed to extract template placeholders: %w", err)
	}

	// Process each template placeholder in reverse order to avoid position conflicts
	// This ensures that earlier positions remain valid after replacements
	for i := len(templatePlaceholders) - 1; i >= 0; i-- {
		placeholder := templatePlaceholders[i]
		err := tr.processTemplatePlaceholder(placeholder)
		if err != nil {
			return fmt.Errorf("failed to process template placeholder %s: %w", placeholder.TemplateContent, err)
		}
	}

	return nil
}

// extractTemplatePlaceholders finds all Go template syntax placeholders in the document
func (tr *TemplateReplacer) extractTemplatePlaceholders() ([]*TemplatePlaceholder, error) {
	var templatePlaceholders []*TemplatePlaceholder

	for fileName := range tr.document.files {
		placeholders, err := ParseTemplatePlaceholders(tr.document.runParsers[fileName].Runs(), tr.document.GetFile(fileName), fileName)
		if err != nil {
			return nil, err
		}
		templatePlaceholders = append(templatePlaceholders, placeholders...)
	}

	return templatePlaceholders, nil
}

// processTemplatePlaceholder processes a single template placeholder
func (tr *TemplateReplacer) processTemplatePlaceholder(placeholder *TemplatePlaceholder) error {
	// Check if the template references missing fields BEFORE executing
	if tr.hasMissingFields(placeholder.TemplateContent) {
		// Skip this placeholder - leave it unchanged in the document
		return nil
	}

	// Parse the template content
	tmpl, err := tr.tmpl.Parse(placeholder.TemplateContent)
	if err != nil {
		return fmt.Errorf("failed to parse template: %w", err)
	}

	// Execute the template with the provided data
	var buf bytes.Buffer
	err = tmpl.Execute(&buf, tr.data)
	if err != nil {
		// Check if the error is due to missing field/property
		// If so, skip this placeholder instead of failing
		if tr.isMissingFieldError(err) {
			// Skip this placeholder - leave it unchanged in the document
			return nil
		}
		return fmt.Errorf("failed to execute template: %w", err)
	}

	// Check if the result contains "<no value>" which indicates missing fields
	result := buf.String()
	if strings.Contains(result, "<no value>") {
		// Skip this placeholder - leave it unchanged in the document
		return nil
	}

	// Replace the placeholder with the executed result
	err = tr.replacePlaceholder(placeholder, result)
	if err != nil {
		return fmt.Errorf("failed to replace placeholder: %w", err)
	}

	return nil
}

// isMissingFieldError checks if the error is due to a missing field/property in the data structure
func (tr *TemplateReplacer) isMissingFieldError(err error) bool {
	if err == nil {
		return false
	}

	errStr := err.Error()

	// Common Go template errors for missing fields
	missingFieldErrors := []string{
		"no such field",
		"can't evaluate field",
		"can't find method",
		"no such method",
		"can't access field",
		"undefined field",
		"nil pointer",
		"invalid value",
	}

	for _, missingErr := range missingFieldErrors {
		if strings.Contains(strings.ToLower(errStr), missingErr) {
			return true
		}
	}

	return false
}

// hasMissingFields checks if the template content references fields that don't exist in the data
func (tr *TemplateReplacer) hasMissingFields(templateContent string) bool {
	if tr.data == nil {
		return true
	}

	// Extract field names from template content like {{.fieldName}}
	// This is a simple regex to find field references
	fieldPattern := `\{\{\.([^}]+)\}\}`
	matches := regexp.MustCompile(fieldPattern).FindAllStringSubmatch(templateContent, -1)

	for _, match := range matches {
		if len(match) > 1 {
			fieldName := match[1]
			if !tr.fieldExists(fieldName) {
				return true
			}
		}
	}

	return false
}

// fieldExists checks if a field exists in the data structure
func (tr *TemplateReplacer) fieldExists(fieldName string) bool {
	if tr.data == nil {
		return false
	}

	// Handle map[string]interface{}
	if dataMap, ok := tr.data.(map[string]interface{}); ok {
		_, exists := dataMap[fieldName]
		return exists
	}

	// Handle structs - use reflection to check if field exists
	// This is a simplified check - for complex nested fields, we'd need more sophisticated logic
	return tr.checkStructField(fieldName)
}

// checkStructField checks if a field exists in a struct using reflection
func (tr *TemplateReplacer) checkStructField(fieldName string) bool {
	// For now, we'll use a simple approach - try to execute a minimal template
	// and see if it fails with a missing field error
	testTemplate := fmt.Sprintf("{{.%s}}", fieldName)
	tmpl, err := template.New("test").Parse(testTemplate)
	if err != nil {
		return false
	}

	var buf bytes.Buffer
	err = tmpl.Execute(&buf, tr.data)
	if err != nil {
		return tr.isMissingFieldError(err)
	}

	// If execution succeeds and doesn't produce "<no value>", the field exists
	result := buf.String()
	return !strings.Contains(result, "<no value>")
}

// replacePlaceholder replaces a template placeholder with the executed result
func (tr *TemplateReplacer) replacePlaceholder(placeholder *TemplatePlaceholder, result string) error {
	// Get the document bytes for the file
	docBytes := tr.document.GetFile(placeholder.FileName)
	if docBytes == nil {
		return fmt.Errorf("file %s not found", placeholder.FileName)
	}

	// Calculate positions
	startPos := int(placeholder.Placeholder.StartPos())
	endPos := int(placeholder.Placeholder.EndPos())

	// Replace the placeholder content
	newBytes := make([]byte, len(docBytes)-(endPos-startPos)+len(result))
	copy(newBytes, docBytes[:startPos])
	copy(newBytes[startPos:], result)
	copy(newBytes[startPos+len(result):], docBytes[endPos:])

	// Update the document
	return tr.document.SetFile(placeholder.FileName, newBytes)
}

// TemplatePlaceholder represents a template placeholder found in the document
type TemplatePlaceholder struct {
	Placeholder     *Placeholder
	FileName        string
	TemplateContent string
	Key             string
}

// Placeholder represents a parsed placeholder from the docx-archive.
type Placeholder struct {
	Fragments []*PlaceholderFragment
}

// StartPos returns the absolute start position of the placeholder.
func (p Placeholder) StartPos() int64 {
	return p.Fragments[0].Run.Text.OpenTag.End + p.Fragments[0].Position.Start
}

// EndPos returns the absolute end position of the placeholder.
func (p Placeholder) EndPos() int64 {
	end := len(p.Fragments) - 1
	return p.Fragments[end].Run.Text.OpenTag.End + p.Fragments[end].Position.End
}

// PlaceholderFragment represents a fragment of a placeholder
type PlaceholderFragment struct {
	Position Position
	Run      *Run
}

// ParseTemplatePlaceholders extracts Go template syntax placeholders from document runs
func ParseTemplatePlaceholders(runs DocumentRuns, docBytes []byte, fileName string) ([]*TemplatePlaceholder, error) {
	var templatePlaceholders []*TemplatePlaceholder

	for _, run := range runs.WithText() {
		runText := run.GetText(docBytes)

		// Find template placeholders using Go template syntax
		templateStarts := findTemplateStarts(runText)
		templateEnds := findTemplateEnds(runText)

		// Match template starts with ends
		for i, start := range templateStarts {
			if i < len(templateEnds) {
				end := templateEnds[i]
				templateContent := runText[start : end+2] // +2 to include }}

				// Create placeholder fragment
				fragment := &PlaceholderFragment{
					Position: Position{int64(start), int64(end + 2)},
					Run:      run,
				}
				placeholder := &Placeholder{Fragments: []*PlaceholderFragment{fragment}}

				// Extract the key (content between {{ and }})
				key := templateContent[2 : len(templateContent)-2] // Remove {{ and }}

				templatePlaceholder := &TemplatePlaceholder{
					Placeholder:     placeholder,
					FileName:        fileName,
					TemplateContent: templateContent,
					Key:             key,
				}

				templatePlaceholders = append(templatePlaceholders, templatePlaceholder)
			}
		}
	}

	return templatePlaceholders, nil
}

// findTemplateStarts finds all positions of "{{" in the text
func findTemplateStarts(text string) []int {
	var starts []int
	for i := 0; i < len(text)-1; i++ {
		if text[i] == '{' && text[i+1] == '{' {
			starts = append(starts, i)
		}
	}
	return starts
}

// findTemplateEnds finds all positions of "}}" in the text
func findTemplateEnds(text string) []int {
	var ends []int
	for i := 0; i < len(text)-1; i++ {
		if text[i] == '}' && text[i+1] == '}' {
			ends = append(ends, i)
		}
	}
	return ends
}

// ExecuteTemplateWithData is a convenience method that combines SetData and ExecuteTemplate
func (tr *TemplateReplacer) ExecuteTemplateWithData(data TemplateData) error {
	tr.SetData(data)
	return tr.ExecuteTemplate()
}

// ExecuteTemplateWithFuncs is a convenience method that adds functions and executes template
func (tr *TemplateReplacer) ExecuteTemplateWithFuncs(data TemplateData, funcMap template.FuncMap) error {
	tr.AddFuncs(funcMap)
	return tr.ExecuteTemplateWithData(data)
}
