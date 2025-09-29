package docx

import (
	"fmt"
	"regexp"
	"strings"
)

// StringReplacer provides string-based placeholder replacement functionality
type StringReplacer struct {
	document *Document
	debug    bool // Enable debug logging
}

// NewStringReplacer creates a new string replacer for the given document
func NewStringReplacer(doc *Document) *StringReplacer {
	return &StringReplacer{
		document: doc,
	}
}

// SetDebug enables or disables debug logging
func (sr *StringReplacer) SetDebug(debug bool) {
	sr.debug = debug
}

// debugLog logs a message if debug mode is enabled
func (sr *StringReplacer) debugLog(format string, args ...interface{}) {
	if sr.debug {
		fmt.Printf("[DEBUG] "+format+"\n", args...)
	}
}

// ReplaceAll replaces all string-based placeholders in the document using the provided PlaceholderMap.
// Placeholders are delimited with { and } and can contain any characters except the delimiters.
func (sr *StringReplacer) ReplaceAll(replaceMap PlaceholderMap) error {
	fmt.Println("Starting ReplaceAll...")

	sr.debugLog("Starting string-based placeholder replacement...")
	sr.debugLog("Found %d placeholders to replace", len(replaceMap))

	// Process each file in the document
	for fileName := range sr.document.files {
		sr.debugLog("Processing file: %s", fileName)

		// Get the current file content
		fileContent := sr.document.GetFile(fileName)
		if fileContent == nil {
			continue
		}

		// Replace placeholders in this file
		newContent, err := sr.replacePlaceholdersInFile(string(fileContent), replaceMap)
		if err != nil {
			return fmt.Errorf("failed to replace placeholders in file %s: %w", fileName, err)
		}

		// Update the file content
		err = sr.document.SetFile(fileName, []byte(newContent))
		if err != nil {
			return fmt.Errorf("failed to update file %s: %w", fileName, err)
		}
	}

	sr.debugLog("String-based placeholder replacement completed successfully")
	return nil
}

// replacePlaceholdersInFile replaces all placeholders in a single file's content
func (sr *StringReplacer) replacePlaceholdersInFile(content string, replaceMap PlaceholderMap) (string, error) {
	result := content

	// Process each placeholder in the replace map
	for placeholder, replacement := range replaceMap {
		sr.debugLog("Replacing placeholder: {%s} with: %s", placeholder, replacement)

		// Create the full placeholder with braces
		fullPlaceholder := "{" + placeholder + "}"

		// Count occurrences for logging
		count := strings.Count(result, fullPlaceholder)
		if count > 0 {
			sr.debugLog("Found %d occurrences of {%s}", count, placeholder)
			result = strings.ReplaceAll(result, fullPlaceholder, replacement)
		} else {
			sr.debugLog("No occurrences found for {%s}", placeholder)
		}
	}

	return result, nil
}

// ExtractPlaceholders extracts all placeholders from the document content
// This is useful for debugging or validation purposes
func (sr *StringReplacer) ExtractPlaceholders() ([]string, error) {
	var allPlaceholders []string
	placeholderRegex := regexp.MustCompile(`\{([^}]+)\}`)

	for fileName := range sr.document.files {
		fileContent := sr.document.GetFile(fileName)
		if fileContent == nil {
			continue
		}

		// Find all placeholders in this file
		matches := placeholderRegex.FindAllStringSubmatch(string(fileContent), -1)
		for _, match := range matches {
			if len(match) > 1 {
				placeholder := match[1] // The content inside the braces
				allPlaceholders = append(allPlaceholders, placeholder)
			}
		}
	}

	return allPlaceholders, nil
}

// ValidatePlaceholders checks if all placeholders in the document have corresponding values in the replace map
func (sr *StringReplacer) ValidatePlaceholders(replaceMap PlaceholderMap) ([]string, error) {
	documentPlaceholders, err := sr.ExtractPlaceholders()
	if err != nil {
		return nil, err
	}

	var missingPlaceholders []string
	for _, placeholder := range documentPlaceholders {
		if _, exists := replaceMap[placeholder]; !exists {
			missingPlaceholders = append(missingPlaceholders, placeholder)
		}
	}

	return missingPlaceholders, nil
}
