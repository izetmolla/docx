package main

import (
	"log"

	"github.com/izetmolla/docx"
)

// Test struct with only some fields
type TestData struct {
	Name  string
	Email string
	// Missing fields: Age, Title, Company
}

func main() {
	// Open document
	doc, err := docx.Open("test/template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer doc.Close()

	// Test data with missing fields
	data := TestData{
		Name:  "John Doe",
		Email: "john@example.com",
		// Age, Title, Company are missing
	}

	log.Println("Testing template processing with missing fields...")
	log.Println("Available fields: Name, Email")
	log.Println("Missing fields: Age, Title, Company")

	// Execute template - this should now skip missing fields instead of failing
	err = doc.ExecuteTemplate(data)
	if err != nil {
		log.Fatal("Template execution failed:", err)
	}

	// Save document
	err = doc.WriteToFile("test_missing_fields_output.docx")
	if err != nil {
		log.Fatal("Failed to write output file:", err)
	}

	log.Println("Document processed successfully!")
	log.Println("Missing fields were skipped instead of causing corruption.")
}
