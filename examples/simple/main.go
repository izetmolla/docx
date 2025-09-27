package main

import (
	"log"

	"github.com/izetmolla/docx"
)

func main() {
	// Open document
	doc, err := docx.Open("template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer doc.Close()

	// Simple data
	data := map[string]interface{}{
		"name":    "John Doe",
		"age":     30,
		"email":   "john@example.com",
		"title":   "Software Engineer",
		"company": "Tech Corp",
	}

	// Execute template
	err = doc.ExecuteTemplate(data)
	if err != nil {
		log.Fatal(err)
	}

	// Save document
	err = doc.WriteToFile("output.docx")
	if err != nil {
		log.Fatal(err)
	}

	log.Println("Document processed successfully!")
}
