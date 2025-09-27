package main

import (
	"log"
	"os"

	"github.com/izetmolla/docx"
)

func main() {
	// Open document
	doc, err := docx.Open("template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer doc.Close()

	// Process templates first
	data := map[string]interface{}{
		"title":      "Company Report",
		"date":       "2024-01-15",
		"author":     "John Smith",
		"department": "Engineering",
		"reportType": "Quarterly Review",
	}

	err = doc.ExecuteTemplate(data)
	if err != nil {
		log.Fatal("Template execution failed:", err)
	}

	log.Println("Template processing completed!")

	// Replace company logo
	logoBytes, err := os.ReadFile("logo.png")
	if err != nil {
		log.Printf("Warning: Could not read logo.png: %v", err)
		log.Println("Continuing without logo replacement...")
	} else {
		err = doc.SetFile("word/media/image1.png", logoBytes)
		if err != nil {
			log.Printf("Warning: Could not replace logo: %v", err)
		} else {
			log.Println("Logo replaced successfully!")
		}
	}

	// Replace signature image
	signatureBytes, err := os.ReadFile("signature.png")
	if err != nil {
		log.Printf("Warning: Could not read signature.png: %v", err)
		log.Println("Continuing without signature replacement...")
	} else {
		err = doc.SetFile("word/media/image2.png", signatureBytes)
		if err != nil {
			log.Printf("Warning: Could not replace signature: %v", err)
		} else {
			log.Println("Signature replaced successfully!")
		}
	}

	// Save the final document
	err = doc.WriteToFile("report_with_images.docx")
	if err != nil {
		log.Fatal("Failed to write output file:", err)
	}

	log.Println("Document with images saved successfully!")
	log.Println()
	log.Println("This example demonstrates:")
	log.Println("✓ Template processing with text data")
	log.Println("✓ Image replacement in Word documents")
	log.Println("✓ Error handling for missing image files")
	log.Println("✓ Multiple image replacements in one document")
}
