package main

import (
	"fmt"
	"log"
	"strings"
	"text/template"
	"time"

	"github.com/izetmolla/docx"
)

// Example data structures
type Person struct {
	Name   string
	Age    int
	Email  string
	Active bool
	Skills []string
}

type Company struct {
	Name      string
	Founded   int
	Revenue   float64
	Employees []Person
}

func main() {
	// Open document
	doc, err := docx.Open("template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer doc.Close()

	// Define custom template functions
	funcMap := template.FuncMap{
		"upper": strings.ToUpper,
		"lower": strings.ToLower,
		"formatCurrency": func(amount float64) string {
			return fmt.Sprintf("$%.2f", amount)
		},
		"formatDate": func(t time.Time) string {
			return t.Format("January 2, 2006")
		},
		"join": func(items []string, sep string) string {
			return strings.Join(items, sep)
		},
		"add": func(a, b int) int {
			return a + b
		},
	}

	// Create comprehensive data
	company := Company{
		Name:    "Tech Innovations Inc",
		Founded: 2015,
		Revenue: 2500000.75,
		Employees: []Person{
			{
				Name:   "Alice Smith",
				Age:    28,
				Email:  "alice@techinnovations.com",
				Active: true,
				Skills: []string{"Go", "Python", "JavaScript"},
			},
			{
				Name:   "Bob Wilson",
				Age:    35,
				Email:  "bob@techinnovations.com",
				Active: true,
				Skills: []string{"DevOps", "AWS", "Docker"},
			},
			{
				Name:   "Carol Davis",
				Age:    31,
				Email:  "carol@techinnovations.com",
				Active: false,
				Skills: []string{"React", "Node.js", "MongoDB"},
			},
		},
	}

	// Additional data for template processing
	data := map[string]interface{}{
		"company":     company,
		"currentDate": time.Now(),
		"year":        2024,
		"version":     "1.0.0",
		"stats": map[string]interface{}{
			"totalEmployees": len(company.Employees),
			"activeEmployees": func() int {
				count := 0
				for _, emp := range company.Employees {
					if emp.Active {
						count++
					}
				}
				return count
			}(),
			"averageAge": func() int {
				total := 0
				for _, emp := range company.Employees {
					total += emp.Age
				}
				return total / len(company.Employees)
			}(),
		},
	}

	log.Println("Processing template with comprehensive data...")

	// Execute template with custom functions
	err = doc.ExecuteTemplateWithFuncs(data, funcMap)
	if err != nil {
		log.Fatal("Template execution failed:", err)
	}

	log.Println("Template processing completed successfully!")

	// Write the final document
	err = doc.WriteToFile("comprehensive_output.docx")
	if err != nil {
		log.Fatal("Failed to write output file:", err)
	}

	log.Printf("Document successfully saved to: comprehensive_output.docx")
	log.Println()
	log.Println("Template features demonstrated:")
	log.Println("✓ Simple variable access: {{.company.name}}")
	log.Println("✓ Nested field access: {{.company.employees.0.name}}")
	log.Println("✓ Conditional logic: {{if .company.employees.0.active}}...{{end}}")
	log.Println("✓ Loops: {{range .company.employees}}...{{end}}")
	log.Println("✓ Function pipelines: {{.company.name | upper}}")
	log.Println("✓ Custom functions: {{.company.revenue | formatCurrency}}")
	log.Println("✓ Complex calculations: {{.stats.averageAge}}")
	log.Println("✓ String functions: {{.company.employees.0.skills | join \", \"}}")
	log.Println("✓ Mathematical operations: {{add .company.employees.0.age 5}}")
}
