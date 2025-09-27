package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"text/template"
)

const (
	// DocumentXml is the relative path where the actual document content resides inside the docx-archive.
	DocumentXml = "word/document.xml"
)

var (
	// HeaderPathRegex matches all header files inside the docx-archive.
	HeaderPathRegex = regexp.MustCompile(`word/header[0-9]*.xml`)
	// FooterPathRegex matches all footer files inside the docx-archive.
	FooterPathRegex = regexp.MustCompile(`word/footer[0-9]*.xml`)
	// MediaPathRegex matches all media files inside the docx-archive.
	MediaPathRegex = regexp.MustCompile(`word/media/*`)
)

// Document exposes the main API of the library for template-based document processing.
// It represents a docx document that will be processed using Go's text/template package.
type Document struct {
	path     string
	docxFile *os.File
	zipFile  *zip.Reader

	// all files from the zip archive which we're interested in
	files FileMap
	// paths to all header files inside the zip archive
	headerFiles []string
	// paths to all footer files inside the zip archive
	footerFiles []string
	// paths to all media files inside the zip archive
	mediaFiles []string
	// The document contains multiple files which eventually need a parser each.
	// The map key is the file path inside the document to which the parser belongs.
	runParsers map[string]*RunParser

	// Template processing components
	templateReplacer *TemplateReplacer
}

// Open will open and parse the file pointed to by path.
// The file must be a valid docx file or an error is returned.
func Open(path string) (*Document, error) {
	fh, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("unable to open .docx file: %s", err)
	}

	rc, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("unable to open zip reader: %s", err)
	}

	return newDocument(&rc.Reader, path, fh)
}

// OpenBytes allows to create a Document from a byte slice.
// It behaves just like Open().
//
// Note: In this case, the docxFile property will be nil!
func OpenBytes(b []byte) (*Document, error) {
	rc, err := zip.NewReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		return nil, fmt.Errorf("unable to open zip reader: %s", err)
	}

	return newDocument(rc, "", nil)
}

// newDocument will create a new document struct given the zipFile.
// The params 'path' and 'docxFile' may be empty/nil in case the document is created from a byte source directly.
//
// newDocument will parse the docx archive and validate that at least a 'document.xml' exists.
// If 'word/document.xml' is missing, an error is returned since the docx cannot be correct.
// Then all files are parsed for their runs before returning the new document.
func newDocument(zipFile *zip.Reader, path string, docxFile *os.File) (*Document, error) {
	doc := &Document{
		docxFile:   docxFile,
		zipFile:    zipFile,
		path:       path,
		files:      make(FileMap),
		runParsers: make(map[string]*RunParser),
	}

	if err := doc.parseArchive(); err != nil {
		return nil, fmt.Errorf("error parsing document: %s", err)
	}

	// a valid docx document should really contain a document.xml :)
	if _, exists := doc.files[DocumentXml]; !exists {
		return nil, fmt.Errorf("invalid docx archive, %s is missing", DocumentXml)
	}

	// parse all files for template processing
	for name, data := range doc.files {
		// find all runs
		doc.runParsers[name] = NewRunParser(data)
		err := doc.runParsers[name].Execute()
		if err != nil {
			return nil, err
		}
	}

	// Initialize template replacer
	doc.templateReplacer = NewTemplateReplacer(doc)

	return doc, nil
}

// ExecuteTemplate processes all template placeholders in the document using the provided data.
// Template placeholders use Go template syntax: {{.field}}, {{if .condition}}...{{end}}, etc.
func (d *Document) ExecuteTemplate(data TemplateData) error {
	return d.templateReplacer.ExecuteTemplateWithData(data)
}

// ExecuteTemplateWithFuncs processes all template placeholders with custom functions.
func (d *Document) ExecuteTemplateWithFuncs(data TemplateData, funcMap template.FuncMap) error {
	return d.templateReplacer.ExecuteTemplateWithFuncs(data, funcMap)
}

// AddTemplateFuncs adds custom functions to the template processor.
func (d *Document) AddTemplateFuncs(funcMap template.FuncMap) {
	d.templateReplacer.AddFuncs(funcMap)
}

// SetTemplateData sets the data to be used for template execution.
func (d *Document) SetTemplateData(data TemplateData) {
	d.templateReplacer.SetData(data)
}

// GetFile returns the content of the given fileName if it exists.
func (d *Document) GetFile(fileName string) []byte {
	if f, exists := d.files[fileName]; exists {
		return f
	}
	return nil
}

// SetFile allows setting the file contents of the given file.
// The fileName must be known, otherwise an error is returned.
func (d *Document) SetFile(fileName string, fileBytes []byte) error {
	if _, exists := d.files[fileName]; !exists {
		return fmt.Errorf("unregistered file %s", fileName)
	}
	d.files[fileName] = fileBytes
	return nil
}

// parseArchive will go through the docx zip archive and read them into the FileMap.
// Files inside the FileMap are those which can be modified by the lib.
// Currently not all files are read, only:
//   - word/document.xml
//   - word/header*.xml
//   - word/footer*.xml
//   - word/media/*
func (d *Document) parseArchive() error {
	readZipFile := func(file *zip.File) []byte {
		readCloser, err := file.Open()
		if err != nil {
			return nil
		}
		defer readCloser.Close()
		fileBytes, err := ioutil.ReadAll(readCloser)
		if err != nil {
			return nil
		}
		return fileBytes
	}

	for _, file := range d.zipFile.File {
		if file.Name == DocumentXml {
			d.files[DocumentXml] = readZipFile(file)
		}
		if HeaderPathRegex.MatchString(file.Name) {
			d.files[file.Name] = readZipFile(file)
			d.headerFiles = append(d.headerFiles, file.Name)
		}
		if FooterPathRegex.MatchString(file.Name) {
			d.files[file.Name] = readZipFile(file)
			d.footerFiles = append(d.footerFiles, file.Name)
		}
		if MediaPathRegex.MatchString(file.Name) {
			d.files[file.Name] = readZipFile(file)
			d.mediaFiles = append(d.mediaFiles, file.Name)
		}
	}
	return nil
}

// WriteToFile will write the document to a new file.
// It is important to note that the target file cannot be the same as the path of this document.
// If the path is not yet created, the function will attempt to MkdirAll() before creating the file.
func (d *Document) WriteToFile(file string) error {
	if file == d.path {
		return fmt.Errorf("WriteToFile cannot write into the original docx archive while it's open")
	}

	err := os.MkdirAll(filepath.Dir(file), 0755)
	if err != nil {
		return fmt.Errorf("unable to ensure path directories: %s", err)
	}

	target, err := os.Create(file)
	if err != nil {
		return err
	}
	defer target.Close()

	return d.Write(target)
}

// Write is responsible for assembling a new .docx file using the modified data as well as all remaining files.
// Docx files are basically zip archives with many XMLs included.
// Files which cannot be modified through this lib will just be read from the original docx and copied into the writer.
func (d *Document) Write(writer io.Writer) error {
	zipWriter := zip.NewWriter(writer)
	defer zipWriter.Close()

	// writeModifiedFile will check if the given zipFile is a file which was modified and writes it.
	// If the file is not one of the modified files, false is returned.
	writeModifiedFile := func(writer io.Writer, zipFile *zip.File) (bool, error) {
		isModified := d.isModifiedFile(zipFile.Name)
		if !isModified {
			return false, nil
		}
		if err := d.files.Write(writer, zipFile.Name); err != nil {
			return false, fmt.Errorf("unable to writeFile %s: %s", zipFile.Name, err)
		}
		return true, nil
	}

	// write all files into the zip archive (docx-file)
	for _, zipFile := range d.zipFile.File {
		fw, err := zipWriter.Create(zipFile.Name)
		if err != nil {
			return fmt.Errorf("unable to create writer: %s", err)
		}

		// write all files which might've been modified by us
		written, err := writeModifiedFile(fw, zipFile)
		if err != nil {
			return err
		}
		if written {
			continue
		}

		// all files which we don't touch here (e.g. _rels.xml) are just copied from the original
		readCloser, err := zipFile.Open()
		if err != nil {
			return fmt.Errorf("unable to open %s: %s", zipFile.Name, err)
		}
		_, err = fw.Write(readBytes(readCloser))
		if err != nil {
			return fmt.Errorf("unable to writeFile zipFile %s: %s", zipFile.Name, err)
		}
		err = readCloser.Close()
		if err != nil {
			return fmt.Errorf("unable to close reader for %s: %s", zipFile.Name, err)
		}
	}
	return nil
}

// isModifiedFile will look through all modified files and check if the searchFileName exists
func (d *Document) isModifiedFile(searchFileName string) bool {
	allFiles := append(d.headerFiles, d.footerFiles...)
	allFiles = append(allFiles, d.mediaFiles...)
	allFiles = append(allFiles, DocumentXml)

	for _, file := range allFiles {
		if searchFileName == file {
			return true
		}
	}
	return false
}

// Close will close everything :)
func (d *Document) Close() {
	if d.docxFile != nil {
		err := d.docxFile.Close()
		if err != nil {
			// Use fmt.Printf instead of log to avoid dependency
			fmt.Printf("Error closing file: %v\n", err)
		}
	}
}

// FileMap is just a convenience type for the map of fileName => fileBytes
type FileMap map[string][]byte

// Write will try to write the bytes from the map into the given writer.
func (fm FileMap) Write(writer io.Writer, filename string) error {
	file, ok := fm[filename]
	if !ok {
		return fmt.Errorf("file not found %s", filename)
	}

	_, err := writer.Write(file)
	if err != nil && err != io.EOF {
		return fmt.Errorf("unable to writeFile '%s': %s", filename, err)
	}
	return nil
}

// readBytes reads an io.Reader into []byte and returns it.
func readBytes(stream io.Reader) []byte {
	buf := new(bytes.Buffer)
	n, err := buf.ReadFrom(stream)

	if n == 0 || err != nil {
		return buf.Bytes()
	}
	return buf.Bytes()
}
