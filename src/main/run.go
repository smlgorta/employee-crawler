package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io/ioutil"
	"strings"
)

const path = "resources/profiles/"

type Document struct {
	XMLName   xml.Name
	Body struct {
		XMLName   xml.Name
		Tbl []struct {
			Tr []Tr `xml:"tr"`
		}`xml:"tbl"`
	}`xml:"body"`
}

type Tr struct {
	Tc []struct {
		InnerXml string `xml:",innerxml"`
	}`xml:"tc>p>r>t"`
}

func main(){
	files, err := ioutil.ReadDir(path)
	if err != nil {
		fmt.Println("Dir reading error", err)
		return
	}
	for _, f := range files {
		byteContents, _, err := readDocxFileContents(path + f.Name())
		if err != nil {
			fmt.Println("File reading error", err)
			return
		}
		content := Document{}
		xml.Unmarshal(byteContents, &content)
		firstName, lastName := parseName(f.Name())
		for _, table := range content.Body.Tbl {
			for _, tr := range table.Tr {
				if containsCompetences(tr) {
					for _, columnContent := range tr.Tc {
						//if index > 1 { // second column
							competences := strings.Split(columnContent.InnerXml, ",")
							for _, competence := range competences {
								createCompetence(firstName, lastName, competence)
							}
						//}
					}
				}
			}
		}
	}
}




func readDocxFileContents(path string) (contents []byte, contentString string, err error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, "", err
	}
	for _, file := range reader.File {
		if file.Name == "word/document.xml" {
			readCloser, err := file.Open()
			if err != nil {
				return nil, "", err
			}
			buf := new(bytes.Buffer)
			buf.ReadFrom(readCloser)
			contents = buf.Bytes()
			contentString = buf.String()
			readCloser.Close()
			break
		}
	}
	return contents, contentString, err
}

func containsCompetences(tr Tr) bool{
	for _, tc := range tr.Tc {
		if strings.Contains(tc.InnerXml, "Fachkompetenz") ||
			strings.Contains(tc.InnerXml, "Methodenkompetenz") ||
			strings.Contains(tc.InnerXml, "Kompetenz"){
			return true
		}
	}
	return false
}

func createCompetence(firstName string, lastName string, competenceName string) {
	trimmedCompetenceName := strings.Trim(competenceName, " ")
	if !strings.Contains(strings.ToLower(trimmedCompetenceName), "kompetenz") &&
		!strings.Contains(strings.ToLower(trimmedCompetenceName), "technische") &&
		trimmedCompetenceName != "" {
		fmt.Println(firstName, lastName, trimmedCompetenceName)
	}
}

func parseName(fileName string) (firstName string, lastName string) {
	split := strings.Split(fileName, "_")
	return split[1], split[0]
}
