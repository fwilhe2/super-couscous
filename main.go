package main

import (
	"encoding/json"
	"encoding/xml"
	"fmt"
	"os"
	"flag"
)

type Cell struct {
	XMLName     xml.Name `xml:"table:table-cell"`
	Text        string   `xml:"text:p"`
	ValueType   string   `xml:"office:value-type,attr,omitempty"`
	CalcExtType string   `xml:"calcext:value-type,attr,omitempty"`
	Value       string   `xml:"office:value,attr,omitempty"`
	DateValue   string   `xml:"office:date-value,attr,omitempty"`
	TimeValue   string   `xml:"office:time-value,attr,omitempty"`
	Currency    string   `xml:"office:currency,attr,omitempty"`
	StyleName   string   `xml:"table:style-name,attr,omitempty"`
}

type Row struct {
	XMLName xml.Name `xml:"table:table-row"`
	Cells   []Cell   `xml:"table:table-cell"`
}

type Table struct {
	XMLName xml.Name `xml:"table:table"`
	Name    string   `xml:"table:name,attr"`
	Rows    []Row    `xml:"table:table-row"`
}

type FODocument struct {
	XMLName           xml.Name        `xml:"office:document"`
	XMLNSOffice       string          `xml:"xmlns:office,attr"`
	XMLNSTable        string          `xml:"xmlns:table,attr"`
	XMLNSText         string          `xml:"xmlns:text,attr"`
	XMLNSStyle        string          `xml:"xmlns:style,attr"`
	XMLNSFo           string          `xml:"xmlns:fo,attr"`
	XMLNSSvg          string          `xml:"xmlns:svg,attr"`
	XMLNSChart        string          `xml:"xmlns:chart,attr"`
	XMLNSDr3d         string          `xml:"xmlns:dr3d,attr"`
	XMLNSMath         string          `xml:"xmlns:math,attr"`
	XMLNSForm         string          `xml:"xmlns:form,attr"`
	XMLNSScript       string          `xml:"xmlns:script,attr"`
	XMLNSConfig       string          `xml:"xmlns:config,attr"`
	XMLNSXlink        string          `xml:"xmlns:xlink,attr"`
	XMLNSDc           string          `xml:"xmlns:dc,attr"`
	XMLNSMeta         string          `xml:"xmlns:meta,attr"`
	XMLNSNumber       string          `xml:"xmlns:number,attr"`
	XMLNSOf           string          `xml:"xmlns:of,attr"`
	XMLNSXforms       string          `xml:"xmlns:xforms,attr"`
	XMLNSXsd          string          `xml:"xmlns:xsd,attr"`
	XMLNSXsi          string          `xml:"xmlns:xsi,attr"`
	XMLNSGrddl        string          `xml:"xmlns:grddl,attr"`
	XMLNSXhtml        string          `xml:"xmlns:xhtml,attr"`
	XMLNSPresentation string          `xml:"xmlns:presentation,attr"`
	XMLNSCss3t        string          `xml:"xmlns:css3t,attr"`
	XMLNSFormx        string          `xml:"xmlns:formx,attr"`
	XMLNSOooc         string          `xml:"xmlns:oooc,attr"`
	XMLNSOoow         string          `xml:"xmlns:ooow,attr"`
	XMLNSRpt          string          `xml:"xmlns:rpt,attr"`
	XMLNSDraw         string          `xml:"xmlns:draw,attr"`
	XMLNSOoo          string          `xml:"xmlns:ooo,attr"`
	XMLNSCalcext      string          `xml:"xmlns:calcext,attr"`
	XMLNSTableooo     string          `xml:"xmlns:tableooo,attr"`
	XMLNSDrawooo      string          `xml:"xmlns:drawooo,attr"`
	XMLNSLoext        string          `xml:"xmlns:loext,attr"`
	XMLNSDom          string          `xml:"xmlns:dom,attr"`
	XMLNSField        string          `xml:"xmlns:field,attr"`
	OfficeVersion     string          `xml:"office:version,attr"`
	OfficeMimetype    string          `xml:"office:mimetype,attr"`
	AutomaticStyles   AutomaticStyles `xml:"office:automatic-styles"`
	Body              Body            `xml:"office:body"`
}

type AutomaticStyles struct {
	XMLName      xml.Name      `xml:"office:automatic-styles"`
	NumberStyles []NumberStyle `xml:"number:number-style"`
	Styles       []Style       `xml:"style:style"`
}

type NumberStyle struct {
	XMLName        xml.Name        `xml:"number:number-style"`
	Name           string          `xml:"style:name,attr"`
	Volatile       string          `xml:"style:volatile,attr,omitempty"`
	Language       string          `xml:"number:language,attr,omitempty"`
	Country        string          `xml:"number:country,attr,omitempty"`
	TextProperties *TextProperties `xml:"style:text-properties,omitempty"`
	NumberElements []NumberElement `xml:",any"`
	Map            *Map            `xml:"style:map,omitempty"`
}

type TextProperties struct {
	XMLName xml.Name `xml:"style:text-properties"`
	Color   string   `xml:"fo:color,attr,omitempty"`
}

type NumberElement struct {
	XMLName          xml.Name `xml:"number:number"`
	DecimalPlaces    string   `xml:"number:decimal-places,attr"`
	MinDecimalPlaces string   `xml:"number:min-decimal-places,attr"`
	MinIntegerDigits string   `xml:"number:min-integer-digits,attr"`
	Grouping         string   `xml:"number:grouping,attr"`
	Language         string
	Country          string
}

type Map struct {
	XMLName        xml.Name `xml:"style:map"`
	Condition      string   `xml:"style:condition,attr"`
	ApplyStyleName string   `xml:"style:apply-style-name,attr"`
}

type Style struct {
	XMLName         xml.Name `xml:"style:style"`
	Name            string   `xml:"style:name,attr"`
	Family          string   `xml:"style:family,attr"`
	ParentStyleName string   `xml:"style:parent-style-name,attr"`
	DataStyleName   string   `xml:"style:data-style-name,attr"`
}

type Body struct {
	XMLName     xml.Name    `xml:"office:body"`
	Spreadsheet Spreadsheet `xml:"office:spreadsheet"`
}

type Spreadsheet struct {
	XMLName xml.Name `xml:"office:spreadsheet"`
	Tables  []Table  `xml:"table:table"`
}

type CellData struct {
	Value     string `json:"value"`
	ValueType string `json:"valueType"`
}

func jsonToFODS(jsonData string) (string, error) {
	var data [][]interface{}
	err := json.Unmarshal([]byte(jsonData), &data)
	if err != nil {
		return "", err
	}

	var rows []Row
	for _, row := range data {
		var cells []Cell
		for _, cell := range row {
			switch v := cell.(type) {
			case string:
				cells = append(cells, Cell{Text: v, ValueType: "string"})
			case map[string]interface{}:
				var cellData CellData
				cellJSON, _ := json.Marshal(v)
				json.Unmarshal(cellJSON, &cellData)
				cells = append(cells, createCell(cellData))
			}
		}
		rows = append(rows, Row{Cells: cells})
	}

	tables := []Table{
		{
			Name: "Sheet1",
			Rows: rows,
		},
	}

	fods := FODocument{
		XMLNSOffice:       "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		XMLNSTable:        "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		XMLNSText:         "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		XMLNSStyle:        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		XMLNSFo:           "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		XMLNSSvg:          "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
		XMLNSChart:        "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
		XMLNSDr3d:         "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
		XMLNSMath:         "http://www.w3.org/1998/Math/MathML",
		XMLNSForm:         "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
		XMLNSScript:       "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
		XMLNSConfig:       "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
		XMLNSXlink:        "http://www.w3.org/1999/xlink",
		XMLNSDc:           "http://purl.org/dc/elements/1.1/",
		XMLNSMeta:         "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
		XMLNSNumber:       "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		XMLNSOf:           "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
		XMLNSXforms:       "http://www.w3.org/2002/xforms",
		XMLNSXsd:          "http://www.w3.org/2001/XMLSchema",
		XMLNSXsi:          "http://www.w3.org/2001/XMLSchema-instance",
		XMLNSGrddl:        "http://www.w3.org/2003/g/data-view#",
		XMLNSXhtml:        "http://www.w3.org/1999/xhtml",
		XMLNSPresentation: "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
		XMLNSCss3t:        "http://www.w3.org/TR/css3-text/",
		XMLNSFormx:        "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0",
		XMLNSOooc:         "http://openoffice.org/2004/calc",
		XMLNSOoow:         "http://openoffice.org/2004/writer",
		XMLNSRpt:          "http://openoffice.org/2005/report",
		XMLNSDraw:         "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
		XMLNSOoo:          "http://openoffice.org/2004/office",
		XMLNSCalcext:      "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
		XMLNSTableooo:     "http://openoffice.org/2009/table",
		XMLNSDrawooo:      "http://openoffice.org/2010/draw",
		XMLNSLoext:        "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
		XMLNSDom:          "http://www.w3.org/2001/xml-events",
		XMLNSField:        "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
		OfficeVersion:     "1.3",
		OfficeMimetype:    "application/vnd.oasis.opendocument.spreadsheet",
		AutomaticStyles: AutomaticStyles{
			NumberStyles: createNumberStyles(),
			Styles:       createStyles(),
		},
		Body: Body{
			Spreadsheet: Spreadsheet{
				Tables: tables,
			},
		},
	}

	output, err := xml.MarshalIndent(fods, "", "  ")
	if err != nil {
		return "", err
	}

	return xml.Header + string(output), nil
}

func createCell(cellData CellData) Cell {
	cell := Cell{
		Text:      cellData.Value,
		ValueType: cellData.ValueType,
	}
	switch cellData.ValueType {
	case "string":
		cell.CalcExtType = "string"
	case "float":
		cell.CalcExtType = "float"
		cell.StyleName = "FLOAT_STYLE"
		cell.Value = cellData.Value
	case "date":
		cell.CalcExtType = "date"
		cell.StyleName = "DATE_STYLE"
		cell.DateValue = cellData.Value
	case "time":
		cell.CalcExtType = "time"
		cell.StyleName = "TIME_STYLE"
		cell.TimeValue = "PT" + cellData.Value + "S"
	case "currency":
		cell.CalcExtType = "currency"
		cell.StyleName = "EUR_STYLE"
		cell.Value = cellData.Value
		cell.Currency = "EUR"
	case "percentage":
		cell.CalcExtType = "percentage"
		cell.StyleName = "PERCENTAGE_STYLE"
		cell.Value = cellData.Value
	}
	return cell
}

func createNumberStyles() []NumberStyle {
	return []NumberStyle{
		{
			Name:     "___FLOAT_STYLE",
			Volatile: "true",
			NumberElements: []NumberElement{
				{
					DecimalPlaces:    "2",
					MinDecimalPlaces: "2",
					MinIntegerDigits: "1",
					Grouping:         "true",
				},
			},
		},
		{
			Name:           "__FLOAT_STYLE",
			TextProperties: &TextProperties{Color: "#ff0000"},
			NumberElements: []NumberElement{
				{
					DecimalPlaces:    "2",
					MinDecimalPlaces: "2",
					MinIntegerDigits: "1",
					Grouping:         "true",
				},
			},
			Map: &Map{Condition: "value()>=0", ApplyStyleName: "___FLOAT_STYLE"},
		},
		{
			Name: "__DATE_STYLE",
			NumberElements: []NumberElement{
				{XMLName: xml.Name{Local: "number:year"}, DecimalPlaces: "long"},
				{XMLName: xml.Name{Local: "number:text"}, DecimalPlaces: "-"},
				{XMLName: xml.Name{Local: "number:month"}, DecimalPlaces: "long"},
				{XMLName: xml.Name{Local: "number:text"}, DecimalPlaces: "-"},
				{XMLName: xml.Name{Local: "number:day"}, DecimalPlaces: "long"},
			},
		},
		{
			Name: "__TIME_STYLE",
			NumberElements: []NumberElement{
				{XMLName: xml.Name{Local: "number:hours"}, DecimalPlaces: "long"},
				{XMLName: xml.Name{Local: "number:text"}, DecimalPlaces: ":"},
				{XMLName: xml.Name{Local: "number:minutes"}, DecimalPlaces: "long"},
				{XMLName: xml.Name{Local: "number:text"}, DecimalPlaces: ":"},
				{XMLName: xml.Name{Local: "number:seconds"}, DecimalPlaces: "long"},
			},
		},
		{
			Name:     "___EUR_STYLE",
			Volatile: "true",
			Language: "en",
			Country:  "DE",
			NumberElements: []NumberElement{
				{
					DecimalPlaces:    "2",
					MinDecimalPlaces: "2",
					MinIntegerDigits: "1",
					Grouping:         "true",
				},
				{XMLName: xml.Name{Local: "number:text"}},
				{XMLName: xml.Name{Local: "number:currency-symbol"}, DecimalPlaces: "€", Language: "de", Country: "DE"},
			},
		},
		{
			Name:           "__EUR_STYLE",
			Language:       "en",
			Country:        "DE",
			TextProperties: &TextProperties{Color: "#ff0000"},
			NumberElements: []NumberElement{
				{XMLName: xml.Name{Local: "number:text"}, DecimalPlaces: "-"},
				{
					DecimalPlaces:    "2",
					MinDecimalPlaces: "2",
					MinIntegerDigits: "1",
					Grouping:         "true",
				},
				{XMLName: xml.Name{Local: "number:text"}},
				{XMLName: xml.Name{Local: "number:currency-symbol"}, DecimalPlaces: "€", Language: "de", Country: "DE"},
			},
			Map: &Map{Condition: "value()>=0", ApplyStyleName: "___EUR_STYLE"},
		},
		{
			Name: "__PERCENTAGE_STYLE",
			NumberElements: []NumberElement{
				{
					DecimalPlaces:    "2",
					MinDecimalPlaces: "2",
					MinIntegerDigits: "1",
				},
				{XMLName: xml.Name{Local: "number:text"}, DecimalPlaces: "%"},
			},
		},
	}
}

func createStyles() []Style {
	return []Style{
		{Name: "FLOAT_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "__FLOAT_STYLE"},
		{Name: "DATE_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "__DATE_STYLE"},
		{Name: "TIME_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "__TIME_STYLE"},
		{Name: "EUR_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "__EUR_STYLE"},
		{Name: "PERCENTAGE_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "__PERCENTAGE_STYLE"},
	}
}

func main() {

	inputFile := flag.String("input", "", "a string")
	outputFile := flag.String("output", "", "a string")

	flag.Parse()

	dat, err := os.ReadFile(*inputFile)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	fmt.Print(string(dat))

	fods, err := jsonToFODS(string(dat))
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	file, err := os.Create(*outputFile)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer file.Close()

	file.WriteString(fods)
	fmt.Println("FODS file created successfully.")
}
