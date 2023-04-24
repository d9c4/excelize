package excelize

import (
	"errors"
	"fmt"
	"reflect"
	"strconv"
)

type modelTableRow struct {
	Index int
	Cols  []modelTableCol
}

type modelTableCol struct {
	ColName   string
	FieldName string
	Value     interface{}
}

type ModelTableOptions struct {
	HasHeader bool
}

// WriteStructsIntoFile
// writes a list of structs into an excel file
// the struct must have the tag "column" with the name of the column to write
// header is optional and can be specified with the tag "columnHeader"
func WriteStructsIntoFile[T any](f *File, structs []T, o *ModelTableOptions) error {
	if f == nil {
		return errors.New("nil file")
	}

	var rows = constructRows(structs)

	index := f.GetActiveSheetIndex()
	sheetName := f.GetSheetName(index)
	if o != nil && o.HasHeader {
		writeHeader[T](f, sheetName)
	}
	writeRows(f, rows, sheetName, o)

	return nil
}

// writeHeader writes the header of the columns putting by default the name of the field or eventually the value
// specified by the columnHeader tag
func writeHeader[T any](f *File, sheetName string) error {
	var t T
	fieldColumnMap := getTagValues(t, "column")
	columnAliasMap := getTagValues(t, "columnHeader")
	for k, v := range fieldColumnMap {
		cellReference := v + strconv.Itoa(1)
		headerColName, found := columnAliasMap[k]
		var err error
		if found {
			f.SetCellValue(sheetName, cellReference, headerColName)
		} else {
			f.SetCellValue(sheetName, cellReference, k)
		}
		if err != nil {
			return err
		}
	}
	return nil
}

// writeRows writes the rows in the file, the row index changes depending on the presence (or not) of the header
func writeRows(f *File, rows []modelTableRow, sheetName string, o *ModelTableOptions) error {
	for i, r := range rows {
		for _, c := range r.Cols {
			var cellReference string
			if o != nil && o.HasHeader {
				cellReference = c.ColName + strconv.Itoa(i+2)
			} else {
				cellReference = c.ColName + strconv.Itoa(i+1)
			}
			cellValue := c.Value
			f.SetCellValue(sheetName, cellReference, cellValue)
		}
	}
	return nil
}

// getFieldValue
// given a struct and the name of a field returns the value of that field for that struct
// and a boolean that indicates whether the field has been found or not
func getFieldValue(i interface{}, fieldName string) (interface{}, bool) {
	v := reflect.ValueOf(i)
	if v.Kind() == reflect.Ptr {
		v = v.Elem()
	}
	if v.Kind() != reflect.Struct {
		return nil, false
	}
	f := v.FieldByName(fieldName)
	if !f.IsValid() {
		return nil, false
	}
	return f.Interface(), true
}

// getTagValues
// returns a map of string to string with the field name as key and the value of the requested tag as value
func getTagValues(i interface{}, tagName string) map[string]string {
	v := reflect.ValueOf(i)
	if v.Kind() == reflect.Ptr {
		v = v.Elem()
	}
	if v.Kind() != reflect.Struct {
		return nil
	}
	t := v.Type()
	fieldCount := v.NumField()
	tagValues := make(map[string]string)
	for i := 0; i < fieldCount; i++ {
		field := t.Field(i)
		tag := field.Tag.Get(tagName)
		if tag != "" {
			tagValues[field.Name] = tag
		}
	}
	return tagValues
}

// constructRows constructs the "rows" of the excel by taking the column name
// and retrieving the value of the struct field from the field name
func constructRows[T any](structs []T) []modelTableRow {
	var rows []modelTableRow

	for i, item := range structs {
		fieldColumnMap := getTagValues(item, "column")
		var cols []modelTableCol
		for fieldName, columnName := range fieldColumnMap {
			value, b := getFieldValue(item, fieldName)
			if b {
				column := constructCol(columnName, fieldName, value)
				cols = append(cols, column)
			}
		}
		toAdd := modelTableRow{
			Index: i,
			Cols:  cols,
		}
		rows = append(rows, toAdd)
	}

	return rows
}

// constructCol constructs a column of the excel by taking the column name
// and retrieving the value of the struct field from the field name
func constructCol(columnName string, fieldName string, value interface{}) modelTableCol {
	var parsed interface{}
	kind := reflect.ValueOf(value).Type().Kind()

	// controllo se il valore è un puntatore vuoto
	if kind == reflect.Pointer {
		if reflect.ValueOf(value).IsNil() {
			parsed = ""
		} else {
			parsed = reflect.ValueOf(value).Elem().Interface()
		}
	} else if kind == reflect.Struct {
		//nel caso di una struct innestata prendo gli eventuali valori da mostrare...
		structItemsMap := getTagValues(value, "columnInnerValue")
		if len(structItemsMap) > 0 {
			parsed = ""
			for k, v := range structItemsMap {
				fieldValue, found := getFieldValue(value, k)
				if found {
					//...e se trovati li formatto come stringa
					parsed = fmt.Sprintf("%s %s %v \n", parsed, v, fieldValue)
				}
			}
		} else {
			parsed = value
		}
	} else {
		parsed = value
	}

	column := modelTableCol{
		ColName:   columnName,
		FieldName: fieldName,
		Value:     parsed,
	}
	return column
}
