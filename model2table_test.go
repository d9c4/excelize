package excelize

import "testing"


func TestWriteStructsIntoFile(t *testing.T) {
	f := NewFile()
	type testStruct struct {
		Column1 string `column:"A" columnHeader:"Column 1"`
		Column2 string `column:"B" columnHeader:"Column 2"`
	}
	var testStructs = []testStruct{
		{Column1: "1", Column2: "2"},
		{Column1: "3", Column2: "4"},
	}
	err := WriteStructsIntoFile(f, testStructs, &ModelTableOptions{HasHeader: true})
	if err != nil {
		t.Errorf("WriteStructsIntoFile returned error: %v", err)
	}
}


func TestGetTagValues(t *testing.T) {
	type testStruct struct {
		Column1 string `column:"A" columnHeader:"Column 1"`
		Column2 string `column:"B" columnHeader:"Column 2"`
	}
	var t1 testStruct
	fieldColumnMap := getTagValues(t1, "column")
	columnAliasMap := getTagValues(t1, "columnHeader")
	if len(fieldColumnMap) != 2 {
		t.Errorf("getTagValues returned wrong number of fields: %v", len(fieldColumnMap))
	}
	if len(columnAliasMap) != 2 {
		t.Errorf("getTagValues returned wrong number of fields: %v", len(columnAliasMap))
	}
}


func TestConstructRows(t *testing.T) {
	type testStruct struct {
		Column1 string `column:"A" columnHeader:"Column 1"`
		Column2 string `column:"B" columnHeader:"Column 2"`
	}
	var testStructs = []testStruct{
		{Column1: "1", Column2: "2"},
		{Column1: "3", Column2: "4"},
	}
	rows := constructRows(testStructs)
	if len(rows) != 2 {
		t.Errorf("constructRows returned wrong number of rows: %v", len(rows))
	}
}
