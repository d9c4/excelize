package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

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
	assert.NoError(t, WriteStructsIntoFile(f, testStructs, &ModelTableOptions{HasHeader: true}))

	rows, err := f.GetRows("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, 3, len(rows))
}

func TestGetTagValues(t *testing.T) {
	type testStruct struct {
		Column1 string `column:"A" columnHeader:"Column 1"`
		Column2 string `column:"B" columnHeader:"Column 2"`
	}
	var t1 testStruct
	fieldColumnMap := getTagValues(t1, "column")
	columnAliasMap := getTagValues(t1, "columnHeader")

	assert.Equal(t, 2, len(fieldColumnMap))
	assert.Equal(t, 2, len(columnAliasMap))
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
	assert.Equal(t, 2, len(rows))
}
