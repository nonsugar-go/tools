package dataframe

import "github.com/xuri/excelize/v2"

// Header represents column headers.
type Header struct {
	Name       string
	ColumnName string
	Col        int
	IsArray    bool
}

// Record represents a single row of data.
type Record []string

// Records represents a collection of multiple rows.
type Records []Record

// DataFrame represents a structured collection of data.
type DataFrame struct {
	Headers []Header
	Records Records
}

// New initializes a DataFrame using specified column letters and their
// corresponding names.
//
// Parameters:
//
//	columns - A list alternating between column letters (e.g., "B") and logical
//	names (e.g., "ID").
//
// Example:
//
//	df := New("B", "ID", "E", "Name", "I", "Age")
func New(columns ...string) *DataFrame {
	df := &DataFrame{Headers: make([]Header, 0)}
	h := Header{}
	for i := 0; i < len(columns); i += 2 {
		h.ColumnName = columns[i]
		var err error
		h.Col, err = excelize.ColumnNameToNumber(h.ColumnName)
		if err != nil {
			return nil
		}
		h.Name = columns[i+1]
		df.Headers = append(df.Headers, h)
	}
	return df
}

// Add adds a new record (row) to the dataset.
//
// Example:
//
//	df.Add("3", "Dog", 4).Add("4", "Cat", 3)
func (df *DataFrame) Add(record ...string) *DataFrame {
	df.Records = append(df.Records, Record(record))
	return df
}
