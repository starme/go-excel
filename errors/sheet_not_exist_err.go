package exerrors

import "fmt"

// ErrSheetNotExist defines an error of sheet that does not exist
type ErrSheetNotExist struct {
	SheetName string
}

func (err ErrSheetNotExist) Error() string {
	return fmt.Sprintf("sheet [%s] does not exist", err.SheetName)
}
