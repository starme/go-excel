package exerrors

import (
	"fmt"
	"strings"
)

type Validated struct {
	Msg []string
}

func (err *Validated) Append(msg string) {
	err.Msg = append(err.Msg, msg)
}

func (err *Validated) Error() string {
	if len(err.Msg) == 0 {
		return ""
	}
	return fmt.Sprintf("validate error: \n%s", strings.Join(err.Msg, "\n"))
}
