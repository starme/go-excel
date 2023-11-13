package exstyle

import "fmt"

type OptionType int

const (
	FontSizeOpt OptionType = iota
	FontFamilyOpt
	FontBoldOpt
	FontItalicOpt
	FontColorOpt

	HorizontalOpt
	VerticalOpt
	WrapTextOpt
	TextRotationOpt

	//BgColorOpt
)

// Option specifies the task processing behavior.
type Option interface {
	// String returns a string representation of the option.
	String() string

	// Type describes the type of the option.
	Type() OptionType

	// Value returns a value used to create this option.
	Value() interface{}
}

// Internal option representations.
type (
	sizeOption   float64
	familyOption string
	boldOption   bool
	italicOption bool
	colorOption  string

	horizontalOption   string
	verticalOption     string
	wrapTextOption     bool
	textRotationOption int

	//bgColorOption []string
)

func Size(size float64) Option {
	if size <= 0 {
		size = DefaultFontSize
	}
	return sizeOption(size)
}

func (n sizeOption) String() string     { return fmt.Sprintf("FontSize(%d)", int(n)) }
func (n sizeOption) Type() OptionType   { return FontSizeOpt }
func (n sizeOption) Value() interface{} { return int(n) }

func Family(family string) Option {
	return familyOption(family)
}

func (n familyOption) String() string     { return fmt.Sprintf("FontFamily(%s)", n) }
func (n familyOption) Type() OptionType   { return FontFamilyOpt }
func (n familyOption) Value() interface{} { return string(n) }

func Bold() Option {
	return boldOption(true)
}

func (n boldOption) String() string     { return fmt.Sprintf("FontBold(%t)", n) }
func (n boldOption) Type() OptionType   { return FontBoldOpt }
func (n boldOption) Value() interface{} { return bool(n) }

func Italic() Option {
	return italicOption(true)
}

func (n italicOption) String() string     { return fmt.Sprintf("FontItalic(%t)", n) }
func (n italicOption) Type() OptionType   { return FontItalicOpt }
func (n italicOption) Value() interface{} { return bool(n) }

func Color(color string) Option {
	return colorOption(color)
}

func (n colorOption) String() string     { return fmt.Sprintf("FontColor(%s)", n) }
func (n colorOption) Type() OptionType   { return FontColorOpt }
func (n colorOption) Value() interface{} { return string(n) }

func Horizontal(align string) Option {
	return horizontalOption(align)
}

func (n horizontalOption) String() string     { return fmt.Sprintf("Horizontal(%s)", n) }
func (n horizontalOption) Type() OptionType   { return HorizontalOpt }
func (n horizontalOption) Value() interface{} { return string(n) }

func Vertical(align string) Option {
	return verticalOption(align)
}

func (n verticalOption) String() string     { return fmt.Sprintf("Vertical(%s)", n) }
func (n verticalOption) Type() OptionType   { return VerticalOpt }
func (n verticalOption) Value() interface{} { return string(n) }

func WrapText() Option {
	return wrapTextOption(true)
}

func (n wrapTextOption) String() string     { return fmt.Sprintf("wrapText(%t)", n) }
func (n wrapTextOption) Type() OptionType   { return WrapTextOpt }
func (n wrapTextOption) Value() interface{} { return bool(n) }

func TextRotation(rotation int) Option {
	return textRotationOption(rotation)
}

func (n textRotationOption) String() string     { return fmt.Sprintf("TextRotation(%d)", n) }
func (n textRotationOption) Type() OptionType   { return TextRotationOpt }
func (n textRotationOption) Value() interface{} { return int(n) }

//func BgColor(color []string) Option {
//	return bgColorOption(color)
//}
//
//func (n bgColorOption) String() string     { return fmt.Sprintf("BgColor(%v)", n) }
//func (n bgColorOption) Type() OptionType   { return BgColorOpt }
//func (n bgColorOption) Value() interface{} { return []string(n) }
