package exstyle

import (
	"github.com/xuri/excelize/v2"
)

const (
	AlignCenter = "center"
	AlignLeft   = "left"
	AlignRight  = "right"
	AlignTop    = "top"
	AlignBottom = "bottom"

	DefaultFontSize = 20.00
	DefaultFontBold = false
)

type Style struct {
	excelize.Style
}

// SetFont 设置字体
func (s *Style) SetFont(opts ...Option) *Style {
	var font = &excelize.Font{}
	for _, opt := range opts {
		switch opt := opt.(type) {
		case sizeOption:
			font.Size = float64(opt)
		case familyOption:
			font.Family = string(opt)
		case boldOption:
			font.Bold = bool(opt)
		case italicOption:
			font.Italic = bool(opt)
		case colorOption:
			font.Color = string(opt)
		default:
			// ignore unexpected option
		}
	}

	if font != nil {
		s.Font = font
	}
	return s
}

func (s *Style) SetAlign(opts ...Option) *Style {
	var align = &excelize.Alignment{
		Indent:          1,
		JustifyLastLine: true,
		ReadingOrder:    0,
		RelativeIndent:  1,
		ShrinkToFit:     true,
	}

	for _, opt := range opts {
		switch opt := opt.(type) {
		case horizontalOption:
			align.Horizontal = string(opt)
		case verticalOption:
			align.Vertical = string(opt)
		case wrapTextOption:
			align.WrapText = bool(opt)
		case textRotationOption:
			align.TextRotation = int(opt)
		default:
			// ignore unexpected option
		}
	}

	if align != nil {
		s.Alignment = align
	}
	return s
}

func (s *Style) SetBgColor(color []string) *Style {
	s.Fill = excelize.Fill{
		Type:    "pattern",
		Pattern: 1,
		Color:   color,
	}
	return s
}

// ApplyStyle 应用样式
/**
 *  @Description: 获取Style对象
 *  @receiver s
 *  @param f
 *  @return style
 *  @return err
 */
func (s *Style) ApplyStyle(f *excelize.File, sheet, hCell, vCell string) error {
	styleId, err := f.NewStyle(&s.Style)
	if err != nil {
		return err
	}

	return f.SetCellStyle(sheet, hCell, vCell, styleId)
}

//func title() (s *Style) {
//	return s.SetAlignCenter().SetWrapText().SetFontBold().SetFontSize(DefaultFontSize)
//}
//
//func DefaultHeaderStyle() (s *Style) {
//	return title()
//}
