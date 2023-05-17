package powerpoint

import (
	"math"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/scjalliance/comshim"
)

const programID = "PowerPoint.Application"

// Specifies a tri-state Boolean value.
const (
	msoTrue  int = -1
	msoFalse int = 0
)

type Screenshot struct {
	ScaleWidth  int
	ScaleHeight int
	Index       int
}

// Export exports the slide, using the specified graphics filter, and saves
// the exported file under the specified output file name.
func (s *Screenshot) Export(input, output string) error {
	comshim.Add(1)
	defer comshim.Done()

	var filterName string
	if s := filepath.Ext(output); len(s) > 1 {
		filterName = strings.ToUpper(s[1:])
	} else {
		filterName = "PNG"
	}

	iface, err := oleutil.CreateObject(programID)
	if err != nil {
		return err
	}
	defer func() {
		iface.Release()
		iface = nil
	}()

	app, err := iface.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}
	defer app.Release()

	presentations := oleutil.MustGetProperty(app, "Presentations")
	defer presentations.Clear()

	presentation, err := oleutil.CallMethod(
		presentations.ToIDispatch(),
		"Open",
		input,
		msoTrue,
		msoFalse,
		msoFalse,
	)
	if err != nil {
		return err
	}
	defer func() {
		if presentation.VT != ole.VT_EMPTY {
			_, _ = oleutil.PutProperty(presentation.ToIDispatch(), "Saved", msoFalse)
			_, _ = oleutil.CallMethod(presentation.ToIDispatch(), "Close")
			_ = presentation.Clear()
		}
	}()

	if s.ScaleHeight == 0 || s.ScaleWidth == 0 {
		sm := oleutil.MustGetProperty(presentation.ToIDispatch(), "SlideMaster")
		s.ScaleWidth = (int)(oleutil.MustGetProperty(sm.ToIDispatch(), "Width").Value().(float32))
		s.ScaleHeight = (int)(oleutil.MustGetProperty(sm.ToIDispatch(), "Height").Value().(float32))
		_ = sm.Clear()
	}

	slides := oleutil.MustGetProperty(presentation.ToIDispatch(), "Slides")
	defer slides.Clear()

	count := (int)(oleutil.MustGetProperty(slides.ToIDispatch(), "Count").Val)
	s.Index = int(math.Min(math.Max(float64(s.Index), 1), float64(count)))

	slide := oleutil.MustCallMethod(slides.ToIDispatch(), "Item", 1)
	defer slide.Clear()

	if _, err := oleutil.CallMethod(
		slide.ToIDispatch(),
		"Export",
		output,
		filterName,
		s.ScaleWidth,
		s.ScaleHeight,
	); err != nil {
		return err
	}
	return nil
}

// vim: set tabstop=4 softtabstop=4 shiftwidth=4 noexpandtab textwidth=78 :
// vim: set fileencoding=utf-8 filetype=go foldenable foldmethod=syntax :