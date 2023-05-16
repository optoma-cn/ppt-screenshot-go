package main

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/scjalliance/comshim"
	"github.com/spf13/pflag"

	"github.com/optoma-cn/pptscreenshot/pkg/version"
)

var instanceName = "main"

// Specifies a tri-state Boolean value.
const (
	msoTrue  int = -1
	msoFalse int = 0
)

func export(input, output string, index, width, height int) error {
	comshim.Add(1)
	defer comshim.Done()

	var filterName string
	ext := filepath.Ext(output)
	if len(ext) > 1 {
		filterName = strings.ToUpper(ext[1:])
	} else {
		filterName = "PNG"
	}

	if index == 0 {
		index = 1
	}

	unknown, err := oleutil.CreateObject("PowerPoint.Application")
	if err != nil {
		return err
	}
	defer func() {
		unknown.Release()
		unknown = nil
	}()

	app, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}
	defer app.Release()

	presentations, err := oleutil.GetProperty(app, "Presentations")
	if err != nil {
		return err
	}
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
		if _, err := oleutil.PutProperty(presentation.ToIDispatch(), "Saved", msoFalse); err != nil {
			logrus.Printf("Failed to set Saved property: %s", err)
		}
		if _, err := oleutil.CallMethod(presentation.ToIDispatch(), "Close"); err != nil {
			logrus.Printf("Failed to close presentation: %s", err)
		}
		_ = presentation.Clear()
	}()

	if width == 0 || height == 0 {
		slideMaster := oleutil.MustGetProperty(presentation.ToIDispatch(), "SlideMaster")
		defer slideMaster.Clear()

		width = int(oleutil.MustGetProperty(slideMaster.ToIDispatch(), "Width").Value().(float32))
		height = int(oleutil.MustGetProperty(slideMaster.ToIDispatch(), "Height").Value().(float32))
	}

	slides := oleutil.MustGetProperty(presentation.ToIDispatch(), "Slides")
	defer slides.Clear()

	slide, err := oleutil.CallMethod(slides.ToIDispatch(), "Item", index)
	if err != nil {
		return err
	}
	defer slide.Clear()

	_ = oleutil.MustCallMethod(
		slide.ToIDispatch(),
		"Export",
		output,
		filterName,
		width,
		height)

	return nil
}

func run(args []string) error {
	fs := pflag.NewFlagSet(instanceName, pflag.ExitOnError)

	fs.Bool("version", false, "print the version number and exit")
	fs.String("input", "", "input presentation filepath")
	fs.String("output", "", "output screenshot image filepath")
	fs.Bool("force", false, "overwrite existing output file")
	fs.Int("width", 0, "output image width")
	fs.Int("height", 0, "output image height")

	_ = fs.Parse(args)

	if v, _ := fs.GetBool("version"); v {
		fmt.Fprintf(os.Stderr, "%s version %s\n", instanceName, version.GetVersion())
		fmt.Fprintf(os.Stderr, "commit %s\n", version.GetGitCommitID())
		fmt.Fprintf(os.Stderr, "build date %s\n", version.GetBuildDate())
		return nil
	}

	var (
		input  string
		output string
	)
	if s, _ := fs.GetString("input"); s == "" {
		input = fs.Arg(0)
	} else {
		input = s
	}

	if input == "" {
		return errors.New("the input file is required")
	}
	if _, err := os.Stat(input); errors.Is(err, os.ErrNotExist) {
		return errors.New("the input file does not exist")
	}
	if !filepath.IsAbs(input) {
		if s, err := filepath.Abs(input); err == nil {
			input = s
		}
	}

	if s, _ := fs.GetString("output"); s == "" {
		return errors.New("the output image file is required")
	} else {
		output = s
	}
	if force, _ := fs.GetBool("force"); !force {
		if _, err := os.Stat(output); errors.Is(err, os.ErrExist) {
			return errors.New("the output image file already exists")
		}
	}
	if !filepath.IsAbs(output) {
		if s, err := filepath.Abs(output); err == nil {
			output = s
		}
	}

	w, _ := fs.GetInt("width")
	h, _ := fs.GetInt("height")

	return export(input, output, 1, w, h)
}

func main() {
	if err := run(os.Args[1:]); err != nil {
		fmt.Fprintln(os.Stderr, err.Error())
		os.Exit(1)
	}
}

// vim: set tabstop=4 softtabstop=4 shiftwidth=4 noexpandtab textwidth=78 :
// vim: set fileencoding=utf-8 filetype=go foldenable foldmethod=syntax :

