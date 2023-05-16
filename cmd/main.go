package main

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"github.com/spf13/pflag"

	"github.com/optoma-cn/pptscreenshot/pkg/powerpoint"
	"github.com/optoma-cn/pptscreenshot/pkg/version"
)

var instanceName = "main"

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

	if s, _ := fs.GetString("input"); s != "" {
		input = s
	} else if s := fs.Arg(0); s != "" {
		input = s
	} else {
		return fmt.Errorf("the input filepath is required")
	}

	if s, _ := fs.GetString("output"); s != "" {
		output = s
	} else {
		return fmt.Errorf("the output filepath is required")
	}

	if _, err := os.Stat(input); errors.Is(err, os.ErrNotExist) {
		return fmt.Errorf("the input filepath does not exist: %s", input)
	}
	if _, err := os.Stat(output); err == nil && !fs.Changed("force") {
		return fmt.Errorf("the output filepath already exists: %s", output)
	}

	if !filepath.IsAbs(input) {
		absPath, err := filepath.Abs(input)
		if err != nil {
			return err
		}
		input = absPath
	}
	if !filepath.IsAbs(output) {
		absPath, err := filepath.Abs(output)
		if err != nil {
			return err
		}
		output = absPath
	}

	w, _ := fs.GetInt("width")
	h, _ := fs.GetInt("height")

	screenshot := &powerpoint.Screenshot{
		ScaleWidth:  w,
		ScaleHeight: h,
	}

	return screenshot.Export(input, output)
}

func main() {
	if err := run(os.Args[1:]); err != nil {
		fmt.Fprintln(os.Stderr, err.Error())
		os.Exit(1)
	}
}

// vim: set tabstop=4 softtabstop=4 shiftwidth=4 noexpandtab textwidth=78 :
// vim: set fileencoding=utf-8 filetype=go foldenable foldmethod=syntax :
