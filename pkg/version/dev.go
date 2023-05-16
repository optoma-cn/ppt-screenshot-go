//go:build !production
// +build !production

package version

import "time"

func init() {
	if gitCommitID == "" {
		gitCommitID = "dev"
	}
	if buildDate == "" {
		buildDate = time.Now().Format(time.RFC3339)
	}
}

// vim: set tabstop=4 softtabstop=4 shiftwidth=4 noexpandtab textwidth=78 :
// vim: set fileencoding=utf-8 filetype=go foldenable foldmethod=syntax :
