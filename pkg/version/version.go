package version

var (
	version     = "v0.0.0-unset"
	buildDate   = ""
	gitCommitID = ""
)

func GetVersion() string {
	return version
}

func GetBuildDate() string {
	return buildDate
}

func GetGitCommitID() string {
	return gitCommitID
}

// vim: set tabstop=4 softtabstop=4 shiftwidth=4 noexpandtab textwidth=78 :
// vim: set fileencoding=utf-8 filetype=go foldenable foldmethod=syntax :
