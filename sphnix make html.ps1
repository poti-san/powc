sphnix\make.bat html
if (Test-Path html -PathType Container) {
	Remove-Item docs
}
Copy-Item -Recurse sphnix/build/html docs
pause