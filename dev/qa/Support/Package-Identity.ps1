# Both EXE flavors must contain the same managed application build.
function Assert-QaPackageBuildIdentity([object[]]$Files) {
    if ($Files.Count -ne 2) { throw 'Expected two production EXE reports' }
    foreach ($file in $Files) {
        if ($file.appBuildId -notmatch '^[0-9a-f]{32}$') { throw 'Missing or invalid application build identity' }
    }
    if ($Files[0].appBuildId -ne $Files[1].appBuildId) { throw 'Production EXEs contain different application builds' }
}
