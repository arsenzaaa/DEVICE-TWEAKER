function Test-QaObsoleteAffinityButton([string]$Name, [string]$AutomationId) {
    if ($Name -notin @('[ ALL ]','[ NONE ]','[ P-CORES ]','[ CCD0 ]')) { return $false }
    # The category toolbar also has [ ALL ]. Identify its actual control,
    # rather than inferring scope from an off-screen heading's coordinates.
    return -not ($Name -eq '[ ALL ]' -and $AutomationId -eq 'FILTER_ALL')
}
