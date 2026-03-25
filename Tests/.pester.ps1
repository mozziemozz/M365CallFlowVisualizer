# Pester test runner configuration
# Run with: pwsh -Command "Invoke-Pester -Configuration (. ./Tests/.pester.ps1)"

$config = New-PesterConfiguration
$config.Run.Path = './Tests'
$config.Output.Verbosity = 'Detailed'
$config.TestResult.Enabled = $true
$config.TestResult.OutputFormat = 'NUnitXml'
$config.TestResult.OutputPath = './Tests/TestResults.xml'
$config.CodeCoverage.Enabled = $true
$config.CodeCoverage.Path = @(
    './Functions/Optimize-DisplayName.ps1'
    './Functions/Get-MsSystemMessage.ps1'
    './Functions/Get-IvrTransferMessage.ps1'
    './Functions/Format-PhoneNumber.ps1'
    './Functions/Get-TruncatedString.ps1'
)
$config.CodeCoverage.OutputFormat = 'JaCoCo'
$config.CodeCoverage.OutputPath = './Tests/Coverage.xml'

return $config
