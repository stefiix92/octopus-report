param(
    [string]$BaseUrl,
    [string]$ApiKey
)

function Get-PackagesUsedByProjects {
    param(
        [string[]]$IgnoredProjects = @(),
        [string[]]$IgnoredPackages = @()
    )
    Write-Output "Getting report of: PackagesUsedByProjects"

    $packagesSheet = $workbook.Worksheets.Add()
    $packagesSheet.Name = 'Packages by project'
    $packagesSheet.Cells.Item(1,1) = 'Project Id'
    $packagesSheet.Cells.Item(1,2) = 'Project Name'
    $packagesSheet.Cells.Item(1,3) = 'Package Name'

    $row = 2
    foreach ($project in $repository.Projects.GetAll()) {
        if ($ignoredProjects -contains $project.Name) {
            continue
        }
        $processId = $project.DeploymentProcessId
        $process = $repository.DeploymentProcesses.FindAll() | Where-Object { $_.Id -eq $processId }
        foreach ($step in $process.Steps) {
            $step.Actions | ForEach-Object { 
                $_.Packages.Where( { $IgnoredPackages -notcontains $_.PackageId } ) | ForEach-Object { 
                    $packagesSheet.Cells.Item($row,1) = $project.Id
                    $packagesSheet.Cells.Item($row,2) = $project.Name
                    $packagesSheet.Cells.Item($row,3) = $_.PackageId

                    $row += 1
                } 
            }
        }
        #break
    }
    $packagesSheet.UsedRange.Columns.Autofit() | Out-Null
}

function Get-TenantInstallParameters {
    Write-Output "Getting report of: TenantInstallParameters"

    $sheetOrder += 1
    $variablesSheet = $workbook.Worksheets.Add()
    $variablesSheet.Name = 'Tenant Project Variables'
    $variablesSheet.Cells.Item(1,1) = 'Project Id'
    $variablesSheet.Cells.Item(1,2) = 'Project Name'
    $variablesSheet.Cells.Item(1,3) = 'Tenant Id'
    $variablesSheet.Cells.Item(1,4) = 'Tenant Name'
    $variablesSheet.Cells.Item(1,5) = 'Environment Id'
    $variablesSheet.Cells.Item(1,6) = 'Environment Name'
    $variablesSheet.Cells.Item(1,7) = 'Variable Label'
    $variablesSheet.Cells.Item(1,8) = 'Variable Name'
    $variablesSheet.Cells.Item(1,9) = 'Variable Value'

    $row = 2
    $tenantVariables = $repository.TenantVariables.GetAll()
    $environments = Get-EnvironmentsAsDictionary
    foreach ($item in $tenantVariables) {

        foreach ($projectId in $item.ProjectVariables.Keys) {
            $project = $item.ProjectVariables[$projectId]

            foreach ($environmentId in $project.Variables.Keys) {
                $projectVariables = $project.Variables[$environmentId]
                foreach ($key in $projectVariables.Keys) {
                    $template = $project.Templates | Where-Object {$_.Id -eq $key}

                    $variablesSheet.Cells.Item($row,1) = $project.ProjectId
                    $variablesSheet.Cells.Item($row,2) = $project.ProjectName
                    $variablesSheet.Cells.Item($row,3) = $item.TenantId
                    $variablesSheet.Cells.Item($row,4) = $item.TenantName
                    $variablesSheet.Cells.Item($row,5) = $environmentId
                    $variablesSheet.Cells.Item($row,6) = $environments[$environmentId]
                    $variablesSheet.Cells.Item($row,7) = $template.Label
                    $variablesSheet.Cells.Item($row,8) = $template.Name
                    $variablesSheet.Cells.Item($row,9) = $projectVariables[$key].Value

                    $row += 1
                }
            }
        }
        #break
    }
    $variablesSheet.UsedRange.Columns.Autofit() | Out-Null
}

function Get-TenantVariables {
    Write-Output "Getting report of: TenantVariables"

    $sheetOrder += 1
    $variablesSheet = $workbook.Worksheets.Add()
    $variablesSheet.Name = 'Tenant Variables'
    $variablesSheet.Cells.Item(1,1) = 'Tenant Id'
    $variablesSheet.Cells.Item(1,2) = 'Tenant Name'
    $variablesSheet.Cells.Item(1,3) = 'Variable Label'
    $variablesSheet.Cells.Item(1,4) = 'Variable Name'
    $variablesSheet.Cells.Item(1,5) = 'Variable Value'

    $row = 2
    $tenantVariables = $repository.TenantVariables.GetAll()
    foreach ($item in $tenantVariables) {
        foreach ($libraryKey in $item.LibraryVariables.Keys) {
            $library = $item.LibraryVariables[$libraryKey]
            foreach ($key in $library.Variables.Keys) {
                $template = $library.Templates | Where-Object {$_.Id -eq $key}

                $variablesSheet.Cells.Item($row,1) = $item.TenantId
                $variablesSheet.Cells.Item($row,2) = $item.TenantName
                $variablesSheet.Cells.Item($row,3) = $template.Label
                $variablesSheet.Cells.Item($row,4) = $template.Name
                $variablesSheet.Cells.Item($row,5) = $library.Variables[$key].Value

                $row += 1
            }
        }
        # break
    }
    $variablesSheet.UsedRange.Columns.Autofit() | Out-Null
}

function Get-TenantsLinkedToProjects {
    Write-Output "Getting report of: TenantsLinkedToProjects"

    $sheetOrder += 1
    $variablesSheet = $workbook.Worksheets.Add()
    $variablesSheet.Name = 'Tenant Projects'
    $variablesSheet.Cells.Item(1,1) = 'Project Id'
    $variablesSheet.Cells.Item(1,2) = 'Project Name'
    $variablesSheet.Cells.Item(1,3) = 'Tenant Id'
    $variablesSheet.Cells.Item(1,4) = 'Tenant Name'
    $variablesSheet.Cells.Item(1,5) = 'Environment Id'
    $variablesSheet.Cells.Item(1,6) = 'Environment Name'

    $tenants = $repository.Tenants.GetAll()

    $row = 2
    $environments = Get-EnvironmentsAsDictionary
    $projects = Get-ProjectsAsDictionary

    foreach($tenant in $tenants) {
        $projectIds = $tenant.ProjectEnvironments.Keys
        
        foreach($projectId in $projectIds) {
            $linkedEnvironmentIds = $tenant.ProjectEnvironments[$projectId] | ForEach-Object { $_ }
            foreach($environmentId in $linkedEnvironmentIds) {
                $variablesSheet.Cells.Item($row,1) = $projectId
                $variablesSheet.Cells.Item($row,2) = $projects[$projectId]
                $variablesSheet.Cells.Item($row,3) = $tenant.Id
                $variablesSheet.Cells.Item($row,4) = $tenant.Name
                $variablesSheet.Cells.Item($row,5) = $environmentId
                $variablesSheet.Cells.Item($row,6) = $environments[$environmentId]

                $row += 1
            }
        }
        #break
    }
    $variablesSheet.UsedRange.Columns.Autofit() | Out-Null
}

function Get-ProjectsAsDictionary {
    $result = @{}
    foreach ($project in $repository.Projects.GetAll()) {
        $result[$project.Id] = $project.Name
    }
    return $result
}

function Get-EnvironmentsAsDictionary {
    $result = @{}
    foreach ($environment in $repository.Environments.GetAll()) {
        $result[$environment.Id] = $environment.Name
    }
    return $result
}

$path = Join-Path (Get-Item ((Get-Package Octopus.Client).source)).Directory.FullName "lib/net452/Octopus.Client.dll"
Add-Type -Path $path

$server = "$BaseUrl/api/reporting/deployments/xml"

$excelFile = "$PSScriptRoot\report.xlsx"

if (Test-Path $excelFile)
{
    Remove-Item $excelFile
}

$excel = New-Object -ComObject excel.application 
$excel.visible = $False
$workbook = $excel.Workbooks.Add()

$endpoint = New-Object Octopus.Client.OctopusServerEndpoint($server, $ApiKey)
$repository = New-Object Octopus.Client.OctopusRepository($endpoint)

Get-PackagesUsedByProjects -IgnoredProjects @('Deploy All', 'Core All', 'EFM All', 'Reporting All') -IgnoredPackages @('HostingScripts')

Get-TenantsLinkedToProjects

Get-TenantInstallParameters

Get-TenantVariables

$workbook.SaveAs($excelFile)

$excel.Workbooks.Close()
$excel.Quit()