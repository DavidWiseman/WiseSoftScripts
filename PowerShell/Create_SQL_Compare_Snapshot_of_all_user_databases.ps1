# Path to RedGate snapper tool (http://labs.red-gate.com/Tools/Details/RedGateSnapper)
$command= "C:\RedGate.SQLSnapper\RedGate.SQLSnapper.exe";
# Root path for snapshots.  Format is [RootPath]\[DatabaseName]\[YEAR-MONTH]\[DatabaseName]_[YEAR][MONTH][DAY]_[HOUR][MIN][SEC].snp
$path="C:\SQL Compare Snapshots\";
# SQL Server instance name. A RedGate schema snapshot is created for all user databases on this instance
$instance="LOCALHOST"

[reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo")
$sqlServer = new-object ("Microsoft.SqlServer.Management.Smo.Server") $instance
foreach($sqlDatabase in $sqlServer.databases) {
    # System databases, offline databases and snapshot databases are excluded
    If (-not $sqlDatabase.IsSystemObject -and $sqlDatabase.IsAccessible -and -not $sqlDatabase.IsDatabaseSnapshot)
    {
    $db=$sqlDatabase.Name;
    $fileName=$db + "_" + (Get-Date).ToString("yyyyMMdd_HHmm") + ".snp";
    $fullPath=[io.path]::Combine($path,$db);
    $fullPath=[io.path]::Combine($fullPath,(Get-Date).ToString("yyyy-MM"));
    
    [io.directory]::CreateDirectory($fullPath);
    $fullPath=[io.path]::Combine($fullPath,"$filename");

    [Array]$args="/server:$instance","/database:$db","/makesnapshot:$fullPath";

    & $command $args;

    }
}
