Describe "Required PowerShell module" {
    Context "Check 'ImportExcel' is installed" {
        It "is Module Found" {
            # Act
            $module = Get-Module -Name ImportExcel -ListAvailable

            # Assert
            $module | Should -Not -BeNullOrEmpty

        }
    }
}
