# PowerShell scripts for SharePoint
PowerShell scripts to work with SharePoint site

## Usage

### Create a SharePoint List
```
New-SPOList -ListName "EmployeeMaster" -ListTitle "Employee Master"
```

### Add a fields to a SharePoint List
```
Add-SPOListColumn-Text -ListName "EmployeeMaster" -FieldName "FirstName" -DisplayName "First Name" -Description "First Name" -IsRequired $true
Add-SPOListColumn-Text -ListName "EmployeeMaster" -FieldName "LastName" -DisplayName "Last Name" -Description "Last Name" -IsRequired $true
Add-SPOListColumn-DateTime -ListName "EmployeeMaster" -FieldName "Birthday" -DisplayName "Birthday" -Description "Birthday" -IsRequired $true
Add-SPOListColumn-Choice -ListName "EmployeeMaster" -FieldName "Gender" -DisplayName "Gender" -Description "Gender" -IsRequired $true -Choices "Male","Female" -DefaultChoice "Male"
Add-SPOListColumn-Boolean -ListName "EmployeeMaster" -FieldName "IsMarried" -DisplayName "Married" -Description "Married" -IsRequired $true -DefaultValue 0
Add-SPOListColumn-Number -ListName "EmployeeMaster" -FieldName "Salary" -DisplayName "Salary" -Description "Salary" -IsRequired $true
Add-SPOListColumn-Url -ListName "EmployeeMaster" -FieldName "Website" -DisplayName "Website" -Description "Website" -IsRequired $false
```
