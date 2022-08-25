#$CSV='C:\Test\DepartmentTests.csv'
$CSV='C:\Test\EmployeeDirectorySync.csv'
$Employees = Import-Csv $CSV -Header EmpNo,FirstName,LastName,MI,Title,Department,Phone,Building,Room,OutsideEmail,TermDate

$Time = Get-Date

$ScriptName = 'Get-Department_Function'
$RunType = 'Test'
$timestamp = Get-Date -Format "MM-dd-yyyy_HHmm"

Start-Transcript -Path "C:\Test\$ScriptName\$RunType.results.$timestamp.txt"

function Set-Department {

    param(
        [ref]$Title,
        [ref]$Division,
        [ref]$Dean,
        [ref]$Department,
        [ref]$Team,
        [ref]$Position,
        [ref]$DeptMdrSG
        )

    #Replace hyphenated instances of Part-Time to keep full position during next data validation phase
    $Title.value = $Title.value.replace('Part-Time','Part Time')
    $Title.value = $Title.value.replace('Part-time','Part Time')
    
    #Remove special characters from titles used to seperate position from department
    $Title.value = $Title.value.Split(",")[0]
    $Title.value = $Title.value.Split("-")[0]
    $Title.value = $Title.value.Split("/")[0]
    $Title.value = $Title.value.Split("\")[0]
    $Title.value = $Title.value.Split("(")[0]
    
    #Trim needs to the last line to get rid of extra white space
    $Title.value = $Title.value.Trim()

    $Team.value = $Department.value
    
    #Start assigning Teams to departments
    Switch ($Department.value) {
        'Highlander Central'                      { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Admissions & Recruitment';                       $DeptMdrSG.value = '' }
        'Student Admissions'                      { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Admissions & Recruitment';                       $DeptMdrSG.value = '' }
        'Student Recruitment'                     { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Admissions & Recruitment';                       $DeptMdrSG.value = '' }
        'Admissions & Recruitment'                { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Admissions & Recruitment';                       $DeptMdrSG.value = '' }
        'Enrollment Systems'                      { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Enrollment Systems';                             $DeptMdrSG.value = '' }
        'Financial Aid'                           { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Financial Aid';                                  $DeptMdrSG.value = '' }
        'Accounts Payable'                        { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Financial Services';                             $DeptMdrSG.value = '' }
        'Business/Financial Reporting'            { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Financial Services';                             $DeptMdrSG.value = '' }
        'Grants & General Ledger'                 { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Financial Services';                             $DeptMdrSG.value = '' }
        'Student Accounts Receivables'            { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Financial Services';                             $DeptMdrSG.value = '' }
        'Financial Services'                      { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Financial Services';                             $DeptMdrSG.value = '' }
        'Human Resources'                         { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Human Resources';                                $DeptMdrSG.value = '' }
        'Administrative Systems'                  { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Information Systems & Services';                 $DeptMdrSG.value = 'Information Systems SG' }
        'Customer Support Services'               { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Information Systems & Services';                 $DeptMdrSG.value = 'Information Systems SG' }
        'Cybersecurity & Online Technology'       { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Information Systems & Services';                 $DeptMdrSG.value = 'Information Systems SG' }
        'Infrastructure'                          { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Information Systems & Services';                 $DeptMdrSG.value = 'Information Systems SG' }
        'Information Systems & Services'          { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Information Systems & Services';                 $DeptMdrSG.value = 'Information Systems SG' }
        'Institutional Resilience Grant'          { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Information Systems & Services';                 $DeptMdrSG.value = 'Information Systems SG' }
        'Website'                                 { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Marketing & Communications';                     $DeptMdrSG.value = '' }
        'Marketing & Communications'              { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Marketing & Communications';                     $DeptMdrSG.value = '' }
        'Marketing & Communications'              { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Marketing & Communications';                     $DeptMdrSG.value = '' }
        'Maintenance'                             { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Physical Plant';                                 $DeptMdrSG.value = '' }
        'Grounds Maintenance'                     { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Physical Plant';                                 $DeptMdrSG.value = '' }
        'Custodial Services'                      { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Physical Plant';                                 $DeptMdrSG.value = '' }
        'Physical Plant'                          { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Physical Plant';                                 $DeptMdrSG.value = '' }
        'Campus Police'                           { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Campus Police';                                  $DeptMdrSG.value = '' }
        'Food Services'                           { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Purchasing & Auxiliary Services';                $DeptMdrSG.value = '' }
        'Purchasing/Auxiliary Services'           { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Purchasing & Auxiliary Services';                $DeptMdrSG.value = ''; $Team.value = 'Purchasing & Auxiliary Services' }
        'Student Records & Registration'          { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Student Records & Registration';                 $DeptMdrSG.value = '' }
        'Finance & Administration'                { $Division.Value = 'Finance & Administration';                                                     $Department.value = 'Finance & Administration';                       $DeptMdrSG.value = '' }
        'Communication Studies'                   { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Language, Literature & Communication';           $DeptMdrSG.value = '' }
        'English'                                 { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Language, Literature & Communication';           $DeptMdrSG.value = '' }
        'French'                                  { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Language, Literature & Communication';           $DeptMdrSG.value = '' }
        'Spanish'                                 { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Language, Literature & Communication';           $DeptMdrSG.value = '' }
        'Language, Literature & Comm'             { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Language, Literature & Communication';           $DeptMdrSG.value = ''; $Team.value = 'Language, Literature & Communication' }
        'Biology'                                 { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Math & Sciences';                                $DeptMdrSG.value = '' }
        'Chemistry'                               { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Math & Sciences';                                $DeptMdrSG.value = '' }
        'Engineering'                             { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Math & Sciences';                                $DeptMdrSG.value = '' }
        'Environmental Science'                   { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Math & Sciences';                                $DeptMdrSG.value = '' }
        'Geography'                               { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Math & Sciences';                                $DeptMdrSG.value = '' }
        'Geology'                                 { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Math & Sciences';                                $DeptMdrSG.value = '' }
        'Mathematics'                             { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Math & Sciences';                                $DeptMdrSG.value = '' }
        'Math & Sciences'                         { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Math & Sciences';                                $DeptMdrSG.value = '' }
        'Physical Education and Health'           { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Physical Education & Health';                    $DeptMdrSG.value = ''; $Team.value = 'Physical Education & Health' }
        'Government'                              { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Social & Behavioral Science';                    $DeptMdrSG.value = '' }
        'History'                                 { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Social & Behavioral Science';                    $DeptMdrSG.value = '' }
        'Philosophy'                              { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Social & Behavioral Science';                    $DeptMdrSG.value = '' }
        'Psychology'                              { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Social & Behavioral Science';                    $DeptMdrSG.value = '' }
        'Sociology'                               { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Social & Behavioral Science';                    $DeptMdrSG.value = '' }
        'Sociology'                               { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Social & Behavioral Science';                    $DeptMdrSG.value = '' }
        'Social/Behavioral Science'               { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Social & Behavioral Science';                    $DeptMdrSG.value = ''; $Team.value = 'Social & Behavioral Science' }
        'Music'                                   { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Visual & Performing Arts';                       $DeptMdrSG.value = '' }
        'Music Industry Careers'                  { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Visual & Performing Arts';                       $DeptMdrSG.value = '' }
        'Theatre'                                 { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Visual & Performing Arts';                       $DeptMdrSG.value = '' }
        'Visual Arts'                             { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Visual & Performing Arts';                       $DeptMdrSG.value = '' }
        'Visual and Performing Arts'              { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Visual & Performing Arts';                       $DeptMdrSG.value = ''; $Team.value = 'Visual & Performing Arts' }
        'Arts & Sciences'                         { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Arts & Sciences';            $Department.value = 'Arts & Sciences';                                $DeptMdrSG.value = '' }
        'Nursing, ADN'                            { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Health Information Technology'           { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Certified Medical Assistant'             { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Medical Laboratory Technician'           { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Occupational Therapy Assistant'          { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Radiologic Technology'                   { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Respiratory Care'                        { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Surgical Technology'                     { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Veterinary Technician Program'           { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Nursing, Vocational'                     { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Health Professions'                      { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Health Professions';         $Department.value = 'Health Professions';                             $DeptMdrSG.value = '' }
        'Adult Education and Literacy'            { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Workforce & Public Service'; $Department.value = 'Adult Education & Literacy';                     $DeptMdrSG.value = ''; $Team.value = 'Adult Education & Literacy'}
        'Alternative Teacher Cert'                { $Division.Value = 'Instruction & Student Engagement'; $Dean.value = 'Workforce & Public Service'; $Department.value = 'Alternative Teacher Certification';              $DeptMdrSG.value = ''; $Team.value = 'Alternative Teacher Certification'}
    
    }

    if(($Department.value -ne $Team.value) -and ($Team.value -ne $null)){
        $Position.value = $Title.value +', '+$Team.value
    }
    else{
        $Position.value = $Title.value
    }
}

#Create an array to store the results
$results = @()
$NoDepartment = @()  

ForEach ($Employee in $Employees)
    {
    $EmpNo       = "0000000" + $Employee.EmpNo
    $EmpNo       = $EmpNo.Substring($EmpNo.Length - 7,7)    
    $TermDate    = $Employee.TermDate
    $FirstName   = $Employee.FirstName
    $LastName    = $Employee.LastName
    $MI          = $Employee.MI
    <#Check for the prcense of a middle name and convert it to an uppercase initial
      Also set the $AccountName variable for use as account name and display name#>
        if ($MI) {
            $MI =($($MI).Substring(0,1)).ToUpper()
        }    
    $Title = $Employee.Title

    $Division = ''
    $Dean = ''
    $Department = $Employee.Department
    $Team = ''
    $Position = ''

    $DeptMdrSG = $null
    
    if(($TermDate -eq '') -or ($TermDate -ge $Time)){
        Set-Department -Title ([ref]$Title) -Division ([ref]$Division) -Dean ([ref]$Dean) -Department ([ref]$Department) -Team ([ref]$Team) -Position ([ref]$Position) -DeptMdrSG ([ref]$DeptMdrSG)
        
        $Data = [PSCustomObject] @{
		    FirstName       = $FirstName
		    MI              = $MI
		    LastName        = $LastName
            Position        = $Position
            Division        = $Division
            Dean            = $Dean
            Department      = $Department
            Team            = $Team
            'M Drive Group' = $DeptMdrSG
        }
        
        If($DeptMdrSG -eq $null){
            #Add data to the $NoDepartment array
            $NoDepartment += $Data            
            
            #We dont have the department from HR mapped anywhere
            Write-Verbose "Send HR a polite email asking them to check the department spelling for: $EmpNo $FirstName $LastName, $Department"

            #Do we continue processing this employee to get them set up in the system or grind everything to a halt because of it like we do for OutsideEmail?
        }

        Else{
            #Add data to the $results array
            $results += $Data          
        }

    }

    else{
        Write-Verbose "$FirstName $LastName - TermDate for this employee is in the past, not processing"
    }
}

Write-Output $results | Sort-Object Division, Dean, Department, Team, Position, LastName | Format-Table -AutoSize
$DepartmetsSorted = $results | Select-Object Team | Sort-Object Team | Get-Unique -AsString
$Sorted = $DepartmetsSorted.count
#Write-Output $DepartmetsSorted.count

Write-Output $NoDepartment | Select-Object Department | Sort-Object Department | Get-Unique -AsString | Format-Table -AutoSize
$DepartmentsLeft = $NoDepartment | Select-Object Department | Sort-Object Department | Get-Unique -AsString
$Left = $DepartmentsLeft.count 
Write-Output "$Sorted teams categorized"
Write-Output "$Left teams left to categorize`n"

Stop-Transcript
Get-ChildItem -Path "C:\Test\$ScriptName" -Recurse -File | Where CreationTime -lt  (Get-Date).AddDays(-1) | Remove-Item -Force