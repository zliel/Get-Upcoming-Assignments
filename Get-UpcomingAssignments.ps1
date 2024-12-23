# Description: This script retrieves upcoming assignments from Canvas LMS and sends an email with the assignments as a PDF attachment.

#Cleanup parameter to remove the markdown and pdf files after sending the email.
param (
    [Alias("c")]
    [switch]$Cleanup = $false,
    [Alias("q")]
    [switch]$Quiet = $false
)

# Load configuration from config.json
$config = Get-Content -Raw -Path "config.json" | ConvertFrom-Json

# Variables for retrieving assignments from the Canvas API
$access_token = $config.access_token
$base_uri = "https://canvas.instructure.com/api/v1"
$courses_uri = "$base_uri/courses?enrollment_state=active&include=[concluded]"
$today = Get-Date
$semester_start = Get-Date -Date $config.semester_start
$headers = @{ Authorization = "Bearer $access_token" }

# Variables for sending email
$From = $config.From
$To = $config.To
$Attachment = $config.Attachment
$Subject = "Assignments for the week of $($today.ToString('MM-dd-yyyy'))"
$Body = $config.Body
$SMTPServer = $config.SMTPServer
$SMTPPort = $config.SMTPPort
$SMTPPassword = $config.SMTPPassword

function Write-Log {
    param (
        [string]$message,
        [string] $context
    )

    if ($Quiet) {
        return
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $log_message = "[$timestamp] From: $context - $message"
    
    Write-Host $log_message
}

function cleanup {
    Write-Log "Cleaning up files" "cleanup"
    Remove-Item "upcoming_assignments.md"
    Remove-Item $Attachment
    Write-Log "Cleanup complete" "cleanup"
}


function Get-CourseData {
    Write-Log "Retrieving course data" "Get-CourseData"
    try {
        # Retrieve the list of courses
        $course_list = Invoke-RestMethod -Uri $courses_uri -Headers $headers -ContentType "application/json"
        
        # Filter out courses that were created after May 1, 2024
        $course_list = $course_list | Where-Object { $_.created_at -gt ($semester_start)}
    
        # Select only the fields we want to display
        $course_list = $course_list | Select-Object id, name, created_at, end_at
        return $course_list

    } catch {
        Write-Log "Error retrieving course data" "Get-CourseData"
        Write-Log "$_.Exception.Message" "Get-CourseData"
        if ($null -ne $_.Exception.Response) {
            $responseBody = $_.Exception.Response.Content.ReadAsStringAsync().Result
            Write-Log "Response Body: $responseBody" "Get-CourseData"
        }
    }
}

function Get-Assignments {
    param (
        [String]$course_id
    )
    $assignments_uri = "$base_uri/courses/$course_id/assignments?include=submission&bucket=unsubmitted"
    Write-Log "Retrieving assignments for course $course_id" "Get-Assignments"
    try {
        $assignments = Invoke-RestMethod -Uri $assignments_uri -Headers $headers -ContentType "application/json"
        
        # Convert the due_at field to a local DateTime object (otherwise it's in +6 hours)
        $assignments | ForEach-Object { $_.due_at = [DateTime]$_.due_at.ToLocalTime() }
        $assignments = $assignments | Where-Object { $null -ne $_.due_at }
        $assignments_due_soon = $assignments | Where-Object { 
            ($_.due_at -ge $today.AddDays(-7)) -and ($_.due_at -le $today.AddDays(10))
            $true
        } | Sort-Object due_at
        

        # Select only the fields we want to display
        $assignments_due_soon = $assignments_due_soon | Select-Object id, name, due_at, submission
        return $assignments_due_soon
    } catch {
        Write-Log "Error retrieving assignments:" "Get-Assignments"
        Write-Log "$_.Exception.Message" "Get-Assignments"
        if ($null -ne $_.Exception.Response) {
            $responseBody = $_.Exception.Response.Content.ReadAsStringAsync().Result
            Write-Log "Response Body: $responseBody" "Get-Assignments"
        }
    }
}


function Show-Table {
    param (
        [System.Collections.Hashtable]$data
    )

    $table = @()

    # Loop through each course and assignment and add it to the table
    Write-Log "Displaying table" "Show-Table"
    $data.GetEnumerator() | ForEach-Object {
        $course_name = $_.Key
        $assignments = $_.Value
        $assignments | ForEach-Object {
            $table += [PSCustomObject]@{
                "Course Name" = $course_name
                "Assignment Name" = $_.name
                "Due Date" = $_.due_at.ToString("yyyy-MM-dd")
            }
        }
    }

    # Sort the table by due date and display it
    $table | Sort-Object "Due Date" | Format-Table -AutoSize

    Write-Log "Total Assignments: $($table.Count)" "Show-Table"
}


function New-MarkdownTable {
    param (
        [System.Collections.Hashtable]$upcoming_assignments
    )
    
    Write-Log "Creating markdown table" "New-MarkdownTable"
    $markdown_table = @()
    $markdown_table += "| Course Name | Assignment Name | Due Date |"
    $markdown_table += "|:--|:--|--:|"
    $upcoming_assignments.GetEnumerator() | ForEach-Object {
        $course_name = $_.Key
        $assignments = $_.Value
        $assignments | ForEach-Object {
            $markdown_table += "| $course_name | $($_.name) | $($_.due_at.ToString("MM-dd-yyyy")) |"
        }
    }

    Write-Log "Markdown table created" "New-MarkdownTable"

    return $markdown_table
}


function Save-MarkdownTable {
    param (
        [System.Collections.Hashtable]$upcoming_assignments
    )

    $markdown_path = "upcoming_assignments.md"
    $markdown_header = @"
---
title: "**Upcoming Assignments**"
---
"@

    $markdown_table = New-MarkdownTable -upcoming_assignments $upcoming_assignments
    Show-Table -data $upcoming_assignments

    Write-Log "Saving markdown table to $markdown_path" "Save-MarkdownTable"
    # Add the header to a markdown file
    $markdown_header | Out-File -FilePath $markdown_path
    # Add the markdown table to the file
    $markdown_table | Out-File -FilePath $markdown_path -Append
    Write-Log "Markdown table saved to $markdown_path" "Save-MarkdownTable"
}


function Get-AssignmentData {
    $all_courses = Get-CourseData

    $upcoming_assignments = [System.Collections.Hashtable]@{}
    # Loop through each course and get the upcoming assignments
    Write-Log "Extracting data for upcoming assignments" "Get-AssignmentData"
    $all_courses | ForEach-Object {
        $course_id = $_.id
        $course_name = $_.name
        $course_num = $course_name -split "-" | Select-Object -First 1
        $short_course_name = ($course_name -split ":" | Select-Object -Last 1).Trim()
        $course_name = "$course_num - $short_course_name"

        $assignments = Get-Assignments -course_id $course_id

        $upcoming_assignments[$course_name] = @($assignments)
    }
    Write-Log "Data extracted for upcoming assignments" "Get-AssignmentData"

    return $upcoming_assignments
}


function Convert-ToPDF {
    param (
        [String]$inputFile,
        [String]$outputFile
    )
    Write-Log "Converting markdown to PDF" "Convert-ToPDF"
    pandoc -f gfm -t pdf -s -o $outputFile $inputFile -V geometry:margin=1in -V text-align:center
    Write-Log "PDF file created: $outputFile" "Convert-ToPDF"
}


function Send-Email {
    Write-Log "Sending email" "Send-Email"
    $credentials = New-Object Management.Automation.PSCredential $From, ($SMTPPassword | ConvertTo-SecureString -AsPlainText -Force)
    Send-MailMessage -From $From -to $To -Subject $Subject `
    -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
    -Attachments $Attachment -Credential $credentials
    Write-Log "Email sent" "Send-Email"
}


function Main {
    $upcoming_assignments = Get-AssignmentData
    Save-MarkdownTable -upcoming_assignments $upcoming_assignments
    Convert-ToPDF -inputFile "upcoming_assignments.md" -outputFile $Attachment
    Send-Email
    if ($Cleanup) {
        cleanup
    }
}


Main