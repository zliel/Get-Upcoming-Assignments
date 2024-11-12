# Description: This script retrieves upcoming assignments from Canvas LMS and sends an email with the assignments as a PDF attachment.

#Cleanup parameter to remove the markdown and pdf files after sending the email.
param (
    [Alias("c")]
    [switch]$Cleanup = $false
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


function Get-CourseData {
    try {
        # Retrieve the list of courses
        $course_list = Invoke-RestMethod -Uri $courses_uri -Headers $headers -ContentType "application/json"
        
        # Filter out courses that were created after May 1, 2024
        $course_list = $course_list | Where-Object { $_.created_at -gt ($semester_start)}
    
        # Select only the fields we want to display
        $course_list = $course_list | Select-Object id, name, created_at, end_at
        return $course_list
    } catch {
        Write-Host "Error retrieving courses:"
        Write-Host $_.Exception.Message
        if ($null -ne $_.Exception.Response) {
            $responseBody = $_.Exception.Response.Content.ReadAsStringAsync().Result
            Write-Host "Response Body: $responseBody"
        }
    }
}

function Get-Assignments {
    param (
        [String]$course_id
    )
    $assignments_uri = "$base_uri/courses/$course_id/assignments?include=submission&bucket=unsubmitted"
    # Write-Host "Requesting assignments from: $assignments_uri"
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
        Write-Host "Error retrieving assignments:"
        Write-Host $_.Exception.Message
        if ($null -ne $_.Exception.Response) {
            $responseBody = $_.Exception.Response.Content.ReadAsStringAsync().Result
            Write-Host "Response Body: $responseBody"
        }
    }
}

# Fetch all course data
$all_courses = Get-CourseData

$upcoming_assignments = [System.Collections.Hashtable]@{}
# Loop through each course and get the upcoming assignments
$all_courses | ForEach-Object {
    $course_id = $_.id
    $course_name = $_.name
    $course_num = $course_name -split "-" | Select-Object -First 1
    $short_course_name = ($course_name -split ":" | Select-Object -Last 1).Trim()
    $course_name = "$course_num - $short_course_name"

    $assignments = Get-Assignments -course_id $course_id

    $upcoming_assignments[$course_name] = @($assignments)
}

# Print the upcoming assignments in a table
function Show-Table {
    param (
        [System.Collections.Hashtable]$data
    )

    $table = @()

    # Loop through each course and assignment and add it to the table
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

    Write-Host "Total Assignments: $($table.Count)"
}

Show-Table -data $upcoming_assignments


function New-MarkdownTable {
    
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

    return $markdown_table
}


function Save-MarkdownTable {
    param (
        [String]$path,
        [String]$header,
        [String[]]$table
    )

    # Add the header to a markdown file
    $header | Out-File -FilePath $path
    # Add the markdown table to the file
    $table | Out-File -FilePath $path -Append
}

# Save the markdown table to a file with a header
$markdown_path = "upcoming_assignments.md"
$markdown_header = @"
---
title: "**Upcoming Assignments**"
---
"@

$markdown_table = New-MarkdownTable
Save-MarkdownTable -path $markdown_path -header $markdown_header -table $markdown_table


# convert markdown to a pdf
pandoc -f gfm -t pdf -s -o upcoming_assignments.pdf upcoming_assignments.md -V geometry:margin=1in -V text-align:center

# Send the email with the PDF attachment
$credentials = New-Object Management.Automation.PSCredential $From, ($SMTPPassword | ConvertTo-SecureString -AsPlainText -Force)
Send-MailMessage -From $From -to $To -Subject $Subject `
-Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
-Attachments $Attachment -Credential $credentials