<#
   Carmahn McCalla

   AUG 2022 

   OBJECTIVE: To read the list of Federation IDs from the Excel Sheet to confirm they match the AD logon names (VERIFY ALL LISTED USERS EXIST IN AD)
      
      1. Access the list of usernames from excel sheet
      2. Compare user emails against the entries in the Active Directory to see if they exist
      3. If a name does not exist, store it in a text file called 'not-found.txt'
      4. Using the case equals operand, verify the email is in the proper case  (First.Last@nspower.ca). If not, add to a file called 'not-match-case.txt'

      The idea is that once you know whose email has to be fixed, someone down the line can develop this script to either:

      1: change the name's capitalization automatically in the excel file
      2:or automate a case to change their logon names in the AD           

#>


Import-Module ActiveDirectory

# PART 1 - Access the list of usernames from excel sheet

   #Run the excel app
   $ExcelObj = New-Object -comobject Excel.Application

   #Open a given excel file (PATH CAN BE CHANGED OR ADD CONFIG FILE)
   $ExcelWorkBook = $ExcelObj.Workbooks.Open("C:\Users\id000\Downloads\Field service users1.xlsx")

   #Get the directory info for the file (located in the same path as this script) to store users who are not found 

      #NOT-FOUND FILE

      $currentdirectory = Get-Location

      $currentfile = "not-found.txt"

      $fullfilenameandlocation = Join-Path $currentdirectory $currentfile

      Write-Host "Current Dir: $fullfilenameandlocation"

   #Create a new file from the string created by $fullfilename and overwrite it each time this script is run (to update it)
   New-Item -Path $currentdirectory -Name $currentfile -ItemType "file" -Force

      #NOT-MATCH-CASE FILE

      $currentdirectory2 = Get-Location

      $currentfile2 = "not-match-case.txt"                                             

      $fullfilenameandlocation2 = Join-Path $currentdirectory2 $currentfile2

      Write-Host "Current Dir: $fullfilenameandlocation2"


   #Create a new file from the string created by $fullfilename and overwrite it each time this script is run (to update it)
   New-Item -Path $currentdirectory2 -Name $currentfile2 -ItemType "file" -Force

   #Access a specific sheet on this excel file
   $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Field Service Users")

   #Use variable 'rowcount' to get the number of filled in rows in the sheet
   $rowcount = $ExcelWorkSheet.UsedRange.Rows.Count

   #output number of emails in the excel sheet
   write-output "Number of Users: $rowcount"


# PART 2 - Compare user emails against the entries in the Active Directory to see if they exist

   <# For Loop Description:

      Loop through Column 3 starting at Row 2 of "Field service users1.xlsx" (these cells contain the usernames)

      1. Verify Member in Specified Group

         - Using the list of names in the excel sheet, query the Active Directory for the matching user (by email).
         - If they are found, list their information: Email, Name, Group, Group MemberOf Status (Yes or No)
         - If MemberOf= Yes, move on to next user. If No, store their email in "not-found.txt" and then move on to next user

      2. Verify Email matches proper case (First.Last@nspower.ca)
      
         - Using the list of names in the excel sheet, query the Active Directory for the matching user (by user principal name).
         - Use an operand to verify whether the cases match or not
         -EX: Proper: First.Last@nspower.ca; Improper: FIRST.LAST/@nspowerca
         - If proper case, move on. Otherwise store user in "not-match-case.txt" file

   #>

   #User (Excel sheet) at iteration $i will be stored in $ADuseremail
   
   for($i=2; $i -le $rowcount; $i++){

      #access fedIDs in the excel sheet (Column Name: Federation ID to be used with SSO)
      $ADuseremail = $ExcelWorkSheet.Columns.Item(3).Rows.Item($i).Text

      #write out user $i email
      write-output $ADuseremail

      # PART 3 - #Try to retrieve user. If a name does not exist in the specified group, store it in a text file called 'not-found.txt'     

         Try{
               #create variable ADUserInfo object to control user's properties within the Active Directory
               $ADUserInfo = Get-ADuser -filter { mail -eq $ADuseremail }

               #write out user $i's principal name
               write-output $ADUserInfo.UserPrincipalName

               #ensure user belongs to the specified AD group ("OneG Default" )
               $group = "OneG Default"  

               #get all members from the specified group
               $groupmembers = Get-ADGroupMember -Identity $group -Recursive | Select -ExpandProperty Name

               #check if the $groupmembers list contains user $i, return yes or no

               If ($groupmembers -contains $ADUserInfo.Name ) {

                  write-output "Group: G $group" "MemberOf: Yes `r" #`r`n formats the screen output spacing

               } Else {

                     write-output "G $group" "MemberOf: No `r"

                     #write user $i's email to the file if they were not found in the specified group
                     $ADuseremail | Out-File -Append -FilePath $fullfilenameandlocation      
               }



               # PART 4 - Verify the email is in the proper case


                  #if statement: if user $i does not match the proper case, store in a file                                       
                  #use the case equal operand (-ceq) to compare the excel email with the AD principal name 

                  If ($ADUserInfo.UserPrincipalName -ceq $ADuseremail){

                     write-output "Proper Case: Yes `r`n"

                  } Else {

                     write-output "Proper Case: No `r`n"

                     $ADUserInfo.UserPrincipalName | Out-File -Append -FilePath $fullfilenameandlocation2

                  }


          }

              Catch {
                     #write to the "not-found" file
                     $ADuseremail + "`r" | Out-File -Append -FilePath $
              }

   }
