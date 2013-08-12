#==================| Satnaam Waheguru Ji |===============================
#           
#            Author  :     Aman Dhally 
#            E-Mail  :      amandhally@gmail.com 
#            website :     www.amandhally.net 
#            twitter :       @AmanDhally 
#            blog    :       http://newdelhipowershellusergroup.blogspot.in/
#            facebook:   http://www.facebook.com/groups/254997707860848/ 
#            Linkedin:     http://www.linkedin.com/profile/view?id=23651495 
# 
#            Creation Date    : 12-08-2013 
#            File    :  		  OutlookTools.psm1
#            Purpose : 
#            Version : 1 
#
#            My Pet Spider :          /^(o.o)^\  
#========================================================================





$ErrorActionPreference = 'stop'

# Create a new appointments using Powershell
$outlookApplication = New-Object -ComObject 'Outlook.Application'



function New-OutlookCalendarMeeting {

	
<#
	.SYNOPSIS
		Create Outlook Calendar Meetings using Powershell Console.

	.DESCRIPTION
		If you are a Powershell Scripter or Programmer, then most of your time is spent
		On the Powershell Console. I want to write a small function which helps me to
		Create a calendar invites from the Powershell console. So that I can add calendar
		Invites on the fly and add them as reminder.

	.PARAMETER  Subject
		Using -Subject parameter please provide the subject of the calendar meeting.

	.PARAMETER  Body
		Using -Body, you can add a more information in to the calendar invite.

	.PARAMETER  Location
		The location of your Meeting, for example can be meeting room1 or any country.

	.PARAMETER  Importance
		By Default the importance is set to 2 which is normal
		You can set to -Importance high by providing 2 as an argument
    	0 = Low
    	1 = Normal
    	2 = High.


	.PARAMETER  AllDayEvent
		If you want to create an all day event mart it as $true.

    .PARAMETER BusyStatus
        To set your status to Busy, Free Tenative, or out of office, By default it is set to Busy
        0 = Free
        1 = Tentative
        2 = Busy
        3 = Out of Office


	.PARAMETER  EnableReminder
		By Default reminders are enabled. If you don’t want to enable Reminder set it to $false.

	.PARAMETER  MeetingStart
		Provide the Date and time of meeting to start from.

	.PARAMETER  MeetingDuration
		By default meeting duration is set to 30 Minutes. You can change the duration Of the meeting using -MeetingDuration Parameter.

	.PARAMETER  Reminder
		'By default you got reminder before 15 minutes of meting starts. 
         You can use -Reminder to set the reminder duration. The value is in Minutes.'


	.EXAMPLE
		PS C:\>New-CalendarMeeting -Subject "Powershell" -Body "Show how to use Powershell with Outlook" -Location "Conf Room 1" -AllDayEvent $true -EnableReminder $false
		

	.EXAMPLE
		PS C:\>New-CalendarMeeting -Subject "Powershell" -Body "Show how to use Powershell with Outlook" -Location "Conf Room 1" -MeetingStart "08/08/2013 22:30" -Reminder 30 
	

	.EXAMPLE
		PS C:\>New-CalendarMeeting -Subject "Powershell" -Body "Show how to use Powershell with Outlook" -Location "Conf Room 1" -Importance 2 


	.NOTES
        I worte this function for adding a quick calender meeting.
        in this fucntion you can't add anyone and sent invites to someone.
        In next version of the same function , i will add these functionality.
        Thanks : Aman Dhally {amandhally@gmail.com}

	.LINK
		www.amandhally.net

	.LINK
		http://newdelhipowershellusergroup.blogspot.in/
#>
	
	
param (

[cmdletBinding()]

	# Subject Parameter	
    [Parameter(
        Mandatory = $True,
        HelpMessage="Please provide a subject of your calendar invite.")]
    [Alias('sub')]
    [string] $Subject,

	#Body parameter
    [Parameter(
        Mandatory = $True,
        HelpMessage="Please provide a description of your calendar invite.")]
    [Alias('bod')]
    [string] $Body,

	#Location Parameter
    [Parameter(
        Mandatory = $True,
        HelpMessage="Please provide the location of your meeting.")]

    [Alias('loc')]
    [string] $Location,

	# Importance Parameter
	[int] $Importance = 1,

	# All Day event Parameter
	[bool] $AllDayEvent = $false,

	# Set Reminder Parameter
	[bool] $EnableReminder = $True,

	# Busy Status Parameter
	[string] $BusyStatus = 2,

	# Metting Start Time Parameter
	[datetime] $MeetingStart =(Get-Date),

	# Meeting time duration parameter
	[int] $MeetingDuration = 30, 

	# Meeting time End parameter
		#[datetime] $MeetingEnd = (Get-Date).AddMinutes(+30),

	# by Default Reminder Duration
	[int] $Reminder = 15



)

BEGIN { 
        
        Write-Verbose " Creating Outlook as an Object"
        

        # Creating a instatance of Calenders
        $newCalenderItem = $outlookApplication.CreateItem('olAppointmentItem')



      }


PROCESS { 
        
         Write-Verbose "Creating Calender Invite"
    
         $newCalenderItem.AllDayEvent = $AllDayEvent
         $newCalenderItem.Subject = $Subject
         $newCalenderItem.Body = $Body
         $newCalenderItem.Location  = $Location
         $newCalenderItem.ReminderSet = $EnableReminder
         $newCalenderItem.Importance = $importance


         if ( ! ($AllDayEvent)) {

         $newCalenderItem.Start = $MeetingStart
         $newCalenderItem.Duration = $MeetingDuration
         
         }
         
         $newCalenderItem.ReminderMinutesBeforeStart = $Reminder
         # 2 is busy, 3 is ou to office
         $newCalenderItem.BusyStatus = $BusyStatus
             
    }

END {
    
        Write-Verbose "Saving Calender Item"
        $newCalenderItem.Save()
      
        # if you want to see the calener invite un-comment the below line
            #un-comment it ==>  $newCalenderItem.Display($True)

       }

	}

function New-OutlookContact {


<#


.SYNOPSIS
	Create Outlook contacts using Powershell console.


	
	
.PARAMETER FirstName
	This is a mandatory prameter, You have to provide a first Name to 
	add a contact.

.PARAMETER Birthday
	To add a birthdate of the new contact use this parameter, and this parameter
	accept MM/DD/YY format.
	
.PARAMETER	BusinessPhone
	To add a Business phone of the conact use this parameter.
	
.PARAMETER	Company
	To add contact's company name use this parameter.
	
.PARAMETER	EmailAddress
	To add contact's email-ID use this parameter.
	
.PARAMETER	HomeAddress
	To add  contact's home address use this parameter.
	
.PARAMETER	JobTitle
	To add new contact's JOB title use this parameter.
	
.PARAMETER	LastName
	To add new contact's last name use this parameter.
	
.PARAMETER	MobileNumber
	To add new contact mobile numner use this parameter.
	
.PARAMETER	Notes
	To add extra notes/description use this parameter.
	
.PARAMETER	Website

.NOTES
        
	Using this Function, you can add an Outlook Contact on the fly. 
	I converted it to the module. So that i can start adding more functionality to 
	it.

	
	
.EXAMPLE
	New-OutlookContact -FirstName "jhonson" -LastName "Smith"
	
	
.EXAMPLE
	New-OutlookContact -FirstName "Ajit" -LastName "Singh" -MobileNumber "9910129889"
	
	
.EXAMPLE	
	New-OutlookContact -FirstName "Jujhar" -LastName "Singh" -Website "www.js.com" -Company "Jujhar Studio"
	
	
.EXAMPLE	
	New-OutlookContact -FirstName "Jorawar" -LastName "Singh" -BusinessAddress "Sarhind, Punjab" -JobTitle "Warrior" -Notes "I saw him in a Sikh Temple"
	
.EXAMPLE	
	New-OutlookContact -FirstName "Fateh" -LastName "Singh" -Birthday "05/01/1999" -BusinessPhone "0999999" -BusinessAddress 'Sarhind, Punjab' -Company "Khalsa" -EmailAddress 'fateh@khalasa.com' -HomeAddress 'Talwandi,Punjab' -JobTitle 'warror' -MobileNumber "90909090'
	
	
.LINK
	www.amandhally.net

.LINK
	http://newdelhipowershellusergroup.blogspot.in/

#>





#Paramerers
param(



#a
#b
[Alias('bday')]
[datetime]$Birthday,

[Alias('bisph')]
[string]$BusinessPhone, 

[Alias('bisadd')]	
[string]$BusinessAddress,
#c
[Alias('comp')]	
[string]$Company,
#d
#e
[Alias('email')]	
[string]$EmailAddress,
#f
[Parameter(
	Mandatory = $True,
	HelpMessage="Please enter a First name of the person.")]
[Alias('first')]	
[string]$FirstName,
#g
#h
[Alias('homead')]	
[string]$HomeAddress,
#i
#j
[Alias('Jobti')]	
[string]$JobTitle,
#k
#l
[Alias('last')]	
[string]$LastName,
#m
[Alias('mobile')]	
[string]$MobileNumber,
#n
[Alias('note')]	
[string]$Notes,
#o
#p
#q
#r
#s
#t
#u
#v
#w
[Alias('web')]	
[string]$Website
#x
#y
#z

) # enf d paramaters
 



BEGIN {

    try 
        {
        
        $contactObject = $outlookApplication.CreateItem('olContactItem')
        }
    
    catch
        {  
        Write-Warning '$_.exception occured'
        }





} # end of begin


PROCESS 
{


    try 
        {
        
        $contactObject.FirstName = $FirstName
        $contactObject.LastName  =$LastName
        $contactObject.MobileTelephoneNumber = $MobileNumber
        $contactObject.Email1Address = $EmailAddress
        $contactObject.WebPage = $Website
        $contactObject.Body = $Notes
        $contactObject.CompanyName = $Company
        $contactObject.JobTitle = $JobTitle  
        if ( ! ( $Birthday -eq $null)){    
        $contactObject.Birthday = $Birthday
        }
        $contactObject.BusinessTelephoneNumber = $BusinessPhone
        $contactObject.BusinessAddress = $BusinessAddress
        $contactObject.HomeAddress = $HomeAddres
		
        

        }

    catch 
        {
        
         Write-Warning '$_.exception occured'

        } 



} # end of process


END 
{

    try 
        {
        $contactObject.Save()
		Write-Host "Contact $FirstName  created" -ForegroundColor 'Yellow'	
        
        }
            

    catch
        {
        Write-Warning '$_.exception occured'
        }





} # end of end





} # end of the functions













# Exporting Module memebers
Export-ModuleMember -alias * -Function *



