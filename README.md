# PowerShellStuff
Various PowerShell scripts and modules that may be useful to the public

My goal is to store and track changes on scripts that I have written, as long as they are generic and well-written.  I will not share company-specific information on here.  
_Some of these items were borrowed from or inspired by other sources, and I have marked that in the NOTES of those functions._  
Hopefully, these can be useful in other contexts.  For example, Active Directory, Hyper-V, vmware, Office 365, file and print servers, etc.

My second goal is to get used to GitHub so that I can contribute with more quality and confidence to public projects. 

Here are available functions:

### add-LocalGroupMember
    
SYNOPSIS
    Remotely adds a user to a local group, assuming adequate permissions
    
    
SYNTAX
    `add-LocalGroupMember [-ComputerName] <String[]> [-Identity] <String> [-Group] <String> [[-Domain] <String>] 
    [<CommonParameters>]`
    
    
DESCRIPTION
    Add users to a local group group remotely; firewall must [allow access to] remote pc
    


###    connect-Office365Session
    
SYNOPSIS
    Use this function to connect to Office 365 PowerShell session; supply credentials, or else it will prompt for them.
    
    
SYNTAX
    `connect-Office365Session [-credential] <PSCredential> [<CommonParameters>]`
    
    
DESCRIPTION
    Given a credential, connects to Office 365 management PowerShell; then can run exchange cmdlets on Office 365.
    Loads a session and a temp module based on user logging on.
    

###    convertTo-ByteString
    
SYNOPSIS
    Converts an integer into a "byte string", like 1GB
    
    
SYNTAX
    `convertTo-ByteString [[-Value] <Int64>] [[-Round] <Int32>] [<CommonParameters>]`
    
    
DESCRIPTION
    Given an integer value, divides by 1024 until the number is between 0 and 1024, then attaches the appropriate byte 
    measurement.  Using Invoke-Expression will convert back to a number. 
    Returns a string with a number and byte measurement abbreviation, the kind that PowerShell resolves.  See examples.
    The largest possible Value is 9223372036854775807 (according to [int64]::maxvalue). In simple terms, most 19-digit 
    numbers and smaller will work.
    

###    get-ComputerNames
    
SYNOPSIS
    Returns an array of "computer name" strings of the same series, such as "000-lab-nn".
    
    
SYNTAX
    `get-ComputerNames [-Prefix] <String> [[-StartNum] <Int32>] [-EndNum] <Int32> [-NoLeadingZero] [[-Exclude] <Int32[]>] [-TestConnection] [<CommonParameters>]`
    
    
DESCRIPTION
    Given a prefix, start number, and end number, returns a list of computer names with the last number part running 
    the range of the start and end numbers (inclusive).
    

###    get-ComputerSpecs
    
SYNOPSIS
    Returns OS, RAM, and hard drive specs for one or more computers.
    
    
SYNTAX
   `get-ComputerSpecs [[-ComputerName] <String[]>] [<CommonParameters>]`
    
    
DESCRIPTION
    This script/function gets specifications about a PC, including OS, RAM, CPU, and hard drive(s), using WMI queries.
    

###    Get-LastBootTime
    
SYNOPSIS
    Gets the last boot-up time for the specified servers
    
    
SYNTAX
    `Get-LastBootTime [[-ComputerName] <String[]>] [<CommonParameters>]`
    
    
DESCRIPTION
    Using CIM, connects to the remote computer(s) and returns a DateTime object representing the moment the server 
    last started and the timespan between now and then.
    Able to resolve IP addresses to names.
    

###    get-LocalGroupMember
    
SYNOPSIS
    Remotely checks membership of a local group, assuming adequate permissions
    
    
SYNTAX
    `get-LocalGroupMember [-Group] <String> [[-ComputerName] <String[]>] [-Indirect] [<CommonParameters>]`
    
    
DESCRIPTION
    Given a computer name and group name, such as administrators, lists those accounts and basic info about them.
    Use -Indirect to show all users and groups nested within the local group. Recommended to pipe results to 
    Format-Table.
    
###    measure-Path
    
SYNOPSIS
    Finds the total size of the given file or folder
    
    
SYNTAX
    `measure-Path [[-Path] <String[]>] [[-UnitSize] <String>] [<CommonParameters>]`
    
    
DESCRIPTION
    Like the File or Folder Properties window, gets the total size of the item and all subitems.  Works on UNC paths 
    also.
    Returns an object with the path (input) and total size.
    Unlike some other cmdlets, this function takes SharePath (as from get-ServerShare) through the pipeline, as well 
    as Path.
    
###    remove-LocalGroupMember
    
SYNOPSIS
    Remotely removes a user from a local group, assuming adequate permissions
    
    
SYNTAX
    `remove-LocalGroupMember [-ComputerName] <String[]> [-Identity] <String> [-Group] <String> [<CommonParameters>]`
    
    
DESCRIPTION
    Remove users from a local group group remotely; firewall must [allow access to] remote pc
    
###   Search-Script
    
SYNOPSIS
    Searches PS1 files for a word or phrase
    
    
SYNTAX
    `Search-Script [-SearchPhrase] <Object> [[-Path] <Object>] [-IncludeAllPSFiles] [<CommonParameters>]`
    
    
DESCRIPTION
    Given a search phrase and location, will look in the code for the search string, then display matches in a grid 
    view.  Any items selected will be opened in PowerShell ISE.
    
###    start-ProgressCountdown
    
SYNOPSIS
    Shows the progress of a second-based countdown
    
    
SYNTAX
    `start-ProgressCountdown [[-seconds] <Object>] [[-activity] <Object>] [<CommonParameters>]`
    
    
DESCRIPTION
    Given a number of seconds, shows a progress bar counting down each second and increasing completion to the end of 
    the countdown.
    
###    Test-ADCredential
    
SYNOPSIS
    Checks whether a certain name and password are valid in a domain.
    
    
SYNTAX
    `Test-ADCredential [[-Credential] <Object>] [<CommonParameters>]`
    
    
DESCRIPTION
    Given a credential (a window asking for username and password), uses PrincipalContext object to validate 
    credentials.
    Will automatically add current domain prefix if not given.
    Returns true or false for whether the credentials are valid, or a warning if a problem with the domain.
    
