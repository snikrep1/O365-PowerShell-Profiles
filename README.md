# Microsoft365 Management PowerShell Profiles
A couple profile options for PowerShell for managing Microsoft 365 services.

These will use Multi-Factor Authentication (MFA) prompts to sign in rather than use a stored credential (basic auth method)

---
## Profile Options

1. Windows PowerShell console
2. Windows PowerShell ISE


### Installation Directions

1. Find your PowerShell Profile location
    
    a. Open **PowerShell**
    
    b. Type **$profile**
        
    ```powershell
    $profile
    
    #result:
    C:\Users\testing\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
    ```

2. The $profile variable will output where the profile for the current user is located. As mentioned previously, if you use ISE more, then it makes the most sense to use the Current User - All Hosts profile location.
    
        Note: Microsoft is confused when they say HOST, they mean a PowerShell program like PowerShell Console or editor ISE.
    
   --- 
   | Description | Path | Command to Open | 
   | --- | --- | ---|
   | Current User - Current Host | $Home\[My ]Documents\PowerShell\Microsoft.PowerShell_profile.ps1 | $profile | 
   | Current User - All Hosts | $Home\[My ]Documents\PowerShell\Profile.ps1 | $profile.CurrentUserAllHosts |
   | All Users - Current Host | $PSHOME\Microsoft.PowerShell_profile.ps1 | $profile.AllUsersCurrentHost |
   | All Users - All Hosts | $PSHOME\Profile.ps1 | $profile.AllUsersAllHosts |

3. Test if you already have a profile:
    
    ```powershell
    test-path $profile
    ```
    If it returns False, then we will need to create a new profile

    ```powershell
    New-Item -Path $profile -Type File -Force
    ```

    You can use this same method to create a new profile for all hosts or all users using the commands in the above table.

4. Lastly, make sure you can run scripts by setting your Execution Policy to Remote Signed.

    ```powershell
    Get-ExecutionPolicy

    #Set the ExecutionPolicy to RemoteSigned
    Set-ExecutionPolicy RemoteSigned
    ```
