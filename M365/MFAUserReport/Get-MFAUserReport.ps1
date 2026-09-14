function Get-MFAUserReport
{
   <#
      .SYNOPSIS
      Get a Azure AD MFA User report

      .DESCRIPTION
      Get a Azure AD MFA User report, the function can export the report as CSV.
      The export is disabled by default.

      .PARAMETER Export
      Export the MFA Report to CSV?

      .PARAMETER Path
      Path of the MFA Export CSV

      .EXAMPLE
      PS> Get-MFAUserReport

      Get a Azure AD MFA User report

      .EXAMPLE
      PS> Get-MFAUserReport -Export

      Get a Azure AD MFA User report and export it to the default report (C:\scripts\PowerShell\exports\MFAUsers.csv)

      .EXAMPLE
      PS> Get-MFAUserReport -Export -Path 'C:\scripts\PowerShell\exports\AllMFAUsers.csv'

      Get a Azure AD MFA User report and export it to given report (C:\scripts\PowerShell\exports\AllMFAUsers.csv)

      .NOTES
      ParameterSet added
   #>
   [CmdletBinding(DefaultParameterSetName = 'Normal',
      SupportsShouldProcess)]
   param
   (
      [Parameter(ParameterSetName = 'Export',
         ValueFromPipeline,
         Position = 1)]
      [Alias('CSV')]
      [switch]
      $Export,
      [Parameter(ParameterSetName = 'Export',
         ValueFromPipeline,
         Position = 2)]
      [string]
      $Path = 'C:\scripts\PowerShell\exports\MFAUsers.csv'
   )

   begin
   {
      # Defaults
      $CNT = 'Continue'
      $STP = 'Stop'

      # Cleanup
      $Report = @()
      $i = 0

      if ($pscmdlet.ShouldProcess('MFA Users', 'Get'))
      {
         # get all Accounts
         try
         {
            $Accounts = (Get-MsolUser -All -ErrorAction $STP -WarningAction $CNT | Where-Object -FilterScript {
                  $_.StrongAuthenticationMethods -ne $Null
               } | Sort-Object -Property DisplayName)
         }
         catch
         {
            $line = ($_.InvocationInfo.ScriptLineNumber)

            # Dump the Info
            Write-Warning -Message ('Error was in Line {0}' -f $line)

            # Dump the Error catched
            Write-Error -Message $_ -ErrorAction $STP

            # Something that should never be reached
            break
         }
      }
   }

   process
   {
      if ($pscmdlet.ShouldProcess('MFA Users', 'Process'))
      {
         foreach ($Account in $Accounts)
         {
            $AccountDisplayName = $Account.DisplayName
            Write-Verbose -Message ('Processing {0}' -f $AccountDisplayName)

            # Counter
            $i++

            # Select Methods
            $Methods = ($Account | Select-Object -ExpandProperty StrongAuthenticationMethods)
            $MFA = ($Account | Select-Object -ExpandProperty StrongAuthenticationUserDetails)
            $State = ($Account | Select-Object -ExpandProperty StrongAuthenticationRequirements)

            $Methods | ForEach-Object -Process {
               if ($_.IsDefault -eq $true)
               {
                  $Method = $_.MethodType
               }
            }

            if ($State.State)
            {
               $MFAStatus = $State.State
            }
            else
            {
               $MFAStatus = 'Disabled'
            }

            $Object = [PSCustomObject][Ordered]@{
               User      = $Account.DisplayName
               UPN       = $Account.UserPrincipalName
               MFAMethod = $Method
               MFAPhone  = $MFA.PhoneNumber
               MFAEmail  = $MFA.Email
               MFAStatus = $MFAStatus
            }

            # Add Obejct to report
            $Report += $Object
         }
      }
   }

   end
   {
      if ($pscmdlet.ShouldProcess('MFA Users', 'Report'))
      {
         Write-Verbose -Message ('{0} accounts are MFA-enabled' -f $i)

         if ($pscmdlet.ParameterSetName -eq 'Export')
         {
            try
            {
               $Null = ($Report | Export-Csv -NoTypeInformation -Path $Path -Force -ErrorAction $STP -WarningAction $CNT)
            }
            catch
            {
               $line = ($_.InvocationInfo.ScriptLineNumber)

               # Dump the Info
               Write-Warning -Message ('Error was in Line {0}' -f $line)

               # Dump the Error catched
               Write-Error -Message $_ -ErrorAction $STP

               # Something that should never be reached
               break
            }
         }
         else
         {
            # Dump to console
            $Report
         }
      }
   }
}