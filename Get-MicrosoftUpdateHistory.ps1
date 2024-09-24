## Microsoft Function Naming Convention: http://msdn.microsoft.com/en-us/library/ms714428(v=vs.85).aspx

#region Function Get-MicrosoftUpdateHistory
Function Get-MicrosoftUpdateHistory
    {
        <#
          .SYNOPSIS
          Gets the list of installed updates
          
          .DESCRIPTION
          Slightly more detailed description of what your function does
                    
          .EXAMPLE
          Get-MicrosoftUpdateHistory

          .EXAMPLE
          Get-MicrosoftUpdateHistory | Group-Object -Property @('UpdateID') | Sort-Object -Property @('Count') -Descending

          .EXAMPLE
          Get-MicrosoftUpdateHistory -Advanced

          .EXAMPLE
          Get-MicrosoftUpdateHistory -Advanced | Group-Object -Property @('UpdateID') | Sort-Object -Property @('Count') -Descending

          .EXAMPLE
          Get-MicrosoftUpdateHistory | Where-Object {([String]::IsNullOrEmpty($_.ArticleID) -eq $False)} | Group-Object -Property @('ArticleID') | Sort-Object -Property @('Count') -Descending
  
          .NOTES
          Any useful tidbits
          
          .LINK
          Place any useful link here where your function or cmdlet can be referenced
        #>
        
        [CmdletBinding(ConfirmImpact = 'Low', DefaultParameterSetName = '__AllParameterSets', HelpURI = '', SupportsShouldProcess = $True, PositionalBinding = $True)]
       
        Param
          (                                                    
              [Parameter(Mandatory=$False)]
              [Switch]$Advanced,
              
              [Parameter(Mandatory=$False)]
              [Switch]$ContinueOnError        
          )
                    
        Begin
          {              
              Try
                {
                    $DateTimeLogFormat = 'dddd, MMMM dd, yyyy @ hh:mm:ss tt'  ###Monday, January 01, 2019 @ 10:15:34 AM###
                    [ScriptBlock]$GetCurrentDateTimeLogFormat = {(Get-Date).ToString($DateTimeLogFormat)}
                    $DateTimeMessageFormat = 'MM/dd/yyyy HH:mm:ss.FFF'  ###03/23/2022 11:12:48.347###
                    [ScriptBlock]$GetCurrentDateTimeMessageFormat = {(Get-Date).ToString($DateTimeMessageFormat)}
                    $DateFileFormat = 'yyyyMMdd'  ###20190403###
                    [ScriptBlock]$GetCurrentDateFileFormat = {(Get-Date).ToString($DateFileFormat)}
                    $DateTimeFileFormat = 'yyyyMMdd_HHmmss'  ###20190403_115354###
                    [ScriptBlock]$GetCurrentDateTimeFileFormat = {(Get-Date).ToString($DateTimeFileFormat)}
                    $TextInfo = (Get-Culture).TextInfo
                    $LoggingDetails = New-Object -TypeName 'System.Collections.Specialized.OrderedDictionary'    
                      $LoggingDetails.Add('LogMessage', $Null)
                      $LoggingDetails.Add('WarningMessage', $Null)
                      $LoggingDetails.Add('ErrorMessage', $Null)
                    $CommonParameterList = New-Object -TypeName 'System.Collections.Generic.List[String]'
                      $CommonParameterList.AddRange([System.Management.Automation.PSCmdlet]::CommonParameters)
                      $CommonParameterList.AddRange([System.Management.Automation.PSCmdlet]::OptionalCommonParameters)

                    [ScriptBlock]$ErrorHandlingDefinition = {
                                                                $ErrorMessageList = New-Object -TypeName 'System.Collections.Specialized.OrderedDictionary'
                                                                  $ErrorMessageList.Add('Message', $_.Exception.Message)
                                                                  $ErrorMessageList.Add('Category', $_.Exception.ErrorRecord.FullyQualifiedErrorID)
                                                                  $ErrorMessageList.Add('Script', $_.InvocationInfo.ScriptName)
                                                                  $ErrorMessageList.Add('LineNumber', $_.InvocationInfo.ScriptLineNumber)
                                                                  $ErrorMessageList.Add('LinePosition', $_.InvocationInfo.OffsetInLine)
                                                                  $ErrorMessageList.Add('Code', $_.InvocationInfo.Line.Trim())

                                                                ForEach ($ErrorMessage In $ErrorMessageList.GetEnumerator())
                                                                  {
                                                                      $LoggingDetails.ErrorMessage = "$($GetCurrentDateTimeMessageFormat.Invoke()) -  ERROR: $($ErrorMessage.Key): $($ErrorMessage.Value)"
                                                                      Write-Warning -Message ($LoggingDetails.ErrorMessage)
                                                                  }

                                                                Switch (($ContinueOnError.IsPresent -eq $False) -or ($ContinueOnError -eq $False))
                                                                  {
                                                                      {($_ -eq $True)}
                                                                        {                  
                                                                            Throw
                                                                        }
                                                                  }
                                                            }
                    
                    #Determine the date and time we executed the function
                      $FunctionStartTime = (Get-Date)
                    
                    [String]$CmdletName = $MyInvocation.MyCommand.Name 
                    
                    $LogMessage = "$($GetCurrentDateTimeMessageFormat.Invoke()) - Function `'$($CmdletName)`' is beginning. Please Wait..."
                    Write-Verbose -Message $LogMessage
              
                    #Define Default Action Preferences
                      $ErrorActionPreference = 'Stop'
                      
                    [String[]]$AvailableScriptParameters = (Get-Command -Name ($CmdletName)).Parameters.GetEnumerator() | Where-Object {($_.Value.Name -inotin $CommonParameterList)} | ForEach-Object {"-$($_.Value.Name):$($_.Value.ParameterType.Name)"}
                    $LoggingDetails.LogMessage = "$($GetCurrentDateTimeMessageFormat.Invoke()) - Available Function Parameter(s) = $($AvailableScriptParameters -Join ', ')"
                    Write-Verbose -Message ($LoggingDetails.LogMessage)

                    [String[]]$SuppliedScriptParameters = $PSBoundParameters.GetEnumerator() | ForEach-Object {"-$($_.Key):$($_.Value.GetType().Name)"}
                    $LoggingDetails.LogMessage = "$($GetCurrentDateTimeMessageFormat.Invoke()) - Supplied Function Parameter(s) = $($SuppliedScriptParameters -Join ', ')"
                    Write-Verbose -Message ($LoggingDetails.LogMessage)
                      
                    $LogMessage = "$($GetCurrentDateTimeMessageFormat.Invoke()) - Execution of $($CmdletName) began on $($FunctionStartTime.ToString($DateTimeLogFormat))"
                    Write-Verbose -Message $LogMessage
                                        
                    #Create an object that will contain the functions output.
                      $OutputObjectList = New-Object -TypeName 'System.Collections.Generic.List[PSObject]'
                }
              Catch
                {
                    $ErrorHandlingDefinition.Invoke()
                }
              Finally
                {
                    
                }
          }

        Process
          {           
              Try
                {  
                    $LogMessage = "$($GetCurrentDateTimeMessageFormat.Invoke()) - Attempting to retrieve the Microsoft update history. Please Wait..." 
                    Write-Verbose -Message $LogMessage 
                                    
                    $MicrosoftUpdateSession = New-Object -ComObject 'Microsoft.Update.Session'
                    $MicrosoftUpdateSearcher = $MicrosoftUpdateSession.CreateUpdateSearcher()
                    $MicrosoftUpdateTotalHistoryCount = $MicrosoftUpdateSearcher.GetTotalHistoryCount()
                    $MicrosoftUpdateList = $MicrosoftUpdateSearcher.QueryHistory(0, $MicrosoftUpdateTotalHistoryCount)

                    $MicrosoftUpdateLoopCounter = 1

                    $MicrosoftUpdateHashSet = New-Object -TypeName 'System.Collections.Generic.HashSet[String]'
                      
                    :MicrosoftUpdateLoop For ($MicrosoftUpdateListIndex = 0; $MicrosoftUpdateListIndex -lt $MicrosoftUpdateTotalHistoryCount; $MicrosoftUpdateListIndex++)
                      {
                          $MicrosoftUpdate = $MicrosoftUpdateList[$MicrosoftUpdateListIndex]

                          $MicrosoftUpdatePropertyDictionary = New-Object -TypeName 'System.Collections.Specialized.OrderedDictionary'
                            $MicrosoftUpdatePropertyDictionary.ArticleID = $Null
                            $MicrosoftUpdatePropertyDictionary.Title = $MicrosoftUpdate.Title
                            $MicrosoftUpdatePropertyDictionary.Description = $MicrosoftUpdate.Description -ireplace '(\s{2,})', ' '
                            $MicrosoftUpdatePropertyDictionary.UpdateID = $MicrosoftUpdate.UpdateIdentity.UpdateID.ToUpper()
                            $MicrosoftUpdatePropertyDictionary.Revision = $MicrosoftUpdate.UpdateIdentity.RevisionNumber
                            $MicrosoftUpdatePropertyDictionary.Category = New-Object -TypeName 'System.Collections.Generic.List[String]'
                            $MicrosoftUpdatePropertyDictionary.UpdateType = New-Object -TypeName 'System.Collections.Generic.List[String]'
                            $MicrosoftUpdatePropertyDictionary.InstallationDate = $MicrosoftUpdate.Date
                            $MicrosoftUpdatePropertyDictionary.SupportURL = $MicrosoftUpdate.SupportURL

                          $LogMessage = "Attempting to retrieve details for update $($MicrosoftUpdateLoopCounter) of $($MicrosoftUpdateTotalHistoryCount) - [Title: $($MicrosoftUpdatePropertyDictionary.Title)] [Update ID: $($MicrosoftUpdatePropertyDictionary.UpdateID)]. Please Wait..." 
                          Write-Verbose -Message $LogMessage
                          
                          $ArticleIDExpressionEvaluation = [Regex]::Match($MicrosoftUpdate.Title, 'KB(\d+)')

                          Switch ($ArticleIDExpressionEvaluation.Success)
                            {
                                {($_ -eq $True)}
                                  {
                                      Switch (([String]::IsNullOrEmpty($ArticleIDExpressionEvaluation.Value) -eq $False) -and ([String]::IsNullOrWhiteSpace($ArticleIDExpressionEvaluation.Value) -eq $False))
                                        {
                                            {($_ -eq $True)}
                                              {
                                                  $MicrosoftUpdatePropertyDictionary.ArticleID = $ArticleIDExpressionEvaluation.Value
                                              }
                                        }
                                  }
                            }

                          ForEach ($Category In $MicrosoftUpdate.Categories)
                            {    
                               Switch (([String]::IsNullOrEmpty($Category.Type) -eq $False) -and ([String]::IsNullOrWhiteSpace($Category.Type) -eq $False))
                                  {
                                      {($_ -eq $True)}
                                        {
                                            $MicrosoftUpdatePropertyDictionary.UpdateType.Add($Category.Type.Trim())
                                        }
                                  }
                                  
                               Switch (([String]::IsNullOrEmpty($Category.Name) -eq $False) -and ([String]::IsNullOrWhiteSpace($Category.Name) -eq $False))
                                  {
                                      {($_ -eq $True)}
                                        {
                                            $MicrosoftUpdatePropertyDictionary.Category.Add($Category.Name.Trim())
                                        }
                                  }  
                            }

                          $MicrosoftUpdatePropertyDictionary.UpdateType = $MicrosoftUpdatePropertyDictionary.UpdateType.ToArray() | Sort-Object -Unique

                          $MicrosoftUpdatePropertyDictionary.Category = $MicrosoftUpdatePropertyDictionary.Category.ToArray() | Sort-Object -Unique
                  
                          $MicrosoftUpdatePropertyDictionary.OperationCode = $MicrosoftUpdate.Operation
                    
                          Switch ($MicrosoftUpdate.Operation)
                            {
                                {($_ -eq 1)} {$MicrosoftUpdatePropertyDictionary.Operation = 'Installation'}
                                {($_ -eq 2)} {$MicrosoftUpdatePropertyDictionary.Operation = 'Uninstallation'}
                                {($_ -eq 3)} {$MicrosoftUpdatePropertyDictionary.Operation = 'Other'}
                                Default {$MicrosoftUpdatePropertyDictionary.Operation = 'Unknown'}
                            }

                          $MicrosoftUpdatePropertyDictionary.ResultCode = $MicrosoftUpdate.ResultCode

                          Switch ($MicrosoftUpdate.ResultCode)
                            {
                                {($_ -eq 0)} {$MicrosoftUpdatePropertyDictionary.Result = 'Not Started'}
                                {($_ -eq 1)} {$MicrosoftUpdatePropertyDictionary.Result = 'In Progress'}
                                {($_ -eq 2)} {$MicrosoftUpdatePropertyDictionary.Result = 'Succeeded'}
                                {($_ -eq 3)} {$MicrosoftUpdatePropertyDictionary.Result = 'Succeeded With Errors'}
                                {($_ -eq 4)} {$MicrosoftUpdatePropertyDictionary.Result = 'Failed'}
                                {($_ -eq 5)} {$MicrosoftUpdatePropertyDictionary.Result = 'Aborted'}
                                Default {$MicrosoftUpdatePropertyDictionary.Result = 'Unknown'}
                            }
          
                          $MicrosoftUpdatePropertyDictionary.Advanced = New-Object -TypeName 'System.Collections.Generic.List[PSObject]'
                          
                          Switch ($Advanced.IsPresent)
                            {
                                {($_ -eq $True)}
                                  {
                                      Switch ($MicrosoftUpdateHashSet.Add($MicrosoftUpdatePropertyDictionary.UpdateID))
                                        {
                                            {($_ -eq $True)}
                                              {
                                                  $LogMessage = "Attempting to retrieve advanced details for update $($MicrosoftUpdateLoopCounter) of $($MicrosoftUpdateTotalHistoryCount) - [Title: $($MicrosoftUpdatePropertyDictionary.Title)] [Update ID: $($MicrosoftUpdatePropertyDictionary.UpdateID)]. Please Wait..." 
                                                  Write-Verbose -Message $LogMessage
            
                                                  $AdvancedUpdateInfo = $MicrosoftUpdateSearcher.Search("UpdateID = '$($MicrosoftUpdatePropertyDictionary.UpdateID)'")

                                                  Switch ($Null -ine $AdvancedUpdateInfo)
                                                    {
                                                        {($_ -eq $True)}
                                                          {
                                                              Switch ($AdvancedUpdateInfo.Updates.Count -gt 0)
                                                                {
                                                                    {($_ -eq $True)}
                                                                      {
                                                                          ForEach ($Update In $AdvancedUpdateInfo.Updates)
                                                                            {
                                                                                $UpdateAdvancedPropertyDictionary = New-Object -TypeName 'System.Collections.Specialized.OrderedDictionary'   
                                                                                  $UpdateAdvancedPropertyDictionary.KBArticleIDs = New-Object -TypeName 'System.Collections.Generic.List[String]'
                                                                                  $UpdateAdvancedPropertyDictionary.CVEIDs = New-Object -TypeName 'System.Collections.Generic.List[String]'
                                                                                  $UpdateAdvancedPropertyDictionary.SecurityBulletinIDs = New-Object -TypeName 'System.Collections.Generic.List[String]'
                                                                                  $UpdateAdvancedPropertyDictionary.SupersededUpdateIDs = New-Object -TypeName 'System.Collections.Generic.List[String]'
                                                                                  $UpdateAdvancedPropertyDictionary.Languages = New-Object -TypeName 'System.Collections.Generic.List[String]'
                                                                                  $UpdateAdvancedPropertyDictionary.RebootRequired = $Update.RebootRequired     
                                                                                  $UpdateAdvancedPropertyDictionary.MinimumDownloadSize = $Update.MinDownloadSize
                                                                                  $UpdateAdvancedPropertyDictionary.MaximumDownloadSize = $Update.MaxDownloadSize
                                                                                  $UpdateAdvancedPropertyDictionary.LastDeploymentChangeTime = $Update.LastDeploymentChangeTime
                                                                                  $UpdateAdvancedPropertyDictionary.DownloadURLs = New-Object -TypeName 'System.Collections.Generic.List[String]'
                                          
                                                                                Switch ($True)
                                                                                  {
                                                                                      {($Update.Languages.Count -gt 0)}
                                                                                        {
                                                                                            $LogMessage = "Attempting to add $($Update.Languages.Count) supported languages. Please Wait..." 
                                                                                            Write-Verbose -Message $LogMessage
                                                      
                                                                                            ForEach ($Language In $Update.Languages)
                                                                                              {
                                                                                                  Switch (([String]::IsNullOrEmpty($Language) -eq $False) -and ([String]::IsNullOrWhiteSpace($Language) -eq $False))
                                                                                                    {
                                                                                                        {($_ -eq $True)}
                                                                                                          {
                                                                                                              $UpdateAdvancedPropertyDictionary.Languages.Add($Language)
                                                                                                          }
                                                                                                    }
                                                                                              }
                                                                                        }
                                                
                                                                                      {($Update.SecurityBulletinIDs.Count -gt 0)}
                                                                                        {
                                                                                            $LogMessage = "Attempting to add $($Update.SecurityBulletinIDs.Count) security bulletin IDs. Please Wait..." 
                                                                                            Write-Verbose -Message $LogMessage
                                                      
                                                                                            ForEach ($SecurityBulletinID In $Update.SecurityBulletinIDs)
                                                                                              {
                                                                                                  Switch (([String]::IsNullOrEmpty($SecurityBulletinID) -eq $False) -and ([String]::IsNullOrWhiteSpace($SecurityBulletinID) -eq $False))
                                                                                                    {
                                                                                                        {($_ -eq $True)}
                                                                                                          {
                                                                                                              $UpdateAdvancedPropertyDictionary.SecurityBulletinIDs.Add($SecurityBulletinID)
                                                                                                          }
                                                                                                    }
                                                                                              }
                                                                                        }

                                                                                      {($Update.SupersededUpdateIDs.Count -gt 0)}
                                                                                        {
                                                                                            $LogMessage = "Attempting to add $($Update.SupersededUpdateIDs.Count) superseded Update IDs. Please Wait..." 
                                                                                            Write-Verbose -Message $LogMessage
                                                      
                                                                                            ForEach ($SupersededUpdateID In $Update.SupersededUpdateIDs)
                                                                                              {
                                                                                                  Switch (([String]::IsNullOrEmpty($SupersededUpdateID) -eq $False) -and ([String]::IsNullOrWhiteSpace($SupersededUpdateID) -eq $False))
                                                                                                    {
                                                                                                        {($_ -eq $True)}
                                                                                                          {
                                                                                                              $UpdateAdvancedPropertyDictionary.SupersededUpdateIDs.Add($SupersededUpdateID)
                                                                                                          }
                                                                                                    }
                                                                                              }
                                                                                        }

                                                                                      {($Update.CVEIDs.Count -gt 0)}
                                                                                        {
                                                                                            $LogMessage = "Attempting to add $($Update.CVEIDs.Count) CVE IDs. Please Wait..." 
                                                                                            Write-Verbose -Message $LogMessage
                                                      
                                                                                            ForEach ($CVEID In $Update.CVEIDs)
                                                                                              {
                                                                                                  Switch (([String]::IsNullOrEmpty($CVEID) -eq $False) -and ([String]::IsNullOrWhiteSpace($CVEID) -eq $False))
                                                                                                    {
                                                                                                        {($_ -eq $True)}
                                                                                                          {
                                                                                                              $UpdateAdvancedPropertyDictionary.CVEIDs.Add($CVEID)
                                                                                                          }
                                                                                                    }
                                                                                              }
                                                                                        }

                                                                                      {($Update.KBArticleIDs.Count -gt 0)}
                                                                                        {
                                                                                            $LogMessage = "Attempting to add $($Update.KBArticleIDs.Count) KB article IDs. Please Wait..." 
                                                                                            Write-Verbose -Message $LogMessage
                                                      
                                                                                            ForEach ($KBArticleID In $Update.KBArticleIDs)
                                                                                              {
                                                                                                  Switch (([String]::IsNullOrEmpty($KBArticleID) -eq $False) -and ([String]::IsNullOrWhiteSpace($KBArticleID) -eq $False))
                                                                                                    {
                                                                                                        {($_ -eq $True)}
                                                                                                          {
                                                                                                              $UpdateAdvancedPropertyDictionary.KBArticleIDs.Add($KBArticleID)
                                                                                                          }
                                                                                                    }
                                                                                              }
                                                                                        }
                                                
                                                                                      {($Update.DownloadContents.Count -gt 0)}
                                                                                        {
                                                                                            $LogMessage = "Attempting to add $($Update.DownloadContents.Count) URLs to the download URL list. Please Wait..." 
                                                                                            Write-Verbose -Message $LogMessage
                                                      
                                                                                            ForEach ($DownloadContent In $Update.DownloadContents)
                                                                                              {
                                                                                                  Switch (([String]::IsNullOrEmpty($DownloadContent.DownloadUrl) -eq $False) -and ([String]::IsNullOrWhiteSpace($DownloadContent.DownloadUrl) -eq $False))
                                                                                                    {
                                                                                                        {($_ -eq $True)}
                                                                                                          {
                                                                                                              $UpdateAdvancedPropertyDictionary.DownloadURLs.Add($DownloadContent.DownloadUrl)
                                                                                                          }
                                                                                                    }
                                                                                              }
                                                                                        }
                                                                                  }

                                                                                $UpdateAdvancedObject = New-Object -TypeName 'PSObject' -Property ($UpdateAdvancedPropertyDictionary)

                                                                                $MicrosoftUpdatePropertyDictionary.Advanced.Add($UpdateAdvancedObject)
                                                                            }
                                                                      }
                                                                }
                                                          }
                                                    }
          
                                                  $MicrosoftUpdateObject = New-Object -TypeName 'PSObject' -Property ($MicrosoftUpdatePropertyDictionary)

                                                  ForEach ($MicrosoftUpdateObjectProperty In $MicrosoftUpdateObject.PSObject.Properties)
                                                    {
                                                        Switch ($Null -ine $MicrosoftUpdateObjectProperty.Value)
                                                          {
                                                              {($_ -eq $True)}
                                                                {                  
                                                                    Switch ($MicrosoftUpdateObjectProperty.Value.GetType().FullName)
                                                                      {
                                                                          {($_ -imatch 'System\.Collections\.Generic\.List.*\[.*\]')}
                                                                            {
                                                                                $MicrosoftUpdateObject.$($MicrosoftUpdateObjectProperty.Name) = $MicrosoftUpdateObjectProperty.Value.ToArray()
                                                                            }
                                                                      }
                                                                }
                                                          }
                                                    }
                                              }

                                            Default
                                              {
                                                  $MicrosoftUpdatePropertyDictionary.Advanced = ($OutputObjectList | Where-Object {($_.UpdateID -ieq $MicrosoftUpdatePropertyDictionary.UpdateID)} | Select-Object -First 1).Advanced
                                              }
                                        }    
                                  }
                            }
                          
                          $MicrosoftUpdateObject = New-Object -TypeName 'PSObject' -Property ($MicrosoftUpdatePropertyDictionary)
            
                          $OutputObjectList.Add($MicrosoftUpdateObject)

                          $MicrosoftUpdateLoopCounter++
                      }   
                }
              Catch
                {
                    $ErrorHandlingDefinition.Invoke()
                }
              Finally
                {
                    
                }
          }
        
        End
          {                                        
              Try
                {                
                    #Determine the date and time the function completed execution
                      $FunctionEndTime = (Get-Date)

                      $LogMessage = "$($GetCurrentDateTimeMessageFormat.Invoke()) - Execution of $($CmdletName) ended on $($FunctionEndTime.ToString($DateTimeLogFormat))"
                      Write-Verbose -Message $LogMessage

                    #Log the total script execution time  
                      $FunctionExecutionTimespan = New-TimeSpan -Start ($FunctionStartTime) -End ($FunctionEndTime)

                      $LogMessage = "$($GetCurrentDateTimeMessageFormat.Invoke()) - Function execution took $($FunctionExecutionTimespan.Hours.ToString()) hour(s), $($FunctionExecutionTimespan.Minutes.ToString()) minute(s), $($FunctionExecutionTimespan.Seconds.ToString()) second(s), and $($FunctionExecutionTimespan.Milliseconds.ToString()) millisecond(s)"
                      Write-Verbose -Message $LogMessage
                    
                    $LogMessage = "$($GetCurrentDateTimeMessageFormat.Invoke()) - Function `'$($CmdletName)`' is completed."
                    Write-Verbose -Message $LogMessage
                }
              Catch
                {
                    $ErrorHandlingDefinition.Invoke()
                }
              Finally
                {
                    $OutputObjectList = $OutputObjectList.ToArray() | Sort-Object -Property @('InstallationDate')
                    
                    Write-Output -InputObject ($OutputObjectList)
                }
          }
    }
#endregion
