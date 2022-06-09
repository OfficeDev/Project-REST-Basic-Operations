function Create-CustomField{
    Param(
            [parameter(Mandatory=$true)]$projContext, 
            [parameter(Mandatory=$true)]$Id,
            [parameter(Mandatory=$true)]$Name,
            [parameter(Mandatory=$false)]$Description,
            [parameter(Mandatory=$true)]$FieldType,
            [parameter(Mandatory=$true)]$EntityType,
            [parameter(Mandatory=$false)]$LookupTable,
            [parameter(Mandatory=$true)]$IsWorkflowControlled,
            [parameter(Mandatory=$true)]$IsRequired,
            [parameter(Mandatory=$false)]$IsEditableInVisibility,
            [parameter(Mandatory=$false)]$IsMultilineText

    )

    [Microsoft.ProjectServer.Client.CustomFieldCreationInformation]$customFieldInfo = New-Object Microsoft.ProjectServer.Client.CustomFieldCreationInformation;
    #Mandatory fields
    $customFieldInfo.Id = $Id
    $customFieldInfo.Name = $Name
    $customFieldInfo.FieldType = $FieldType
    $customFieldInfo.EntityType = $EntityType
    $customFieldInfo.IsWorkflowControlled = $IsWorkflowControlled
    $customFieldInfo.IsRequired = $IsRequired

    #Non Mandatory fields
    if($Description -ne $null)
    {
        $customFieldInfo.Description = $Description
    }

    if($IsEditableInVisibility -ne $null)
    {
        $customFieldInfo.IsEditableInVisibility = $IsEditableInVisibility
    }

    if($IsMultilineText -ne $null)
    {
        $customFieldInfo.IsMultilineText = $IsMultilineText
    }

    if($LookupTable -ne $null)
    {
        $customFieldInfo.LookupTable = $LookupTable
    }
    
    $newCustomField = $projContext.CustomFields.Add($customFieldInfo);
    $projContext.CustomFields.Update()
    $projContext.ExecuteQuery()
}
