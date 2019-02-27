function New-HTMLReport{
    <#     
    .SYNOPSIS      
        Generates an HTML report for the information supplied in $information

    .DESCRIPTION    
        Builds an HTML table from a supplied hash table.  The PrimaryColumn property should be the 1st column, typically a computer name or user etc.,
        and will only be output once per occurance.  Each successive line for the PrimaryColumn will have all information except for the PrimaryColumn.
        This creates an indent effect.  ColorCoding, if used, should be a hash table of column names and hexadecimal colors for the output for each row.

    .EXAMPLE    
        New-HTMLReport -Data $computersToReport
        $userGroupsToReport | ForEach-Object {[pscustomobject]$_ } | Sort-Object -Property Name | New-HTMLReport

    .NOTES
        Written by Jason Dillman on 2-20-2018
        Rev. 1.0
    #>
    Param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName)]
            [Alias('Data')]
            [PSObject]$information
    )

    begin{
        <#  Create function variables  #>
        $previousPrimaryColumn = ''
        $insertHeaders = $true
        
        '<!DOCTYPE html>
        <html>
        <head>
        <style>
        table {
            font-family: arial, sans-serif;
            border-collapse: collapse;
            width: 100%;
        }

        table.bottomBorder { 
            border-collapse: collapse; 
        }
        table.bottomBorder td, 
        table.bottomBorder th { 
            border-bottom: 1px solid #000000; 
            padding: 8px; 
            text-align: left;
        }

        </style>
        </head>
        <body>

        <table class="bottomBorder">
            <tr>'
    } # End of begin block

    process{
        # when the function is called on the pipeline the being{} block does not have access to the current pipeline object,
        # so the $insertHeaders variable is used to determine if the current pipeline object is the table header
        if ($insertHeaders){
            $information | ForEach-Object{
                    $_.PSObject.Properties.Name | Where-Object {$_ -ne 'Color Coding' -and $_ -ne 'Primary Column Name'}
                } | Select-Object -Unique | ForEach-Object {
                    '
                    <th>{0}</th>' -f $_
                }
            '
            </tr>' # Table and Header created
            $insertHeaders = $false
        }
        foreach ($line in $information ){
            # Begin new table row
            '
            <tr>'
            foreach ($column in $line.PSObject.Properties.Name | Where-Object {$_ -ne 'Color Coding' -and $_ -ne 'Primary Column Name'}){
                # if the previous lines 'Primary Column' equals this lines 'Primary Column' (computer name, username, etc.), then insert blank cell
                # for indent effect and then proceed to the next column
                if ($previousPrimaryColumn -eq $line.$column){
                    '
                    <td></td>'
                    continue
                }
                # if the column is listed in the 'Color Coding' field the insert the cell with the value in 'Color Coding'
                if ($column -in $line.'Color Coding'.Key){
                    "
                    <td><font color=#{1}>{0}</font></td>" -f $line.$column, ($line.'Color Coding' | Where-Object {$_.key -eq $column}).value
                    continue                
                }
                # if the loop is still processing then insert the value of the column
                '
                    <td>{0}</td>' -f $line.$column
            }
            '
            </tr>'
            # Set $previousPrimaryColumn to the current 'Primary Column Name' before next iteration of the loop for future comparison
            $previousPrimaryColumn = $line."$($line.'Primary Column Name' | Select-Object -First 1)"
        }
    } # End of process block

    end {
        '
        </table>
        
        </body>
        </html>'
    } # End of end block
} # End function New-HTMLReport