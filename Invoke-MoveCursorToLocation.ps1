function Invoke-MoveCursorToLocation {

<#

.SYNOPSIS
Moves the cursor from its current location to a specific X,Y coordinate

.DESCRIPTION
Queries the current location of the cursor and slowly moves it to the requested destination

.EXAMPLE
  # PS C:\> Invoke-MoveCursorToLocation -X 100 -Y 100

  # This command would move the cursor from the current location to the X,Y coordinate 100,100

.PARAMETER X
The X coordinate of the desired cursor position

.PARAMETER Y
The Y coordinate of the desired cursor position

.PARAMETER NumberOfMoves
This is the number of times the cursor is actually moved.  The higher the number, the smoother the movement appears to the user.  However, if the cursor is moving a short or long distance, it can either appear to move extremely fast or slow

.PARAMETER CursorDelay
This parameter controls the amount of wait time between moves.  It can be adjusted along with the NumberOfMoves Parameter to 

.NOTES
#>


[CmdletBinding()]

PARAM (

    [Parameter(Mandatory=$true)]
    [uint32]$X,

    [Parameter(Mandatory=$true)]
    [uint32]$Y,

    [Parameter()]
    [uint32]$NumberOfMoves = 50,

    [Parameter()]
    [uint32]$CursorDelay = 5

)

    Try {

        $currentCursorPosition = [System.Windows.Forms.Cursor]::Position

        $xDeltaMovement = ($X - $currentCursorPosition.X) / $NumberOfMoves
        $yDeltaMovement = ($Y - $currentCursorPosition.Y) / $NumberOfMoves

        for ( $i = 0 ; $i -lt $NumberOfMoves ; $i++ ) {

            $currentCursorPosition.X += $xDeltaMovement
            $currentCursorPosition.Y += $yDeltaMovement

            [System.Windows.Forms.Cursor]::Position = $currentCursorPosition

            Start-Sleep -Milliseconds $CursorDelay
        
        }        
        

    }

    Catch {

        Write-Error $_.Exception.Message

    }

}
