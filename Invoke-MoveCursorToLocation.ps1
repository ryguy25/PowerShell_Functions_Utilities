Add-Type -AssemblyName System.Windows.Forms

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
    [uint32]$WaitBetweenMoves = 50

)

    Try {

        $currentCursorPosition = [System.Windows.Forms.Cursor]::Position

        #region - Calculate positiveXChange
        if ( ( $currentCursorPosition.X - $X ) -ge 0 ) {
            
            $positiveXChange = $false
        
        }
        
        else {
        
            $positiveXChange = $true
        
        }
        #endregion - Calculate positiveXChange

        #region - Calculate positiveYChange
        if ( ( $currentCursorPosition.Y - $Y ) -ge 0 ) {
        
            $positiveYChange = $false
        
        }
        
        else {
        
            $positiveYChange = $true
        
        }
        #endregion - Calculate positiveYChange

        #region - Setup Trig Values

        ### We're always going to use Tan and ArcTan to calculate the movement increments because we know the x/y values which are always
        ### going to be the adjacent and opposite values of the triangle
        $xTotalDelta = [Math]::Abs( $X - $currentCursorPosition.X )
        $yTotalDelta = [Math]::Abs( $Y - $currentCursorPosition.Y )
        
        ### To avoid any strange behavior, we're always going to calculate our movement values using the larger delta value
        if ( $xTotalDelta -ge $yTotalDelta ) {

            $tanAngle = [Math]::Tan( $yTotalDelta / $xTotalDelta )

            if ( $NumberOfMoves -gt $xTotalDelta ) {
            
                $NumberOfMoves = $xTotalDelta
            
            }

            $xMoveIncrement = $xTotalDelta / $NumberOfMoves
            $yMoveIncrement = $yTotalDelta - ( ( [Math]::Atan($tanAngle) * ( $xTotalDelta - $xMoveIncrement ) ) )

        }
        else {

            $tanAngle = [Math]::Tan( $xTotalDelta / $yTotalDelta )

            if ( $NumberOfMoves -gt $yTotalDelta ) {

                $NumberOfMoves = $yTotalDelta

            }

            $yMoveIncrement = $yTotalDelta / $NumberOfMoves
            $xMoveIncrement = $xTotalDelta - ( ( [Math]::Atan($tanAngle) * ( $yTotalDelta - $yMoveIncrement ) ) )
        }
        #endregion - Setup Trig Values

        #region Verbose Output (Before for loop)
        Write-Verbose "StartingX: $($currentCursorPosition.X)`t`t`t`t`t`tStartingY: $($currentCursorPosition.Y)"
        Write-Verbose "Total X Delta: $xTotalDelta`t`t`t`t`tTotal Y Delta: $yTotalDelta"
        Write-Verbose "Positive X Change: $positiveXChange`t`t`tPositive Y Change: $positiveYChange"
        Write-Verbose "X Move Increment: $xMoveIncrement`t`t`tY Move Increment: $yMoveIncrement"
        #endregion

        for ( $i = 0 ; $i -lt $NumberOfMoves ; $i++) {
            
            ##$yPos = [Math]::Atan($tanAngle) * ($xTotalDelta - $currentCursorPosition.X)
            ##$yMoveIncrement = $yTotalDelta - $yPos

            #region Calculate X movement direction
            switch ( $positiveXChange ) {

                $true    { $currentCursorPosition.X += $xMoveIncrement }
                $false   { $currentCursorPosition.X -= $xMoveIncrement }
                default  { $currentCursorPosition.X = $currentCursorPosition.X }

            }
            #endregion Calculate X movement direction

            #region Calculate Y movement direction
            switch ( $positiveYChange ) {

                $true    { $currentCursorPosition.Y += $yMoveIncrement }
                $false   { $currentCursorPosition.Y -= $yMoveIncrement }
                default  { $currentCursorPosition.Y = $currentCursorPosition.Y }

            }
            #endregion Calculate Y movement direction

            [System.Windows.Forms.Cursor]::Position = $currentCursorPosition
            Start-Sleep -Milliseconds $WaitBetweenMoves

            #region Verbose Output (During Loop)
            Write-Verbose "Current X Position:`t $($currentCursorPosition.X)`tCurrent Y Position: $($currentCursorPosition.Y)"
            #endregion Verbose Output (During Loop)
        }
        
        $currentCursorPosition.X = $X
        $currentCursorPosition.Y = $Y
        [System.Windows.Forms.Cursor]::Position = $currentCursorPosition
        Write-Verbose "End X Position: $($currentCursorPosition.X)`tEnd Y Position: $($currentCursorPosition.Y)"
    
    }

    Catch {

        Write-Error $_.Exception.Message

    }

}
