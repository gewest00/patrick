Patrick				
Microscopy output organisation tool

## NAME
    patrick


## SYNOPSIS
    patrick [-h] [-c COLUMN] input [input ...] [-r]


## DESCRIPTION
    Take Guava files and reformat for easier use.

    -h
        display help information and exit

    -c, --column
        Specify the column/s containing the data you wish to summarise
        Use the name/s from the header row, as displayed when the input file(s) is(are)
        viewed with Excel. Defaults are 'Children_Nuclei_Count' and 'AreaShape_Area'
	
    -r
	Remove rows where Children_Nuclei_Count = 0


## INPUT FILE

        Input files must be .csv formatted

        If your input files do not match this format, contact the author!


## COMPATIBILITY
        Patrick was developed with:
        "Python 3.8.8"


## NOTES
    Output is written to a .csv file named output_[input].xlsx
    Data from all the input files will be reformatted into corresponding output file.

    Data is presented as four sheets containing data according to infection and nucleation
    status, plus summary sheets at the end based on specified categories.


## AUTHOR
    Written by gewest00


## COPYRIGHT
    Copyright © 2025 gewest00
	
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.


PATRICK				April 2025
