    VB_xBase Class - A VB6 Class for accessing DBF Files without ADO
    Copyright (C) 2008,  Achmad Junus

    This library is free software; you can redistribute it and/or
    modify it under the terms of the GNU Lesser General Public
    License as published by the Free Software Foundation; either
    version 2.1 of the License, or (at your option) any later version.

    This library is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
    Lesser General Public License for more details.

    You should have received a copy of the GNU Lesser General Public
    License along with this library; if not, write to the Free Software
    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


=================================================================================


VB_xBase Class History
======================
Version 1.1.2

Author Name:   Achmad Junus
Author Email:  seruling_m4l4m \\at\\ yahoo.com


----------------------------------------
List of update and fix for version 1.1.2 
----------------------------------------
All updates and fixes in previous version, and the following:

1.  Function SortData() will sort data type Double as well.

    Negative value on data type double still can not maintain relative position.
    For positive value in data type Double, the sort is remain stable.
    Still using Radix Sort for sorting algorithm.
    Sorting get faster by optimizing Disk IO.

2.  Improve search speed on function SearchAll(). It now can search from 1% to 30% faster, depends on file size and number of entries found.

		**Note: The improved speed is still slower than Locate() function in Foxpro ver 7 with "SET OPTIMIZE = ON".
		My test searching in 750,000 records, my class about 3+ time slower (2100ms vs 650ms).
		
3.  User Interface - Demo project:

    - Several layout changes from old version.
    - Browse records no longer limited by FlexGrid capacity.
      This is done by scaling VScrool bar to emulate value larger than 32767,
      and loading text array to FlexGrid based on visible rows.
    - Modified mousewheel procedure when FlexGrid receive focus.


--------------------------------------------------------
List of update and fix for version 1.1.0  (Dec 12, 2008)  *** NOT UPLOADED TO PSC ***
--------------------------------------------------------
1.  Removed some garbage code

2.  Added function SortData() for sorting all data type except data type Double. Available for sort ascending only.

    The sort is Stable, using Radix Sort algorithm.
    The sort function is Not as fast as FoxPro.
    
3.  Fixed bug in function DBFClose() that still leave some information that prevent creating new file correctly.

    Now can create a new file correctly after calling DBFClose() without has to destroy the class instance.
    
4.  User Interface - Demo project:

    - Random records generator create data with better randomness
    - Added mousewheel event when FlexGrid receive focus.

5.  !!!Bug found!!!: Function GetDistinctValues() takes very long time when the field is sorted.

    Reason: No tree rebalancing routine in this function, and fall to sequential search


-----------------------------
Version 1.0.0 (March 9, 2008)
-----------------------------
My first attempt rewriting my C code into VB6 for accessing dbf files without ADO.

Known Bug and Limitation: 1. Only support ANSI and Single Byte Caharacter Set (SBCS)
                             Character data that contain DBCS not read correctly.
                          2. File size limited to 2GB ((2^31) - 2)
