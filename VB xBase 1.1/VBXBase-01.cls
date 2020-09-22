VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VB_xBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' ------------------------------------------------------------------------------
' All code in this class were writen by: Yunus
'
'   Email:    seruling_m4l4m@yahoo.com
'
' ------------------------------------------------------------------------------

'    VB_xBase Class - A VB6 Class for accessing DBF Files without ADO
'    Copyright (C) 2008, Achmad Junus
'
'    This library is free software; you can redistribute it and/or
'    modify it under the terms of the GNU Lesser General Public
'    License as published by the Free Software Foundation; either
'    version 2.1 of the License, or (at your option) any later version.
'
'    This library is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'    Lesser General Public License for more details.
'
'    You should have received a copy of the GNU Lesser General Public
'    License along with this library; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'
'=================================================================================

' ================================================================================
'           AVAILABLE DATA TYPES THAT COMPATIBLE WITH FoxPro or xBase:           '
' ================================================================================
Private Const DT_B As Byte = 66 ' "B"   Double,                 IEEE 64bit (8 byte) floating point
Private Const DT_C As Byte = 67 ' "C"   Character String,       Max 255 characters (dBase, FoxBase, XBase default)
Private Const DT_D As Byte = 68 ' "D"   Date,                   YYYYMMDD
Private Const DT_I As Byte = 73 ' "I"   32bit Signed Integer    Signed 32bit integer
Private Const DT_L As Byte = 76 ' "L"   Logical,                T or F
Private Const DT_N As Byte = 78 ' "N"   Numeric String,         Up to 20 characters, include decimal separator and decimal digit
Private Const DT_F As Byte = 70 ' "F"   Float,                  Up to 20 characters in table, 8 byte in memory (double data type)
                                '                               Float stored to file in same way as numeric (N) data type
                                '                               when doing arithmetic operation, it converted to Double
Private Const DT_Y As Byte = 89 ' "Y"   Currency,               Currency data type, 64bit signed integer (large integer) scaled by 10000

' =========================================================
'           EXTRA DATA TYPES THAT YOU MAY NEED:           '
' =========================================================
Private Const DT_Z As Byte = 90 ' "Z"   Byte,                   Byte data type, 8 bit unsigned integer, 1 byte
Private Const DT_S As Byte = 83 ' "S"   Short,                  Short data type, 16 bit signed integer, 2 byte

' Unsupported data types are:
'   T   Date time   8 byte binary >> Not supported because VB Datetime not compatible with FoxPro Datetime
'   M   Memo        4 byte Binary
'   G   General     ...
'   P   Picture     ...
' ==============================================================================

' ========================= '
'   SUPPORTED FILE TYPES:   '
' ========================= '
Private Const FILE_DBASE As Byte = 3    ' FoxBase+, FoxPro, dBaseIII+, dBaseIV, no memo
                                        ' Operation supported: CREATE, Open, Read, Write.

Private Const FILE_FOXPRO As Byte = 48  ' Visual Foxpro III no memo
                                        ' Operation supported: Open, Read, Write, NO CREATE
                                        
' For file that already indexed by dBase, FoxPro or othe xBase application, the index will ignored.
' This class will use its own index system
' ------------------------------------------------------------------------------

' ///////////// CONSTANTS THAT FOLLOW DBF FORMAT
Private Const MAXFIELDLEN As Integer = 254          ' Max width, in byte, for Character string ("C") data type. This is apply for DBF file.
                                                    ' YOU MAY CHANGE TO Possible Maximum Allowed: 32766. But the file is no longer a valid DBF file
Private Const MAXNUMERICDIGIT As Byte = 20      ' Max width, in byte, for Numeric ("N") and Float ("F") data type
Private Const MAX_RECORDLEN As Long = 65500     ' Maximum record size, in byte, include delete flag
Private Const MAXFIELD As Integer = 255         ' Maximum number of fields in table
Private Const MAXFIELDNAMEWIDTH As Byte = 11    ' Max width of field name, in byte, include null terminator
' Delete flag field:
Private Const DELETED As Byte = 42              ' Asterisk (*)
Private Const NOTDELETED As Byte = 32           ' [Space]
' Other DBF constants
Private Const HEADER_TERMINATOR As Byte = 13    ' 0x0d
Private Const EOF_FLAG As Byte = 26             ' 0x1a
' ///////////// END CONSTANTS THAT FOLLOW DBF FORMAT

Private RecNumChange As Boolean
Private RecEdited As Boolean

' ======================================
' UDT for using with Unsigned integer (2 byte integer value greater than 32767)
Private Type MY_INT
    MyInt1 As Integer       ' 2 * 2 byte
    MyInt2 As Integer
End Type
Private Type MY_LONG
    MyLong1 As Long         ' 4 byte
End Type
Private MyLong As MY_LONG
Private myUint As MY_INT
' ======================================

' ======================================
' UDT to convert integer array to string
Private Type MY_INTSTRING
    sInt(0 To MAXFIELDLEN - 1) As Integer
End Type
Private Type MY_STRING
    iString As String * MAXFIELDLEN
End Type
' ======================================


' ======= Public Type for sending information about fields structure
' ------- And for processing current opened table
Private Type DBF_FILEINFO
    FileDescription As Byte     ' 0x03 - No memo
    LastYearUpd As Byte
    LastMonthUpd As Byte
    LastDayUpd As Byte
    NumberOfRecords As Long     ' Total records in file, include deleted records
    HeaderSize As Integer       ' Header size. Not offset for first data in file!
    RecordSize As Integer       ' Unsigned integer. Record size include delete flag
    ' Below are fields that we did not use
'    zReserved1 As Integer       ' Let it 0
'    zReserved2 As Integer       ' Let it 0
'    zMultiUse(0 To 11) As Byte  ' Let it 0
'    TableFlag As Byte       ' Let it 0
'    CodePageMark As Byte    ' For file created with this class will sets to 0x02 ( MS-DOS 850 )
'    zReserved3 As Integer   ' Let it 0
    ' An important information that not exist in DBF file header was: Number of fields  :)
End Type
' Header
Private FHeader As DBF_FILEINFO
' Blank Header, declared separately from UDT
Private zHeaderBlank As String             ' 20 byte fill with Chr(0)
Private FHRecordSize As Long    ' Recordsize include delete flag

'++++++++++++++++++++++++++++++++++++++++++++ DBF FIELD DESCRIPTOR
'Private Type DBF_FIELD_DESCRIPTOR
Private FIFieldName() As String         ' 11 byte string include null terminator
                                        ' When created, fieldname converted to UPPERCASE
                                        ' Space is not allowed for field name, any space will converted to underscore

Private FIFieldType() As Byte          ' 1 character string, OR 1 byte
    
Private FIFieldDisplacement() As Long   ' start byte offset for each field in a record
                                        ' We need this for processing record with byte array

Private FIFieldWidth() As Byte
Private FIFieldDecimal() As Byte
        ' Note When creating new field:
        '   Field width are used for data type C, N and F, and ignored for other type.
        '   Field decimal use only for data type B, N and F, and ignored for other type
        '   For data type B (double), the decimal will rounded to decimal width before writen to file
    
Private FIFieldFlag() As Byte       ' 0x01 - System column (not visible to user, but visible to programmer :))
                                    ' 0x02 - Column can store null values <I don't have any clue how to use this flag>
                                    ' 0x04 - Binary column (flag for BINARY CHAR, DOUBLE, INTEGER, CURRENCY and MEMO)

' Below are field information that we did not use
'    zReserved as byte
'    zWorkAreaID As Byte
'    zReserved2(0 To 9) As Byte
'    zProductionIndex As Byte
Private FIFieldBlank As String      ' 13 byte fill with Chr(0).
                                    ' This is to replace unused field informations with one variable

'End Type
'++++++++++++++++++++++++++++++++++++++++++++ END DBF FIELD DESCRIPTOR

Public Enum xCompareOperator
    IsEqual = 5001
    IsNotEqual
    IsLessThan
    IsLessThanOrEqual
    IsGreaterThan
    IsGreaterThanOrEqual
    IsBeginWith
End Enum

Public Enum xRangeOperator
    IsBetween = 6000
    IsFrom = 7000
    IsTo = 7001
End Enum

' DBF File
Private nNomorFile As Integer
Private NamaFile As String

' File use flag
Private FileUsed As Boolean     ' True if a .dbf file already opened or new file is built

' Delete flag variables
Private RecordDeleted As Byte

' Internal Information
Private nNumberOfFields As Long
Private DataOffset As Long

' String for comunicate with caller or send string array as record result
Private RBlank() As String      ' String array without data

' One Byte
Private OneByte As Byte

' IO Read/Write buffer for special data type
Private dString As String
Private bDouble As Double
Private yCurrency As Currency
Private iInt32 As Long
Private sShort As Integer
Private zByte As Byte

' Date format keeper
Private DefDateFormat As String

Private oAutoUpdateHeader As Boolean
Private oAutoGetUpdatedHeader As Boolean

' Keep last position where search and find function find something
Private LastSearchPos As Long

Private nCurrentPos As Long

Dim s As String
Private Type BKeyInfo
    BytePos As Long
    Start As Long
    End As Long
End Type
Private Type bCounter
    pos(0 To 255) As Long
End Type

Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Sub GetMem1 Lib "Msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub GetMem2 Lib "Msvbvm60" (ByVal Addr As Long, RetVal As Integer)
Private Declare Sub GetMem4 Lib "Msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Sub PutMem1 Lib "Msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "Msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "Msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
' ==================================== INVOKE WINDOWS API ====================================
' Use copy memory to convert byte array into other data type, such double or currency
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemoryBV Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

' ==================================== E V E N T S ====================================
Public Event FileBuilt()    ' Fired when a new .dbf file successfully created
Public Event FileOpened()   ' Fired when a .dbf file successfully opened
' --------------------------------------------------------
' These events is used when sorting and indexing records.
' It need for user feedback since these process takes long time
Public Event RecRead(ByVal Total As Long)   ' Fired when start reading records
                            ' Parameter Total is total records read
Public Event SortStarted()  ' Fired when sorting record starts
Public Event RecWriten(ByVal Total As Long) ' Fired when some records writen
                                            ' to destination file
                                            ' Parameter Total is total records writen
Public Event SortCompleted()    ' Fired when sorting is complete

' -------------------------------------------------
Public Property Let AutoGetUpdatedHeader(ByVal yy As Boolean)
    oAutoGetUpdatedHeader = yy
End Property

' -------------------------------------------------
Public Property Get AutoGetUpdatedHeader() As Boolean
    AutoGetUpdatedHeader = oAutoGetUpdatedHeader
End Property

' -------------------------------------------------
Public Property Let AutoUpdateHeader(ByVal yy As Boolean)
    oAutoUpdateHeader = yy
End Property

' -------------------------------------------------
Public Property Get AutoUpdateHeader() As Boolean
    AutoUpdateHeader = oAutoUpdateHeader
End Property

' -------------------------------------------------
Public Property Get MaxPossibleRecords() As Long
    If FileUsed Then
        MaxPossibleRecords = ((2 ^ 31) - FHeader.HeaderSize - 2) \ FHRecordSize
    Else
        MaxPossibleRecords = 0
    End If
End Property

' -------------------------------------------------
Public Property Get RecordCount() As Long
    If FileUsed Then
        RecordCount = FHeader.NumberOfRecords
    Else
        RecordCount = -1
    End If
End Property

' -------------------------------------------------
Public Property Get FieldCount() As Long
    If FileUsed Then
        FieldCount = nNumberOfFields
    Else
        FieldCount = -1
    End If
End Property

' -------------------------------------------------
Public Property Get CurrentPosition() As Long
    Dim i As Long
    If FileUsed Then
        i = Seek(nNomorFile)
        If i = LOF(nNomorFile) Then
            CurrentPosition = -1
        Else
            CurrentPosition = ((i - FHeader.HeaderSize) / FHRecordSize) + 1
        End If
    Else
        CurrentPosition = -1
    End If
End Property


' -------------------------------------------------
' Function: CreateTable
' Description: Create as blank file
' Parameter:
'       sNamaFile as String:   complete file name with extension
'                   If no extension given, file will created without extension
'
' Return value: If success return file number. Otherwise -1
'
' Note:
' -----
    ' If file already exist, the file will overwriten
' -------------------------------------------------
Public Function CreateTable(sNamaFile As String) As Integer
    
    If Len(sNamaFile) > 0 Then
        If Not FileUsed Then
            nNomorFile = FreeFile
            Open sNamaFile For Binary Access Read Write Shared As #nNomorFile
            NamaFile = sNamaFile
            CreateTable = nNomorFile
        Else
            CreateTable = -1
        End If
    Else
        CreateTable = -1
    End If
    
End Function


' -------------------------------------------------
' Function: Create Field
' Description: Create or add a field for DBF file
' Parameters:
'       FieldName as String:  Field name
'                   Maximum 10 byte character string, without space
'                   The string will converted to UPPERCASE
'       sDataType as String:  Data types, one of these: C N D L F I Y B Z S
'
'       Optional parameters:
'       --------------------
'       FieldWidth as Byte:     Field width.
'                       If data type is D,L,I,Y,B,Z and S,  field width will ignored
'                       Data type C, N and F, when not assigned will sets to 10
'       DecimalPlace as Byte:   Decimal place. When not assigned, will sets to 0
'                       If data type is D,L,I,Y,B,Z and S decimal place will ignored
'       FieldIndexed as Boolean:   Default is False.
'
' Return Value: 1 if success, -1 if an error occur, and last field definition cancelled
'
' Extra note:
' -----------
    ' The first created field must begin with 1, and increase by 1 for next field
    ' I didn't write error checking for creating fields with field number not properly ordered
'
' -------------------------------------------------
Public Function CreateField(FieldName As String, sDataType As String, Optional FieldWidth As Byte = 0, Optional ByVal DecimalPlace As Byte = 0) As Long
'    On Error Resume Next
    
    Dim i As Long
    Dim TempB As Byte
    Dim Jenis As String
    Dim sDummy As String
    
    If FileUsed Then
        Debug.Print "File already open"
        CreateField = -1
        Exit Function
    End If
    
    If nNumberOfFields < MAXFIELD Then
        nNumberOfFields = nNumberOfFields + 1
    Else
        Debug.Print "Too many fields"
        CreateField = -2
        Exit Function
    End If
    
    ReDim Preserve FIFieldName(0 To nNumberOfFields - 1)
    ReDim Preserve FIFieldType(0 To nNumberOfFields - 1)
    ReDim Preserve FIFieldDisplacement(0 To nNumberOfFields - 1)
    ReDim Preserve FIFieldWidth(0 To nNumberOfFields - 1)
    ReDim Preserve FIFieldDecimal(0 To nNumberOfFields - 1)
    ReDim Preserve FIFieldFlag(0 To nNumberOfFields - 1)
    
    ' ============ FIELD NAME ============
    If Len(FieldName) = 0 Then
        CreateField = -3
        Debug.Print "fieldname len is 0"
        Exit Function
    End If
    
    ' Manage spaces and string length on field name
    sDummy = Replace(Trim(FieldName), " ", "_")
    FieldName = UCase(sDummy)
    Jenis = String(10, Chr(0))
    LSet Jenis = FieldName
    
    ' Check if field name already used by other fields
    If nNumberOfFields > 1 Then
        For i = 0 To nNumberOfFields - 2
            If StrComp(Jenis, FIFieldName(i)) = 0 Then
                nNumberOfFields = nNumberOfFields - 1
                CreateField = -4    ' Field name already used
                Debug.Print "Fieldname already used"    ' or Raise Error
                Exit Function
            End If
        Next i
    End If
    
    ' Copy and add null terminator for field name
    FIFieldName(nNumberOfFields - 1) = String(MAXFIELDNAMEWIDTH, Chr(0))
    LSet FIFieldName(nNumberOfFields - 1) = Jenis
    
    ' ============ FIELD TYPE ============
    '       ============ FIELD WIDTH ============
    '               ============ FIELD DECIMAL WIDTH ============
    If Len(sDataType) = 1 Then
        
        sDataType = UCase(sDataType)
        Select Case sDataType
            
            Case "C"    ' Character String
                FIFieldType(nNumberOfFields - 1) = DT_C
                
                If FieldWidth = 0 Then
                    FIFieldWidth(nNumberOfFields - 1) = 10
                ElseIf FieldWidth > MAXFIELDLEN Then
                    FIFieldWidth(nNumberOfFields - 1) = MAXFIELDLEN
                Else
                    FIFieldWidth(nNumberOfFields - 1) = FieldWidth
                End If
            
                FIFieldDecimal(nNumberOfFields - 1) = 0
                
                
            Case "N"    ' Numeric String
                FIFieldType(nNumberOfFields - 1) = DT_N
            
                If FieldWidth = 0 Then
                    FIFieldWidth(nNumberOfFields - 1) = 10
                ElseIf FieldWidth > MAXNUMERICDIGIT Then
                    FIFieldWidth(nNumberOfFields - 1) = MAXNUMERICDIGIT
                Else
                    FIFieldWidth(nNumberOfFields - 1) = FieldWidth
                End If
            
                If DecimalPlace > 0 Then
                    If DecimalPlace >= FIFieldWidth(nNumberOfFields - 1) Then
                        FIFieldDecimal(nNumberOfFields - 1) = FIFieldWidth(nNumberOfFields - 1) - 1
                    Else
                        FIFieldDecimal(nNumberOfFields - 1) = DecimalPlace
                    End If
                Else
                    FIFieldDecimal(nNumberOfFields - 1) = 0
                End If
                
            Case "L"    ' Logical... Perhaps, index not needed for this type
                FIFieldType(nNumberOfFields - 1) = DT_L
                FIFieldWidth(nNumberOfFields - 1) = 1
                FIFieldDecimal(nNumberOfFields - 1) = 0
            
            Case "F"    ' Float (Single)
                FIFieldType(nNumberOfFields - 1) = DT_F
                
                If FieldWidth = 0 Then
                    FIFieldWidth(nNumberOfFields - 1) = 10
                ElseIf FieldWidth > MAXNUMERICDIGIT Then
                    FIFieldWidth(nNumberOfFields - 1) = MAXNUMERICDIGIT
                Else
                    FIFieldWidth(nNumberOfFields - 1) = FieldWidth
                End If
            
                If DecimalPlace > 0 Then
                    If DecimalPlace >= FIFieldWidth(nNumberOfFields - 1) Then
                        FIFieldDecimal(nNumberOfFields - 1) = FIFieldWidth(nNumberOfFields - 1) - 1
                    Else
                        FIFieldDecimal(nNumberOfFields - 1) = DecimalPlace
                    End If
                Else
                    FIFieldDecimal(nNumberOfFields - 1) = 0
                End If
            
            Case "I"    ' Integer (Long, signed)
                FIFieldType(nNumberOfFields - 1) = DT_I
                FIFieldWidth(nNumberOfFields - 1) = 4
                FIFieldDecimal(nNumberOfFields - 1) = 0
                FIFieldFlag(nNumberOfFields - 1) = 4
            
            Case "Y"    ' Currency
                FIFieldType(nNumberOfFields - 1) = DT_Y
                FIFieldWidth(nNumberOfFields - 1) = 8
                FIFieldDecimal(nNumberOfFields - 1) = 4
                FIFieldFlag(nNumberOfFields - 1) = 4
            
            Case "B"    ' Double
                FIFieldType(nNumberOfFields - 1) = DT_B
                FIFieldWidth(nNumberOfFields - 1) = 8
                ' in data type double, decimal place is used for rounding values and
                ' convert to text string
                If DecimalPlace > 0 And DecimalPlace < 20 Then
                    FIFieldDecimal(nNumberOfFields - 1) = DecimalPlace
                Else
                    FIFieldDecimal(nNumberOfFields - 1) = 0
                End If
                FIFieldFlag(nNumberOfFields - 1) = 4
            
            Case "D"    ' Date, YYYYMMDD
                FIFieldType(nNumberOfFields - 1) = DT_D
                FIFieldWidth(nNumberOfFields - 1) = 8
                FIFieldDecimal(nNumberOfFields - 1) = 0
            
            Case "Z"    ' 8 bit unsigned Integer (byte)
                FIFieldType(nNumberOfFields - 1) = DT_Z
                FIFieldWidth(nNumberOfFields - 1) = 1
                FIFieldDecimal(nNumberOfFields - 1) = 0
                FIFieldFlag(nNumberOfFields - 1) = 4
            
            Case "S"    ' 16 bit signed Integer (Integer)
                FIFieldType(nNumberOfFields - 1) = DT_S
                FIFieldWidth(nNumberOfFields - 1) = 2
                FIFieldDecimal(nNumberOfFields - 1) = 0
                FIFieldFlag(nNumberOfFields - 1) = 4
            
            Case Else
                nNumberOfFields = nNumberOfFields - 1
                CreateField = -5
                Exit Function
        End Select
    
    Else ' Len(sDataType) Not Equal 1
        
        nNumberOfFields = nNumberOfFields - 1
        Debug.Print "invalid data type"
        CreateField = -6
        Exit Function
    
    End If
    
    ' ============ FIELD DISPLACEMENT ============
    If nNumberOfFields = 1 Then
        FIFieldDisplacement(nNumberOfFields - 1) = 1
    Else
        FIFieldDisplacement(nNumberOfFields - 1) = FIFieldDisplacement(nNumberOfFields - 2) + FIFieldWidth(nNumberOfFields - 2)
    End If
    
    ' ============ CHECKING TOTAL FIELD WIDTH ============
    If CLng(FIFieldDisplacement(nNumberOfFields - 1) + FIFieldWidth(nNumberOfFields - 1)) > MAX_RECORDLEN Then
        nNumberOfFields = nNumberOfFields - 1
        Debug.Print "Fieldname size too big"
        CreateField = -7
        Exit Function
    End If
    
    CreateField = 1
End Function


' -------------------------------------------------
' Function: FieldNameMatch
' Description: Compare given string to field name
' Parameter:
'       sString as String: String to evaluate
'
' Return value:
'       If sString match with one of field names in array, index of field name will returned
'       Otherwise, return -1
' -------------------------------------------------
Private Function FieldNameMatch(sString As String) As Long
    On Error GoTo SALAH
    Dim i As Long
    Dim crByte() As Byte
    
    If Not FileUsed Then
        FieldNameMatch = -1
        Exit Function
    End If
    
    If nNumberOfFields < 1 Then
        FieldNameMatch = -1
        Exit Function
    End If
    
    sString = Trim$(UCase$(sString))
    If Len(sString) < 1 Then
        FieldNameMatch = -1
    ElseIf Len(sString) > 10 Then
        FieldNameMatch = -1
    Else
        s = String$(MAXFIELDNAMEWIDTH, Chr(0))
        Mid(s, 1) = sString
        For i = 0 To nNumberOfFields - 1
            If StrComp(s, FIFieldName(i)) = 0 Then
                FieldNameMatch = i
                Exit Function
            End If
        Next i
    End If

SALAH:
    FieldNameMatch = -1
    Exit Function
End Function


' -------------------------------------------------
' Function: BuildFile
' Description: Writing new file to Disk
' Parameter: None
'
' Return value: If success return 1, otherwise -1
' -------------------------------------------------
Public Function BuildFile() As Long
    Dim BuildOK As Long
    Dim X() As Long
    Dim i, j As Long
    Dim HariIni As String
    Dim crByte() As Byte
    
    If Not FileUsed Then
        
        If nNumberOfFields > 0 Then
            
            HariIni = Format(Now, "YYMMDD")
            ' ============ UPDATE HEADER INFORMATION ============
            With FHeader
                .FileDescription = 3    ' DBF III, III+, IV, No Memo
                .LastYearUpd = CByte(Val(Left(HariIni, 2)))
                .LastMonthUpd = CByte(Val(Mid(HariIni, 3, 2)))
                .LastDayUpd = CByte(Val(Right(HariIni, 2)))
                .NumberOfRecords = 0
                FHRecordSize = FIFieldDisplacement(nNumberOfFields - 1) + FIFieldWidth(nNumberOfFields - 1)
                MyLong.MyLong1 = FHRecordSize
                LSet myUint = MyLong
                .RecordSize = myUint.MyInt1
                .HeaderSize = (nNumberOfFields * 32) + 33
            End With
            
            ' Writing new file header
            Seek #nNomorFile, 1
            Put #nNomorFile, , FHeader
            HariIni = String(17, 0)
            ' We need this only once. Doesn't need it anymore
            Put #nNomorFile, , HariIni
            HariIni = String(1, 2)          ' Codepage MS-DOS 850
            Put #nNomorFile, , HariIni
            HariIni = String(2, 0)
            Put #nNomorFile, , HariIni
            
            ' Writing new fields descriptions
            For i = 0 To nNumberOfFields - 1
                crByte = StrConv(Trim(FIFieldName(i)), vbFromUnicode)
                ReDim Preserve crByte(0 To 10)
                Put #nNomorFile, , crByte()
                FIFieldName(i) = StrConv(crByte, vbUnicode)
                
                Put #nNomorFile, , FIFieldType(i)
                Put #nNomorFile, , FIFieldDisplacement(i)
                Put #nNomorFile, , FIFieldWidth(i)
                Put #nNomorFile, , FIFieldDecimal(i)
                Put #nNomorFile, , FIFieldFlag(i)
                Put #nNomorFile, , FIFieldBlank
            Next i
            
            OneByte = HEADER_TERMINATOR
            Put #nNomorFile, , OneByte
        Else
            BuildFile = -1  ' No fields descriptions
            Exit Function
        End If
        
        ' =============================================
        ' ========= SOME EXTRA WORK HERE ==============
        ' =============================================
        ' preparing IO Buffer so that we can use it without need to resize
        ReDim RBlank(0 To nNumberOfFields - 1)
    
        For i = 0 To nNumberOfFields - 1
            RBlank(i) = ""
        Next i
        ' =============================================
    
    Else
        BuildFile = -1
        Exit Function
    End If
    
    ' Update class flag and raise some event
    DataOffset = FHeader.HeaderSize + 1
    FileUsed = True
    RaiseEvent FileBuilt
    RaiseEvent FileOpened
    nCurrentPos = 1
    BuildFile = 1
    
End Function

' -------------------------------------------------
' Sub: DBFOpenFile
' Description: Open a table file in this class instance
' Parameters:
'       SNamaFile As String: Complete path and file name with extension
'       [Optional] UseMode As Long = 1: Open mode, shared or not shared.
'               Value for open mode are:    1 (default), file is shared
'                                           7   File opened exclusive
'
' Return value: File number of opened file upon successful. Otherwise value less than 0 is returned
'
' -------------------------------------------------
Public Function DBFOpenFile(ByVal sNamaFile As String, Optional UseMode As Long = 1) As Integer
    On Error GoTo SALAH
    Dim i As Long
    Dim fn As Integer
    Dim FileType As Byte
    
    If Len(sNamaFile) < 1 Then  ' File name not valid
        DBFOpenFile = -9  ' File name not valid
        Exit Function
    Else    ' Check if file exist
        If Dir(sNamaFile) = "" Then
            DBFOpenFile = -9  ' File name not valid
            Exit Function
        End If
    End If
    
    If Not FileUsed Then
        fn = FreeFile
        If UseMode = 7 Then     ' Exclusive
            Open sNamaFile For Binary Access Read Write Lock Read Write As #fn
        Else                    ' Shared
            Open sNamaFile For Binary Access Read Write Shared As #fn
        End If
        
    Else
        DBFOpenFile = -99  ' Still opening other file. Can't open other file without closing it
        Exit Function
    End If
    
    ' Checking if opened file a valid DBF file
    If LOF(fn) < 65 Then
        DBFOpenFile = -19  ' File format not valid or has no field definition
        Exit Function
    End If
    
    Get #fn, 1, OneByte
    Select Case OneByte
        Case 3      ' FoxBase+, FoxPro, dBaseIII+, dBaseIV, no memo
            FileType = FILE_DBASE
        Case 48     ' Visual Foxpro III no memo
            FileType = FILE_FOXPRO
        Case 2, 67, 99, 131, 139, 142, 203, 245, 251
            DBFOpenFile = -2  ' File format not yet supported
            Exit Function
        Case Else
            DBFOpenFile = -19  ' Invalid file format
            Exit Function
    End Select
    
    ' Now checking entire file header
    Seek #fn, 1
    Get #fn, , FHeader
    Get #fn, , zHeaderBlank     ' adjust file pointer
    
    ' Convert unsigned integer to Long
    myUint.MyInt1 = FHeader.RecordSize
    LSet MyLong = myUint
    FHRecordSize = MyLong.MyLong1
    
    ' Checking Month update
    If FHeader.LastMonthUpd > 12 Or FHeader.LastMonthUpd = 0 Then
        DBFOpenFile = -19  ' Invalid file format
        Exit Function
    End If
    
    If FHeader.HeaderSize < 33 Or FHeader.HeaderSize > 8456 Then
        DBFOpenFile = -19  ' Invalid file format
        Exit Function
    Else
        Select Case FileType
            Case FILE_DBASE
                nNumberOfFields = (FHeader.HeaderSize - 33) \ 32
            Case FILE_FOXPRO
                nNumberOfFields = (FHeader.HeaderSize - (33 + 263)) \ 32
        End Select
        
        If nNumberOfFields < 1 Then
            DBFOpenFile = -19  ' Invalid file format
            Exit Function
        Else
            ReDim FIFieldName(0 To nNumberOfFields - 1)
            ReDim FIFieldType(0 To nNumberOfFields - 1)
            ReDim FIFieldDisplacement(0 To nNumberOfFields - 1)
            ReDim FIFieldWidth(0 To nNumberOfFields - 1)
            ReDim FIFieldDecimal(0 To nNumberOfFields - 1)
            ReDim FIFieldFlag(0 To nNumberOfFields - 1)
        End If
    End If
        
    ' Now loading field descriptor
    NamaFile = sNamaFile
    nNomorFile = fn
    
    For i = 0 To nNumberOfFields - 1
        FIFieldName(i) = Space(MAXFIELDNAMEWIDTH)
        Get #nNomorFile, , FIFieldName(i)
        Get #nNomorFile, , FIFieldType(i)
        Get #nNomorFile, , FIFieldDisplacement(i)
        Get #nNomorFile, , FIFieldWidth(i)
        Get #nNomorFile, , FIFieldDecimal(i)
        Get #nNomorFile, , FIFieldFlag(i)
        Get #nNomorFile, , FIFieldBlank
    Next i
    
    ' =============================================
    ' ========= SOME EXTRA WORK HERE ==============
    ' =============================================
    ' preparing IO Buffer so that we can use it without need to resize
    ReDim RBlank(0 To nNumberOfFields - 1)

    For i = 0 To nNumberOfFields - 1
        RBlank(i) = ""
    Next i
    ' =============================================
    
    ' File is opened and ready for use
    DataOffset = FHeader.HeaderSize + 1
    FileUsed = True
    MoveFirst       ' Sets position to first record
    DBFOpenFile = fn
    RaiseEvent FileOpened
    
    Exit Function

SALAH:
    
    If Err.Number = 52 Or Err.Number = 67 Then
        Debug.Print Err.Number; Err.Description
        fn = CInt(Err.Number)
        Err.Clear
        DBFOpenFile = fn * -1
        Exit Function
    Else
        Debug.Print Err.Number; Err.Description
        Err.Clear
        DBFOpenFile = -1
        Exit Function
    End If
End Function

' -------------------------------------------------
' Sub: DBFClose
' Description: Close all opened files in this class instance
'
' Return value: None
'
' -------------------------------------------------
Public Sub DBFClose()
    On Error GoTo E_CLOSE
    Dim i As Long
    
    If FileUsed Then
        UpdateHeader
        If RecNumChange Then
            OneByte = EOF_FLAG
            Seek #nNomorFile, LOF(nNomorFile) + 1
            Put #nNomorFile, , OneByte
        End If
        Close #nNomorFile
        FileUsed = False
    End If
    ' Clear fields information
    nNumberOfFields = 0
    Erase FIFieldName
    Erase FIFieldType
    Erase FIFieldDisplacement
    Erase FIFieldWidth
    Erase FIFieldDecimal
    Erase FIFieldFlag
    
    ' Cleanup header information
    With FHeader
        .FileDescription = 0
        .HeaderSize = 0
        .LastDayUpd = 0
        .LastMonthUpd = 0
        .LastYearUpd = 0
        .NumberOfRecords = 0
        .RecordSize = 0
    End With
    
    Exit Sub

E_CLOSE:
    FileUsed = False
End Sub


' -------------------------------------------------
' Function: AppendFromStrings
' Description:  Add a row (add a record) to table
'               This function is similar to Insert data
' Parameter:
'       [ParamArray] sStrings() As Variant: One or more character strings represent each field data
'       All data type must presented as string even if data declared as numeric, double or anything else
'
'       Example:    There are 2 fields of Date data type and Currency data type
'                   This command example will successful for those data types:
'                   Result = AppendFromStrings("May 31 2007", "$5216542.20")    ' All data must in string format
'
' Return value: 1=Success, -1=Failed
'
' Note:
    ' Faster way for append data is using Append() function
' -------------------------------------------------
Public Function AppendFromStrings(ParamArray sStrings()) As Long
    On Error GoTo AFS_ERROR
    Dim i As Long
    Dim sNewString() As String
    
    If FileUsed Then
        If UBound(sStrings) <> nNumberOfFields - 1 Then
            AppendFromStrings = -1
            Exit Function
        
        Else
            ReDim sNewString(LBound(sStrings) To UBound(sStrings))
            For i = LBound(sStrings) To UBound(sStrings)
                sNewString(i) = CStr(sStrings(i))
            Next i
            i = Append(sNewString)
            AppendFromStrings = i
        End If
    'Else
        ' Do nothing or raise error
    End If
    Exit Function
    
AFS_ERROR:
    AppendFromStrings = -1
    Err.Clear
End Function


' -------------------------------------------------
' Function: Append
' Description:  Add a row (add a record) to table
'               This function is similar to Insert data
' Parameter:
'       Datanya() As String: Array od character string contains field data.
'       All data type must presented as string
'
' Return value: 1=Success, -1=Failed
'
' -------------------------------------------------
Public Function Append(Datanya() As String) As Long
    On Error GoTo E_APPEND
    Dim i As Long
    Dim String1 As String
    Dim StringX As String
    
    If FileUsed Then
        If oAutoGetUpdatedHeader Then GetUpdatedHeader
        
        If LBound(Datanya) <> 0 Then
            Append = -1
            Exit Function
        End If
        
        If UBound(Datanya) > nNumberOfFields - 1 Then
            Append = -1
            Exit Function
        End If
        
        If FHeader.NumberOfRecords > 0 Then
            If (LOF(nNomorFile) - FHeader.HeaderSize) Mod FHRecordSize = 1 Then
                i = LOF(nNomorFile)
            Else
                i = LOF(nNomorFile) + 1
            End If
        Else
            i = LOF(nNomorFile) + 1
        End If
        
        RecordDeleted = NOTDELETED
        Seek #nNomorFile, i
        Put #nNomorFile, , RecordDeleted

        If UBound(Datanya) = nNumberOfFields - 1 Then
            
            For i = 0 To nNumberOfFields - 1
                Select Case FIFieldType(i)
                    
                    Case DT_C ' "C"
                        
                        If Len(Datanya(i)) = FIFieldWidth(i) Then
                            Put #nNomorFile, , Datanya(i)
                        ElseIf Len(Datanya(i)) < FIFieldWidth(i) Then
                            String1 = Space(FIFieldWidth(i))
                            LSet String1 = Datanya(i)
                            Put #nNomorFile, , String1
                        ElseIf Len(Datanya(i)) > FIFieldWidth(i) Then
                            String1 = Left(Datanya(i), FIFieldWidth(i))
                            Put #nNomorFile, , String1
                        End If
                    
                    Case DT_N, DT_F ' NUMERIC and FLOAT
                        String1 = Space(FIFieldWidth(i))
                        If IsNumeric(Datanya(i)) Then
                            If FIFieldDecimal(i) = 0 Then
                                StringX = Val(Datanya(i))
                            Else
                                If FIFieldWidth(i) - FIFieldDecimal(i) = 1 Then
                                    StringX = Format(Val(Datanya(i)), "." & String(FIFieldDecimal(i), "0"))
                                Else
                                    StringX = Format(Val(Datanya(i)), String(FIFieldWidth(i) - FIFieldDecimal(i) - 1, "#") & "." & String(FIFieldDecimal(i), "0"))
                                End If
                            End If
                            
                        Else
                            If FIFieldDecimal(i) = 0 Then
                                StringX = "0"
                            Else
                                If FIFieldWidth(i) - FIFieldDecimal(i) = 1 Then
                                    StringX = Format(0, "." & String(FIFieldDecimal(i), "0"))
                                Else
                                    StringX = Format(0, String(FIFieldWidth(i) - FIFieldDecimal(i) - 1, "#") & "." & String(FIFieldDecimal(i), "0"))
                                End If
                            End If
                        End If
                        
                        RSet String1 = StringX      ' Right align
                        ' replace decimal separator with period (.)
                        'If FIFieldDecimal(i) > 0 Then Mid(String1, FIFieldWidth(i) - FIFieldDecimal(i)) = "."
                        If FIFieldDecimal(i) > 0 Then PutMem1 StrPtr(String1) + ((FIFieldWidth(i) - FIFieldDecimal(i) - 1) * 2), 46
                        Put #nNomorFile, , String1
                    
                    Case DT_D ' "D"
                        If IsDate(Datanya(i)) Then
                            String1 = Format$(DateValue(Datanya(i)), DefDateFormat)
                        Else
                            String1 = Space(8)
                        End If
                        Put #nNomorFile, , String1
                    
                    Case DT_L ' "L"
                        
                        If Len(Datanya(i)) = 1 Then
                            Select Case Datanya(i)
                                Case "T", "t"
                                    String1 = "T"
                                Case Else
                                    String1 = "F"
                            End Select
                        Else
                            String1 = "F"
                        End If
                        Put #nNomorFile, , String1
                    
                    Case DT_B ' "B"
                        bDouble = IIf(IsNumeric(Datanya(i)), Round(CDbl(Val(Datanya(i))), FIFieldDecimal(i)), 0)
                        Put #nNomorFile, , bDouble
                    
                    Case DT_I ' "I"
                        iInt32 = IIf(IsNumeric(Datanya(i)), CLng(Val(Datanya(i))), 0)
                        Put #nNomorFile, , iInt32
                    
                    Case DT_Y ' "Y"
                        yCurrency = IIf(IsNumeric(Datanya(i)), CCur(Val(Datanya(i))), 0)
                        Put #nNomorFile, , yCurrency
                    
                    Case DT_Z ' "Z"
                        zByte = IIf(IsNumeric(Datanya(i)), CByte(Val(Datanya(i))), 0)
                        Put #nNomorFile, , zByte
                    
                    Case DT_S ' "S"
                        sShort = IIf(IsNumeric(Datanya(i)), CInt(Val(Datanya(i))), 0)
                        Put #nNomorFile, , sShort
                    
                    Case Else       ' "G"
                        If Len(Datanya(i)) = FIFieldWidth(i) Then
                            Put #nNomorFile, , Datanya(i)
                        ElseIf Len(Datanya(i)) < FIFieldWidth(i) Then
                            String1 = String(FIFieldWidth(i), Chr(0))
                            LSet String1 = StringX
                            Put #nNomorFile, , String1
                        ElseIf Len(Datanya(i)) > FIFieldWidth(i) Then
                            String1 = Left(Datanya(i), FIFieldWidth(i))
                            Put #nNomorFile, , String1
                        End If
                
                End Select
            
            Next i
            
        End If
        
        ' update header in memory
        FHeader.NumberOfRecords = FHeader.NumberOfRecords + 1      ' Di update tapi tidak perlu di tulis dulu
        nCurrentPos = -1    ' At the end of file
        RecNumChange = True
        If oAutoUpdateHeader Then UpdateHeader      ' Update header in disk
        Append = 1
    Else
        Append = -1
    End If
    
    Exit Function

E_APPEND:
    Append = -1
    Err.Clear
End Function

' -------------------------------------------------
' Function: ReplaceRecord
' Description: Replace record of current position with new data (Edit data)
' Parameters:
'       NomorRecord As Long: Record position number.
'       Datanya() As String: Array of character string to write to update record
'
' Return value: 1=Success, -1=Failed
'
' -------------------------------------------------
Public Function ReplaceRecord(NomorRecord As Long, Datanya() As String) As Long
    ' This function is similar with append, except the write position is
    ' not in end of file
    On Error GoTo E_REPLACERCD
    Dim i As Long
    Dim String1 As String
    Dim StringX As String
    
    If FileUsed Then
        If LBound(Datanya) <> 0 Then
            ReplaceRecord = -1
            Exit Function
        End If
        
        If UBound(Datanya) > nNumberOfFields - 1 Then
            ReplaceRecord = -1
            Exit Function
        End If
        
        If NomorRecord > FHeader.NumberOfRecords Or NomorRecord < 1 Then
            ReplaceRecord = -1
            Exit Function
        End If
        
        Seek #nNomorFile, DataOffset + ((NomorRecord - 1) * FHRecordSize)
        Get #nNomorFile, , RecordDeleted
        
        If RecordDeleted = DELETED Then
            ReplaceRecord = -1      ' Can't replace on deleted record. Undelete first
            Exit Function
        End If
        
        Seek #nNomorFile, DataOffset + ((NomorRecord - 1) * FHRecordSize) + 1

        If UBound(Datanya) = nNumberOfFields - 1 Then
            
            For i = 0 To nNumberOfFields - 1
                Select Case FIFieldType(i)
                    
                    Case DT_C ' "C"
                        
                        If Len(Datanya(i)) = FIFieldWidth(i) Then
                            Put #nNomorFile, , Datanya(i)
                        ElseIf Len(Datanya(i)) < FIFieldWidth(i) Then
                            String1 = Space(FIFieldWidth(i))
                            LSet String1 = Datanya(i)
                            Put #nNomorFile, , String1
                        ElseIf Len(Datanya(i)) > FIFieldWidth(i) Then
                            String1 = Left(Datanya(i), FIFieldWidth(i))
                            Put #nNomorFile, , String1
                        End If
                    
                    Case DT_N, DT_F ' "N"
                        String1 = Space(FIFieldWidth(i))
                        If IsNumeric(Datanya(i)) Then
                            If FIFieldDecimal(i) = 0 Then
                                StringX = Val(Datanya(i))
                            Else
                                If FIFieldWidth(i) - FIFieldDecimal(i) = 1 Then
                                    StringX = Format(Val(Datanya(i)), "." & String(FIFieldDecimal(i), "0"))
                                Else
                                    StringX = Format(Val(Datanya(i)), String(FIFieldWidth(i) - FIFieldDecimal(i) - 1, "#") & "." & String(FIFieldDecimal(i), "0"))
                                End If
                            End If
                            
                        Else
                            If FIFieldDecimal(i) = 0 Then
                                StringX = "0"
                            Else
                                If FIFieldWidth(i) - FIFieldDecimal(i) = 1 Then
                                    StringX = Format(0, "." & String(FIFieldDecimal(i), "0"))
                                Else
                                    StringX = Format(0, String(FIFieldWidth(i) - FIFieldDecimal(i) - 1, "#") & "." & String(FIFieldDecimal(i), "0"))
                                End If
                            End If
                        End If
                        
                        RSet String1 = StringX      ' Right align
                        Put #nNomorFile, , String1
                    
                    Case DT_D ' "D"
                        If IsDate(Datanya(i)) Then
                            String1 = Format(Datanya(i), DefDateFormat)
                        Else
                            String1 = Space(8)
                        End If
                        Put #nNomorFile, , String1
                    
                    Case DT_L ' "L"
                        
                        If Len(Datanya(i)) = 1 Then
                            Select Case Datanya(i)
                                Case "T", "t"
                                    String1 = "T"
                                Case Else
                                    String1 = "F"
                            End Select
                        Else
                            String1 = "?"
                        End If
                        Put #nNomorFile, , String1
                    
                    Case DT_B ' "B"
                        bDouble = IIf(IsNumeric(Datanya(i)), Round(CDbl(Datanya(i)), FIFieldDecimal(i)), 0)
                        Put #nNomorFile, , bDouble
                    
                    Case DT_I ' "I"
                        iInt32 = IIf(IsNumeric(Datanya(i)), CLng(Datanya(i)), 0)
                        Put #nNomorFile, , iInt32
                    
                    Case DT_Y ' "Y"
                        yCurrency = IIf(IsNumeric(Datanya(i)), CCur(Datanya(i)), 0)
                        Put #nNomorFile, , yCurrency
                    
                    Case DT_Z ' "Z"
                        zByte = IIf(IsNumeric(Datanya(i)), CByte(Val(Datanya(i))), 0)
                        Put #nNomorFile, , zByte
                    
                    Case DT_S ' "S"
                        sShort = IIf(IsNumeric(Datanya(i)), CInt(Val(Datanya(i))), 0)
                        Put #nNomorFile, , sShort
                    
                    Case Else       ' "G"
                        If Len(Datanya(i)) = FIFieldWidth(i) Then
                            Put #nNomorFile, , Datanya(i)
                        ElseIf Len(Datanya(i)) < FIFieldWidth(i) Then
                            String1 = String(FIFieldWidth(i), Chr(0))
                            LSet String1 = StringX
                            Put #nNomorFile, , String1
                        ElseIf Len(Datanya(i)) > FIFieldWidth(i) Then
                            String1 = Left(Datanya(i), FIFieldWidth(i))
                            Put #nNomorFile, , String1
                        End If
                
                End Select
            
            Next i
            
        End If
        
        nCurrentPos = IIf(NomorRecord = FHeader.NumberOfRecords, -1, NomorRecord + 1)
        RecEdited = True
        ReplaceRecord = 1
    Else
        ReplaceRecord = -1
    End If
    
    Exit Function

E_REPLACERCD:
    ReplaceRecord = -1
    Err.Clear
End Function


' -------------------------------------------------
' Function: GetRow
' Description:  Read row record and send the result as string array
' Parameter:
'       RecordNumber as Long:   Record number position
'       ByRef strResult() As String: Buffer to fill with row result
'
' Return value: 1 = Success, and strResult is filled by data
'               -1 = failed, and strResult contain zero length string
'
'   Note:
'   -----
'        Data type Date will presented as "MM/DD/YYYY"
'        Other data type will presented as it original in string form, except Numeric will
'        formatted as defined on field width and decimal width.
'        Empty date is "  /  /    "
'
' -------------------------------------------------
Public Function GetRow(RecordNumber As Long, ByRef strResult() As String) As Long
On Error GoTo E_GETROW
    Dim i As Long
    Dim Cok As Long
    Dim dDouble As Double
    Dim cCurrency As Currency
    Dim lLong As Long
    Dim sdString As String
    Dim varByte As Byte
    Dim varInteger As Integer
    
    
    If FileUsed Then
        
        If oAutoGetUpdatedHeader Then GetUpdatedHeader
        
        ReDim strResult(0 To nNumberOfFields - 1)
        If RecordNumber < 1 Or RecordNumber > FHeader.NumberOfRecords Then
            GetRow = -1
            Exit Function
        End If
        
        sdString = Space(8)
        
        Seek #nNomorFile, DataOffset + ((RecordNumber - 1) * FHRecordSize)
        Get #nNomorFile, , RecordDeleted
        
        If RecordDeleted = NOTDELETED Then
            For i = 0 To nNumberOfFields - 1
                Select Case FIFieldType(i)
                    Case DT_B   ' Double
                        Get #nNomorFile, , dDouble
                        If FIFieldDecimal(i) > 0 Then
                            strResult(i) = Format$(dDouble, "#." & String(FIFieldDecimal(i), "0"))
                            PutMem1 StrPtr(strResult(i)) + (((Len(strResult(i)) - FIFieldDecimal(i)) - 1) * 2), 46
                        Else
                            strResult(i) = CStr(dDouble)
                        End If
                                                            'strResult(i) = CStr(bDouble)
                    Case DT_C, DT_L ' Character string
                        strResult(i) = Space$(FIFieldWidth(i))
                        Get #nNomorFile, , strResult(i)
                    Case DT_D   ' Date
                        strResult(i) = String(10, "/")
                        Get #nNomorFile, , sdString
                        
                        ' Convert YYYYMMDD to MM/DD/YYYY
                        'strResult(i) = Mid$(sdString, 5, 2) & "/" & Right(sdString, 2) & "/" & Left(sdString, 4)
                        GetMem4 StrPtr(sdString) + 8, lLong
                        PutMem4 StrPtr(strResult(i)), lLong  ' MM

                        GetMem4 StrPtr(sdString) + 12, lLong
                        PutMem4 StrPtr(strResult(i)) + 6, lLong  ' DD

                        GetMem4 StrPtr(sdString), lLong
                        PutMem4 StrPtr(strResult(i)) + 12, lLong  ' YY
                        GetMem4 StrPtr(sdString) + 4, lLong
                        PutMem4 StrPtr(strResult(i)) + 16, lLong  ' YY
                        
                    
                    Case DT_I   ' Int32
                        Get #nNomorFile, , lLong
                        strResult(i) = CStr(lLong)
                    
                    Case DT_N, DT_F   ' Numeric string
                        strResult(i) = Space$(FIFieldWidth(i))
                        Get #nNomorFile, , strResult(i)
                        strResult(i) = LTrim$(strResult(i))
                    
                    Case DT_Y   ' Currency
                        Get #nNomorFile, , cCurrency
                        strResult(i) = Format$(cCurrency, "#.0000")
                        PutMem1 StrPtr(strResult(i)) + ((Len(strResult(i)) - 5) * 2), 46
                    
                    Case DT_Z   ' Byte
                        Get #nNomorFile, , varByte
                        strResult(i) = CStr(varByte)
                    
                    Case DT_S   ' Integer
                        Get #nNomorFile, , varInteger
                        strResult(i) = CStr(varInteger)
                    
                    Case Else
                        strResult(i) = Space$(FIFieldWidth(i))
                        Get #nNomorFile, , strResult(i)
                End Select
            Next i
            
            GetRow = 1
        
        Else    ' Record is deleted
            Seek #nNomorFile, DataOffset + (RecordNumber * FHRecordSize) ' Advance to next position
            GetRow = -1
        End If
    Else
        GetRow = -1
    End If
    Exit Function
    
E_GETROW:
    Err.Clear
    GetRow = -1
End Function

' -------------------------------------------------
' Function: GetRow2DA
' Description:  Read row record and send the result as 2 Dimension Array of string
'               This function is provided for faster IO operation when need
'               to retrieve many records at once
' Parameter:
'       RecordNumbers() as Long:   Array of record number position to read
'
' Return value: A 2 Dimensional Array of string, represent data that read from disk
'               if failed the contents of array are zero length strings
'
'               2D array for: 1st dimension for fields
'                             2nd dimension for rows
'   Note:
'   -----
'      - Data type Date will presented as "MM/DD/YYYY",
'      - Other data type will presented as it original in string form, except Numeric will
'        formatted as defined on field width and decimal width.
'      - Empty date is "  /  /    "
'
' -------------------------------------------------
Public Function GetRow2DA(RecordNumbers() As Long) As String()
On Error GoTo E_GETROW2DA
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim dDouble As Double
    Dim lpDouble As Long
    Dim cCurrency As Currency
    Dim lpCurrency As Long
    Dim vvLong As Long
    Dim varInteger As Integer
    Dim strResult() As String
    Dim lpstrResult As Long
    Dim Tempat As Long
    
    Dim ReadByte() As Byte
    Dim FieldPtr() As Long
    Dim TWidth() As Long
    Dim lpmyIntStr As Long
    Dim my_iString As MY_INTSTRING      ' UDT of fixed array of integers used for converting Byte array to String
                                        ' This method combined with CopyMemory is faster than StrConv
    '================================
    If Not FileUsed Then
        GetRow2DA = RBlank
        Exit Function
    End If
    
    If oAutoGetUpdatedHeader Then GetUpdatedHeader
    
'   Change array dimension to 2DA (2 Dimension Array)
'               2D array for: 1st dimension for fields
'                             2nd dimension for rows
    ReDim strResult(0 To nNumberOfFields - 1, 0 To UBound(RecordNumbers))
    
    ReDim ReadByte(0 To FHRecordSize - 1)
    ReDim FieldPtr(0 To nNumberOfFields - 1)
    ReDim TWidth(0 To nNumberOfFields - 1)
    lpmyIntStr = VarPtr(my_iString.sInt(0))
    For i = 0 To nNumberOfFields - 1
        FieldPtr(i) = VarPtr(ReadByte(FIFieldDisplacement(i)))
        TWidth(i) = CLng(FIFieldWidth(i))
    Next i
    
    dDouble = 0
    cCurrency = 0
    lpDouble = VarPtr(dDouble)
    lpCurrency = VarPtr(cCurrency)
    
    For i = 0 To UBound(RecordNumbers)
        
        Get #nNomorFile, DataOffset + ((RecordNumbers(i) - 1) * FHRecordSize), ReadByte
                
                If ReadByte(0) = NOTDELETED Then
                    
                    For j = 0 To nNumberOfFields - 1
                       
                        Select Case FIFieldType(j)
                            
                            Case DT_B
                                CopyMemoryBV lpDouble, FieldPtr(j), 8
                                If FIFieldDecimal(j) > 0 Then
                                    strResult(j, i) = Format$(dDouble, "#." & String(FIFieldDecimal(j), "0"))
                                    'Mid(strResult(j, i), Len(strResult(j, i)) - FIFieldDecimal(j)) = "."
                                    PutMem1 StrPtr(strResult(j, i)) + (((Len(strResult(j, i)) - FIFieldDecimal(j)) - 1) * 2), 46
                                Else
                                    strResult(j, i) = CStr(dDouble)
                                End If
                                
                            Case DT_C
                                strResult(j, i) = Space$(TWidth(j))
                                Tempat = FIFieldDisplacement(j)
                                For k = 0 To TWidth(j) - 1
                                    my_iString.sInt(k) = CInt(ReadByte(Tempat + k))
                                Next k
                                CopyMemoryBV StrPtr(strResult(j, i)), lpmyIntStr, TWidth(j) * 2

                            Case DT_D
                                strResult(j, i) = Space$(10)
                                Tempat = FIFieldDisplacement(j)
                                ' YYYY MM DD
                                ' 0123 45 67
                                ' M M / D D / Y Y Y Y
                                ' 0123456789012345678
                                ' 0 1 2 3 4 5 6 7 8 9
                                For k = 6 To 9
                                    my_iString.sInt(k) = CInt(ReadByte(Tempat + k - 6))
                                Next k
                                For k = 0 To 1
                                    my_iString.sInt(k) = CInt(ReadByte(Tempat + k + 4))
                                Next k
                                For k = 3 To 4
                                    my_iString.sInt(k) = CInt(ReadByte(Tempat + k + 3))
                                Next k
                                my_iString.sInt(2) = 47     ' Date separator "/"
                                my_iString.sInt(5) = 47     ' Date separator "/"
                                CopyMemoryBV StrPtr(strResult(j, i)), lpmyIntStr, 20
                            
                            Case DT_F, DT_N
                                strResult(j, i) = Space$(TWidth(j))
                                Tempat = FIFieldDisplacement(j)
                                For k = 0 To TWidth(j) - 1
                                    my_iString.sInt(k) = CInt(ReadByte(Tempat + k))
                                Next k
                                CopyMemoryBV StrPtr(strResult(j, i)), lpmyIntStr, TWidth(j) * 2
                                
                                ' Always remove leading spaces
                                strResult(j, i) = LTrim$(strResult(j, i))
                            
                            Case DT_I
                                GetMem4 FieldPtr(j), vvLong
                                strResult(j, i) = CStr(vvLong)
                            
                            Case DT_L
                                strResult(j, i) = Chr$(ReadByte(FIFieldDisplacement(j)))
                            
                            Case DT_Y
                                CopyMemoryBV lpCurrency, FieldPtr(j), 8
                                strResult(j, i) = Format$(cCurrency, "#0.0000")
                                'Mid(strResult(j, i), Len(strResult(j, i)) - 4) = "."
                                PutMem1 StrPtr(strResult(j, i)) + ((Len(strResult(j, i)) - 5) * 2), 46
                            
                            Case DT_Z
                                strResult(j, i) = CStr(ReadByte(FIFieldDisplacement(j)))
                            
                            Case DT_S
                                GetMem2 FieldPtr(j), varInteger
                                strResult(j, i) = CStr(varInteger)
                            
                            Case Else
                                strResult(j, i) = Space$(FIFieldWidth(j))
                                Tempat = FIFieldDisplacement(j)
                                For k = 0 To FIFieldWidth(j) - 1
                                    my_iString.sInt(k) = CInt(ReadByte(Tempat + k))
                                Next k
                                CopyMemoryBV StrPtr(strResult(j, i)), lpmyIntStr, TWidth(j) * 2
                                
                        End Select
                    Next j
                End If
    Next i
    Erase ReadByte
    Erase FieldPtr
    Erase TWidth
    GetRow2DA = strResult
    Exit Function
    
E_GETROW2DA:
    Err.Clear
    Erase ReadByte
    Erase FieldPtr
    Erase TWidth
    GetRow2DA = strResult

End Function

' -------------------------------------------------
' MoveFirst: Sets the Read-Write pointer to the first record
' Return value: none
'   Note: This function only move the record pointer (cursor), and not read any bytes
' -------------------------------------------------
Public Sub MoveFirst()
    If FileUsed Then
        Seek #nNomorFile, DataOffset
    End If
End Sub

' -------------------------------------------------
' MoveLast: Sets the Read-Write pointer to the last record
' Return value: none
'   Note: This function only move the record pointer (cursor), and not read any bytes
' -------------------------------------------------
Public Sub MoveLast()
    If FileUsed Then
        If oAutoGetUpdatedHeader Then GetUpdatedHeader
        Seek #nNomorFile, DataOffset + ((FHeader.NumberOfRecords - 1) * FHRecordSize)
    End If
End Sub

' -------------------------------------------------
' MoveNext: Sets the Read-Write pointer to next record
'
' Return value: none
'   Note: This function only move the record pointer (cursor), and not read any bytes
' -------------------------------------------------
Public Sub MoveNext()
    Dim i As Long
    If FileUsed Then
        i = TellPosition
        If oAutoGetUpdatedHeader Then GetUpdatedHeader
        If i = -1 Or i = FHeader.NumberOfRecords Then
            MoveLast
        Else
            MoveTo i
        End If
    End If
End Sub

' -------------------------------------------------
' MovePrevious: Sets the Read-Write pointer to previous record
'
' Return value: none
'   Note: This function only move the record pointer (cursor), and not read any bytes
' -------------------------------------------------
Public Sub MovePrevious()
    Dim i As Long
    If FileUsed Then
        i = TellPosition
        If i > 2 And i <= FHeader.NumberOfRecords + 1 Then
            MoveTo i - 2
        ElseIf i = -1 Then
            MoveTo FHeader.NumberOfRecords - 1
        Else
            MoveFirst
        End If
    End If
End Sub

' -------------------------------------------------
' MoveTo: Sets the Read-Write pointer to the specified record
' Return value: None
'
'   Note: This function only move the record pointer (cursor), and not read any bytes
' -------------------------------------------------
Public Sub MoveTo(nRecNum As Long)
    If FileUsed Then
        If nRecNum > FHeader.NumberOfRecords Then
            Seek #nNomorFile, LOF(nNomorFile)
        ElseIf nRecNum < 1 Then
            MoveFirst
        Else
            Seek #nNomorFile, DataOffset + ((nRecNum - 1) * FHRecordSize)
        End If
    Else
        Exit Sub
    End If
End Sub

' -------------------------------------------------
' TellPosition:
'       Return record number position where "NEXT" read/write operation will occur
'
' -------------------------------------------------
Public Function TellPosition() As Long
    Dim i As Long
    If FileUsed Then
        i = Seek(nNomorFile)
        If i = LOF(nNomorFile) Then
            TellPosition = -1
        Else
            TellPosition = ((i - FHeader.HeaderSize - 1) / FHRecordSize) + 1
        End If
    Else
        TellPosition = -1
    End If
End Function

' -------------------------------------------------
'
' Since the file by default opened in SHared mode
' It is possible that other process change the file contents.
' Function UpdateHeader and GetUpdatedHeader is made to prevent conflict occured by such situation
'
' It is recommended that each time a record appended, call UpdateHeader function immediately
'
' And it is recommended that call GetUpdatedHeader right before running search function or Append function
'
' By default, this class will update header and get updated header right before or after doinf critical operation
'
' You can set/get this clas properties AutoUpdateHeader and AutoGetUpdatedHeader
' -------------------------------------------------
Public Sub UpdateHeader()
    Dim HariIni As String
    
    If RecNumChange Or RecEdited Then
        HariIni = Format(Now, "YYMMDD")
        With FHeader
            .LastYearUpd = CByte(Val(Left(HariIni, 2)))
            .LastMonthUpd = CByte(Val(Mid(HariIni, 3, 2)))
            .LastDayUpd = CByte(Val(Right(HariIni, 2)))
        End With
        nCurrentPos = TellPosition
        Put #nNomorFile, 1, FHeader
        MoveTo nCurrentPos
    End If
    
End Sub

' -------------------------------------------------
Public Sub GetUpdatedHeader()
    
    nCurrentPos = TellPosition
    Get #nNomorFile, 1, FHeader
    MoveTo nCurrentPos
    
End Sub

' -------------------------------------------------
' Function: DeleteRecord
' Description: Put the delete flag for the record
' Parameter:
'       nRecNumber As Long: Record number position
'
' Return value: 1 Success. If failed return -1
'
' -------------------------------------------------
Public Function DeleteRecord(nRecNumber As Long) As Long
    RecordDeleted = DELETED
    ' Go to record number
    If FileUsed Then
        If oAutoGetUpdatedHeader Then GetUpdatedHeader
        If nRecNumber > FHeader.NumberOfRecords Or nRecNumber < 1 Then
            DeleteRecord = -1
        Else
            Put #nNomorFile, DataOffset + ((nRecNumber - 1) * FHRecordSize), RecordDeleted    ' Write a byte for delete flag
            RecEdited = True
            DeleteRecord = 1
            MoveTo nRecNumber + 1
        End If
        Exit Function
    Else
        DeleteRecord = -1
        Exit Function
    End If
End Function

' -------------------------------------------------
' Function: RecallRecord
' Description:  Remove the delete flag for the record
'               This is opposite of DeleteRecord function
' Parameter:
'       nRecNumber As Long: Record number to undelete
'
' Return value: number of record undeleted
'
' -------------------------------------------------
Public Function RecallRecord(nRecNumber As Long) As Long
    RecordDeleted = NOTDELETED
    ' Go to record number
    If FileUsed Then
        If oAutoGetUpdatedHeader Then GetUpdatedHeader
        If nRecNumber > FHeader.NumberOfRecords Or nRecNumber < 1 Then
            RecallRecord = -1
            Exit Function
        Else
            Put #nNomorFile, DataOffset + ((nRecNumber - 1) * FHRecordSize), RecordDeleted    ' Write a byte for delete flag
            RecEdited = True
            RecallRecord = 1
            MoveTo nRecNumber
            Exit Function
        End If
    Else
        RecallRecord = -1
        Exit Function
    End If
End Function


' -------------------------------------------------
' Function: FindText
' Description:  Find string or substring inside file
'               This function will evaluate all bytes in data, ignoring data type
' Parameter:
'       TextToFind As String    Text to find
'       MatchCase As Boolean    Optional. Search is case sensitive or case insensitive.
'                               Default is case sensitive
'       StartPos As Long        Optional. Record position where the search will stared.
'                               The first record is 1
'                               Default is -1 (Start from current position)
'
' Return value: If success return Record position where the text is found.
'               If fail, -1 is returned
'
'   Note:
'   I was forced to use integer, because VB doesn't support negative value for byte data type (signed char, or signed byte)
'
' -------------------------------------------------
Public Function FindText(TextToFind As String, Optional MatchCase As Boolean = True, Optional StartPos As Long = -1) As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Ujung As Long
    Dim ReadByte() As Byte
    Dim BufferSize As Long
    Dim RecLeft As Long
    Dim rPos As Long
    Dim BigLoop As Long
    Dim jLoop As Long
    Dim TinyLoop As Long
    Dim TTFInt() As Byte        ' Byte array for convert string to array of byte
    Dim TTFByte() As Integer    ' integer array for convert byte array to integer array
    Dim Match As Integer
    Dim PosKetemu As Long
    Dim T As Long
    Dim UB As Long
    
    If Not FileUsed Then FindText = -1
    If Len(TextToFind) >= FHRecordSize Then FindText = -1
    If Len(TextToFind) < 1 Then FindText = -1
    If StartPos = 0 Then StartPos = -1
    If FHeader.NumberOfRecords = 0 Then FindText = -1

    If oAutoGetUpdatedHeader Then GetUpdatedHeader
    
    BufferSize = IIf(FHRecordSize < 32767, 262144 \ FHRecordSize, 10)   ' 256 KB buffer (better with multi-core processor)
'    BufferSize = IIf(FHRecordSize < 6550, 65535 \ FHRecordSize, 10)     ' 64 KB buffer
    
    If StartPos = -1 Then
        LastSearchPos = TellPosition() - 1
        rPos = DataOffset + (LastSearchPos * FHRecordSize)
        RecLeft = (FHeader.NumberOfRecords - LastSearchPos)
        PosKetemu = LastSearchPos + 1
    ElseIf StartPos <= FHeader.NumberOfRecords And StartPos > 0 Then
        rPos = DataOffset + ((StartPos - 1) * FHRecordSize)
        RecLeft = (FHeader.NumberOfRecords - StartPos) + 1
        PosKetemu = StartPos
    Else
        FindText = -1
        Exit Function       ' wrong sStartpos parameter
    End If

    Seek #nNomorFile, rPos      ' move file pointer
    
    If RecLeft < BufferSize Then BufferSize = RecLeft       ' Resize cache for tiny table, or start search near end of file
    
    If BufferSize > 0 Then
        ReDim ReadByte(0 To (FHRecordSize * BufferSize) - 1)
        BigLoop = IIf(RecLeft Mod BufferSize = 0, RecLeft \ BufferSize, (RecLeft \ BufferSize) + 1)
    Else
        FindText = -1
        Exit Function
    End If
    
    ' convert TextToFind to byte array
    TTFInt = StrConv(TextToFind, vbFromUnicode)
    ReDim TTFByte(LBound(TTFInt) To UBound(TTFInt))
    UB = UBound(TTFByte)
    
    For i = LBound(TTFByte) To UB
        TTFByte(i) = TTFInt(i)
    Next i
    Erase TTFInt
    If IsNumeric(TextToFind) Then   ' Numeric expression is always case sensitive, to keep in best performance
        If Not MatchCase Then MatchCase = True
    End If
    
    ' convert to lowercase for case insensitive
    If Not MatchCase Then
        For i = LBound(TTFByte) To UB
            If TTFByte(i) >= 65 And TTFByte(i) <= 90 Then
                TTFByte(i) = TTFByte(i) + 32
            End If
        Next i
    End If
    
    ' /////////////////
    ' ////// In this FIndText, case insensitive is slightly slower than case sensitive
    ' /////////////////
    
    Ujung = (FHRecordSize - Len(TextToFind))
    rPos = BufferSize            ' small loop
    
    For jLoop = 0 To BigLoop - 1
        Get #nNomorFile, , ReadByte
        
        For j = 0 To rPos - 1
            Match = 1
            T = FHRecordSize * j
            
            If ReadByte(T) <> DELETED Then       ' record not deleted
                If RecLeft > 0 Then
                    
                    ' Search in bytes array one by one
                    If MatchCase Then   ' Case sensitive
                        For k = 1 To Ujung
                            If CInt(TTFByte(0) - ReadByte(T + k)) = 0 Then  ' Compare first character
                                If CInt(TTFByte(UB) - ReadByte(T + k + UB)) = 0 Then  ' Compare last character
                                    For TinyLoop = 1 To UB - 1
                                        If CInt(TTFByte(TinyLoop) - ReadByte(T + k + TinyLoop)) <> 0 Then   ' Compare characters between first and last
                                            Match = 0
                                            Exit For
                                        End If
                                    Next TinyLoop
                                    If Match = 1 Then
                                        PosKetemu = PosKetemu + (jLoop * BufferSize) + j
                                        Erase ReadByte
                                        Erase TTFByte
                                        FindText = PosKetemu
                                        Exit Function       ' Quit function
                                    End If
                                End If
                            End If
                        Next k
                    '/
                    'Else
                    ElseIf Not MatchCase Then    ' Ignore Case (Case insensitive)
                    '/
                        For k = 1 To Ujung
'==================================================================================
                            If CInt(TTFByte(0) - ReadByte(T + k)) = 0 Then
                                If CInt(TTFByte(UB) - ReadByte(T + k + UB)) = 0 Then
                                    For TinyLoop = 1 To UB - 1
                                        If CInt(TTFByte(TinyLoop) - ReadByte(T + k + TinyLoop)) <> 0 Then
                                            If CInt(TTFByte(TinyLoop) - ReadByte(T + k + TinyLoop)) <> 32 Then
                                                Match = 0
                                                Exit For
                                            End If
                                        End If
                                    Next TinyLoop
                                    If Match = 1 Then
                                        PosKetemu = PosKetemu + (jLoop * BufferSize) + j
                                        Erase ReadByte
                                        Erase TTFByte
                                        FindText = PosKetemu
                                        Exit Function       ' Quit function
                                    End If
                                Else 'TTFByte(UBound(TTFByte)) <> ReadByte(T + k + UBound(TTFByte)) Then
                                    If CInt(TTFByte(UB) - ReadByte(T + k + UB)) = 32 Then
                                        For TinyLoop = 1 To UB - 1
                                            If CInt(TTFByte(TinyLoop) - ReadByte(T + k + TinyLoop)) <> 0 Then
                                                If CInt(TTFByte(TinyLoop) - ReadByte(T + k + TinyLoop)) <> 32 Then
                                                    Match = 0
                                                    Exit For
                                                End If
                                            End If
                                        Next TinyLoop
                                        If Match = 1 Then
                                            PosKetemu = PosKetemu + (jLoop * BufferSize) + j
                                            Erase ReadByte
                                            Erase TTFByte
                                            FindText = PosKetemu
                                            Exit Function       ' Quit function
                                        End If
                                    End If
                                End If
                            
                            Else    ' If TTFByte(0) <> ReadByte(T + k)
                                
                                If CInt(TTFByte(0) - ReadByte(T + k)) = 32 Then
                                    If CInt(TTFByte(UB) - ReadByte(T + k + UB)) = 0 Then
                                        For TinyLoop = 1 To UB - 1
                                            If CInt(TTFByte(TinyLoop) - ReadByte(T + k + TinyLoop)) <> 0 Then
                                                If CInt(TTFByte(TinyLoop) - ReadByte(T + k + TinyLoop)) <> 32 Then
                                                    Match = 0
                                                    Exit For
                                                End If
                                            End If
                                        Next TinyLoop
                                        If Match = 1 Then
                                            PosKetemu = PosKetemu + (jLoop * BufferSize) + j
                                            Erase ReadByte
                                            Erase TTFByte
                                            FindText = PosKetemu
                                            Exit Function       ' Quit function
                                        End If
                                    Else
                                        If CInt(TTFByte(UB) - ReadByte(T + k + UB)) = 32 Then
                                            For TinyLoop = 1 To UB - 1
                                                If CInt(TTFByte(TinyLoop) - ReadByte(T + k + TinyLoop)) <> 0 Then
                                                    If CInt(TTFByte(TinyLoop) - ReadByte(T + k + TinyLoop)) <> 32 Then
                                                        Match = 0
                                                        Exit For
                                                    End If
                                                End If
                                            Next TinyLoop
                                            If Match = 1 Then
                                                PosKetemu = PosKetemu + (jLoop * BufferSize) + j
                                                Erase ReadByte
                                                Erase TTFByte
                                                FindText = PosKetemu
                                                Exit Function       ' Quit function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
'==================================================================================
                        Next k
                    End If
                End If  ' RecLeft > 0 Then
            End If  ' ReadByte(FHRecordSize * j) <> DELETED Then       ' record not deleted
            
            RecLeft = RecLeft - 1
        Next j
        
        If RecLeft < 1 Then        ' All records has evaluated, not found
            FindText = -1
            Exit Function
        ElseIf RecLeft < BufferSize Then    ' Rest of record is smaller than cache block: resize cache
            ReDim ReadByte(0 To (FHRecordSize * RecLeft) - 1)
            rPos = RecLeft
        End If
    Next jLoop

End Function


' -------------------------------------------------
' Functions:
'       SearchRecord
'
' Description:  Perform sequential search
'               Each data type has it's own search function
'               If requsted, for data type "C", this function will search substring in a field as well
' Parameters:
'       FieldName As String:    Field name to search
'       ValurToSearch As String:          Value to search, as string
'                               If the value is numeric with decimal, the decimal separator must use period (.)
'       sStartPos As Long:      From where search is begin.
'                               1  =    Start search from BOF
'                               -1 =    Start search from current position
'       lOperator As Long:
'                           lOperator                     Value (Long)
'                           -------                     ------------
'                           Equal                       5001
'                           Not Equal                   5002
'                           Less Than                   5003
'                           Less Than or Equal          5004
'                           Greater Than                5005
'                           Greater Than or Equal       5006
'                           Begin With (string)         5007        Available for only Character String
'
'
' Return value:
'       When success record position returned, and LastSearchPos is sets to record position
'       If failed, any number less than 0 is returned, and LastSearchPos sets to 0
'
' Note:
' -----
    ' For data type "C", String to search is case sensitive
' -------------------------------------------------

Public Function SearchRecord(FieldName As String, ValueToSearch As String, sStartPos As Long, lOperator As Integer) As Long
On Error GoTo E_SRCHRECORD
    Dim i As Long, j As Long
    Dim sX As String
    Dim PosKetemu As Long
    Dim rPos As Long
    Dim jLoop As Long
    Dim BigLoop As Long
    Dim BufferSize As Long
    Dim RecLeft As Long
    Dim ReadByte() As Byte
    Dim Hasil As Long
    Dim FldNum As Integer
    Dim Awal As Long
    
    Dim varDouble As Double
    Dim varString As String
    Dim varLong As Long
    Dim varCurrency As Currency
    Dim varByte As Byte
    Dim varByteArray() As Byte
    Dim varInteger As Integer
    
    ' == Checking if file is opened ==
    If Not FileUsed Then
        SearchRecord = -1
        Exit Function       ' File not open
    End If
    If oAutoGetUpdatedHeader Then GetUpdatedHeader
    
    ' Checking parameters
    FldNum = CInt(FieldNameMatch(FieldName))
    If FldNum < 0 Then
        SearchRecord = -2
        Exit Function           ' Field name not match
    End If
    
    sX = Trim$(ValueToSearch)
    If FIFieldType(FldNum) <> DT_C Then
        If Len(sX) < 1 Then
            SearchRecord = -3
            Exit Function       ' String Length = 0, data type not C
        End If
    End If
    
    Select Case FIFieldType(FldNum)
        Case DT_C, DT_N, DT_L, DT_F
            If Len(sX) > FIFieldWidth(FldNum) Then
                SearchRecord = -33
                Exit Function           ' String length > field width (data type: C,N,F,L)
            End If
    End Select
    ' ============== end parameter checking ============

    ' Checking operator "BeginWith"
    If lOperator = 5007 Then
        If FIFieldType(FldNum) <> DT_C Then
            SearchRecord = -34
            Exit Function           ' Can't apply Operator
        End If
    End If
    
    ' ============== convert string to proper type for comparison purpose ============
    Select Case FIFieldType(FldNum)
        Case DT_B
            If IsNumeric(ValueToSearch) Then
                varDouble = CDbl(Val(ValueToSearch))
            Else
                SearchRecord = -1
                Exit Function           ' string is not numeric string
            End If

        Case DT_C
            If lOperator <> 5007 Then
                varString = Space$(FIFieldWidth(FldNum))
                LSet varString = ValueToSearch
                varByteArray = StrConv(varString, vbFromUnicode)
            Else
                If Len(ValueToSearch) = 0 Then
                    SearchRecord = -1
                    Exit Function           ' string is not numeric string
                Else
                    varByteArray = StrConv(ValueToSearch, vbFromUnicode)
                End If
            End If
            
        Case DT_D
            If IsDate(ValueToSearch) Then
                varString = Format$(DateValue(ValueToSearch), DefDateFormat)
                varByteArray = StrConv(varString, vbFromUnicode)
            Else
                If ValueToSearch = "-" Then
                    varString = Space$(8)
                    varByteArray = StrConv(varString, vbFromUnicode)
                Else
                    SearchRecord = -1
                    Exit Function           ' string not a date, or not represent blank date
                End If
            End If
        
        Case DT_F, DT_N
            If IsNumeric(ValueToSearch) Then
                varString = Space$(FIFieldWidth(FldNum))
                
                If FIFieldDecimal(FldNum) = 0 Then
                    sX = CStr(Val(ValueToSearch))
                Else
                    sX = Format$(Val(ValueToSearch), "#." & String(FIFieldDecimal(FldNum), "0"))
                End If
                RSet varString = sX
                If FIFieldDecimal(FldNum) > 0 Then Mid$(varString, FIFieldWidth(FldNum) - FIFieldDecimal(FldNum)) = "."
                varByteArray = StrConv(varString, vbFromUnicode)
                
            Else
                SearchRecord = -1
                Exit Function       ' string is not numeric string
            End If

        Case DT_I
            If IsNumeric(ValueToSearch) Then
                ' Assume no overflow, so there is no overflow checking. Please be careful
                varLong = CLng(Val(ValueToSearch))
            Else
                SearchRecord = -1
                Exit Function       ' string is not numeric string
            End If
        
        Case DT_L
            Select Case ValueToSearch
                Case "T", "t", "Y", "y", "1"
                    varByte = 84
                Case Else
                    varByte = 70        ' other case is assume as False
            End Select
        
        Case DT_Y
            If IsNumeric(ValueToSearch) Then
                varCurrency = CCur(Val(ValueToSearch))
            Else
                SearchRecord = -1
                Exit Function       ' string is not numeric string
            End If
        
        Case DT_Z
            If IsNumeric(ValueToSearch) Then
                ' Assume no overflow, so there is no overflow checking. Please be careful
                varByte = CByte(Val(ValueToSearch))
            Else
                SearchRecord = -1
                Exit Function       ' string is not numeric string
            End If
        
        Case DT_S
            If IsNumeric(ValueToSearch) Then
                ' Assume no overflow, so there is no overflow checking. Please be careful
                varInteger = CInt(Val(ValueToSearch))
            Else
                SearchRecord = -1
                Exit Function       ' string is not numeric string
            End If
        
        Case Else
                SearchRecord = -1
                Exit Function       ' Data type not available for searcheing
    End Select
    
    ' Read buffer for several records. Less disk access result in faster search.
    ' You may choose any value, multiply by 4096.
'    BufferSize = IIf(FHRecordSize < 16384, 65536 \ FHRecordSize, 10)   ' 64 KB read buffer -> This is better for accessing file in network drive
    BufferSize = IIf(FHRecordSize < 32766, 262144 \ FHRecordSize, 10)   ' 256 KB read buffer -> This is better for local disk.
    
    If sStartPos = -1 Then
        LastSearchPos = TellPosition() - 1
        rPos = DataOffset + (LastSearchPos * FHRecordSize)
        RecLeft = (FHeader.NumberOfRecords - LastSearchPos)
        PosKetemu = LastSearchPos + 1
    ElseIf sStartPos <= FHeader.NumberOfRecords And sStartPos > 0 Then
        rPos = DataOffset + ((sStartPos - 1) * FHRecordSize)
        RecLeft = (FHeader.NumberOfRecords - sStartPos) + 1
        PosKetemu = sStartPos
    Else
        SearchRecord = -1
        Exit Function       ' wrong sStartpos parameter
    End If
    
    If RecLeft < BufferSize Then BufferSize = RecLeft       ' Resize cache for tiny table, or start search near end of file
    
    If BufferSize > 0 Then
        ReDim ReadByte(0 To (FHRecordSize * BufferSize) - 1)
        BigLoop = IIf(RecLeft Mod BufferSize = 0, RecLeft \ BufferSize, (RecLeft \ BufferSize) + 1)
    Else
        SearchRecord = -1
        Exit Function
    End If
    
    Seek #nNomorFile, rPos      ' move file pointer
    rPos = BufferSize            ' small loop
    
    ' ====================================== Search loop ======================================
    For jLoop = 0 To BigLoop - 1
        
        Get #nNomorFile, , ReadByte     ' read several records and cache to byte array
        
        For j = 0 To rPos - 1
            If ReadByte(FHRecordSize * j) <> DELETED Then       ' record not deleted
                Awal = (FHRecordSize * j) + FIFieldDisplacement(FldNum)
                
                If RecLeft > 0 Then
                    Select Case FIFieldType(FldNum)
                        
                        Case DT_B   ' Double
                            Hasil = SearchxB(FldNum, lOperator, varDouble, ReadByte, Awal)
                        
                        Case DT_C, DT_D, DT_F, DT_N ' Char, Date, Float, Numeric
                            Hasil = SearchxCDFN(FldNum, lOperator, varByteArray, ReadByte, Awal)
                        
                        Case DT_I   ' Signed Integer32
                            Hasil = SearchxI(FldNum, lOperator, varLong, ReadByte, Awal)
                        
                        Case DT_L   ' Logical
                            Hasil = SearchxL(FldNum, lOperator, varByte, ReadByte, Awal)
                        
                        Case DT_Y   ' Currency
                            Hasil = SearchxY(FldNum, lOperator, varCurrency, ReadByte, Awal)
                    
                        Case DT_Z   ' Byte
                            Hasil = SearchxZ(FldNum, lOperator, varByte, ReadByte, Awal)
                        
                        Case DT_S   ' Short
                            Hasil = SearchxS(FldNum, lOperator, varInteger, ReadByte, Awal)
                        
                    End Select
                End If
                
                If Hasil > 0 Then   ' found
                    PosKetemu = PosKetemu + (jLoop * BufferSize) + j
                    SearchRecord = PosKetemu
                    Exit Function       ' Quit function
                End If
            End If  ' Record not deleted
            
            RecLeft = RecLeft - 1         ' RecLeft--
        Next j
        
        If RecLeft < 1 Then        ' All records has evaluated, not found
            Erase ReadByte
            SearchRecord = -1
            Exit Function
        ElseIf RecLeft < BufferSize Then    ' Rest of record is smaller than cache block: resize cache
            ReDim ReadByte(0 To (FHRecordSize * RecLeft) - 1)
            rPos = RecLeft
        End If
    
    Next jLoop
    ' ====================================== Search loop end ======================================
    Exit Function

E_SRCHRECORD:
    Err.Clear
    Erase ReadByte
    SearchRecord = -1
End Function

Private Function SearchxB(FieldNumber As Integer, lOperator As Integer, ByRef ValToSearch As Double, ByRef ArrSource() As Byte, PosAwal As Long) As Long
    Dim Found As Boolean
    Dim ValDouble As Double
    
    Found = False
    CopyMemory ValDouble, ArrSource(PosAwal), 8
    Select Case lOperator
                Case 5001   '"="  "=="
                    If ValDouble - ValToSearch = 0 Then Found = True
                Case 5002   '"<>", "!=", "#"
                    If ValDouble - ValToSearch <> 0 Then Found = True
                Case 5003   '"<"
                    If ValDouble - ValToSearch < 0 Then Found = True
                Case 5004   '"<=", "=<"
                    If ValDouble - ValToSearch <= 0 Then Found = True
                Case 5005   '">"
                    If ValDouble - ValToSearch > 0 Then Found = True
                Case 5006   '">=", "=>"
                    If ValDouble - ValToSearch >= 0 Then Found = True
                'Case Else
                    ' do nothing
    End Select
    
    SearchxB = IIf(Found, 1, 0)

End Function

Private Function SearchxCDFN(FieldNumber As Integer, lOperator As Integer, ByRef ValToSearch() As Byte, ByRef ArrSource() As Byte, PosAwal As Long) As Long
    Dim i As Long
    Dim Length As Long
    Dim Found As Boolean
    Dim Sesuai As Byte
    
    Found = False
    Select Case lOperator
                Case 5001   '"="  "=="
                    Sesuai = 1
                    Length = FIFieldWidth(FieldNumber) - 1
                    For i = 0 To Length
                        If ArrSource(i + PosAwal) <> ValToSearch(i) Then
                            Sesuai = 0
                            Exit For
                        End If
                    Next i
                    
                    If Sesuai = 1 Then Found = True
                
                Case 5002   '"<>", "!=", "#"
                    Sesuai = 1
                    Length = FIFieldWidth(FieldNumber) - 1
                    For i = 0 To Length
                        If ArrSource(i + PosAwal) <> ValToSearch(i) Then
                            Sesuai = 0
                            Exit For
                        End If
                    Next i
                    If Sesuai = 0 Then Found = True

                Case 5003   '"<"
                    Length = FIFieldWidth(FieldNumber) - 1
                    If ArrSource(PosAwal) < ValToSearch(i) Then
                        Found = True
                    ElseIf ArrSource(PosAwal) = ValToSearch(i) Then
                        For i = 1 To Length
                            If ArrSource(i + PosAwal) > ValToSearch(i) Then
                                Exit For
                            ElseIf ArrSource(i + PosAwal) < ValToSearch(i) Then
                                Found = True
                                Exit For
                            End If
                        Next i
                    End If
                    
                
                Case 5004   '"<=", "=<"
                    Sesuai = 1
                    Length = FIFieldWidth(FieldNumber) - 1
                    If ArrSource(PosAwal) < ValToSearch(i) Then
                        Found = True
                    ElseIf ArrSource(PosAwal) = ValToSearch(i) Then
                        For i = 0 To Length
                            If ArrSource(i + PosAwal) < ValToSearch(i) Then
                                Sesuai = 1
                                Exit For
                            ElseIf ArrSource(i + PosAwal) > ValToSearch(i) Then
                                Sesuai = 0
                                Exit For
                            End If
                        Next i
                        If Sesuai > 0 Then Found = True
                    End If
                    
                Case 5005   '">"
                    Length = FIFieldWidth(FieldNumber) - 1
                    If ArrSource(PosAwal) > ValToSearch(i) Then
                        Found = True
                    ElseIf ArrSource(PosAwal) = ValToSearch(i) Then
                        For i = 1 To Length
                            If ArrSource(i + PosAwal) < ValToSearch(i) Then
                                Exit For
                            ElseIf ArrSource(i + PosAwal) > ValToSearch(i) Then
                                Found = True
                                Exit For
                            End If
                        Next i
                    End If
                
                Case 5006   '">=", "=>"
                    
                    Sesuai = 1
                    Length = FIFieldWidth(FieldNumber) - 1
                    If ArrSource(PosAwal) > ValToSearch(i) Then
                        Found = True
                    ElseIf ArrSource(PosAwal) = ValToSearch(i) Then
                        For i = 0 To Length
                            If ArrSource(i + PosAwal) > ValToSearch(i) Then
                                Sesuai = 1
                                Exit For
                            ElseIf ArrSource(i + PosAwal) < ValToSearch(i) Then
                                Sesuai = 0
                                Exit For
                            End If
                        Next i
                        If Sesuai > 0 Then Found = True
                    End If
                    
                    
                Case 5007   ' BEGINWITH         **** only for character string (Data type "C")
                    Sesuai = 1
                    Length = UBound(ValToSearch)
                    For i = 0 To Length
                        If ArrSource(i + PosAwal) <> ValToSearch(i) Then
                            Sesuai = 0
                            Exit For
                        End If
                    Next i
                    If Sesuai = 1 Then Found = True
                
                'Case Else
                    ' do nothing
    End Select
    
    SearchxCDFN = IIf(Found, 1, 0)

End Function

Private Function SearchxI(FieldNumber As Integer, lOperator As Integer, ByRef ValToSearch As Long, ByRef ArrSource() As Byte, PosAwal As Long) As Long
    Dim Found As Boolean
    Dim ValLong As Long
    
    Found = False
    GetMem4 VarPtr(ArrSource(PosAwal)), ValLong

    Select Case lOperator
                Case 5001   '"="  "=="
                    If ValLong - ValToSearch = 0 Then Found = True
                Case 5002   '"<>", "!=", "#"
                    If ValLong - ValToSearch <> 0 Then Found = True
                Case 5003   '"<"
                    If ValLong - ValToSearch < 0 Then Found = True
                Case 5004   '"<=", "=<"
                    If ValLong - ValToSearch <= 0 Then Found = True
                Case 5005   '">"
                    If ValLong - ValToSearch > 0 Then Found = True
                Case 5006   '">=", "=>"
                    If ValLong - ValToSearch >= 0 Then Found = True
                'Case Else
                    ' do nothing
    End Select

    SearchxI = IIf(Found, 1, 0)
End Function

Private Function SearchxL(FieldNumber As Integer, lOperator As Integer, ByRef ValToSearch As Byte, ByRef ArrSource() As Byte, PosAwal As Long) As Long
    Dim Found As Boolean
    Found = False
    Select Case lOperator
                Case 5001   '"="  "=="
                    If ValToSearch = ArrSource(PosAwal) Then Found = True
                Case 5002   '"<>", "!=", "#"
                    If ValToSearch <> ArrSource(PosAwal) Then Found = True
                'Case Else
                    ' do nothing
    End Select
    SearchxL = IIf(Found, 1, 0)
End Function

Private Function SearchxY(FieldNumber As Integer, lOperator As Integer, ByRef ValToSearch As Currency, ByRef ArrSource() As Byte, PosAwal As Long) As Long
    Dim Found As Boolean
    Dim ValCurrency As Currency
    
    Found = False
    CopyMemory ValCurrency, ArrSource(PosAwal), 8
    Select Case lOperator
                Case 5001   '"="  "=="
                    If ValCurrency - ValToSearch = 0 Then Found = True
                
                Case 5002   '"<>", "!=", "#"
                    If ValCurrency - ValToSearch <> 0 Then Found = True
                    
                Case 5003   '"<"
                    If ValCurrency - ValToSearch < 0 Then Found = True
                    
                Case 5004   '"<=", "=<"
                    If ValCurrency - ValToSearch <= 0 Then Found = True
                    
                Case 5005   '">"
                    If ValCurrency - ValToSearch > 0 Then Found = True
                    
                Case 5006   '">=", "=>"
                    If ValCurrency - ValToSearch >= 0 Then Found = True
                
                'Case Else
                    ' do nothing
    End Select

    SearchxY = IIf(Found, 1, 0)

End Function


Private Function SearchxZ(FieldNumber As Integer, lOperator As Integer, ByRef ValToSearch As Byte, ByRef ArrSource() As Byte, PosAwal As Long) As Long
    Dim Found As Boolean
    Found = False
    Select Case lOperator
                Case 5001   '"="  "=="
                    If ArrSource(PosAwal) = ValToSearch Then Found = True
                Case 5002   '"<>", "!=", "#"
                    If ArrSource(PosAwal) <> ValToSearch Then Found = True
                Case 5003   '"<"
                    If ArrSource(PosAwal) < ValToSearch Then Found = True
                Case 5004   '"<=", "=<"
                    If ArrSource(PosAwal) <= ValToSearch Then Found = True
                Case 5005   '">"
                    If ArrSource(PosAwal) > ValToSearch Then Found = True
                Case 5006   '">=", "=>"
                    If ArrSource(PosAwal) >= ValToSearch Then Found = True
                'Case Else
                    ' do nothing
    End Select

    SearchxZ = IIf(Found, 1, 0)
End Function

Private Function SearchxS(FieldNumber As Integer, lOperator As Integer, ByRef ValToSearch As Integer, ByRef ArrSource() As Byte, PosAwal As Long) As Long
    Dim Found As Boolean
    Dim ValInt As Integer
    
    Found = False
    GetMem2 VarPtr(ArrSource(PosAwal)), ValInt
    Select Case lOperator
                Case 5001   '"="  "=="
                    If ValInt - ValToSearch = 0 Then Found = True
                Case 5002   '"<>", "!=", "#"
                    If ValInt - ValToSearch <> 0 Then Found = True
                Case 5003   '"<"
                    If ValInt - ValToSearch < 0 Then Found = True
                Case 5004   '"<=", "=<"
                    If ValInt - ValToSearch <= 0 Then Found = True
                Case 5005   '">"
                    If ValInt - ValToSearch > 0 Then Found = True
                Case 5006   '">=", "=>"
                    If ValInt - ValToSearch >= 0 Then Found = True
                'Case Else
                    ' do nothing
    End Select

    SearchxS = IIf(Found, 1, 0)
End Function


' -------------------------------------------------
' Functions:
'       SearchAllRecord
'
' Description:  Perform sequential search
'               Each data type has it's own search function
'               If requsted, for data type "C", this function will search substring in a field as well
'               This function is similar to SearchRecord except that this function will allways
'               Search from first record, and continue search until reach EOF and position of all records that
'               match with criteria will put into array of long.
' Parameters:
'       FieldName As String:    Field name to search
'       ValurToSearch As String:          Value to search, as string
'                               If the value is numeric with decimal, the decimal separator must use period (.)
'       lBuffer() as long:      Buffer to hold array of record position
'       lOperator As Long:
'                           lOperator                     Value (Long)
'                           -------                     ------------
'                           Equal                       5001
'                           Not Equal                   5002
'                           Less Than                   5003
'                           Less Than or Equal          5004
'                           Greater Than                5005
'                           Greater Than or Equal       5006
'                           Begin With (string)         5007        Available for only Character String
'       [optional] StartPos as long:    Start position in where search will start. Defaulr is 1 (start search form first record)
'
' Return value:
'       When success it will return number of records match, and lBuffer() will filled by array of record positions,
'       If failed, any number less than 0 is returned, and lBuffer contains 1 element value sets to -1
'
' Note:
' -----
    ' For data type "C", String to search is case sensitive
' -------------------------------------------------
Public Function SearchAllRecord(FieldName As String, ValueToSearch As String, ByRef lBuffer() As Long, lOperator As Integer, Optional StartPos As Long = 1) As Long
On Error GoTo E_SRCHALLRECORD
    Dim i As Long, j As Long
    Dim sX As String
    Dim PosKetemu As Long
    Dim rPos As Long
    Dim jLoop As Long
    Dim BigLoop As Long
    Dim BufferSize As Long
    Dim RecLeft As Long
    Dim ReadByte() As Byte
    Dim Hasil As Long
    Dim FldNum As Integer
    Dim Awal As Long
    
    Dim varDouble As Double
    Dim varString As String
    Dim varLong As Long
    Dim varCurrency As Currency
    Dim varByte As Byte
    Dim varByteArray() As Byte
    Dim varInteger As Integer
    
    Dim JumlahKetemu As Long
    Dim UA As Long
    
    ReDim lBuffer(0 To 0)
    lBuffer(0) = -1
    ' == Checking if file is opened ==
    If Not FileUsed Then
        SearchAllRecord = -1
        Exit Function       ' File not open
    End If
    If oAutoGetUpdatedHeader Then GetUpdatedHeader
    
    ' Checking parameters
    FldNum = CInt(FieldNameMatch(FieldName))
    If FldNum < 0 Then
        SearchAllRecord = -2
        Exit Function           ' Field name not match
    End If
    
    sX = Trim$(ValueToSearch)
    If FIFieldType(FldNum) <> DT_C Then
        If Len(sX) < 1 Then
            SearchAllRecord = -3
            Exit Function       ' String Length = 0, data type not C
        End If
    End If
    
    Select Case FIFieldType(FldNum)
        Case DT_C, DT_N, DT_L, DT_F
            If Len(sX) > FIFieldWidth(FldNum) Then
                SearchAllRecord = -33
                Exit Function           ' String length > field width (data type: C,N,F,L)
            End If
    End Select
    ' ============== end parameter checking ============
    ' Checking operator "BeginWith"
    If lOperator = 5007 Then
        If FIFieldType(FldNum) <> DT_C Then
            SearchAllRecord = -34
            Exit Function           ' Can't apply Operator
        End If
    End If

    
    ' ============== convert string to proper type for comparison purpose ============
    Select Case FIFieldType(FldNum)
        Case DT_B
            If IsNumeric(ValueToSearch) Then
                varDouble = CDbl(Val(ValueToSearch))
            Else
                SearchAllRecord = -1
                Exit Function           ' string is not numeric string
            End If
        
        Case DT_C
            If lOperator <> 5007 Then
                varString = Space$(FIFieldWidth(FldNum))
                LSet varString = ValueToSearch
                varByteArray = StrConv(varString, vbFromUnicode)
            Else
                If Len(ValueToSearch) = 0 Then
                    SearchAllRecord = -1
                    Exit Function           ' string is not numeric string
                Else
                    varByteArray = StrConv(ValueToSearch, vbFromUnicode)
                End If
            End If
        
        Case DT_D
            If IsDate(ValueToSearch) Then
                varString = Format$(DateValue(ValueToSearch), DefDateFormat)
                varByteArray = StrConv(varString, vbFromUnicode)
            
            Else
                If ValueToSearch = "-" Then
                    varString = Space$(8)
                    varByteArray = StrConv(varString, vbFromUnicode)
                    
                Else
                    SearchAllRecord = -1
                    Exit Function           ' string not a date, or not represent blank date
                End If
            End If
        
        Case DT_F, DT_N
            If IsNumeric(ValueToSearch) Then
                varString = Space$(FIFieldWidth(FldNum))
                
                If FIFieldDecimal(FldNum) = 0 Then
                    sX = CStr(Val(ValueToSearch))
                Else
                    sX = Format$(Val(ValueToSearch), "#." & String(FIFieldDecimal(FldNum), "0"))
                End If
                RSet varString = sX
                If FIFieldDecimal(FldNum) > 0 Then Mid$(varString, FIFieldWidth(FldNum) - FIFieldDecimal(FldNum)) = "."
                varByteArray = StrConv(varString, vbFromUnicode)
                
            Else
                SearchAllRecord = -1
                Exit Function       ' string is not numeric string
            End If
        
        Case DT_I
            If IsNumeric(ValueToSearch) Then
                varLong = CLng(Val(ValueToSearch))
            Else
                SearchAllRecord = -1
                Exit Function       ' string is not numeric string
            End If
        
        Case DT_L
            Select Case ValueToSearch
                Case "T", "t", "Y", "y", "1"
                    varByte = 84
                Case Else
                    varByte = 70        ' other case is assume as False
            End Select
        
        Case DT_Y
            If IsNumeric(ValueToSearch) Then
                varCurrency = CCur(Val(ValueToSearch))
            Else
                SearchAllRecord = -1
                Exit Function       ' string is not numeric string
            End If
        
        Case DT_Z
            If IsNumeric(ValueToSearch) Then
                ' Assume no overflow, so there is no overflow checking. Please be careful
                varByte = CByte(Val(ValueToSearch))
            Else
                SearchAllRecord = -1
                Exit Function       ' string is not numeric string
            End If
        
        Case DT_S
            If IsNumeric(ValueToSearch) Then
                ' Assume no overflow, so there is no overflow checking. Please be careful
                varInteger = CInt(Val(ValueToSearch))
            Else
                SearchAllRecord = -1
                Exit Function       ' string is not numeric string
            End If
        
        Case Else
                SearchAllRecord = -1
                Exit Function       ' data type not available for search
    End Select

    ' Read buffer for several records. Less disk access result in faster search.
    BufferSize = IIf(FHRecordSize < 16384, 65536 \ FHRecordSize, 10) ' 64 KB  or (16 * 4096) read buffer -> This is best for almost all cases

    RecLeft = (FHeader.NumberOfRecords - StartPos) + 1
    PosKetemu = 1

    If RecLeft < BufferSize Then BufferSize = RecLeft       ' Resize cache for tiny table, or start search near end of file

    If BufferSize > 0 Then
        ReDim ReadByte(0 To (FHRecordSize * BufferSize) - 1)
        BigLoop = IIf(RecLeft Mod BufferSize = 0, RecLeft \ BufferSize, (RecLeft \ BufferSize) + 1)
    Else
        SearchAllRecord = -1
        Exit Function
    End If
    
'    Seek #nNomorFile, DataOffset      ' move file pointer.
    Seek #nNomorFile, DataOffset + ((StartPos - 1) * FHRecordSize)  ' Move file pointer
    
    ReDim lBuffer(0 To 1000)
    
    ' ====================================== Search loop ======================================
    rPos = BufferSize
    For jLoop = 0 To BigLoop - 1
        
        Get #nNomorFile, , ReadByte     ' read several records to byte array

        For j = 0 To rPos - 1       ' Evaluate each records in read buffer
            
            If ReadByte(FHRecordSize * j) <> DELETED Then       ' record not deleted
                Awal = (FHRecordSize * j) + FIFieldDisplacement(FldNum)
                
                If RecLeft > 0 Then
                    Select Case FIFieldType(FldNum)
                        
                        Case DT_B   ' Double
                            Hasil = SearchxB(FldNum, lOperator, varDouble, ReadByte, Awal)
                        
                        Case DT_C, DT_D, DT_F, DT_N ' Char, Date, Float, Numeric
                            Hasil = SearchxCDFN(FldNum, lOperator, varByteArray, ReadByte, Awal)
                        
                        Case DT_I   ' Signed Integer32
                            Hasil = SearchxI(FldNum, lOperator, varLong, ReadByte, Awal)
                        
                        Case DT_L   ' Logical
                            Hasil = SearchxL(FldNum, lOperator, varByte, ReadByte, Awal)
                        
                        Case DT_Y   ' Currency
                            Hasil = SearchxY(FldNum, lOperator, varCurrency, ReadByte, Awal)
                    
                        Case DT_Z   ' Byte
                            Hasil = SearchxZ(FldNum, lOperator, varByte, ReadByte, Awal)
                        
                        Case DT_S   ' Short
                            Hasil = SearchxS(FldNum, lOperator, varInteger, ReadByte, Awal)
                        
                    End Select
                End If  ' RecLeft > 0
                
                If Hasil > 0 Then   ' found
                    'JumlahKetemu = JumlahKetemu + 1
                    'ReDim Preserve lBuffer(0 To JumlahKetemu - 1)
                    'lBuffer(JumlahKetemu - 1) = PosKetemu + (jLoop * BufferSize) + j
                    
                    ' The above 3 lines are slow, because calling 'Redim Preserve' too often
                    ' Replaced by these line below, that gives same result, but faster:
                    If UBound(lBuffer) = JumlahKetemu - 1 Then
                        ReDim Preserve lBuffer(0 To JumlahKetemu + 1000) ' saves extra 1000 loop
                    End If
                    lBuffer(JumlahKetemu) = PosKetemu + (jLoop * BufferSize) + j
                    JumlahKetemu = JumlahKetemu + 1
                End If  ' Hasil > 0
            End If  ' Record not deleted
            
            RecLeft = RecLeft - 1
        Next j
        
        If RecLeft < 1 Then        ' All records has evaluated
            If JumlahKetemu = 0 Then
                ReDim lBuffer(0 To 0)
                lBuffer(0) = -1
                Erase ReadByte
                SearchAllRecord = -1
                Exit Function
            Else
                Erase ReadByte
                ReDim Preserve lBuffer(0 To JumlahKetemu - 1)   ' Resize if array too large
                SearchAllRecord = JumlahKetemu
                Exit Function
            End If
        ElseIf RecLeft < BufferSize Then    ' Rest of record is smaller than cache block: resize cache
            ReDim ReadByte(0 To (FHRecordSize * RecLeft) - 1)
            rPos = RecLeft
        End If
    
    Next jLoop
    ' ====================================== Search loop end ======================================
    
    Exit Function
    
E_SRCHALLRECORD:
    Err.Clear
    Erase ReadByte
    SearchAllRecord = JumlahKetemu
End Function

' -------------------------------------------------
' Function: GetFieldInfo
' Description:  Retrieve fields information for current opened table
'               The information will stored to 2D array of STRING
' Parameters:
'       FIBuffer() As String:   Array to fill with information
'
' Return value:
'       None
'       When success FIBuffer() filled with these informations in 2D array:
'           - Field name
'           - Field type
'           - Field width
'           - Decimal width
'       If failed, FIbuffer() contains nothing
'
' FoxPro command: DISPLAY STRUCTURE
' -------------------------------------------------
Public Sub GetFieldInfo(FIBuffer() As String)
    Dim i As Long
    Dim s As String

    ReDim FIBuffer(0 To nNumberOfFields - 1, 0 To 3)
    
    If FileUsed Then
        For i = 0 To nNumberOfFields - 1
            ' Replace trailing Chr(0)s with spaces
            s = Replace(FIFieldName(i), Chr(0), " ")
            FIBuffer(i, 0) = Trim(s)
            FIBuffer(i, 1) = Chr(FIFieldType(i))
            FIBuffer(i, 2) = CStr(FIFieldWidth(i))
            FIBuffer(i, 3) = CStr(FIFieldDecimal(i))
        Next i
    End If

End Sub


' -------------------------------------------------
' Function: GetFileInfo
' Description:  Retrieve file information for current opened table
'               The information will stored to array of LONG
' Parameters:
'       FBuffer() As Long:   Variable to fill with information
'
' Return value:
'       None
'       When success FBuffer filled with these informations:
'           - Number of records
'           - Number of fields
'           - Record size (include delete flag)
'       If failed Fbuffer contains nothing
'
' -------------------------------------------------
Public Sub GetFileInfo(ByRef FBuffer() As Long)
    
    ReDim FBuffer(0 To 2)
    
    If FileUsed Then
        FBuffer(0) = FHeader.NumberOfRecords
        FBuffer(1) = nNumberOfFields
        FBuffer(2) = FHeader.RecordSize
    End If
End Sub

' -------------------------------------------------
Private Sub Class_Initialize()
    FileUsed = False
    nNumberOfFields = 0
    FIFieldBlank = String(13, 0)
    zHeaderBlank = String(20, 0)
    dString = Space(8)
    DefDateFormat = "YYYYMMDD"
    LastSearchPos = 0
    
    RecNumChange = False
    RecEdited = False
    s = String(MAXFIELDNAMEWIDTH, 0)
    
    oAutoUpdateHeader = True
    oAutoGetUpdatedHeader = True

End Sub

' -------------------------------------------------
Private Sub Class_Terminate()
    DBFClose
End Sub

' ------------------------------------------------- Some stories when testing this class
'
'   Everything here is done by sequential disk access.
'   I have run test and compare with FoxPro in my Office.
'   I am sattisfied with this class' performance against FoxPro
'
'   But when I test and compare with other VB project using MySQL, this class is 2 to 3 time slower than MySQL
'   Everyone know that VB must be slower. That is right.
'   Comparing this class with MySQL doesn't sattisfy me.
'   The string operation in VB is slightly slower. Very slow.
'   MySQL (With MyISAM storage engine) browse 18,000 records in FlexGrid in less than 1 second (0.8 seconds when compiled to Exe, and 0.9 seconds run from VB IDE). Command invoked for MySQL in the test is "Select * from Accounting.Transdtl01" (without any condition)
'   and this class require almost 2 seconds to browse. (1.8 to 2.0 seconds, compiled in Exe). ( I use equal data structure but in DBF format, and equal record size, that is 360 bytes)
'
'   One important thing that sattisfy me about this class is:
'   When doing a project, I know exactly the string returned from this class.
'   Since it not bound with control I can manipulate it, change its font in Grid,
'   I can decide in which column or row a string placed in Grid,
'   or make this string a caption for command button or caption for form, doing mathematical operation with numbers, etc.
'
'   More work to do... mostly to improve search speed and speed of functin "GetRow2DA"
'   or...
'   Create a new record manager that not compatible with DBF file, and that new record manager must faster than this class... (I have some idea for that, but net yet start building a new one)
'
' -------------------------------------------------













' -------------------------------------------------
'
' Function: SearchWithFilter
'
' Description:  By using Search With Filter, this function will not search entire table.
'               This function search only in record positions described in parameter lFilter().
'
' Parameters:
'       FieldName As String:    Field name to search
'       ValueToSearch As String:          Value to search, as string
'                               If the value is numeric with decimal, the decimal separator must use period (.)
'       lOperator As Long:
'                           lOperator                     Value (Long)
'                           -------                     ------------
'                           Equal                       5001
'                           Not Equal                   5002
'                           Less Than                   5003
'                           Less Than or Equal          5004
'                           Greater Than                5005
'                           Greater Than or Equal       5006
'                           Begin With (string)         5007        Available for only Character String
'
'       lFilter() As Long:  Array of record positions where search is performed.
'                           EXAMPLE: there are 1000 records, but you want to search only in records 1 to 100, and records 201 to 250
'                           then fill the array lFilter with number from 1 to 100, and 201 to 250, so there are 150 elements in array lFilter ( 1, 2, 3, ..., 100, 201, 202, ..., 250 )
'
'                           With filter, search not evaluate the whole table. Only on specified record positions.
'
' Return Value: Array of long of records position match with criteria.
'               If failed, the array contain 1 element and the value is -1
'
' -------------------------------------------------
Public Function SearchWithFilter(FieldName As String, ValueToSearch As String, lOperator As Integer, lFilter() As Long) As Long()
On Error GoTo E_SRCHFILTER
    Dim i As Long, j As Long
    Dim sX As String
    Dim JumlahKetemu As Long
    Dim ReadByte() As Byte
    Dim Hasil As Long
    Dim FldNum As Integer
    Dim Awal As Long
    
    Dim varDouble As Double
    Dim varString As String
    Dim varLong As Long
    Dim varCurrency As Currency
    Dim varByte As Byte
    Dim varByteArray() As Byte
    Dim varInteger As Integer
    
    Dim lResult() As Long       ' Variable to send after search is finished
    
    ' == Checking if file is opened ==
    ReDim lResult(0 To 0)
    lResult(0) = -1
    If Not FileUsed Then
        SearchWithFilter = lResult
        Exit Function       ' File not open
    End If
    If oAutoGetUpdatedHeader Then GetUpdatedHeader
    
    ' Checking parameters
    FldNum = CInt(FieldNameMatch(FieldName))
    If FldNum < 0 Then
        SearchWithFilter = lResult
        Exit Function           ' Field name not match
    End If
    
    sX = Trim(ValueToSearch)
    If FIFieldType(FldNum) <> DT_C Then
        If Len(sX) < 1 Then
            SearchWithFilter = lResult
            Exit Function       ' String Length = 0, data type not C
        End If
    End If
    
    Select Case FIFieldType(FldNum)
        Case DT_C, DT_N, DT_L, DT_F
            If Len(sX) > FIFieldWidth(FldNum) Then
                SearchWithFilter = lResult
                Exit Function           ' String length > field width (data type: C,N,F,L)
            End If
    End Select
    
    ' Scan for bad record number..
    For i = LBound(lFilter) To UBound(lFilter)
        If lFilter(i) < 0 Or lFilter(i) > FHeader.NumberOfRecords Then
            SearchWithFilter = lResult
            Exit Function
        End If
    Next i
    ' ============== end parameter checking ============

    ' Checking operator "BeginWith"
    If lOperator = 5007 Then
        If FIFieldType(FldNum) <> DT_C Then
            SearchWithFilter = lResult
            Exit Function           ' Can't apply Operator
        End If
    End If
    
    ' ============== convert string to proper type for comparison purpose ============
    Select Case FIFieldType(FldNum)
        Case DT_B
            If IsNumeric(ValueToSearch) Then
                varDouble = CDbl(Val(ValueToSearch))
            Else
                SearchWithFilter = lResult
                Exit Function           ' string is not numeric string
            End If

        Case DT_C
            If lOperator <> 5007 Then
                varString = Space(FIFieldWidth(FldNum))
                LSet varString = ValueToSearch
                varByteArray = StrConv(varString, vbFromUnicode)
            Else
                If Len(ValueToSearch) = 0 Then
                    SearchWithFilter = lResult
                    Exit Function           ' string is not numeric string
                Else
                    varByteArray = StrConv(ValueToSearch, vbFromUnicode)
                End If
            End If
            
        Case DT_D
            If IsDate(ValueToSearch) Then
                varString = Format(DateValue(ValueToSearch), DefDateFormat)
                varByteArray = StrConv(varString, vbFromUnicode)
            Else
                If ValueToSearch = "-" Then
                    varString = Space(8)
                    varByteArray = StrConv(varString, vbFromUnicode)
                Else
                    SearchWithFilter = lResult
                    Exit Function           ' string not a date, or not represent blank date
                End If
            End If
        
        Case DT_F, DT_N
            If IsNumeric(ValueToSearch) Then
                varString = Space(FIFieldWidth(FldNum))
                
                If FIFieldDecimal(FldNum) = 0 Then
                    sX = CStr(Val(ValueToSearch))
                Else
                    sX = Format(Val(ValueToSearch), "#." & String(FIFieldDecimal(FldNum), "0"))
                End If
                RSet varString = sX
                If FIFieldDecimal(FldNum) > 0 Then Mid(varString, FIFieldWidth(FldNum) - FIFieldDecimal(FldNum)) = "."
                varByteArray = StrConv(varString, vbFromUnicode)
                
            Else
                SearchWithFilter = lResult
                Exit Function       ' string is not numeric string
            End If

        Case DT_I
            If IsNumeric(ValueToSearch) Then
                ' Assume no overflow, so there is no overflow checking. Please be careful
                varLong = CLng(Val(ValueToSearch))
            Else
                SearchWithFilter = lResult
                Exit Function       ' string is not numeric string
            End If
        
        Case DT_L
            Select Case ValueToSearch
                Case "T", "t", "Y", "y", "1"
                    varByte = 84
                Case Else
                    varByte = 70        ' other case is assume as False
            End Select
        
        Case DT_Y
            If IsNumeric(ValueToSearch) Then
                varCurrency = CCur(Val(ValueToSearch))
            Else
                SearchWithFilter = lResult
                Exit Function       ' string is not numeric string
            End If
        
        Case DT_Z
            If IsNumeric(ValueToSearch) Then
                ' Assume no overflow, so there is no overflow checking. Please be careful
                varByte = CByte(Val(ValueToSearch))
            Else
                SearchWithFilter = lResult
                Exit Function       ' string is not numeric string
            End If
        
        Case DT_S
            If IsNumeric(ValueToSearch) Then
                ' Assume no overflow, so there is no overflow checking. Please be careful
                varInteger = CInt(Val(ValueToSearch))
            Else
                SearchWithFilter = lResult
                Exit Function       ' string is not numeric string
            End If
        
        Case Else
                SearchWithFilter = lResult
                Exit Function       ' Data type not available for searcheing
    End Select
    
    ReDim ReadByte(0 To FHRecordSize - 1)
    j = UBound(lFilter)
    JumlahKetemu = 0
    ' ====================================== Search loop ======================================
    For i = 0 To j
        Get #nNomorFile, DataOffset + ((lFilter(i) - 1) * FHRecordSize), ReadByte    ' read a record
    
        If ReadByte(0) <> DELETED Then       ' record not deleted
            Awal = FIFieldDisplacement(FldNum)
            
            Select Case FIFieldType(FldNum)
                Case DT_B   ' Double
                    Hasil = SearchxB(FldNum, lOperator, varDouble, ReadByte, Awal)
                Case DT_C, DT_D, DT_F, DT_N ' Char, Date, Float, Numeric
                    Hasil = SearchxCDFN(FldNum, lOperator, varByteArray, ReadByte, Awal)
                Case DT_I   ' Signed Integer32
                    Hasil = SearchxI(FldNum, lOperator, varLong, ReadByte, Awal)
                Case DT_L   ' Logical
                    Hasil = SearchxL(FldNum, lOperator, varByte, ReadByte, Awal)
                Case DT_Y   ' Currency
                    Hasil = SearchxY(FldNum, lOperator, varCurrency, ReadByte, Awal)
                Case DT_Z   ' Byte
                    Hasil = SearchxZ(FldNum, lOperator, varByte, ReadByte, Awal)
                Case DT_S   ' Short
                    Hasil = SearchxS(FldNum, lOperator, varInteger, ReadByte, Awal)
            End Select
            
            If Hasil > 0 Then   ' found
                JumlahKetemu = JumlahKetemu + 1
                ReDim Preserve lResult(0 To JumlahKetemu - 1)
                lResult(JumlahKetemu - 1) = lFilter(i)
            End If
        End If  ' Record not deleted
            
    Next i
    
    ' ====================================== Search loop end ======================================
    Erase ReadByte
    SearchWithFilter = lResult
    Exit Function

E_SRCHFILTER:
    Err.Clear
    Erase ReadByte
    SearchWithFilter = lResult
End Function


' -------------------------------------------------
' Function: GetDistinctValues
'
' Description:  Get distinct values in a field
'
' Parameter:
'       FieldName As String:    Field to evaluate
'       [Optional] CaseSens As Boolean:     Set case sensitive or not.
'                                           Default is case sensitive
'
' Return Value:
'       Array of string contains distinct values.
'       If fail, array contain 1 element with zero length string
'
' Note:
' -----
'   I don't know exactly the purpose of distinct value, but someone ask me
'   to create function to retrieve such thing.
'   ...Ehm... , now I know the purpose of this function :)
'
' Known Bug:
' ----------
'   Searching on field that already sorted:
'   Will doing so many comparisons and, of course, take very long time
'   It is 2 to 200 time slower on character data type :(
'
'   Reason: No tree rebalancing routine. The search fall to sequential search
' -------------------------------------------------
Public Function GetDistinctValues(FieldName As String, ByRef strResult() As String, Optional CaseSens As Boolean = True) As Double
    Dim i As Long, j As Long
    Dim k As Long, f As Long
    Dim m As Long, n As Long
'    Dim FldNum As Integer
    Dim Length As Long      ' length of string or field
    Dim Tempat As Long
    Dim RecString As String ' Read-write buffer as string
    Dim lpRecString As Long
    Dim BuffSize As Long
    Dim RecLeft As Long
    Dim BigLoop As Long
    Dim Olah As Long
    Dim CompareMethod As VbCompareMethod
    Dim CompResultA As Integer
    Dim CompResultN As Integer
    Dim pTambah As Long
    Dim DistinctFound As Long
    
    Dim RecRead() As String ' Array of strings
    Dim lpRecRead() As Long
    Dim alAkar(0 To 255) As Long
    Dim vAkar() As Byte

'    Dim strResult() As String
    Dim sRKi() As Long
    Dim sRKa() As Long
    
    Dim CompareCount As Double

    ' -----------------------------------------
    ReDim strResult(0 To 0)
    strResult(0) = ""
    CompareCount = 0
    GetDistinctValues = CompareCount
    
    If Not FileUsed Then Exit Function
    
    f = FieldNameMatch(FieldName)
    If f < 0 Then Exit Function
    
'    If FIFieldFlag(f) >= 4 Then     ' This function is not available for binary data
'        GetDistinctValues = strResult
'        Exit Function
'    End If

    Select Case FIFieldType(f)
        Case DT_D
            pTambah = 14
        Case DT_N, DT_F
            pTambah = IIf(FIFieldWidth(f) - FIFieldDecimal(f) < FIFieldDecimal(f), (FIFieldWidth(f) - FIFieldDecimal(f)) * 2, ((FIFieldWidth(f) - FIFieldDecimal(f)) \ 2) * 2)
        Case DT_C
            pTambah = (FIFieldWidth(f) \ 4) * 2
        Case DT_Y
            pTambah = 2
        Case DT_B
            pTambah = 14
        Case Else
            pTambah = 0
    End Select
    
    ' Working with strings, I cannot sets buffer too large
    BuffSize = IIf(FHRecordSize < 16384, 32766 \ FHRecordSize, 4)
    RecLeft = FHeader.NumberOfRecords
    If RecLeft < BuffSize Then BuffSize = RecLeft
    
    If BuffSize > 0 Then
        RecString = Space$(BuffSize * FHRecordSize)
        ReDim RecRead(0 To BuffSize - 1)
        ReDim lpRecRead(0 To BuffSize - 1)
        ReDim vAkar(0 To BuffSize - 1)
        lpRecString = StrPtr(RecString)
        BigLoop = IIf(RecLeft Mod BuffSize = 0, FHeader.NumberOfRecords \ BuffSize, (FHeader.NumberOfRecords \ BuffSize) + 1)
    Else
        Exit Function
    End If
    
    For i = 0 To BuffSize - 1
        RecRead(i) = Space$(FIFieldWidth(f))
        lpRecRead(i) = StrPtr(RecRead(i))
    Next i
    
    For i = 0 To 255
        alAkar(i) = -1
    Next i
    
    Length = FIFieldWidth(f) * 2
    Tempat = FIFieldDisplacement(f) * 2
    Olah = FHRecordSize * 2
    
    ' Allocate a lot of ram
    ReDim sRKi(0 To FHeader.NumberOfRecords - 1)    ' Index to the left
    ReDim sRKa(0 To FHeader.NumberOfRecords - 1)    ' Index to the right
    ReDim strResult(0 To FHeader.NumberOfRecords - 1)   ' String array
    
    CompareMethod = IIf(CaseSens, vbBinaryCompare, vbTextCompare)
    
    ' Now, read the whole table sequentialy and build distinct values
    Seek #nNomorFile, DataOffset
    DistinctFound = 0
    '================================================================================
    For i = 0 To BigLoop - 1
        Get #nNomorFile, , RecString
        
        '--------------------------------------------------------------------------------
        For j = 0 To BuffSize - 1
            ' copy to string to string array.
            CopyMemoryBV lpRecRead(j), lpRecString + (j * Olah) + Tempat, Length
            GetMem1 lpRecRead(j) + pTambah, vAkar(j)
        Next j
        '--------------------------------------------------------------------------------
        
        '================================================================================
        ' Main loop for Search distinct
        For j = 0 To BuffSize - 1
            m = alAkar(vAkar(j))
            
            If m >= 0 Then
                ' compare string
                CompResultA = StrComp(RecRead(j), strResult(m), CompareMethod)
                CompareCount = CompareCount + 1     ' Count comparisons
                
                ' Ignoring equal string  (CompResult = 0)
                If CompResultA < 0 Then
                    If sRKi(m) <> 0 Then
                        n = sRKi(m)
                        
                        For k = 0 To 2000000000
                            CompResultN = StrComp(RecRead(j), strResult(n), CompareMethod)
                            CompareCount = CompareCount + 1     ' Count comparisons
                            
                            If CompResultN < 0 Then
                                If sRKi(n) <> 0 Then
                                    n = sRKi(n)
                                Else
                                    DistinctFound = DistinctFound + 1
                                    sRKi(n) = DistinctFound - 1
                                    strResult(DistinctFound - 1) = RecRead(j)
                                    Exit For
                                End If
                                
                            ElseIf CompResultN > 0 Then
                                If sRKa(n) <> 0 Then
                                    n = sRKa(n)
                                Else
                                    DistinctFound = DistinctFound + 1
                                    sRKa(n) = DistinctFound - 1
                                    strResult(DistinctFound - 1) = RecRead(j)
                                    Exit For
                                End If
                                
                            Else
                                Exit For
                                
                            End If
                        Next k
                    Else
                        DistinctFound = DistinctFound + 1
                        sRKi(m) = DistinctFound - 1
                        strResult(DistinctFound - 1) = RecRead(j)
                    End If
                
                ElseIf CompResultA > 0 Then
                    If sRKa(m) <> 0 Then
                        n = sRKa(m)
                        
                        For k = 0 To 2000000000
                            CompResultN = StrComp(RecRead(j), strResult(n), CompareMethod)
                            CompareCount = CompareCount + 1     ' Count comparisons
                            
                            If CompResultN < 0 Then
                                If sRKi(n) <> 0 Then
                                    n = sRKi(n)
                                Else
                                    DistinctFound = DistinctFound + 1
                                    sRKi(n) = DistinctFound - 1
                                    strResult(DistinctFound - 1) = RecRead(j)
                                    Exit For
                                End If
                            
                            ElseIf CompResultN > 0 Then
                                If sRKa(n) <> 0 Then
                                    n = sRKa(n)
                                Else
                                    DistinctFound = DistinctFound + 1
                                    sRKa(n) = DistinctFound - 1
                                    strResult(DistinctFound - 1) = RecRead(j)
                                    Exit For
                                End If
                            
                            Else
                                Exit For
                            
                            End If
                            
                        Next k
                    Else
                        DistinctFound = DistinctFound + 1
                        sRKa(m) = DistinctFound - 1
                        strResult(DistinctFound - 1) = RecRead(j)
                    End If
                End If
            
            Else  ' m < 0
                DistinctFound = DistinctFound + 1
                alAkar(vAkar(j)) = DistinctFound - 1
                strResult(DistinctFound - 1) = RecRead(j)
                sRKi(DistinctFound - 1) = 0
                sRKa(DistinctFound - 1) = 0
            End If
        
            RecLeft = RecLeft - 1
        Next j
        '================================================================================
        
        If RecLeft < 1 Then
            Exit For
        ElseIf RecLeft < BuffSize Then
            BuffSize = RecLeft
            RecString = Space$(BuffSize * FHRecordSize)
            lpRecString = StrPtr(RecString)
            ReDim vAkar(0 To BuffSize - 1)
            ReDim RecRead(0 To BuffSize - 1)
            ReDim lpRecRead(0 To BuffSize - 1)
            For j = 0 To BuffSize - 1
                RecRead(j) = Space$(FIFieldWidth(f))
                lpRecRead(j) = StrPtr(RecRead(j))
            Next j
        End If
    Next i
    '================================================================================
    
    ' Cleanup
    Erase RecRead
    Erase lpRecRead
    Erase sRKi
    Erase sRKa
    RecString = vbNullString
    
    DoEvents
    
    ' Resize array
    If DistinctFound > 0 Then
        ReDim Preserve strResult(0 To DistinctFound - 1)
    Else
        ReDim strResult(0 To 0)
        strResult(0) = ""
    End If
    
    ' Return result
    GetDistinctValues = CompareCount

End Function



' -------------------------------------------------
' Function: SortData
'
' Description: Sorts records in ascending, based on a key field.
'              This sort is stable, except in condition explained in the note.
'              Sort method used: MSD radix/counting sort
'
' Parameter:
'           sTargetFile As String:  Target file name of sorted table
'           sField As String:       Field to sort
'           Optional ReOrderNegative As Boolean:    Used for field with data type binary numeric
'                                                   True: Negative value appear before positive value, and
'                                                   False: Negative value appear after positive vaule
'                                                   Default value is True
'
' Note:
' =====
'   When sorting binary floating numer (data type 'Double' or 'B") and ReOrderNegative sets to 'True',
'   the negatives double values does not maintain relative order (not stable),
'   but the positive values remain stable.
'
'   For binary integer (Byte, Sort, Long, Currency), both negative and positive values are stable.
'
' -------------------------------------------------
Public Sub SortData(sTargetFile As String, sField As String, Optional ReOrderNegative As Boolean = True)  ' , Optional SortOrder As MySortOrder = Ascending)
    Dim FldNum As Long
    
    If Not FileUsed Then Exit Sub
    
    FldNum = CLng(FieldNameMatch(sField))
    If FldNum < 0 Then Exit Sub
    
    Select Case FIFieldType(FldNum)
        Case DT_C, DT_N, DT_L, DT_Z, DT_F, DT_D     ' String like fields
            SortDataUrut sTargetFile, FldNum
        Case DT_I, DT_Y, DT_S                    ' Binary numeric fields
            SortDataLong sTargetFile, FldNum, ReOrderNegative, False
        Case DT_B                       ' Binary floating numeric fields
            SortDataLong sTargetFile, FldNum, ReOrderNegative, True
    End Select

End Sub

' -------------------------------------------------
' Sort for binary numeric
' -------------------------------------------------
Private Sub SortDataLong(sTargetFile As String, FldNum As Long, FlipBit As Boolean, isDouble As Boolean)    ' For sorting Short, Long and Currency
    Dim InfoMemory As MEMORYSTATUS
    Dim srtNomorFile As Integer
    Dim FLebar As Long  ' Field width

    Dim RWBuffer() As Byte      ' Buffer for one record
    Dim IsiData() As Byte    ' Buffer file as much as memory
    
    Dim Tembolok() As Byte      ' Cache for large file
    Dim PTembolok As Long       ' Pointer to cache
    Dim IsiTembolok As Long
    
    Dim lpA() As Long
    Dim lpB() As Long
    
    Dim ByteCount As bCounter
    Dim EmptyCounter As bCounter
    
    Dim Histo() As Long
    
    Dim aStack() As BKeyInfo
    Dim bMin As Byte
    Dim bMax As Byte
    Dim bVal As Byte
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim n As Long
    
    Dim PosisiStack As Long
    
    Dim JumRec As Long
    Dim XX As Long
    Dim wWw As Long

    ' --------------------------------------------
    ' Menyiapkan memory untuk menampung byte array data-data
    ' Asumsi: Memorynya bisa menampung semua key
    JumRec = FHeader.NumberOfRecords
    FLebar = FIFieldWidth(FldNum)
    
    ReDim aStack(0 To (FLebar * 255))
    ReDim Histo(0 To 255)
    
    ' --------------------------------------------
    ' Membaca records
    
    Seek #nNomorFile, DataOffset
    
    ' Baca dan copy alamat........
    ReDim IsiData(0 To (FLebar * FHeader.NumberOfRecords) - 1)  ' Ini untuk menampung data-data dalam bentuk byte array
    ReDim lpA(0 To JumRec - 1)
    
    lpA(0) = VarPtr(IsiData(0))
    
    ' Jumlah buffer maksimal ditentukan maksimal sebesar "wWw" records
    wWw = 65535 \ FHRecordSize
    If JumRec >= wWw Then
        ReDim RWBuffer(0 To (wWw * FHRecordSize) - 1)
        ReDim lpB(0 To wWw - 1)
        
        ' Jumlah pengulangan loop
        n = IIf(JumRec Mod wWw = 0, JumRec \ wWw, (JumRec \ wWw) + 1)
    Else
        ReDim RWBuffer(0 To (JumRec * FHRecordSize) - 1)
        ReDim lpB(0 To JumRec - 1)
        
        ' Jumlah pengulangan loop
        n = 1
    End If
    
    ' Mengisi alamat memori array untuk source
    lpB(0) = VarPtr(RWBuffer(0)) + FIFieldDisplacement(FldNum)
    k = UBound(lpB)
    For i = 1 To k
        lpB(i) = lpB(i - 1) + FHRecordSize
    Next i
    
    m = VarPtr(IsiData(0))
    
    XX = 0      ' Mewakili jumlah record yang dibaca
    
    ' Mengulang baca tulis
    For j = 0 To n - 1
        ' Baca dari file sumber...
        Get #nNomorFile, , RWBuffer
        For i = 0 To k
            lpA(XX) = XX * FLebar
            
            ' Copy memory to byte array
            CopyMemoryBV lpA(XX) + m, lpB(i), FLebar
            
            XX = XX + 1     ' Counter
        Next i
        
        ' Hitung yang sudah ditulis dan tampilkan
        ' Feedback for user, because this proccess require a lot of time
        ' Time needed mostly for disk read and write.
        ' This event helps user know how long they must wait
        '----------------------
        If XX Mod (wWw * 100) = 0 Then
            RaiseEvent RecRead(XX)
        '----------------------
            DoEvents
        End If
        
        ' Mengecek sisa
        If XX < JumRec Then
            If JumRec - XX < wWw Then
                ReDim RWBuffer(0 To ((JumRec - XX) * FHRecordSize) - 1)
                ReDim lpB(0 To (JumRec - XX) - 1)
                lpB(0) = VarPtr(RWBuffer(0)) + FIFieldDisplacement(FldNum)
                k = UBound(lpB)
                For i = 1 To k
                    lpB(i) = lpB(i - 1) + FHRecordSize
                Next i
            End If
        Else
            Exit For
        End If
    Next j
    
    '----------------------
    RaiseEvent RecRead(XX)
    '----------------------
    DoEvents
    
    ReDim lpB(0 To JumRec - 1)
    
    ' Start the real Sorting
    '----------------------
    RaiseEvent SortStarted
    '----------------------
    DoEvents
    ' --------------------------------------------
    ' Initialising
    aStack(0).BytePos = FLebar - 1
    aStack(0).Start = 0
    aStack(0).End = JumRec - 1
    
    PosisiStack = 0
    Do
        j = aStack(PosisiStack).Start
        k = aStack(PosisiStack).End
        m = aStack(PosisiStack).BytePos

        If m >= 0 Then  ' Bytepos to evaluate still in field width
            If j = k Then
                PosisiStack = PosisiStack - 1
            Else
                If j >= k Then
                    PosisiStack = PosisiStack - 1
                ElseIf k - j = 1 Then   ' Only 2 records
                    ' Compare... byte by byte
                    For i = FLebar - 1 To 0 Step -1
                        If IsiData(lpA(j) + i) < IsiData(lpA(k) + i) Then
                            Exit For
                        ElseIf IsiData(lpA(j) + i) > IsiData(lpA(k) + i) Then
                            ' Swap
                            XX = lpA(j)
                            lpA(j) = lpA(k)
                            lpA(k) = XX
                            
                            Exit For    ' Exit For i
                        
                        End If
                    Next i
                    PosisiStack = PosisiStack - 1
                Else
                    LSet ByteCount = EmptyCounter
                    ' ASCENDING
                    If k - j < 128 Then
                        bMin = 255
                        bMax = 0
                        For i = j To k
                            lpB(i) = lpA(i)
                            ByteCount.pos(IsiData(lpB(i) + m)) = ByteCount.pos(IsiData(lpB(i) + m)) + 1
                            If IsiData(lpB(i) + m) > bMax Then bMax = IsiData(lpB(i) + m)
                            If IsiData(lpB(i) + m) < bMin Then bMin = IsiData(lpB(i) + m)
                        Next i
                    Else
                        bMin = 0
                        bMax = 255
                        For i = j To k
                            lpB(i) = lpA(i)
                            ByteCount.pos(IsiData(lpB(i) + m)) = ByteCount.pos(IsiData(lpB(i) + m)) + 1
                        Next i
                    End If
                    
                    ' Build histogram
                    ' ASCENDING
                    Histo(bMin) = j
                    For i = bMin To bMax - 1
                        Histo(i + 1) = ByteCount.pos(i) + Histo(i)
                    Next i
                    
                    ' Update stack info
                    n = 0
                    For i = bMax To bMin Step -1
                        If ByteCount.pos(i) > 1 Then
                            With aStack(PosisiStack + n)
                                .BytePos = m - 1
                                .Start = Histo(i)
                                .End = Histo(i) + ByteCount.pos(i) - 1
                            End With
                            n = n + 1
                        End If
                    Next i
                    
                    ' Updating Index info/Pointer info
                    For i = j To k
                        bVal = IsiData(lpB(i) + m)
                        lpA(Histo(bVal)) = lpB(i)
                        Histo(bVal) = Histo(bVal) + 1
                    Next i
                    
                    ' "IIf" not good inside a loop. It slower than "If...Else...End If"
                    'PosisiStack = IIf(n > 0, PosisiStack + n - 1, PosisiStack - 1)
                    If n > 0 Then
                        PosisiStack = PosisiStack + n - 1
                    Else
                        PosisiStack = PosisiStack - 1
                    End If
                End If
            
            End If
        Else        ' Exceed field width
            PosisiStack = PosisiStack - 1
        End If
    
    Loop Until PosisiStack < 0
    
    ' ================================= EXTRA 1 PASS for negative values =========================
    If FlipBit Then
        ' Reverse the order of negative and positive value
        ' Single pass counting sort
        j = 0
        k = JumRec - 1
        m = FLebar - 1
        
                If k - j = 1 Then   ' Only 2 records
                    ' Compare... byte by byte
                    For i = FLebar - 1 To 0 Step -1
                        If IsiData(lpA(j) + i) < IsiData(lpA(k) + i) Then
                            Exit For
                        ElseIf IsiData(lpA(j) + i) > IsiData(lpA(k) + i) Then
                            ' Swap
                            XX = lpA(j)
                            lpA(j) = lpA(k)
                            lpA(k) = XX
                            Exit For    ' Exit For i
                        End If
                    Next i
                Else
                    LSet ByteCount = EmptyCounter
                    For i = j To k
                        lpB(i) = lpA(i)
                        If IsiData(lpB(i) + m) > 128 Then       ' Negative values
                            ByteCount.pos(0) = ByteCount.pos(0) + 1
                        Else
                            ByteCount.pos(100) = ByteCount.pos(100) + 1
                        End If
                    Next i
                    
                    ' Build histogram
                    ' ASCENDING
                    Histo(0) = 0
                    Histo(100) = ByteCount.pos(0)
                    
                    If ByteCount.pos(0) > 0 Then
                        If isDouble Then
                            ' Updating Index info/Pointer info: Floating Number
                            XX = ByteCount.pos(0) - 1
                            wWw = ByteCount.pos(0)
                            For i = j To k
                                If IsiData(lpB(i) + m) > 128 Then   ' Negative
                                    lpA(XX) = lpB(i)
                                    XX = XX - 1
                                Else                                ' Positive
                                    lpA(wWw) = lpB(i)
                                    wWw = wWw + 1
                                End If
                            Next i
                        Else
                            ' Updating Index info/Pointer info: Integer
                            XX = 0
                            wWw = ByteCount.pos(0)
                            For i = j To k
                                If IsiData(lpB(i) + m) > 128 Then   ' Negative
                                    lpA(XX) = lpB(i)
                                    XX = XX + 1
                                Else                                ' Positive
                                    lpA(wWw) = lpB(i)
                                    wWw = wWw + 1
                                End If
                            Next i
                        End If  ' Floating or integer condition
                    End If
                End If
        
    End If      ' FlipBit condition
    
    '---------------------- INTERNAL SORT COMPLETED ----------------------
    RaiseEvent SortCompleted
    '----------------------
    DoEvents
    
    ' ================================= INTERNAL SORT FINISHED =========================
    ' ================================= INTERNAL SORT FINISHED =========================
    ' ================================= INTERNAL SORT FINISHED =========================
    ' ================================= INTERNAL SORT FINISHED =========================
    ' ================================= INTERNAL SORT FINISHED =========================
    ' menuliskan hasil ke file .HS
    srtNomorFile = FreeFile
    Open sTargetFile For Binary As #srtNomorFile
    ReDim RWBuffer(0 To FHeader.HeaderSize - 1) ' buffer for file header to copy
    Get #nNomorFile, 1, RWBuffer
    Put #srtNomorFile, 1, RWBuffer
    
    DoEvents
    ' Menghitung untuk cache data
    GlobalMemoryStatus InfoMemory
    
    IsiTembolok = (InfoMemory.dwAvailPhys \ FHRecordSize) - ((InfoMemory.dwAvailPhys \ FHRecordSize) Mod 5000)
    IsiTembolok = (IsiTembolok \ 2)
    
    If IsiTembolok >= JumRec Then IsiTembolok = JumRec
    
    ' --------------------------------------------
    ' Membaca dan meload ke cache
    ReDim Tembolok(0 To (IsiTembolok * FHRecordSize) - 1)
    DoEvents
    
    PTembolok = VarPtr(Tembolok(0))
    Seek #nNomorFile, DataOffset
    Get #nNomorFile, , Tembolok
    DoEvents
    
    ' --------------------------------------------
    ' Menuliskan hasil ke file tujuan
    Seek #srtNomorFile, DataOffset
    
    ' Jumlah buffer maksimal ditentukan maksimal sebesar "wWw" records
    wWw = 65535 \ FHRecordSize
    
    If JumRec >= wWw Then
        ReDim IsiData(0 To (wWw * FHRecordSize) - 1)
        ReDim lpB(0 To wWw - 1)
        
        ' Jumlah pengulangan loop
        n = IIf(JumRec Mod wWw = 0, JumRec \ wWw, (JumRec \ wWw) + 1)
    Else
        ReDim IsiData(0 To (JumRec * FHRecordSize) - 1)
        ReDim lpB(0 To JumRec - 1)
        
        ' Jumlah pengulangan loop
        n = 1
    End If
    
    ' Mengisi alamat memori array untuk source
    lpB(0) = VarPtr(IsiData(0))
    k = UBound(lpB)
    For i = 1 To k
        lpB(i) = lpB(i - 1) + FHRecordSize
    Next i
    
    ReDim RWBuffer(0 To FHRecordSize - 1)
    ' Mengisi alamat memory array untuk RW buffer
    m = VarPtr(RWBuffer(0))
    
    XX = 0  ' Mewakili banyaknya records
    ' Memulai proses Baca-Tulis secara berulang...
    For j = 0 To n - 1
        For i = 0 To k
            ' Baca dari file sumber... atau dari cache, tergantung
            If (lpA(XX) \ FLebar) > IsiTembolok Then
                Get #nNomorFile, DataOffset + ((lpA(XX) \ FLebar) * FHRecordSize), RWBuffer
                ' Copy memory
                CopyMemoryBV lpB(i), m, FHRecordSize
            Else
                ' langsung Copy memory
                CopyMemoryBV lpB(i), PTembolok + ((lpA(XX) \ FLebar) * FHRecordSize), FHRecordSize
            End If
            
            
            XX = XX + 1     ' Counter
        Next i
        ' Tuliskan ke file tujuan
        Put #srtNomorFile, , IsiData
        
        ' Hitung yang sudah ditulis dan tampilkan
        ' Feedback for user, because this proccess require a lot of time
        ' Time needed mostly for disk read and write.
        ' This event helps user know how long they must wait
        '----------------------
        If XX Mod (wWw * 40) = 0 Then
            RaiseEvent RecWriten(XX)
        '----------------------
            DoEvents
        End If
        
        ' Mengecek sisa
        If XX < JumRec Then
            If JumRec - XX < wWw Then
                ReDim IsiData(0 To ((JumRec - XX) * FHRecordSize) - 1)
                ReDim lpB(0 To (JumRec - XX) - 1)
                lpB(0) = VarPtr(IsiData(0))
                k = UBound(lpB)
                For i = 1 To k
                    lpB(i) = lpB(i - 1) + FHRecordSize
                Next i
            End If
        Else
            Exit For
        End If
    Next j
    
    '----------------------
    RaiseEvent RecWriten(XX)
    '----------------------
    DoEvents
    
    ' Closes the target file
    Close #srtNomorFile
    DoEvents
    
    
    Erase IsiData
    Erase lpA
    Erase lpB
    Erase RWBuffer
    Erase Histo
    
    DoEvents
    
End Sub

' -------------------------------------------------
' Sort for character string
' -------------------------------------------------
Private Sub SortDataUrut(sTargetFile As String, FldNum As Long)    ' For sorting Char, Logical, Byte, Numeric, and Floating
    Dim srtNomorFile As Integer
    Dim InfoMemory As MEMORYSTATUS
    Dim FLebar As Long  ' Field width

    Dim RWBuffer() As Byte      ' Buffer for one record
    Dim IsiData() As Byte       ' Buffer file 64KB
    Dim Tembolok() As Byte      ' Cache for large file
    Dim PTembolok As Long       ' Pointer to cache
    Dim IsiTembolok As Long
    
    Dim lpData As Long
    Dim lpA() As Long
    Dim lpB() As Long
    Dim plpA As Long
    Dim plpB As Long
    
    Dim ByteCount As bCounter
    Dim EmptyCounter As bCounter
    
    Dim Histo() As Long
    
    Dim aStack() As BKeyInfo
    Dim bMin As Byte
    Dim bMax As Byte
    Dim bVal As Byte
    
    Dim i As Long, j As Long, k As Long, m As Long, n As Long
    Dim p As Long, q As Long
    
    Dim PosisiStack As Long
    
    Dim JumRec As Long
    Dim XX As Long
    Dim wWw As Long


    ' --------------------------------------------
    ' Menyiapkan memory untuk menampung byte array data-data
    ' Asumsi: Memorynya bisa menampung semua key
    JumRec = FHeader.NumberOfRecords
    FLebar = FIFieldWidth(FldNum)
    
    ReDim aStack(0 To (FLebar * 255))
    ReDim Histo(0 To 255)
    
    ' --------------------------------------------
    ' Membaca record satu persatu
    
    Seek #nNomorFile, DataOffset
    
    ' Baca dan copy alamat........
    ReDim IsiData(0 To (FLebar * FHeader.NumberOfRecords) - 1)  ' Ini untuk menampung data-data dalam bentuk byte array
    ReDim lpA(0 To JumRec - 1)
    lpA(0) = VarPtr(IsiData(0))
    
    ' Jumlah buffer maksimal ditentukan maksimal sebesar "wWw" records
    wWw = 65535 \ FHRecordSize
    If JumRec >= wWw Then
        ReDim RWBuffer(0 To (wWw * FHRecordSize) - 1)
        ReDim lpB(0 To wWw - 1)
        
        ' Jumlah pengulangan loop
        n = IIf(JumRec Mod wWw = 0, JumRec \ wWw, (JumRec \ wWw) + 1)
    Else
        ReDim RWBuffer(0 To (JumRec * FHRecordSize) - 1)
        ReDim lpB(0 To JumRec - 1)
        
        ' Jumlah pengulangan loop
        n = 1
    End If
    
    ' Mengisi alamat memori array untuk source
    lpB(0) = VarPtr(RWBuffer(0)) + FIFieldDisplacement(FldNum)
    k = UBound(lpB)
    For i = 1 To k
        lpB(i) = lpB(i - 1) + FHRecordSize
    Next i
    
    m = VarPtr(IsiData(0))
    
    XX = 0      ' Mewakili jumlah record yang dibaca
    
    ' Mengulang baca tulis
    For j = 0 To n - 1
        ' Baca dari file sumber...
        Get #nNomorFile, , RWBuffer
        For i = 0 To k
            lpA(XX) = XX * FLebar
            
            ' Copy memory to byte array
            CopyMemoryBV lpA(XX) + m, lpB(i), FLebar
            
            XX = XX + 1     ' Counter
        Next i
        
        ' Hitung yang sudah ditulis dan tampilkan
        ' Feedback for user, because this proccess require a lot of time
        ' Time needed mostly for disk read and write.
        ' This event helps user know how long they must wait
        '----------------------
        If XX Mod (wWw * 100) = 0 Then
            RaiseEvent RecRead(XX)
        '----------------------
            DoEvents
        End If
        
        ' Mengecek sisa
        If XX < JumRec Then
            If JumRec - XX < wWw Then
                ReDim RWBuffer(0 To ((JumRec - XX) * FHRecordSize) - 1)
                ReDim lpB(0 To (JumRec - XX) - 1)
                lpB(0) = VarPtr(RWBuffer(0)) + FIFieldDisplacement(FldNum)
                k = UBound(lpB)
                For i = 1 To k
                    lpB(i) = lpB(i - 1) + FHRecordSize
                Next i
            End If
        Else
            Exit For
        End If
    Next j
    
    '----------------------
    RaiseEvent RecRead(XX)
    '----------------------
    ' Start the real Sorting
    '----------------------
    RaiseEvent SortStarted
    '----------------------
    DoEvents
    DoEvents
    
    ReDim lpB(0 To JumRec - 1)
    plpA = VarPtr(lpA(0))
    plpB = VarPtr(lpB(0))
    
    ' --------------------------------------------
    ' Mengurutkan
    aStack(0).BytePos = 0
    aStack(0).Start = 0
    aStack(0).End = JumRec - 1
    
    PosisiStack = 0
    Do
        j = aStack(PosisiStack).Start
        k = aStack(PosisiStack).End
        m = aStack(PosisiStack).BytePos

        If m < FLebar Then  ' Bytepos to evaluate still in field width
            If j = k Then
                PosisiStack = PosisiStack - 1
            Else
                If j >= k Then
                    PosisiStack = PosisiStack - 1
                ElseIf k - j = 1 Then   ' Only 2 records
                    ' Compare...
                    p = lpA(j)
                    q = lpA(k)
                    For i = m To FLebar - 1
'                    For i = m + 1 To FLebar - 1
                        If IsiData(p + i) < IsiData(q + i) Then
                            Exit For
                        ElseIf IsiData(p + i) > IsiData(q + i) Then
                            ' Swap
                            XX = lpA(j)
                            lpA(j) = lpA(k)
                            lpA(k) = XX
'                            XX = p
'                            p = q
'                            q = XX
                            Exit For    ' Exit For i
                        End If
                    Next i
                    PosisiStack = PosisiStack - 1
                Else
                    LSet ByteCount = EmptyCounter
                    If k - j < 128 Then
                        bMin = 255
                        bMax = 0
                        
                        'CopyMemoryBV plpB + (j * 4), plpA + (j * 4), ((k - j) + 1) * 4
                        For i = j To k
                            lpB(i) = lpA(i)
                            XX = IsiData(lpA(i) + m)
                            ByteCount.pos(XX) = ByteCount.pos(XX) + 1
                            If XX > bMax Then bMax = XX
                            If XX < bMin Then bMin = XX
                        Next i
                        
                        
                    Else
                        bMin = 0
                        bMax = 255
                        
                        CopyMemoryBV plpB + (j * 4), plpA + (j * 4), ((k - j) + 1) * 4
                        For i = j To k
                            'lpB(i) = lpA(i)
                            XX = IsiData(lpA(i) + m)
                            ByteCount.pos(XX) = ByteCount.pos(XX) + 1
                        Next i
                        
                    End If
                    
                    ' Build histogram
                    ' ASCENDING
                    Histo(bMin) = j
                    For i = bMin To bMax - 1
                        Histo(i + 1) = ByteCount.pos(i) + Histo(i)
                    Next i
                    
                    ' Update stack info
                    n = 0
                    For i = bMax To bMin Step -1
                        If ByteCount.pos(i) > 1 Then
                            With aStack(PosisiStack + n)
                                .BytePos = m + 1
                                .Start = Histo(i)
                                .End = Histo(i) + ByteCount.pos(i) - 1
                            End With
                            n = n + 1
                        End If
                    Next i
                    
                    ' Updating Index info/Pointer info
                    If n <> 1 Or aStack(PosisiStack).Start <> j Or aStack(PosisiStack).End <> k Then
                        For i = j To k
                            bVal = IsiData(lpB(i) + m)
                            lpA(Histo(bVal)) = lpB(i)
                            Histo(bVal) = Histo(bVal) + 1
                        Next i
                    End If
                    
                    ' "IIf" not good inside a loop. It slower than "If...Else...End If"
                    'PosisiStack = IIf(n > 0, PosisiStack + n - 1, PosisiStack - 1)
                    If n > 0 Then
                        PosisiStack = PosisiStack + n - 1
                    Else
                        PosisiStack = PosisiStack - 1
                    End If
                
                End If
            
            End If
        Else        ' Exceed field width
            PosisiStack = PosisiStack - 1
        End If
    
    Loop Until PosisiStack < 0
    
    '----------------------
    RaiseEvent SortCompleted
    '----------------------
    DoEvents
    ' Added This Single Line to free some memory for next caching
    Erase IsiData
    ' End of freeing some memory
    
    DoEvents
    DoEvents
    
    ' menuliskan hasil ke file .HS
    srtNomorFile = FreeFile
    Open sTargetFile For Binary As #srtNomorFile
    ReDim RWBuffer(0 To FHeader.HeaderSize - 1) ' buffer for file header to copy
    Get #nNomorFile, 1, RWBuffer
    Put #srtNomorFile, 1, RWBuffer
    
    DoEvents
    ' Menghitung untuk cache data
    GlobalMemoryStatus InfoMemory
    
    IsiTembolok = (InfoMemory.dwAvailPhys \ FHRecordSize) - ((InfoMemory.dwAvailPhys \ FHRecordSize) Mod 5000)
    IsiTembolok = (IsiTembolok \ 2)
    
    If IsiTembolok >= JumRec Then IsiTembolok = JumRec
    
    ' --------------------------------------------
    ' Membaca dan meload ke cache
    ReDim Tembolok(0 To (IsiTembolok * FHRecordSize) - 1)
    DoEvents
    
    PTembolok = VarPtr(Tembolok(0))
    Seek #nNomorFile, DataOffset
    Get #nNomorFile, , Tembolok
    DoEvents
    
    ' --------------------------------------------
    ' Menuliskan hasil ke file tujuan
    Seek #srtNomorFile, DataOffset
    
    ' Jumlah buffer maksimal ditentukan maksimal sebesar "wWw" records
    wWw = 65535 \ FHRecordSize
    
    If JumRec >= wWw Then
        ReDim IsiData(0 To (wWw * FHRecordSize) - 1)
        ReDim lpB(0 To wWw - 1)
        
        ' Jumlah pengulangan loop
        n = IIf(JumRec Mod wWw = 0, JumRec \ wWw, (JumRec \ wWw) + 1)
    Else
        ReDim IsiData(0 To (JumRec * FHRecordSize) - 1)
        ReDim lpB(0 To JumRec - 1)
        
        ' Jumlah pengulangan loop
        n = 1
    End If
    
    ' Mengisi alamat memori array untuk source
    lpB(0) = VarPtr(IsiData(0))
    k = UBound(lpB)
    For i = 1 To k
        lpB(i) = lpB(i - 1) + FHRecordSize
    Next i
    
    ReDim RWBuffer(0 To FHRecordSize - 1)
    ' Mengisi alamat memory array untuk RW buffer
    m = VarPtr(RWBuffer(0))
    
    XX = 0  ' Mewakili banyaknya records
    ' Memulai proses Baca-Tulis secara berulang...
    For j = 0 To n - 1
        For i = 0 To k
            ' Baca dari file sumber... atau dari cache, tergantung
            If (lpA(XX) \ FLebar) > IsiTembolok Then
                Get #nNomorFile, DataOffset + ((lpA(XX) \ FLebar) * FHRecordSize), RWBuffer
                ' Copy memory
                CopyMemoryBV lpB(i), m, FHRecordSize
            Else
                ' langsung Copy memory
                CopyMemoryBV lpB(i), PTembolok + ((lpA(XX) \ FLebar) * FHRecordSize), FHRecordSize
            End If
            
            
            XX = XX + 1     ' Counter
        Next i
        ' Tuliskan ke file tujuan
        Put #srtNomorFile, , IsiData
        
        ' Hitung yang sudah ditulis dan tampilkan
        ' Feedback for user, because this proccess require a lot of time
        ' Time needed mostly for disk read and write.
        ' This event helps user know how long they must wait
        '----------------------
        If XX Mod (wWw * 40) = 0 Then
            RaiseEvent RecWriten(XX)
        '----------------------
            DoEvents
        End If
        
        ' Mengecek sisa
        If XX < JumRec Then
            If JumRec - XX < wWw Then
                ReDim IsiData(0 To ((JumRec - XX) * FHRecordSize) - 1)
                ReDim lpB(0 To (JumRec - XX) - 1)
                lpB(0) = VarPtr(IsiData(0))
                k = UBound(lpB)
                For i = 1 To k
                    lpB(i) = lpB(i - 1) + FHRecordSize
                Next i
            End If
        Else
            Exit For
        End If
    Next j
    
    '----------------------
    RaiseEvent RecWriten(XX)
    '----------------------
    DoEvents
    
    ' Closes the target file
    Close #srtNomorFile
    
    Erase IsiData
    Erase lpA
    Erase lpB
    Erase RWBuffer
    Erase Histo

End Sub


