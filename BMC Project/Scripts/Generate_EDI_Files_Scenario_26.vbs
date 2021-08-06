Dim rootPath, lineOfBusiness, scenarioName
Dim segmentLoopDict
' Declaring variables for file system object
Dim FSO
' this integer will indicate the total number of insurance segments to be created
Dim numberOfRecordsNeeded
Dim val_separator
Dim randFirstName, randLastName, randSSN, randCity, randState, randDOB, randAddr1, randAddr2, randMiddleInitial, randUsed
Dim randomDataDict
' currentRandKey will be used to indicate the current index of edi file being created,
' this will be incremented by one in each loop creating an edi file based on the numberOfRecordsNeeded
Dim currentRandKey
Dim template_folder, supportingFilesFolder, outputFolder, logFolder, endLoopNameIndicator, beginSegmentSeparator, endSegmentSeparator
Dim randIndicator, randChooseIndicator, refIndicator, randNumIndicator, randDateIndicator, randTimeIndicator, randFNIndicator, randLNIndicator, randMIIndicator, randSSNIndicator, randDOBIndicator, randADDR1Indicator, randADDR2Indicator
Dim combIndicator
Dim memberIdLength ' variable that will indicate the length of the unique member id for the line of business, whose value should be changed in config file
Dim isa06, isa08
' variable to save the values to be inserted into database
Dim dataToBeSavedToDbDict
' array variable to store the list of dataToBeSavedToDbDict created for each edi record
Dim dataToBeSavedToDbArr
' array variable to store the list of dataToBeSavedToDbDict created for each edi record for dependants
Dim depDataToBeSavedToDbArr


' variable that will indicate if template file should be saved as a new file after making changes to deal with loop segments with special logic
Dim saveAsNewFile
' variable to hold the path of temporary template file created
Dim newTemplateFilePath
Dim logFilePath
' this variable will store the EDI template excel file name for the given line of business
Dim templateFileName
' this variable will be used to store the unique insurance id number to be generated for each member being enrolled, this is used in function for storing values to DB
Dim MEME_MEDCD_NO
Dim storedDataDict
Dim oShell
Dim dependantDataRequired, how_many_random_data, dependantRandomDataDict, dep_ct, segments_req_count


Sub MainFunction()
    
    st_time = Time
    rootPath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project"
    fnCreateEDIFile rootPath, "DCH", 1, 1
'    Application.Wait Now() + TimeValue("00:00:03")
'    fnCreateEDIFile rootPath, "CCA", 1, 1
'    Application.Wait Now() + TimeValue("00:00:03")
'    MsgBox "Done scenario 1 - take copy of DB"
'    fnCreateEDIFile rootPath, "CCA", 1, 4
'    Application.Wait Now() + TimeValue("00:00:03")
'    fnCreateEDIFile rootPath, "HSA", 1, 4
'    Application.Wait Now() + TimeValue("00:00:03")
'    MsgBox "Done scenario 4 - take copy of DB"
'    fnCreateEDIFile rootPath, "CCA", 1, 9
'    Application.Wait Now() + TimeValue("00:00:03")
'    fnCreateEDIFile rootPath, "HSA", 1, 9
'    fnCreateEDIFile rootPath, "BACO", 1, 17
'    Application.Wait Now() + TimeValue("00:00:03")
'    fnCreateEDIFile rootPath, "MERCY", 1, 17
'    Application.Wait Now() + TimeValue("00:00:03")
'    fnCreateEDIFile rootPath, "SIGNA", 1, 17
'    Application.Wait Now() + TimeValue("00:00:03")
'    fnCreateEDIFile rootPath, "SOUTHCOAST", 1, 17
'    Application.Wait Now() + TimeValue("00:00:03")
'    fnCreateEDIFile rootPath, "NH", 1, 17
    'fnCreateEDIFile rootPath, "SOUTHCOAST", 1, 5
    'fnCreateEDIFile rootPath, "BACO", 1, 5
    'fnCreateEDIFile rootPath, "NH", 1, 5
    'fnCreateEDIFile rootPath, "CCA", 1, 5
    'fnCreateEDIFile rootPath, "HSA", 1, 5
    
    MsgBox "Done"
End Sub
Sub fnCreateEDIFile(rootPath, line_of_business, noOfRecordsNeeded, ScenName)
    'wscript.echo "inside main " & rootPath & " - " & noOfRecordsNeeded
'    ' values to be passed as arguments
'    rootPath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project"
'    lineOfBusiness = "MH"
'
'
'    scenarioName = "Scenario 1"
'    numberOfRecordsNeeded = 3
'    '***************
' delimiter that will be used to separate the values for the segment
    val_separator = ","
    ' indicators used to represent what random value is required
    randIndicator = "RAND:"
    randChooseIndicator = "ANY:"
    refIndicator = "REF:"
    randNumIndicator = "NUMBER"
    randDateIndicator = "DATE"
    randTimeIndicator = "TIME"
    randFNIndicator = "FN"
    randLNIndicator = "LN"
    randMIIndicator = "MI"
    randSSNIndicator = "SSN"
    randDOBIndicator = "DOB"
    randADDR1Indicator = "ADDR1"
    randADDR2Indicator = "ADDR2"
    combIndicator = "COMB"
    'default values for folders used in project
    template_folder = "Scenarios"
    supportingFilesFolder = "Supporting Files"
    outputFolder = "Output"
    logFolder = "Logs"
    'loop name that will indicate the end of the last row of the EDI Scenario Template file
    endLoopNameIndicator = "10000"
    beginSegmentSeparator = "INS*"
    endSegmentSeparator = "~SE*"

'    '***************
'    'default values for folders used in project
'    template_folder = "Scenarios"
'    supportingFilesFolder = "Supporting Files"
'    outputFolder = "Output"
'    logFolder = "Logs"
'    '******************
'
    '''On Error Resume Next
    'Creating a object for writing a text file.
    lineOfBusiness = line_of_business
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set oShell = CreateObject("WScript.Shell")
    ' assigning value to numberOfRecordsNeeded based on the value passed as argument by user
    numberOfRecordsNeeded = noOfRecordsNeeded
    'calling function to read the config file
    'On Error Resume Next
'    ReadConfigFile rootPath
'    If Err.Number <> 0 Then
'        MsgBox "Error: " & Err.Description & vbNewLine & "Stopping script execution."
'        wcript.Quit
'    End If
    On Error GoTo 0
    fileSuffix = CStr(GenerateUniqueRandomNumber(""))
    'wscript.echo "fileSuffix " & fileSuffix
    '*******************
    templateFileName = "EDITemplate"
    
    'creating file paths using the arguments passed
    Call getscenarioName(ScenName)
    
    Call getMemberIdLength
    
    Call getIsaValues
    
    logFileName = "Log_File.txt"
    logFilePath = rootPath & Chr(92) & logFileName
    
    ' deleting Log file created in previous run
    On Error Resume Next
    If FSO.FileExists(logFilePath) Then
        FSO.DeleteFile logFilePath
    End If
    On Error GoTo 0
    
    strDataPath = rootPath & Chr(92) & templateFileName & ".xlsx" ' chr(92) is for Backslash
    LogStatement "strDataPath: " & strDataPath
    'wscript.echo "logFilePath " & logFilePath
    randomDataFilePath = rootPath & Chr(92) & "Random_Data_DB.xlsx"
    LogStatement "randomDataFilePath: " & randomDataFilePath
    'finalOutFilePath = rootPath & Chr(92) & "Output_EDI_File" & ".txt"
    If lineOfBusiness <> "DCH" Then
        finalOutFilePath = rootPath & Chr(92) & "Output_EDI_File_" & lineOfBusiness & "_" & Split(scenarioName, "_", 2)(1) & "_" & fileSuffix & ".txt"
    Else
        finalOutFilePath = rootPath & Chr(92) & "834_" & Year(Date) & Month(Date) & Day(Date) & getRandomNumberOfLength(9) & "Z_BMCHP_C_E_S_1" & ".txt"
    
    End If
    ' deleting output file created in previous run
    On Error Resume Next
    If FSO.FileExists(finalOutFilePath) Then
        FSO.DeleteFile finalOutFilePath
    End If
    On Error GoTo 0
    LogStatement "finalOutFilePath: " & finalOutFilePath
    updateDbPath = rootPath & Chr(92) & "Random_Updated_DB.xlsx"
    LogStatement "updateDbPath: " & updateDbPath
    'On Error Resume Next
    Call PreprocessTemplateFile(rootPath, lineOfBusiness, scenarioName)
    
    If Err.Number <> 0 Then
        LogStatement "Error: " & Err.Description
        MsgBox "Error: " & Err.Description & vbNewLine & "Stopping script execution."
        'Call cleanup
    End If
    On Error GoTo 0
    ' changing the path of the template file to point to temporary template file
    If saveAsNewFile = True Then
        LogStatement "changed the path of the template file to point to temporary template file"
        strDataPath = newTemplateFilePath
    End If
    ''''''''deug.print "template file path is " & strDataPath
    ' Declaring variables for Excel object
    Dim objExcel
    ' variable for Excel workbook
    Dim objTemplateWorkbook
    ' variable for Excel worksheet
    Dim objTemplateWorksheet
    ' variable for Excel worksheet Range
    Dim objTemplateWorksheetRange, objTemplateWorksheetRangeRows
    ' creating a Dictionary object that will store all segment values with Loop + SegmentName as key
    Set segmentLoopDict = CreateObject("Scripting.Dictionary")
    'Set randomDataDict = CreateObject("Scripting.Dictionary")
'    ' dictionary object that will store values to be inserted to db for each record
'    Set dataToBeSavedToDbDict = CreateObject("Scripting.Dictionary")
    ' variable that will store the dictionary object dataToBeSavedToDbDict created for each record
    ' this dict will be used finally to update Database
    Set dataToBeSavedToDbArr = CreateObject("Scripting.Dictionary")
    
    If InStr(1, scenarioName, "Scenario_7") <> 0 Or InStr(1, scenarioName, "Scenario_8") <> 0 Or InStr(1, scenarioName, "Scenario_9") <> 0 Or InStr(1, scenarioName, "Scenario_10") <> 0 _
        Or InStr(1, scenarioName, "Scenario_11") <> 0 Or InStr(1, scenarioName, "Scenario_12") <> 0 Or InStr(1, scenarioName, "Scenario_13") <> 0 Or InStr(1, scenarioName, "Scenario_14") <> 0 Then
            dependantDataRequired = True
            Set depDataToBeSavedToDbArr = CreateObject("Scripting.Dictionary") 'overall dictionary to save all dependants dictionary
            how_many_random_data = 3
            dep_ct = 0
    Else
        dependantDataRequired = False
        how_many_random_data = 1
        dep_ct = 0
    End If
    
    
    LogStatement "calling function CreateSaveDataDictionary that will create dictionary object and insert keys with nulll value"
    ' calling function that will create dictionary object and insert keys with nulll value
    ' values will be updated as and when the value is read from template file in Function create_segment
    Call CreateSaveDataDictionary(strDataPath, lineOfBusiness)
    LogStatement "CreateSaveDataDictionary Function completed"
    Set objExcel = CreateObject("Excel.Application")
    If FSO.FileExists(strDataPath) Then
        Set objTemplateWorkbook = objExcel.Workbooks.Open(strDataPath)
        Set objTemplateWorksheet = objTemplateWorkbook.Worksheets(scenarioName)
    Else
        LogStatement "Template excel file not found " & strDataPath
        'Call cleanup
        Exit Sub
    End If
    
    Set objTemplateWorksheetRange = objTemplateWorksheet.UsedRange
    objTemplateWorksheetRangeRows = objTemplateWorksheetRange.Rows.Count
    totalColumnCount = 7
    ' loop for processing each edi file
        'Call create_beginning_segment(objTemplateWorksheet, objTemplateWorksheetRangeRows)
        
       
        
        
        
        LogStatement "calling function GetRandomDataValues from Main function"
        '************************
        If Split(scenarioName, "_", 2)(1) = "Scenario_1" Or InStr(1, scenarioName, "Scenario_7") <> 0 Or InStr(1, scenarioName, "Scenario_8") <> 0 Or InStr(1, scenarioName, "Scenario_16") <> 0 Then
            'GetRandomDataValues numberOfRecordsNeeded, randomDataFilePath
            'On Error Resume Next
            GetRandomDataValues_2 numberOfRecordsNeeded, randomDataFilePath
            'MsgBox "Completed GetRandomDataValues_2"
            If Err.Number <> 0 Then
                LogStatement "Error: " & Err.Description
                MsgBox "Error: " & Err.Description & vbNewLine & "Stopping script execution."
                'Call cleanup
            End If
            On Error GoTo 0

            
        Else
            'On Error Resume Next
            GetStoredRandomDataValues numberOfRecordsNeeded, updateDbPath
            
            If Err.Number <> 0 Then
                LogStatement "Error: " & Err.Description
                MsgBox "Error: " & Err.Description & vbNewLine & "Stopping script execution."
                'Call cleanup
            End If
            On Error GoTo 0
            
        End If
        'variable to hold the row ranges for each segment
        Dim segment_loop_ranges
        LogStatement "calling function FindSegments from Main function"
        'On Error Resume Next
        segment_loop_ranges = FindSegments(objTemplateWorksheet, "0", endLoopNameIndicator)
        If Err.Number <> 0 Then
            LogStatement "Error: " & Err.Description
            MsgBox "Error: " & Err.Description & vbNewLine & "Stopping script execution."
            'Call cleanup
        End If
        On Error GoTo 0
        
        
        ''msgbox segment_loop_ranges
        '*********************
        ' below code is to remove duplicate ranges
        row_numbers_range_arr = Split(segment_loop_ranges, ":")
        new_row_numbers_range = ""
        For Each Rng_txt In row_numbers_range_arr
            If InStr(1, new_row_numbers_range, Rng_txt) = 0 Then
                If Len(new_row_numbers_range) = 0 Then
                    new_row_numbers_range = Rng_txt
                Else
                    new_row_numbers_range = new_row_numbers_range & ":" & Rng_txt
                End If
            End If
        Next
        segment_loop_ranges = new_row_numbers_range
        LogStatement "Final segment_loop_ranges " & segment_loop_ranges
        ''''''''deug.print segment_loop_ranges
        '*********************
        'array variable to hold the split row range for each segment
        Dim segment_loop_ranges_array
        segment_loop_ranges_array = Split(segment_loop_ranges, ":")
        LogStatement "segment_loop_ranges_array created"
        ''msgbox "Length of array segment_loop_ranges_array is " & UBound(segment_loop_ranges_array)
        ' variable to hold the complete edi text for specified number of records
        Dim allEdiTextDict
        ' variable to hold only the beginning segment of the edi text
        Dim BeginSegmentText
        ' variable to hold only the end segment of the edi text
        Dim EndSegmentText
        ' variable to hold the entire text of all edi records, this variable will be used to produce a consolidated edi text file
        Dim fullEdiText
        ' assigning default values to variables
        BeginSegmentText = ""
        EndSegmentText = ""
        fullEdiText = ""
        Set allEdiTextDict = CreateObject("Scripting.Dictionary")
        'variables for each row of the template specifying the segment and values
        Dim tmplLoop, tmplSubLoop, tmplSegmentName, tmplRequirement, tmplLoopCount, tmplSegmentValue, tmplValueLength
        'variable for complete edi string, beginning_segment and ending segment of one record
        Dim complete_edi_text, beginning_segment, ending_segment
        If dependantDataRequired = True Then
            how_many_segments = 3 ' member plus 2 dependants
        Else
            how_many_segments = 1
        End If
    
    For edi_file_no = 1 To numberOfRecordsNeeded
        MEME_MEDCD_NO = ""
        For segments_req_count = 1 To how_many_segments
            LogStatement "Starting to create segment for edi_file_no " & edi_file_no
            outFileName = lineOfBusiness & "_" & scenarioName & "_" & fileSuffix & "_" & edi_file_no & ".txt"
            outFilePath = rootPath & Chr(92) & outFileName
            currentRandKey = edi_file_no
            '''''''''deug.print "currentRandKey value 1 " & currentRandKey
            'clearing the dictionary that holds the segment values to be used in references during each loop of creating a new edi record
            segmentLoopDict.RemoveAll
            complete_edi_text = ""
            For Each arr In segment_loop_ranges_array
                arr_range = Split(arr, "$")(0)
                arr_place_value = Split(arr, "$")(1)
                objSegmentStartRow = Split(arr_range, "to")(0)
                objSegmentEndRow = Split(arr_range, "to")(1)
                If Split(scenarioName, "_", 2)(1) = "Scenario_1" Or InStr(1, scenarioName, "Scenario_16") <> 0 Then
                    complete_edi_text = complete_edi_text + create_segment(objTemplateWorksheet, objSegmentStartRow, objSegmentEndRow, arr_place_value)
                ElseIf InStr(1, scenarioName, "Scenario_7") <> 0 Or InStr(1, scenarioName, "Scenario_8") <> 0 Then
                    
                    
                    complete_edi_text = complete_edi_text + create_segment_7(objTemplateWorksheet, objSegmentStartRow, objSegmentEndRow, arr_place_value)
                ElseIf InStr(1, scenarioName, "Scenario_9") <> 0 Or InStr(1, scenarioName, "Scenario_11") <> 0 Or InStr(1, scenarioName, "Scenario_12") <> 0 Or InStr(1, scenarioName, "Scenario_10") <> 0 Then
                    
                    
                    complete_edi_text = complete_edi_text + create_segment_scenario9(objTemplateWorksheet, objSegmentStartRow, objSegmentEndRow, arr_place_value)
                    
                Else
                    complete_edi_text = complete_edi_text + create_segment_scenario2(objTemplateWorksheet, objSegmentStartRow, objSegmentEndRow, arr_place_value)
                End If
            Next
            LogStatement "complete_edi_text created"
            'removing extra ~ added at the end of the edi text
            'If Right(complete_edi_text, 1) = "~" Then
            '    complete_edi_text = Left(complete_edi_text, Len(complete_edi_text) - 1)
            'End If
            '''''''''deug.print complete_edi_text
            CreateOutputTextFile outFilePath, complete_edi_text, "Y"
            'LogStatement "complete_edi_text written to output file"
            currentEdiFileText = complete_edi_text
            
            currentEdiFileTextArr = Split(currentEdiFileText, "INS*")
            MiddleSegmentText = Split(currentEdiFileTextArr(1), "~SE*")(0)
            MiddleSegmentText = "INS*" & MiddleSegmentText
            If Not allEdiTextDict.exists(edi_file_no) Then
                allEdiTextDict.Add edi_file_no, MiddleSegmentText
            Else
                allEdiTextDict(edi_file_no) = allEdiTextDict(edi_file_no) & "~" & MiddleSegmentText
            End If
            If edi_file_no = 1 And segments_req_count = 1 Then
                BeginSegmentText = currentEdiFileTextArr(0)
                EndSegmentText = Split(currentEdiFileTextArr(1), "~SE*")(1)
                EndSegmentText = "~SE*" & EndSegmentText
                LogStatement "Steps completed to get BeginSegmentText and EndSegmentText"
            End If
        Next
    Next
    'adding the beginning segment to the full edi text
    fullEdiText = BeginSegmentText
    ' adding middle segment of each edi record in loop to full edi text
    txtItemCt = 0
    For Each txtItem In allEdiTextDict.items
        ' below condition is to add "~" while concatenating segments except the first iteration
        ' as "~" is getting remmoved during splitting of edi text
        If txtItemCt = 0 Then
            fullEdiText = fullEdiText + txtItem
        Else
            fullEdiText = fullEdiText + "~" + txtItem
        End If
        txtItemCt = txtItemCt + 1
    Next
    LogStatement "Loop completed to concatenate the middle segment texts"
    ' getting the total number of segments in the full edi text
    totalNoOfSegments = UBound(Split(fullEdiText, "~"))
    LogStatement "totalNoOfSegments " & totalNoOfSegments
    EndSegmentText = Replace(EndSegmentText, "CALC:NO_OF_SEGMENTS", CStr(totalNoOfSegments))
    LogStatement "Final EndSegmentText created"
    'adding the ending segment to the full edi text
    fullEdiText = Trim(fullEdiText + EndSegmentText)
    fullEdiText = UCase(fullEdiText)
    'removing any extra spaces or newlines that might be present at the end of the text
    Do Until Right(fullEdiText, 1) = "~"
        fullEdiText = Left(fullEdiText, Len(fullEdiText) - 1)
    Loop
    'finalOutFilePath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Output\MH\MH_Scenario 1.txt"
    'finalOutFilePath = outFilePath = rootPath & Chr(92) & outputFolder & Chr(92) & lineOfBusiness & Chr(92) & lineOfBusiness & "_" & scenarioName & ".txt"
    CreateOutputTextFile finalOutFilePath, fullEdiText, "N"
    LogStatement "fullEdiText written to output file"
    LogStatement "Calling function InsertDictValuesIntoDb from Main function"
    ' inserting values used in generating edi file to database for use in other scenarios for same member
    If Split(scenarioName, "_", 2)(1) = "Scenario_1" Or InStr(1, scenarioName, "Scenario_7") <> 0 Or InStr(1, scenarioName, "Scenario_8") <> 0 Or InStr(1, scenarioName, "Scenario_16") <> 0 Then
        'MsgBox "Enable after testing"
        Call InsertDictValuesIntoDb_1(updateDbPath, lineOfBusiness)
    Else
        'MsgBox "Enable after testing"
        Call InsertDictValuesIntoDb_2(updateDbPath, lineOfBusiness)
    End If
'    For Each dVal In segmentLoopDict.items
'        'msgbox dVal
'    Next
'
    objTemplateWorkbook.Close
    objExcel.Quit
    Set objExcel = Nothing
    Set objTemplateWorkbook = Nothing
    Set objTemplateWorksheet = Nothing
    Set allEdiTextDict = Nothing
    LogStatement "Completed"
    'msgbox "Completed"
End Sub

Function getIsaValues()
    If lineOfBusiness = "MH" Then
        isa08 = "110025617D"
    ElseIf lineOfBusiness = "BACO" Then
        isa08 = "110104314B"
    ElseIf lineOfBusiness = "MERCY" Then
        isa08 = "110104314C"
    ElseIf lineOfBusiness = "SIGNA" Then
        isa08 = "110104314D"
    ElseIf lineOfBusiness = "SOUTHCOAST" Then
        isa08 = "110104314E"
    ElseIf lineOfBusiness = "NH" Then
        isa08 = "NH100832"
    ElseIf lineOfBusiness = "CCA" Then
        isa08 = "043373331"
    ElseIf lineOfBusiness = "HSA" Then
        isa08 = "043373331"
    ElseIf lineOfBusiness = "DCH" Then
        isa08 = "043373331"
    ElseIf lineOfBusiness = "SCO" Then
        isa08 = "110025617H"
    End If

End Function

Function getPCPValue(req_loop_segment)
    pcp_value = ""
    If lineOfBusiness = "MH" Then
'        If req_loop_segment = "2000-DTP03" Then
'            pcp_value = "20200701"
'
'        ElseIf req_loop_segment = "2300-REF02" Then
'            pcp_value = "110004164A"
'        ElseIf req_loop_segment = "2310-NM109" Then
'            pcp_value = "1396740551"
'        End If
        Debug.Print "values to be provided to us"
        
        
    ElseIf lineOfBusiness = "BACO" Then
        If req_loop_segment = "2000-DTP03" Then
            pcp_value = "20200701"
        
        ElseIf req_loop_segment = "2300-REF02" Then
            pcp_value = "110004164A"
        ElseIf req_loop_segment = "2310-NM109" Then
            pcp_value = "1396740551"
        End If
    ElseIf lineOfBusiness = "MERCY" Then
        If req_loop_segment = "2000-DTP03" Then
            pcp_value = "20200701"
        
        ElseIf req_loop_segment = "2300-REF02" Then
            pcp_value = "110004164A"
        ElseIf req_loop_segment = "2310-NM109" Then
            pcp_value = "1396740551"
        End If
    ElseIf lineOfBusiness = "SIGNA" Then
        If req_loop_segment = "2000-DTP03" Then
            pcp_value = "20200807"
        
        ElseIf req_loop_segment = "2300-REF02" Then
            pcp_value = "110058993A"
        ElseIf req_loop_segment = "2310-NM109" Then
            pcp_value = "1881879591"
        End If
    ElseIf lineOfBusiness = "SOUTHCOAST" Then
        If req_loop_segment = "2000-DTP03" Then
            pcp_value = "20180710"
        
        ElseIf req_loop_segment = "2300-REF02" Then
            pcp_value = "110054688A"
        ElseIf req_loop_segment = "2310-NM109" Then
            pcp_value = "1255328365"
        End If
    ElseIf lineOfBusiness = "NH" Then
        If req_loop_segment = "2300-DTP03" Then
            pcp_value = "20190901"
        
        ElseIf req_loop_segment = "2300-REF02" Then
            pcp_value = ""
        
        ElseIf req_loop_segment = "2310-NM103" Then
            pcp_value = ""
        ElseIf req_loop_segment = "2310-NM104" Then
            pcp_value = ""
        ElseIf req_loop_segment = "2310-NM109" Then
            pcp_value = "1588017446"
        End If
    ElseIf lineOfBusiness = "CCA" Then
        If req_loop_segment = "2300-DTP03" Then
            pcp_value = "20190401"
        
        ElseIf req_loop_segment = "2300-REF02" Then
            pcp_value = ""
        ElseIf req_loop_segment = "2310-NM103" Then
            pcp_value = "ANNE"
        ElseIf req_loop_segment = "2310-NM104" Then
            pcp_value = "TREADUP"
        ElseIf req_loop_segment = "2310-NM109" Then
            pcp_value = "1568438513"
        End If
    ElseIf lineOfBusiness = "HSA" Then
        If req_loop_segment = "2300-DTP03" Then
            pcp_value = "20190401"
        
        ElseIf req_loop_segment = "2300-REF02" Then
            pcp_value = ""
        ElseIf req_loop_segment = "2310-NM103" Then
            pcp_value = "ANNE"
        ElseIf req_loop_segment = "2310-NM104" Then
            pcp_value = "TREADUP"
        ElseIf req_loop_segment = "2310-NM109" Then
            pcp_value = "1568438513"
        End If
    End If
    
    getPCPValue = pcp_value

End Function

Function getMemberIdLength()
    If lineOfBusiness = "HSA" Or lineOfBusiness = "NH" Then
        memberIdLength = 11
    ElseIf lineOfBusiness = "DCH" Then
        memberIdLength = 6
    Else
        memberIdLength = 12
    End If

End Function

Function getscenarioName(ScenName)
    
    If lineOfBusiness = "HSA" Or lineOfBusiness = "NH" Or lineOfBusiness = "CCA" Or lineOfBusiness = "DCH" Then
        scenarioName = lineOfBusiness & "_Scenario_" & ScenName
    Else
        scenarioName = "MH" & "_Scenario_" & ScenName
    End If



End Function

Function ReadConfigFile(rootPath)
'rootPath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project"
    configFolder = "Configs"
    configFileName = "Config - " & lineOfBusiness & ".txt"
    configFile = rootPath & Chr(92) & configFolder & Chr(92) & configFileName
    'wscript.echo "configFile " & configFile
    configText = ReadEdiTextFile(configFile)
    'wscript.echo configText
    configTextArr = Split(configText, vbNewLine)
    For Each c_t In configTextArr
        If InStr(1, c_t, "=") <> 0 Then
            c_t_arr = Split(c_t, "=")
            c_t_var = Trim(c_t_arr(0))
            c_t_val = Trim(c_t_arr(1))
            If c_t_var = "rootPath" Then
                rootPath = Trim(CStr(c_t_val))
            ElseIf c_t_var = "templateFileName" Then
                templateFileName = Trim(CStr(c_t_val))
            'ElseIf c_t_var = "lineOfBusiness" Then
            '    lineOfBusiness = Trim(CStr(c_t_val))
            ElseIf c_t_var = "scenarioName" Then
                scenarioName = Trim(CStr(c_t_val))
            ElseIf c_t_var = "numberOfRecordsNeeded" Then
                numberOfRecordsNeeded = CInt(Trim(CStr(c_t_val)))
            ElseIf c_t_var = "val_separator" Then
                val_separator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randIndicator" Then
                randIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randChooseIndicator" Then
                randChooseIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "combIndicator" Then
                combIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "refIndicator" Then
                refIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randNumIndicator" Then
                randNumIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randDateIndicator" Then
                randDateIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randTimeIndicator" Then
                randTimeIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randFNIndicator" Then
                randFNIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randLNIndicator" Then
                randLNIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randMIIndicator" Then
                randMIIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randSSNIndicator" Then
                randSSNIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randDOBIndicator" Then
                randDOBIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randADDR1Indicator" Then
                randADDR1Indicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "randADDR2Indicator" Then
                randADDR2Indicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "template_folder" Then
                template_folder = Trim(CStr(c_t_val))
            ElseIf c_t_var = "supportingFilesFolder" Then
                supportingFilesFolder = Trim(CStr(c_t_val))
            ElseIf c_t_var = "outputFolder" Then
                outputFolder = Trim(CStr(c_t_val))
            ElseIf c_t_var = "logFolder" Then
                logFolder = Trim(CStr(c_t_val))
            ElseIf c_t_var = "endLoopNameIndicator" Then
                endLoopNameIndicator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "beginSegmentSeparator" Then
                beginSegmentSeparator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "endSegmentSeparator" Then
                endSegmentSeparator = Trim(CStr(c_t_val))
            ElseIf c_t_var = "isa06" Then
                isa06 = Trim(CStr(c_t_val))
            ElseIf c_t_var = "isa08" Then
                isa08 = Trim(CStr(c_t_val))
            ElseIf c_t_var = "memberIdLength" Then
                memberIdLength = CInt(Trim(CStr(c_t_val)))
            End If
        End If
    Next
    '''''''''deug.print "Done"
End Function
Function LogStatement(statement)
    On Error Resume Next
        Set logfile = FSO.OpenTextFile(logFilePath, 8, True) ' 8 for appending to file
        logfile.WriteLine statement
        logfile.Close
    On Error GoTo 0
End Function
Function CreateOutputTextFile(out_FilePath, outText, formatted)
    'values for testing
'    outFilePath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Output\MH\MH_Scenario 1_260521132638.txt"
'    outText = "Hello"
    LogStatement "starting function CreateOutputTextFile  with argument " & out_FilePath & " - " & formatted
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Set objFile = FSO.CreateTextFile(out_FilePath, True)
    If formatted = "N" Then
        objFile.Write outText
    Else
        outTextArr = Split(outText, "~")
        For o_arr_ct = 0 To UBound(outTextArr)
            If o_arr_ct <> UBound(outTextArr) Then
                objFile.WriteLine outTextArr(o_arr_ct) & "~"
            Else
                objFile.WriteLine outTextArr(o_arr_ct)
            End If
        Next
    End If
    objFile.Close
    LogStatement "Function completed"
End Function
Function create_segment(objTemplateWorksheet, objTemplateWorksheetStartRow, objTemplateWorksheetEndRow, valuePlace)
    'objTemplateWorksheet - Template Scenario worksheet reference
    'objTemplateWorksheetStartRow -  row from which segment needs to be parsed for creating segment edi text
    'objTemplateWorksheetEndRow -  row till which segment needs to be parsed for creating segment edi text
    'valuePlace - integer that will indicate index of the value to be used for edi text when multiple values are present for the segment
    Dim tmplSegmentValuearr
    LogStatement "starting function create_segment with argument valuePlace - " & valuePlace
    Dim tmplLoop, tmplSubLoop, tmplSegmentName, tmplRequirement, tmplLoopCount, tmplSegmentValue, tmplValueLength
    Dim SegmentName
    Dim beginning_segment, beginning_segment_text
    ' variable to hold the entire beginning segment text that will be returned as output by the function
    beginning_segment_text = ""
    ' variable to hold one segment text
    segment_text = ""
    ' variable that will be crosschecked to identify the change in segments as we loop through the excel rows
    startSegmentName = ""
    For Row = objTemplateWorksheetStartRow To objTemplateWorksheetEndRow
        LogStatement "Starting to loop row " & Row
        ' checking if the value in first column indicates a beginning segment
        tmplSegmentLoop = Trim(objTemplateWorksheet.Cells(Row, 1))
        tmplSegmentLoopName = Trim(objTemplateWorksheet.Cells(Row, 3))
        tmplSegmentName = Trim(objTemplateWorksheet.Cells(Row, 4))
        tmplRequirement = Trim(objTemplateWorksheet.Cells(Row, 5))
        loopSegmentKey = tmplSegmentLoopName & "-" & tmplSegmentName
        ' moving to the next row if the current row segment is not to be included in the edi text,
        ' indicated by absence of "Y" in column
        ' proceeding to include the segment value only if "Y"
        If tmplRequirement = "Y" Then
            tmplValueLength = objTemplateWorksheet.Cells(Row, 8)
            tmplLoopCount = objTemplateWorksheet.Cells(Row, 6)
            tmplSegmentValue = Trim(objTemplateWorksheet.Cells(Row, 7))
            ' code to update segment values based on values passed in config values
'            If tmplSegmentName = "ISA06" Then
'                tmplSegmentValue = isa06
'            End If
            If tmplSegmentName = "ISA08" Then
                tmplSegmentValue = isa08
            End If
            ' checking if multiple values are present for a segment and fetching the required value using the valuePlace argument passed
            ' valuePlace will be an integer and it will indicate the index position in the split values array
            If InStr(1, tmplSegmentValue, val_separator) <> 0 Then
                tmplSegmentValue = Split(tmplSegmentValue, val_separator)(valuePlace - 1)
            End If
            ' Numbers that must be prefixed with zeros will be placed inside Single quotes in the EDI template file, the below code will remove those quotes
            If InStr(1, tmplSegmentValue, Chr(39)) <> 0 Then
                tmplSegmentValue = CStr(Replace(tmplSegmentValue, Chr(39), ""))
            End If
            ' working on fetching random values based on the information from template
            If InStr(1, Left(tmplSegmentValue, 5), randIndicator) <> 0 Then
                    ' example - RAND:MEME_BIRTH_DT or RAND:MEME_BIRTH_DT~YYYYMMDD
                    ' checking if "|" is present to get the second part of the segment value
                    ' example - "RAND:MEME_BIRTH_DT|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
                    ' tmplSegmentValue_suffix will be set to -22991231
                    tmplSegmentValue_suffix = ""
                    If InStr(1, tmplSegmentValue, "|") <> 0 Then
                        tmplSegmentValue_suffix = Split(tmplSegmentValue, "|")(1)
                        tmplSegmentValue = Split(tmplSegmentValue, "|")(0)
                    End If
                    format_type = ""
                    If InStr(1, tmplSegmentValue, "~") <> 0 Then
                        format_type = Split(tmplSegmentValue, "~")(1)
                        tmplSegmentValue = Split(tmplSegmentValue, "~")(0)
                    End If
                    tmplSegmentValue = CreateRandomValue(tmplSegmentValue, valuePlace)
                    tmplSegmentValue = Trim(tmplSegmentValue)
                    If Len(format_type) > 0 Then
                        If InStr(1, format_type, "YY") <> 0 Or InStr(1, format_type, "DD") <> 0 Then
                            tmplSegmentValue = GetFormattedDate(tmplSegmentValue, format_type)
                        ElseIf InStr(1, format_type, "HH") <> 0 Then
                            tmplSegmentValue = GetFormattedTime(tmplSegmentValue, format_type)
                        End If
                    End If
                    If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
                        tmplSegmentValue = tmplSegmentValue & tmplSegmentValue_suffix
                    End If
            End If
            ' working on choosing random values from a given set of values
            If InStr(1, Left(tmplSegmentValue, 4), randChooseIndicator) <> 0 Then
                    tmplSegmentValue = ChooseRandomValue(tmplSegmentValue)
                    tmplSegmentValue = Trim(tmplSegmentValue)
            End If
            If InStr(1, Left(tmplSegmentValue, 4), refIndicator) <> 0 Then
                ' checking if "|" is present to get the second part of the segment value
                ' example - "REF:2000-DTP03|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
                ' tmplSegmentValue_suffix will be set to -22991231
                tmplSegmentValue_suffix = ""
                If InStr(1, tmplSegmentValue, "|") <> 0 Then
                    tmplSegmentValue_suffix = Split(tmplSegmentValue, "|")(1)
                End If
                tmplSegmentValue = GetReferenceSegmentValue(tmplSegmentValue, valuePlace)
                tmplSegmentValue = Trim(tmplSegmentValue)
                If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
                    tmplSegmentValue = tmplSegmentValue & tmplSegmentValue_suffix
                End If
            End If
            ' checking if the indicator that indicates creation of a combined text is present
            ' by checking if the first 5 characters is COMB:
            If InStr(1, Left(tmplSegmentValue, 5), combIndicator & ":") <> 0 Then
                tmplSegmentValue = GetCombinationValue(tmplSegmentValue, valPlace)
            End If
            
            ' getting the correct dob as per plan for MH and ACO Lobs
            If loopSegmentKey = "2300-HD04" Then
                    plan_change_lobs = "MH,BACO,MERCY,SIGNA,SOUTHCOAST"
                    
                    If InStr(1, plan_change_lobs, lineOfBusiness) <> 0 Then
                    
                        new_tmplSegment_Value = ""
                        
                        new_tmplSegment_Value = change_dob_as_per_plan(UCase(tmplSegmentValue))
        
                        
                        If new_tmplSegment_Value <> "" Then
                            dataToBeSavedToDbArr(currentRandKey)("2100A-DMG02") = new_tmplSegment_Value
                            segmentLoopDict("2100A-DMG02") = new_tmplSegment_Value
                        End If
                    End If
                
            
            End If
            
            If InStr(1, tmplSegmentValue, "USE_PCP_VALUE") <> 0 Then
                tmplSegmentValue = getPCPValue(loopSegmentKey)
            End If
            
            If InStr(1, tmplSegmentValue, "SBSB_EMAIL") <> 0 Then
                tmplSegmentValue = getRandomEmail(valuePlace)
            End If
            
            If InStr(1, tmplSegmentValue, "NUMBER_OF_LEN") <> 0 Then
                
                tmplSegmentValue = getRandomNumberOfLength(Split(tmplSegmentValue, ":")(1))
            End If
            
            ' checking if the segment value should be of specific length and calling the function to meet the requirement
            ' by adding spaces to the end of the string
            If tmplValueLength <> "" Then
                tmplValueLength = CInt(tmplValueLength)
                tmplSegmentValue = RequiredLengthString(tmplSegmentValue, tmplValueLength)
            End If
            LogStatement "Starting to add segment value to text"
            SegmentName = RootSegment(tmplSegmentName)
            If startSegmentName <> SegmentName Then
                startSegmentName = SegmentName
                If Len(beginning_segment_text) > 0 Then
                    beginning_segment_text = beginning_segment_text + "~" + SegmentName & "*" & tmplSegmentValue
                    segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                    If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                        segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                    Else
                        tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                        tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                        segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                    End If
                Else
                    beginning_segment_text = SegmentName & "*" & tmplSegmentValue
                    segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                    If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                        ''On Error Resume Next
                        segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                        ' the below code is to handle cases when script is throwing error even when
                        ' the key segmentLoopDictKey is not present in the dictionary segmentLoopDict
                        If Err.Number <> 0 Then
                            segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                        End If
                        'on error goto 0
                    Else
                        tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                        tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                        segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                    End If
                End If
            Else
                beginning_segment_text = beginning_segment_text & "*" & tmplSegmentValue
                segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                    If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                        segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                    Else
                        tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                        tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                        segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                    End If
            End If
            LogStatement "Completed to add segment value to text"

            
            ' adding the tmplSegmentValue to the dictionary that will contain values to be saved to database for use in other scenarios
            If dataToBeSavedToDbArr(currentRandKey).exists(segmentLoopDictKey) Then
                dataToBeSavedToDbArr(currentRandKey)(segmentLoopDictKey) = tmplSegmentValue
            End If
        End If
        LogStatement "Completed loop for row " & Row
    Next
    LogStatement "Function completed"
    ''msgbox beginning_segment_text
    If Len(beginning_segment_text) > 0 Then
        create_segment = beginning_segment_text + "~"
    Else
        create_segment = ""
    End If
End Function

Function create_segment_7(objTemplateWorksheet, objTemplateWorksheetStartRow, objTemplateWorksheetEndRow, valuePlace)
    'objTemplateWorksheet - Template Scenario worksheet reference
    'objTemplateWorksheetStartRow -  row from which segment needs to be parsed for creating segment edi text
    'objTemplateWorksheetEndRow -  row till which segment needs to be parsed for creating segment edi text
    'valuePlace - integer that will indicate index of the value to be used for edi text when multiple values are present for the segment
    Dim tmplSegmentValuearr
    LogStatement "starting function create_segment with argument valuePlace - " & valuePlace
    Dim tmplLoop, tmplSubLoop, tmplSegmentName, tmplRequirement, tmplLoopCount, tmplSegmentValue, tmplValueLength
    Dim SegmentName
    Dim beginning_segment, beginning_segment_text
    
    
    any_but_indicator = "ANY_BUT:"
    any_but_separator = "@@@"
    any_but_val_separator = "|"
    
    ' variable to hold the entire beginning segment text that will be returned as output by the function
    beginning_segment_text = ""
    ' variable to hold one segment text
    segment_text = ""
    ' variable that will be crosschecked to identify the change in segments as we loop through the excel rows
    startSegmentName = ""
    
    
    
    
        For Row = objTemplateWorksheetStartRow To objTemplateWorksheetEndRow
            LogStatement "Starting to loop row " & Row
            ' checking if the value in first column indicates a beginning segment
            tmplSegmentLoop = Trim(objTemplateWorksheet.Cells(Row, 1))
            tmplSegmentLoopName = Trim(objTemplateWorksheet.Cells(Row, 3))
            tmplSegmentName = Trim(objTemplateWorksheet.Cells(Row, 4))
            tmplRequirement = Trim(objTemplateWorksheet.Cells(Row, 5))
            
            
            
            If InStr(1, tmplRequirement, ",") <> 0 Then
                tmplRequirement = Split(tmplRequirement, ",")(segments_req_count - 1)
                dep_ct = segments_req_count - 1
            End If
            
            loopSegmentKey = tmplSegmentLoopName & "-" & tmplSegmentName
            
            
            
            
            ' moving to the next row if the current row segment is not to be included in the edi text,
            ' indicated by absence of "Y" in column
            ' proceeding to include the segment value only if "Y"
            If tmplRequirement = "Y" Then
                tmplValueLength = objTemplateWorksheet.Cells(Row, 8)
                tmplLoopCount = objTemplateWorksheet.Cells(Row, 6)
                tmplSegmentValue = Trim(objTemplateWorksheet.Cells(Row, 7))
                If InStr(1, tmplSegmentValue, "!@!") <> 0 Then
                    tmplSegmentValue = Split(tmplSegmentValue, "!@!")(segments_req_count - 1)
                End If
                
                
                
                
                ' code to update segment values based on values passed in config values
    '            If tmplSegmentName = "ISA06" Then
    '                tmplSegmentValue = isa06
    '            End If
                If tmplSegmentName = "ISA08" Then
                    tmplSegmentValue = isa08
                End If
                ' checking if multiple values are present for a segment and fetching the required value using the valuePlace argument passed
                ' valuePlace will be an integer and it will indicate the index position in the split values array
                If InStr(1, tmplSegmentValue, val_separator) <> 0 Then
                    tmplSegmentValue = Split(tmplSegmentValue, val_separator)(valuePlace - 1)
                End If
                ' Numbers that must be prefixed with zeros will be placed inside Single quotes in the EDI template file, the below code will remove those quotes
                If InStr(1, tmplSegmentValue, Chr(39)) <> 0 Then
                    tmplSegmentValue = CStr(Replace(tmplSegmentValue, Chr(39), ""))
                End If
                
                If InStr(1, tmplSegmentValue, "USE_MEMBER_SEGMENT_VAL") <> 0 Then
                    tmplSegmentValue = dataToBeSavedToDbArr(currentRandKey)(loopSegmentKey)
                    If InStr(1, tmplSegmentValue, ":") <> 0 Then
                        tmplSegmentValue = Trim(Str(Split(tmplSegmentValue, ":")(0)))
                        
                    End If
                    If InStr(1, tmplSegmentValue, "@@@@@") <> 0 Then
                        tmplSegmentValue = Split(tmplSegmentValue, "@@@@@")(valuePlace - 1)
                    End If
                End If
                
                
                
                ' checking if the indicator that indicates getting random value from a list of values in template is present
                If InStr(1, tmplSegmentValue, any_but_indicator) <> 0 Then
                    
                    tmplSegmentValue = GetAnyButValue(tmplSegmentValue, any_but_indicator, valPlace, any_but_separator, any_but_val_separator)
                End If
                If InStr(1, tmplSegmentValue, "NUMBER_OF_LEN") <> 0 Then
                
                    tmplSegmentValue = getRandomNumberOfLength(Split(tmplSegmentValue, ":")(1))
                End If
                
                ' working on fetching random values based on the information from template
                If InStr(1, Left(tmplSegmentValue, 5), randIndicator) <> 0 Then
                    ' below condition is for the beginning segments marked by "B" in first column of template
                    ' this is to ensure that Random data is generated only once for segment that will be part of beginning segment
                    If (tmplSegmentLoop = "B" And segments_req_count = 1) Or tmplSegmentLoop <> "B" Then
                        ' example - RAND:MEME_BIRTH_DT or RAND:MEME_BIRTH_DT~YYYYMMDD
                        ' checking if "|" is present to get the second part of the segment value
                        ' example - "RAND:MEME_BIRTH_DT|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
                        ' tmplSegmentValue_suffix will be set to -22991231
                        tmplSegmentValue_suffix = ""
                        If InStr(1, tmplSegmentValue, "|") <> 0 Then
                            tmplSegmentValue_suffix = Split(tmplSegmentValue, "|")(1)
                            tmplSegmentValue = Split(tmplSegmentValue, "|")(0)
                        End If
                        format_type = ""
                        If InStr(1, tmplSegmentValue, "~") <> 0 Then
                            format_type = Split(tmplSegmentValue, "~")(1)
                            tmplSegmentValue = Split(tmplSegmentValue, "~")(0)
                        End If
                        tmplSegmentValue = CreateRandomValue(tmplSegmentValue, valuePlace)
                        tmplSegmentValue = Trim(tmplSegmentValue)
                        If Len(format_type) > 0 Then
                            If InStr(1, format_type, "YY") <> 0 Or InStr(1, format_type, "DD") <> 0 Then
                                tmplSegmentValue = GetFormattedDate(tmplSegmentValue, format_type)
                            ElseIf InStr(1, format_type, "HH") <> 0 Then
                                tmplSegmentValue = GetFormattedTime(tmplSegmentValue, format_type)
                            End If
                        End If
                        If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
                            tmplSegmentValue = tmplSegmentValue & tmplSegmentValue_suffix
                        End If
                        
                    
                        
                    ElseIf tmplSegmentLoop = "B" And segments_req_count <> 1 Then
                        tmplSegmentValue = dataToBeSavedToDbArr(currentRandKey)(loopSegmentKey)
                        If InStr(1, tmplSegmentValue, "@@@@@") <> 0 Then
                            tmplSegmentValue = Trim(Split(tmplSegmentValue, "@@@@@")(valuePlace - 1))
                        End If
                    
                    End If
                End If
                ' working on choosing random values from a given set of values
                If InStr(1, Left(tmplSegmentValue, 4), randChooseIndicator) <> 0 Then
                        tmplSegmentValue = ChooseRandomValue(tmplSegmentValue)
                        tmplSegmentValue = Trim(tmplSegmentValue)
                End If
                If InStr(1, Left(tmplSegmentValue, 4), refIndicator) <> 0 Then
                    ' checking if "|" is present to get the second part of the segment value
                    ' example - "REF:2000-DTP03|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
                    ' tmplSegmentValue_suffix will be set to -22991231
                    tmplSegmentValue_suffix = ""
                    If InStr(1, tmplSegmentValue, "|") <> 0 Then
                        tmplSegmentValue_suffix = Split(tmplSegmentValue, "|")(1)
                    End If
                    tmplSegmentValue = GetReferenceSegmentValue(tmplSegmentValue, valuePlace)
                    tmplSegmentValue = Trim(tmplSegmentValue)
                    If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
                        tmplSegmentValue = tmplSegmentValue & tmplSegmentValue_suffix
                    End If
                End If
                ' checking if the indicator that indicates creation of a combined text is present
                ' by checking if the first 5 characters is COMB:
                If InStr(1, Left(tmplSegmentValue, 5), combIndicator & ":") <> 0 Then
                    ' below condition is for the beginning segments marked by "B" in first column of template
                    ' this is to ensure that Random data is generated only once for segment that will be part of beginning segment
                    If tmplSegmentLoop = "B" And segments_req_count = 1 Then
                        tmplSegmentValue = GetCombinationValue(tmplSegmentValue, valPlace)
                    ElseIf tmplSegmentLoop = "B" And segments_req_count <> 1 Then
                        tmplSegmentValue = dataToBeSavedToDbArr(currentRandKey)(loopSegmentKey)
                        If InStr(1, tmplSegmentValue, "@@@@@") <> 0 Then
                            tmplSegmentValue = Trim(Split(tmplSegmentValue, "@@@@@")(valuePlace - 1))
                        End If
                    Else
                        tmplSegmentValue = GetCombinationValue(tmplSegmentValue, valPlace)
                    End If
                End If
                
                If InStr(1, tmplSegmentValue, "USE_PCP_VALUE") <> 0 Then
                    tmplSegmentValue = getPCPValue(loopSegmentKey)
                End If
                
                ' checking if the segment value should be of specific length and calling the function to meet the requirement
                ' by adding spaces to the end of the string
                If tmplValueLength <> "" Then
                    tmplValueLength = CInt(tmplValueLength)
                    tmplSegmentValue = RequiredLengthString(tmplSegmentValue, tmplValueLength)
                End If
                LogStatement "Starting to add segment value to text"
                SegmentName = RootSegment(tmplSegmentName)
                If startSegmentName <> SegmentName Then
                    startSegmentName = SegmentName
                    If Len(beginning_segment_text) > 0 Then
                        beginning_segment_text = beginning_segment_text + "~" + SegmentName & "*" & tmplSegmentValue
                        segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                        If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                            segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                        Else
                            tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                            tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                            segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                        End If
                    Else
                        beginning_segment_text = SegmentName & "*" & tmplSegmentValue
                        segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                        If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                            ''On Error Resume Next
                            segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                            ' the below code is to handle cases when script is throwing error even when
                            ' the key segmentLoopDictKey is not present in the dictionary segmentLoopDict
                            If Err.Number <> 0 Then
                                segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                            End If
                            'on error goto 0
                        Else
                            tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                            tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                            segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                        End If
                    End If
                Else
                    beginning_segment_text = beginning_segment_text & "*" & tmplSegmentValue
                    segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                        If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                            segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                        Else
                            tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                            tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                            segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                        End If
                End If
                LogStatement "Completed to add segment value to text"
    
                
                ' adding the tmplSegmentValue to the dictionary that will contain values to be saved to database for use in other scenarios
                If segments_req_count = 1 Then
                    If dataToBeSavedToDbArr(currentRandKey).exists(segmentLoopDictKey) Then
                        dataToBeSavedToDbArr(currentRandKey)(segmentLoopDictKey) = tmplSegmentValue
                    End If
                Else
                    
                    If depDataToBeSavedToDbArr(currentRandKey)(segments_req_count - 1).exists(segmentLoopDictKey) Then
                        depDataToBeSavedToDbArr(currentRandKey)(segments_req_count - 1)(segmentLoopDictKey) = tmplSegmentValue
                    End If
                End If
            
            End If
            LogStatement "Completed loop for row " & Row
        Next
        LogStatement "Function completed"
        
    
    
    ''msgbox beginning_segment_text
    If Len(beginning_segment_text) > 0 Then
        create_segment_7 = beginning_segment_text + "~"
    Else
        create_segment_7 = ""
    End If
    
End Function




Function create_segment_scenario2(objTemplateWorksheet, objTemplateWorksheetStartRow, objTemplateWorksheetEndRow, valuePlace)
    'objTemplateWorksheet - Template Scenario worksheet reference
    'objTemplateWorksheetStartRow -  row from which segment needs to be parsed for creating segment edi text
    'objTemplateWorksheetEndRow -  row till which segment needs to be parsed for creating segment edi text
    'valuePlace - integer that will indicate index of the value to be used for edi text when multiple values are present for the segment
    Dim tmplSegmentValuearr
    LogStatement "starting function create_segment with argument valuePlace - " & valuePlace
    Dim tmplLoop, tmplSubLoop, tmplSegmentName, tmplRequirement, tmplLoopCount, tmplSegmentValue, tmplValueLength
    Dim SegmentName
    Dim beginning_segment, beginning_segment_text
    prev_val_indicator = "PREVIOUS_VALUE"
    current_date_indicator = "CURRENT_DATE"
    any_but_indicator = "ANY_BUT:"
    any_but_separator = "@@@"
    any_but_val_separator = "|"
    
    ' variable to hold the entire beginning segment text that will be returned as output by the function
    beginning_segment_text = ""
    ' variable to hold one segment text
    segment_text = ""
    ' variable that will be crosschecked to identify the change in segments as we loop through the excel rows
    startSegmentName = ""
    For Row = objTemplateWorksheetStartRow To objTemplateWorksheetEndRow
        LogStatement "Starting to loop row " & Row
        ' checking if the value in first column indicates a beginning segment
        tmplSegmentLoop = Trim(objTemplateWorksheet.Cells(Row, 1))
        tmplSegmentModifyValue = Trim(objTemplateWorksheet.Cells(Row, 2))
        tmplSegmentLoopName = Trim(objTemplateWorksheet.Cells(Row, 3))
        tmplSegmentName = Trim(objTemplateWorksheet.Cells(Row, 4))
        tmplRequirement = Trim(objTemplateWorksheet.Cells(Row, 5))
        loopSegmentKey = tmplSegmentLoopName & "-" & tmplSegmentName

        
        ' moving to the next row if the current row segment is not to be included in the edi text,
        ' indicated by absence of "Y" in column
        ' proceeding to include the segment value only if "Y"
        If tmplRequirement = "Y" Then
            tmplValueLength = objTemplateWorksheet.Cells(Row, 8)
            tmplLoopCount = objTemplateWorksheet.Cells(Row, 6)
            ' getting the value from storedDataDict which contains the value obtained from database having the original values used in enrolling a member
            If storedDataDict(currentRandKey).exists(loopSegmentKey) Then
                tmplSegmentValue = storedDataDict(currentRandKey)(loopSegmentKey)
                If InStr(1, tmplSegmentValue, "@@@@@") <> 0 Then
                    tmplSegmentValue = Replace(tmplSegmentValue, "@@@@@", ",")
                End If
                tmplSegmentValue_1 = Trim(objTemplateWorksheet.Cells(Row, 7))
            Else
                tmplSegmentValue = Trim(objTemplateWorksheet.Cells(Row, 7))
                tmplSegmentValue_1 = tmplSegmentValue
            End If
            ' code to update segment values based on values passed in config values
'            If tmplSegmentName = "ISA06" Then
'                tmplSegmentValue = isa06
'            End If
            If tmplSegmentName = "ISA08" Then
                tmplSegmentValue = isa08
            End If
            ' this condition will be true when the segment had only one value during enrollment but segment needs to be included more than once
            ' in other scenarios with different values
            If InStr(1, tmplSegmentValue_1, val_separator) <> 0 And InStr(1, tmplSegmentValue, val_separator) = 0 Then
                
                n_tmplSegmentValue = Split(tmplSegmentValue_1, val_separator)(valuePlace - 1)
                If n_tmplSegmentValue <> "" Then
                    tmplSegmentValue = n_tmplSegmentValue
                End If
                
            ElseIf InStr(1, tmplSegmentValue_1, val_separator) = 0 And InStr(1, tmplSegmentValue, val_separator) <> 0 And InStr(1, tmplSegmentValue_1, "USE_OLD_VAL") = 0 Then
                ' this condition helps to handle scenario where separate segments needs to be considered as one such as LX segments
                ' in such cases the tmplSegmentValue_1 contains the correct value to be placed into the segment, while tmplSegmentValue will have all values that goes into the segment separated by comma
                tmplSegmentValue = tmplSegmentValue_1
            End If
            ' checking if multiple values are present for a segment and fetching the required value using the valuePlace argument passed
            ' valuePlace will be an integer and it will indicate the index position in the split values array
            'on error goto 0
            If InStr(1, tmplSegmentValue, val_separator) <> 0 Then
                ''On Error Resume Next
                tmplSegmentValue = Split(tmplSegmentValue, val_separator)(valuePlace - 1)
                If Err.Number <> 0 Then
                    If InStr(1, tmplSegmentValue_1, val_separator) <> 0 Then
                        tmplSegmentValue = Split(tmplSegmentValue_1, val_separator)(valuePlace - 1)
                    End If
                End If
            End If
            'on error goto 0
            If InStr(1, tmplSegmentValue_1, val_separator) <> 0 Then
                tmplSegmentValue_1 = Split(tmplSegmentValue_1, val_separator)(valuePlace - 1)
            End If
            'Introduced the prefix USE_THIS_VAL_ to override the value in variable tmplSegmentValue with value from variable tmplSegmentValue_1.
            ' This is to use any hard coded values from template rather than using the value from Database
            If InStr(1, tmplSegmentValue_1, "USE_THIS_VAL_") <> 0 Then
                tmplSegmentValue = Split(tmplSegmentValue_1, "USE_THIS_VAL_")(1)
            End If
            
            If InStr(1, tmplSegmentValue_1, "CHANGE_DOB") <> 0 Then
                new_tmplSegment_Value = ""
                If storedDataDict(currentRandKey).exists("2300-HD04") Then
                    new_tmplSegment_Value = change_dob_as_per_plan(UCase(storedDataDict(currentRandKey)("2300-HD04")))

                End If
                If new_tmplSegment_Value <> "" Then
                    tmplSegmentValue = new_tmplSegment_Value
                End If
            End If
            
            If InStr(1, tmplSegmentValue, "NUMBER_OF_LEN") <> 0 Then
                
                tmplSegmentValue = getRandomNumberOfLength(Split(tmplSegmentValue, ":")(1))
            End If
            
            If InStr(1, tmplSegmentValue_1, "START_OF_MONTH") <> 0 Then
                new_tmplSegment_Value = ""
                new_tmplSegment_Value = getStartOfMonthDate()
                If new_tmplSegment_Value <> "" Then
                    tmplSegmentValue = new_tmplSegment_Value
                End If
            
            End If
'            If InStr(1, tmplSegmentValue_1, "USE_OLD_VAL") <> 0 Then
'                If storedDataDict(currentRandKey).exists(loopSegmentKey) Then
'                    tmplSegmentValue = storedDataDict(currentRandKey)(loopSegmentKey)
'                End If
'            End If
            
            ' Numbers that must be prefixed with zeros will be placed inside Single quotes in the EDI template file, the below code will remove those quotes
            If InStr(1, tmplSegmentValue, Chr(39)) <> 0 Then
                tmplSegmentValue = CStr(Replace(tmplSegmentValue, Chr(39), ""))
            End If
            ' working on fetching random values based on the information from template
            If InStr(1, Left(tmplSegmentValue, 5), randIndicator) <> 0 Then
                    ' example - RAND:MEME_BIRTH_DT or RAND:MEME_BIRTH_DT~YYYYMMDD
                    ' checking if "|" is present to get the second part of the segment value
                    ' example - "RAND:MEME_BIRTH_DT|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
                    ' tmplSegmentValue_suffix will be set to -22991231
                    tmplSegmentValue_suffix = ""
                    If InStr(1, tmplSegmentValue, "|") <> 0 Then
                        tmplSegmentValue_suffix = Split(tmplSegmentValue, "|")(1)
                        tmplSegmentValue = Split(tmplSegmentValue, "|")(0)
                    End If
                    format_type = ""
                    If InStr(1, tmplSegmentValue, "~") <> 0 Then
                        format_type = Split(tmplSegmentValue, "~")(1)
                        tmplSegmentValue = Split(tmplSegmentValue, "~")(0)
                    End If
                    tmplSegmentValue = CreateRandomValue(tmplSegmentValue, valuePlace)
                    tmplSegmentValue = Trim(tmplSegmentValue)
                    If Len(format_type) > 0 Then
                        If InStr(1, format_type, "YY") <> 0 Or InStr(1, format_type, "DD") <> 0 Then
                            tmplSegmentValue = GetFormattedDate(tmplSegmentValue, format_type)
                        ElseIf InStr(1, format_type, "HH") <> 0 Then
                            tmplSegmentValue = GetFormattedTime(tmplSegmentValue, format_type)
                        End If
                    End If
                    If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
                        tmplSegmentValue = tmplSegmentValue & tmplSegmentValue_suffix
                    End If
            End If
            ' working on choosing random values from a given set of values
            If InStr(1, Left(tmplSegmentValue, 4), randChooseIndicator) <> 0 Then
                    tmplSegmentValue = ChooseRandomValue(tmplSegmentValue)
                    tmplSegmentValue = Trim(tmplSegmentValue)
            End If
            If InStr(1, Left(tmplSegmentValue_1, 4), refIndicator) <> 0 Then
                tmplSegmentValue = tmplSegmentValue_1
                ' checking if "|" is present to get the second part of the segment value
                ' example - "REF:2000-DTP03|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
                ' tmplSegmentValue_suffix will be set to -22991231
                tmplSegmentValue_suffix = ""
                If InStr(1, tmplSegmentValue, "|") <> 0 Then
                    tmplSegmentValue_suffix = Split(tmplSegmentValue, "|")(1)
                End If
                tmplSegmentValue = GetReferenceSegmentValue(tmplSegmentValue, valuePlace)
                tmplSegmentValue = Trim(tmplSegmentValue)
                If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
                    tmplSegmentValue = tmplSegmentValue & tmplSegmentValue_suffix
                End If
            End If
            ' checking if the indicator that indicates creation of a combined text is present
            ' by checking if the first 5 characters is COMB:
            If InStr(1, Left(tmplSegmentValue, 5), combIndicator & ":") <> 0 Then
                tmplSegmentValue = GetCombinationValue(tmplSegmentValue, valPlace)
            End If
            
            ' checking if the indicator that indicates getting previous value in the same segment is present
            If InStr(1, tmplSegmentValue_1, prev_val_indicator) <> 0 Then
                tmplSegmentValue = GetPreviousValue(loopSegmentKey, tmplSegmentValue_1, valPlace)
            End If
            
            ' checking if the indicator that indicates getting previous value in the same segment is present
            If InStr(1, tmplSegmentValue_1, current_date_indicator) <> 0 And InStr(1, tmplSegmentValue_1, "(") <> 0 Then
                tmplSegmentValue = FormattedDate("YYYYMMDD", "Y", "P", 1, "d")
            ElseIf InStr(1, tmplSegmentValue_1, current_date_indicator) <> 0 And InStr(1, tmplSegmentValue_1, "(") = 0 Then
                tmplSegmentValue = FormattedDate("YYYYMMDD", "N", "", "", "")
            End If
            
            If InStr(1, tmplSegmentValue_1, "START_DATE_OF_YEAR") <> 0 And InStr(1, tmplSegmentValue_1, "RAND:") = 0 Then
                If InStr(1, tmplSegmentValue_1, "(") <> 0 Then
                    date_change_value = Split(tmplSegmentValue_1, "(")(1)
                    date_change_value = CInt(Replace(date_change_value, ")", ""))
                    tmplSegmentValue_1 = Split(tmplSegmentValue_1, "(")(0)
                    yr_start_dt = CDate("01/01/" & Year(Date))
                    tmplSegmentValue = DateAdd("d", -1, yr_start_dt)
                    tmplSegmentValue = GetFormattedDate(tmplSegmentValue, "YYYYMMDD")
                End If
            
            End If
            
            ' checking if the indicator that indicates getting random value from a list of values in template is present
            If InStr(1, tmplSegmentValue_1, any_but_indicator) <> 0 Then
                
                tmplSegmentValue = GetAnyButValue(tmplSegmentValue_1, any_but_indicator, valPlace, any_but_separator, any_but_val_separator)
            End If
            
            ' checking if the segment value needs to be modified based on value of the variable tmplSegmentModifyValue = "Y"
            If tmplSegmentModifyValue = "Y" Then
                If tmplSegmentValue = "" Then
                    'msgbox "Need to handle this block"
                Else
                    If InStr(1, tmplSegmentValue_1, "LAST_NAME") <> 0 Or InStr(1, tmplSegmentValue_1, "FIRST_NAME") <> 0 Or InStr(1, tmplSegmentValue_1, "MID_INIT") <> 0 Then
                        tmplSegmentValue = changeName(tmplSegmentValue)
                    ElseIf InStr(1, tmplSegmentValue_1, "ADDR1") <> 0 Then
                        tmplSegmentValue = changeAddr1(tmplSegmentValue)
                    ElseIf InStr(1, tmplSegmentValue_1, "_SSN") <> 0 Then
                        tmplSegmentValue = changeSSN(tmplSegmentValue)
                    ElseIf InStr(1, tmplSegmentValue_1, "_DT") <> 0 Then
                        tmplSegmentValue = changeDOB(tmplSegmentValue)
                    Else
                        ' this will take value from template and use it for current EDI file creation rather than using value from stored database
                        tmplSegmentValue = tmplSegmentValue_1
                    End If
                End If
            End If
            If InStr(1, tmplSegmentValue_1, "MEME_MEDCD_NO") <> 0 Then
                
                MEME_MEDCD_NO = tmplSegmentValue
            End If
            
                        
            ' checking if the segment value should be of specific length and calling the function to meet the requirement
            ' by adding spaces to the end of the string
            If tmplValueLength <> "" Then
                tmplValueLength = CInt(tmplValueLength)
                tmplSegmentValue = RequiredLengthString(tmplSegmentValue, tmplValueLength)
            End If
            LogStatement "Starting to add segment value to text"
            SegmentName = RootSegment(tmplSegmentName)
            If startSegmentName <> SegmentName Then
                startSegmentName = SegmentName
                If Len(beginning_segment_text) > 0 Then
                    beginning_segment_text = beginning_segment_text + "~" + SegmentName & "*" & tmplSegmentValue
                    segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                    
                    If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                    ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                        End If
                    Else
                        tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                        tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                        ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                        End If
                    End If
                Else
                    beginning_segment_text = SegmentName & "*" & tmplSegmentValue
                    segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                    If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                    ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                        End If
                    Else
                        tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                        tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                        ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                        End If
                    End If
                End If
            Else
                beginning_segment_text = beginning_segment_text & "*" & tmplSegmentValue
                segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                    If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                    ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                        End If
                    Else
                        tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                        tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                        ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                        End If
                    End If
            End If
            LogStatement "Completed to add segment value to text"
            ' adding the tmplSegmentValue to the dictionary that will contain values to be saved to database for use in other scenarios
            If dataToBeSavedToDbArr(currentRandKey).exists(segmentLoopDictKey) Then
                ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                If tmplSegmentLoop <> "N" Then
                    dataToBeSavedToDbArr(currentRandKey)(segmentLoopDictKey) = tmplSegmentValue
                End If
            End If
        End If
        LogStatement "Completed loop for row " & Row
    Next
    LogStatement "Function completed"
    ''msgbox beginning_segment_text
    If Len(beginning_segment_text) > 0 Then
        create_segment_scenario2 = beginning_segment_text + "~"
    Else
        create_segment_scenario2 = ""
    End If
End Function



Function create_segment_scenario9(objTemplateWorksheet, objTemplateWorksheetStartRow, objTemplateWorksheetEndRow, valuePlace)
    'objTemplateWorksheet - Template Scenario worksheet reference
    'objTemplateWorksheetStartRow -  row from which segment needs to be parsed for creating segment edi text
    'objTemplateWorksheetEndRow -  row till which segment needs to be parsed for creating segment edi text
    'valuePlace - integer that will indicate index of the value to be used for edi text when multiple values are present for the segment
    Dim tmplSegmentValuearr
    LogStatement "starting function create_segment with argument valuePlace - " & valuePlace
    Dim tmplLoop, tmplSubLoop, tmplSegmentName, tmplRequirement, tmplLoopCount, tmplSegmentValue, tmplValueLength
    Dim SegmentName
    Dim beginning_segment, beginning_segment_text
    prev_val_indicator = "PREVIOUS_VALUE"
    current_date_indicator = "CURRENT_DATE"
    any_but_indicator = "ANY_BUT:"
    any_but_separator = "@@@"
    any_but_val_separator = "|"
    
    ' variable to hold the entire beginning segment text that will be returned as output by the function
    beginning_segment_text = ""
    ' variable to hold one segment text
    segment_text = ""
    ' variable that will be crosschecked to identify the change in segments as we loop through the excel rows
    startSegmentName = ""
    For Row = objTemplateWorksheetStartRow To objTemplateWorksheetEndRow
        LogStatement "Starting to loop row " & Row
        ' checking if the value in first column indicates a beginning segment
        tmplSegmentLoop = Trim(objTemplateWorksheet.Cells(Row, 1))
        tmplSegmentModifyValue = Trim(objTemplateWorksheet.Cells(Row, 2))
        tmplSegmentLoopName = Trim(objTemplateWorksheet.Cells(Row, 3))
        tmplSegmentName = Trim(objTemplateWorksheet.Cells(Row, 4))
        tmplRequirement = Trim(objTemplateWorksheet.Cells(Row, 5))
        loopSegmentKey = tmplSegmentLoopName & "-" & tmplSegmentName
        
        If InStr(1, loopSegmentKey, "2100A-N301") <> 0 And segments_req_count <> 1 Then
            Debug.Print loopSegmentKey
        End If
        
        If InStr(1, tmplRequirement, ",") <> 0 Then
            tmplRequirement = Split(tmplRequirement, ",")(segments_req_count - 1)
            
            dep_ct = segments_req_count - 1
        End If
        
        ' moving to the next row if the current row segment is not to be included in the edi text,
        ' indicated by absence of "Y" in column
        ' proceeding to include the segment value only if "Y"
        If tmplRequirement = "Y" Then
            tmplValueLength = objTemplateWorksheet.Cells(Row, 8)
            tmplLoopCount = objTemplateWorksheet.Cells(Row, 6)
            ' getting the value from storedDataDict which contains the value obtained from database having the original values used in enrolling a member
            
            If segments_req_count = 1 Then
            
                If storedDataDict(currentRandKey).exists(loopSegmentKey) Then
                    tmplSegmentValue = storedDataDict(currentRandKey)(loopSegmentKey)
                    If InStr(1, tmplSegmentValue, "@@@@@") <> 0 Then
                        tmplSegmentValue = Replace(tmplSegmentValue, "@@@@@", ",")
                    End If
                    tmplSegmentValue_1 = Trim(objTemplateWorksheet.Cells(Row, 7))
                Else
                    tmplSegmentValue = Trim(objTemplateWorksheet.Cells(Row, 7))
                    tmplSegmentValue_1 = tmplSegmentValue
                End If
            Else
                
                If dependantRandomDataDict(currentRandKey)(segments_req_count - 1).exists(loopSegmentKey) Then
                    tmplSegmentValue = dependantRandomDataDict(currentRandKey)(segments_req_count - 1)(loopSegmentKey)
                    If InStr(1, tmplSegmentValue, "@@@@@") <> 0 Then
                        tmplSegmentValue = Replace(tmplSegmentValue, "@@@@@", ",")
                    End If
                    tmplSegmentValue_1 = Trim(objTemplateWorksheet.Cells(Row, 7))
                Else
                    tmplSegmentValue = Trim(objTemplateWorksheet.Cells(Row, 7))
                    tmplSegmentValue_1 = tmplSegmentValue
                End If
                
                
            End If
            If InStr(1, tmplSegmentValue, "!@!") <> 0 Then
                tmplSegmentValue = Split(tmplSegmentValue, "!@!")(segments_req_count - 1)
            End If
            
            If InStr(1, tmplSegmentValue_1, "!@!") <> 0 Then
                tmplSegmentValue_1 = Split(tmplSegmentValue_1, "!@!")(segments_req_count - 1)
            End If
            
            ' code to update segment values based on values passed in config values
'            If tmplSegmentName = "ISA06" Then
'                tmplSegmentValue = isa06
'            End If
            If tmplSegmentName = "ISA08" Then
                tmplSegmentValue = isa08
            End If
            ' this condition will be true when the segment had only one value during enrollment but segment needs to be included more than once
            ' in other scenarios with different values
            If InStr(1, tmplSegmentValue_1, val_separator) <> 0 And InStr(1, tmplSegmentValue, val_separator) = 0 And InStr(1, tmplSegmentValue_1, "ANY:") = 0 Then
                
                n_tmplSegmentValue = Split(tmplSegmentValue_1, val_separator)(valuePlace - 1)
                If n_tmplSegmentValue <> "" Then
                    tmplSegmentValue = n_tmplSegmentValue
                End If
                
            ElseIf InStr(1, tmplSegmentValue_1, val_separator) = 0 And InStr(1, tmplSegmentValue, val_separator) <> 0 And InStr(1, tmplSegmentValue_1, "USE_OLD_VAL") = 0 Then
                ' this condition helps to handle scenario where separate segments needs to be considered as one such as LX segments
                ' in such cases the tmplSegmentValue_1 contains the correct value to be placed into the segment, while tmplSegmentValue will have all values that goes into the segment separated by comma
                tmplSegmentValue = tmplSegmentValue_1
            End If
            ' checking if multiple values are present for a segment and fetching the required value using the valuePlace argument passed
            ' valuePlace will be an integer and it will indicate the index position in the split values array
            'on error goto 0
            If InStr(1, tmplSegmentValue, val_separator) <> 0 Then
                ''On Error Resume Next
                tmplSegmentValue = Split(tmplSegmentValue, val_separator)(valuePlace - 1)
                If Err.Number <> 0 Then
                    If InStr(1, tmplSegmentValue_1, val_separator) <> 0 Then
                        tmplSegmentValue = Split(tmplSegmentValue_1, val_separator)(valuePlace - 1)
                    End If
                End If
            End If
            'on error goto 0
            If InStr(1, tmplSegmentValue_1, val_separator) <> 0 Then
                tmplSegmentValue_1 = Split(tmplSegmentValue_1, val_separator)(valuePlace - 1)
            End If
            'Introduced the prefix USE_THIS_VAL_ to override the value in variable tmplSegmentValue with value from variable tmplSegmentValue_1.
            ' This is to use any hard coded values from template rather than using the value from Database
            If InStr(1, tmplSegmentValue_1, "USE_THIS_VAL_") <> 0 Then
                tmplSegmentValue = Split(tmplSegmentValue_1, "USE_THIS_VAL_")(1)
            End If

            
            If InStr(1, tmplSegmentValue_1, "START_DATE_OF_YEAR") <> 0 And InStr(1, tmplSegmentValue_1, "RAND:") = 0 Then
                If InStr(1, tmplSegmentValue_1, "(") <> 0 Then
                    date_change_value = Split(tmplSegmentValue_1, "(")(1)
                    date_change_value = CInt(Replace(date_change_value, ")", ""))
                    tmplSegmentValue_1 = Split(tmplSegmentValue_1, "(")(0)
                    yr_start_dt = CDate("01/01/" & Year(Date))
                    tmplSegmentValue = DateAdd("d", -1, yr_start_dt)
                    tmplSegmentValue = GetFormattedDate(tmplSegmentValue, "YYYYMMDD")
                End If
            
            End If
            
            If InStr(1, tmplSegmentValue_1, "USE_MEMBER_SEGMENT_VAL") <> 0 Then
                tmplSegmentValue = dataToBeSavedToDbArr(currentRandKey)(loopSegmentKey)
                If InStr(1, tmplSegmentValue, ":") <> 0 Then
                    tmplSegmentValue = Trim(Str(Split(tmplSegmentValue, ":")(0)))
                    
                End If
                If InStr(1, tmplSegmentValue, "@@@@@") <> 0 Then
                    tmplSegmentValue = Split(tmplSegmentValue, "@@@@@")(valuePlace - 1)
                End If
            End If
            
            If InStr(1, tmplSegmentValue_1, "START_OF_MONTH") <> 0 Then
                new_tmplSegment_Value = ""
                new_tmplSegment_Value = getStartOfMonthDate()
                If new_tmplSegment_Value <> "" Then
                    tmplSegmentValue = new_tmplSegment_Value
                End If
            
            End If
            
            If InStr(1, tmplSegmentValue, "NUMBER_OF_LEN") <> 0 Then
                
                tmplSegmentValue = getRandomNumberOfLength(Split(tmplSegmentValue, ":")(1))
            End If
            
            ' Numbers that must be prefixed with zeros will be placed inside Single quotes in the EDI template file, the below code will remove those quotes
            If InStr(1, tmplSegmentValue, Chr(39)) <> 0 Then
                tmplSegmentValue = CStr(Replace(tmplSegmentValue, Chr(39), ""))
            End If
            ' working on fetching random values based on the information from template
            If InStr(1, Left(tmplSegmentValue, 5), randIndicator) <> 0 Then
                    ' example - RAND:MEME_BIRTH_DT or RAND:MEME_BIRTH_DT~YYYYMMDD
                    ' checking if "|" is present to get the second part of the segment value
                    ' example - "RAND:MEME_BIRTH_DT|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
                    ' tmplSegmentValue_suffix will be set to -22991231
                    tmplSegmentValue_suffix = ""
                    If InStr(1, tmplSegmentValue, "|") <> 0 Then
                        tmplSegmentValue_suffix = Split(tmplSegmentValue, "|")(1)
                        tmplSegmentValue = Split(tmplSegmentValue, "|")(0)
                    End If
                    format_type = ""
                    If InStr(1, tmplSegmentValue, "~") <> 0 Then
                        format_type = Split(tmplSegmentValue, "~")(1)
                        tmplSegmentValue = Split(tmplSegmentValue, "~")(0)
                    End If
                    tmplSegmentValue = CreateRandomValue(tmplSegmentValue, valuePlace)
                    tmplSegmentValue = Trim(tmplSegmentValue)
                    If Len(format_type) > 0 Then
                        If InStr(1, format_type, "YY") <> 0 Or InStr(1, format_type, "DD") <> 0 Then
                            tmplSegmentValue = GetFormattedDate(tmplSegmentValue, format_type)
                        ElseIf InStr(1, format_type, "HH") <> 0 Then
                            tmplSegmentValue = GetFormattedTime(tmplSegmentValue, format_type)
                        End If
                    End If
                    If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
                        tmplSegmentValue = tmplSegmentValue & tmplSegmentValue_suffix
                    End If
            End If
            ' working on choosing random values from a given set of values
            If InStr(1, Left(tmplSegmentValue_1, 4), randChooseIndicator) <> 0 And tmplSegmentModifyValue = "Y" Then
                    tmplSegmentValue = ChooseRandomValue(tmplSegmentValue_1)
                    tmplSegmentValue = Trim(tmplSegmentValue)
            End If
            If InStr(1, Left(tmplSegmentValue_1, 4), refIndicator) <> 0 Then
                tmplSegmentValue = tmplSegmentValue_1
                ' checking if "|" is present to get the second part of the segment value
                ' example - "REF:2000-DTP03|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
                ' tmplSegmentValue_suffix will be set to -22991231
                tmplSegmentValue_suffix = ""
                If InStr(1, tmplSegmentValue, "|") <> 0 Then
                    tmplSegmentValue_suffix = Split(tmplSegmentValue, "|")(1)
                End If
                tmplSegmentValue = GetReferenceSegmentValue(tmplSegmentValue, valuePlace)
                tmplSegmentValue = Trim(tmplSegmentValue)
                If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
                    tmplSegmentValue = tmplSegmentValue & tmplSegmentValue_suffix
                End If
            End If
            ' checking if the indicator that indicates creation of a combined text is present
            ' by checking if the first 5 characters is COMB:
            If InStr(1, Left(tmplSegmentValue, 5), combIndicator & ":") <> 0 Then
                tmplSegmentValue = GetCombinationValue(tmplSegmentValue, valPlace)
            End If
            
            ' checking if the indicator that indicates getting previous value in the same segment is present
            If InStr(1, tmplSegmentValue_1, prev_val_indicator) <> 0 Then
                tmplSegmentValue = GetPreviousValue(loopSegmentKey, tmplSegmentValue_1, valPlace)
            End If
            
            ' checking if the indicator that indicates getting previous value in the same segment is present
            If InStr(1, tmplSegmentValue_1, current_date_indicator) <> 0 And InStr(1, tmplSegmentValue_1, "(") <> 0 Then
                tmplSegmentValue = FormattedDate("YYYYMMDD", "Y", "P", 1, "d")
            ElseIf InStr(1, tmplSegmentValue_1, current_date_indicator) <> 0 And InStr(1, tmplSegmentValue_1, "(") = 0 Then
                tmplSegmentValue = FormattedDate("YYYYMMDD", "N", "", "", "")
            End If
            
            
            ' checking if the segment value needs to be modified based on value of the variable tmplSegmentModifyValue = "Y"
            If tmplSegmentModifyValue = "Y" Then
                If tmplSegmentValue = "" Then
                    'msgbox "Need to handle this block"
                Else
                    
                    If InStr(1, tmplSegmentValue_1, "LAST_NAME") <> 0 Or InStr(1, tmplSegmentValue_1, "FIRST_NAME") <> 0 Or InStr(1, tmplSegmentValue_1, "MID_INIT") <> 0 Then
                        tmplSegmentValue = changeName(tmplSegmentValue)
                        
                    ElseIf InStr(1, tmplSegmentValue_1, "ADDR1") <> 0 Then
                        If segments_req_count = 1 Then
                            ' this condition is to make sure that change to demo data is done only for policy holder as the dependants should have the same information as policy holder for the given demographics
                            tmplSegmentValue = changeAddr1(tmplSegmentValue)
                        Else
                            ' getting the value of the segment for the policy holder
                            tmplSegmentValue = dataToBeSavedToDbArr(currentRandKey)(segmentLoopDictKey)
                        
                        End If
                    ElseIf InStr(1, tmplSegmentValue_1, "_SSN") <> 0 Then
                        tmplSegmentValue = changeSSN(tmplSegmentValue)
                    ElseIf InStr(1, tmplSegmentValue_1, "_DT") <> 0 Then
                        tmplSegmentValue = changeDOB(tmplSegmentValue)
                    Else
                        If InStr(1, tmplSegmentValue_1, "USE_MEMBER_SEGMENT_VAL") = 0 Then
                        ' this will take value from template and use it for current EDI file creation rather than using value from stored database
                            tmplSegmentValue = tmplSegmentValue_1
                        End If
                    End If
                End If
            End If
            If InStr(1, tmplSegmentValue_1, "MEME_MEDCD_NO") <> 0 Then
                If segments_req_count = 1 Then ' saving only the policy holders policy number as the MEME_MEDCD_NO value, this variable is used in function that inserts segment data into database.
                    MEME_MEDCD_NO = tmplSegmentValue
                End If
            End If
            
                        
            ' checking if the segment value should be of specific length and calling the function to meet the requirement
            ' by adding spaces to the end of the string
            If tmplValueLength <> "" Then
                tmplValueLength = CInt(tmplValueLength)
                tmplSegmentValue = RequiredLengthString(tmplSegmentValue, tmplValueLength)
            End If
            LogStatement "Starting to add segment value to text"
            SegmentName = RootSegment(tmplSegmentName)
            If startSegmentName <> SegmentName Then
                startSegmentName = SegmentName
                If Len(beginning_segment_text) > 0 Then
                    beginning_segment_text = beginning_segment_text + "~" + SegmentName & "*" & tmplSegmentValue
                    segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                    
                    If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                    ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                        End If
                    Else
                        tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                        tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                        ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                        End If
                    End If
                Else
                    beginning_segment_text = SegmentName & "*" & tmplSegmentValue
                    segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                    If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                    ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                        End If
                    Else
                        tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                        tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                        ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                        End If
                    End If
                End If
            Else
                beginning_segment_text = beginning_segment_text & "*" & tmplSegmentValue
                segmentLoopDictKey = tmplSegmentLoopName & "-" & tmplSegmentName
                    If Not segmentLoopDict.exists(segmentLoopDictKey) Then
                    ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict.Add segmentLoopDictKey, tmplSegmentValue
                        End If
                    Else
                        tmplSegmentValuearr = segmentLoopDict(segmentLoopDictKey)
                        tmplSegmentValue = tmplSegmentValuearr & "@@@@@" & tmplSegmentValue
                        ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                        If tmplSegmentLoop <> "N" Then
                            segmentLoopDict(segmentLoopDictKey) = tmplSegmentValue
                        End If
                    End If
            End If
            LogStatement "Completed to add segment value to text"
            ' adding the tmplSegmentValue to the dictionary that will contain values to be saved to database for use in other scenarios
            

            
            If segments_req_count = 1 Then
                If dataToBeSavedToDbArr(currentRandKey).exists(segmentLoopDictKey) Then
                    ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                    If tmplSegmentLoop <> "N" Then
                        dataToBeSavedToDbArr(currentRandKey)(segmentLoopDictKey) = tmplSegmentValue
                    End If
                End If
            Else
                
                
                If depDataToBeSavedToDbArr(currentRandKey)(segments_req_count - 1).exists(segmentLoopDictKey) Then
                    ' checking if this segment can be stored into the segmentLoopDict by checking if value from column 1 is not "N"
                    If tmplSegmentLoop <> "N" Then
                        depDataToBeSavedToDbArr(currentRandKey)(segments_req_count - 1)(segmentLoopDictKey) = tmplSegmentValue
                    End If
                End If
            
            End If
                    
        End If
        LogStatement "Completed loop for row " & Row
    Next
    LogStatement "Function completed"
    ''msgbox beginning_segment_text
    If Len(beginning_segment_text) > 0 Then
        create_segment_scenario9 = beginning_segment_text + "~"
    Else
        create_segment_scenario9 = ""
    End If
End Function


Function Get_Refernced_Value(tmplSegmentValue_1, refIndicator, valuePlace)
    tmplSegmentValue = ""
    If InStr(1, Left(tmplSegmentValue_1, 4), refIndicator) <> 0 Then
        tmplSegmentValue = tmplSegmentValue_1
        ' checking if "|" is present to get the second part of the segment value
        ' example - "REF:2000-DTP03|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
        ' tmplSegmentValue_suffix will be set to -22991231
        tmplSegmentValue_suffix = ""
        If InStr(1, tmplSegmentValue, "|") <> 0 Then
            tmplSegmentValue_suffix = Split(tmplSegmentValue, "|")(1)
        End If
        
        tmplSegmentValue = GetReferenceSegmentValue(tmplSegmentValue, valuePlace)
        tmplSegmentValue = Trim(tmplSegmentValue)
        If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
            tmplSegmentValue = tmplSegmentValue & tmplSegmentValue_suffix
        End If
    End If
    Get_Refernced_Value = tmplSegmentValue

End Function
Function GetAnyButValue(tmplSegmentValue_1, any_but_indicator, valPlace, any_but_separator, any_but_val_separator)
    ' example tmplSegmentValue_1 = "ANY_BUT(MGIM,MGIC,MGIP@@@REF:2300-HD04)"
    'anyButValue - value to be returned by function
    anyButValue = ""
    'lines of businesses that have a specific requirement related to scenario 6
    scenario_6_lobs = "MH,BACO,MERCY,SIGNA,SOUTHCOAST"
    
    'any_but_value - the value that needs to be excluded while choosing value from a list of given values
    If InStr(1, tmplSegmentValue_1, any_but_separator) <> 0 Then
        any_but_value = Trim(Split(tmplSegmentValue_1, any_but_separator)(1))
        If InStr(1, any_but_value, refIndicator) <> 0 Then
            
            any_but_value = Get_Refernced_Value(any_but_value, refIndicator, valuePlace)
        End If
        Debug.Print any_but_value
    End If
    'any_but_values_list - the list of values from which one value can be chosen other than the any_but_value
    any_but_values_list = Replace(tmplSegmentValue_1, any_but_indicator, "")
    
    any_but_values_list = Split(any_but_values_list, any_but_separator)(0)
    
    ' below code is to remove the any_but_value from the any_but_values_list as it should not be included
    If InStr(1, any_but_values_list, any_but_val_separator) <> 0 Then
        any_but_values_list_split = Split(any_but_values_list, any_but_val_separator)
        new_any_but_values_list = ""
        For Each v In any_but_values_list_split
            If v <> any_but_value Then
                'the below condition is specific for scenario 6 and lines of businesses MH,MERCY, BACO, SIGNA and Southcoast, in which the type of plan change is plan class change i.e. from PC to PM or PM to PC
                If InStr(1, scenarioName, "Scenario_6") <> 0 And InStr(1, scenario_6_lobs, lineOfBusiness) <> 0 Then
                    If Right(any_but_value, 2) <> Right(v, 2) Then
                        If new_any_but_values_list = "" Then
                            new_any_but_values_list = v & ","
                        Else
                            new_any_but_values_list = new_any_but_values_list & v & ","
                        End If
                    End If
                'the below condition is specific for scenario 17 and lines of businesses MH,MERCY, BACO, SIGNA and Southcoast, in which the type of plan change should have the same plan class i.e. from PC to PC or PM to PM
                
                ElseIf (InStr(1, scenarioName, "Scenario_17") <> 0 Or InStr(1, scenarioName, "Scenario_15") <> 0) And InStr(1, scenario_6_lobs, lineOfBusiness) <> 0 Then
                    If Right(any_but_value, 2) = Right(v, 2) Then
                        If new_any_but_values_list = "" Then
                            new_any_but_values_list = v & ","
                        Else
                            new_any_but_values_list = new_any_but_values_list & v & ","
                        End If
                    End If
                Else
                    If new_any_but_values_list = "" Then
                        new_any_but_values_list = v & ","
                    Else
                        new_any_but_values_list = new_any_but_values_list & v & ","
                    End If
                End If
            End If
        Next
        If Right(new_any_but_values_list, 1) = "," Then
            new_any_but_values_list = Left(new_any_but_values_list, Len(new_any_but_values_list) - 1)
        End If
    End If
    
    
    Debug.Print new_any_but_values_list
    new_any_but_values_list = randChooseIndicator & new_any_but_values_list
    anyButValue = ChooseRandomValue(new_any_but_values_list)
    Debug.Print anyButValue
    GetAnyButValue = anyButValue
End Function


Function change_dob_as_per_plan(plan_value)
'    plan_value = "MBAPM"
'    plan_value_suffix = Right(plan_value, 2)
'    dob_val = get_dob_as_per_plan(plan_value_suffix)
'    'debug.print dob_val
'    Exit Sub


    If plan_value <> "" Then
        
        plan_value_suffix = Right(plan_value, 2)
        ' for scenario 1 which is enrollment of member we need the DOB as per plan
        'get_dob_as_per_plan function returns a DOB that is opposite of the current plan
        ' so changing the plan_value_suffix to opposite plan so DOB returned matches the current plan in case of scenario 1
        If Split(scenarioName, "_", 2)(1) = "Scenario_1" Then
            If plan_value_suffix = "PM" Then
                plan_value_suffix = "PC"
            ElseIf plan_value_suffix = "PC" Then
                plan_value_suffix = "PM"
            End If
        End If
        
        dob_val = get_dob_as_per_plan(plan_value_suffix)
        'debug.print dob_val
        If dob_val <> "" Then
            dob_val = GetFormattedDate(dob_val, "YYYYMMDD")
        End If
        change_dob_as_per_plan = dob_val
        
    Else
        change_dob_as_per_plan = GetFormattedDate(DateAdd("d", -1, Date), "YYYYMMDD")
    End If

End Function
Function get_dob_as_per_plan(plan_val)
    dt = Date
    Randomize
    If plan_val = "PM" Then
        r_1 = CInt(Rnd() * 100) Mod 18
        dt = DateAdd("yyyy", -r_1, dt)
        dt = DateAdd("m", -1, dt)
        dt_diff = DateDiff("yyyy", dt, Date)
        ' checking if the new dob generated results in age being greater than 21 and providing a new dob that is 3 years less then current date
        If dt_diff >= 21 Then
            'debug.print "Greater than 21"
            dt = DateAdd("yyyy", -3, Date)
            dt_diff = DateDiff("yyyy", dt, Date)
        End If
        get_dob_as_per_plan = dt
    ElseIf plan_val = "PC" Then

        r_1 = CInt(Rnd() * 100) Mod 65
        dt = DateAdd("yyyy", -r_1, dt)
        dt = DateAdd("m", -1, dt)
        dt_diff = DateDiff("yyyy", dt, Date)
        ' checking if the new dob generated results in age being less than 21 and providing a new dob that is atleast 22 years more then generated date
        ' and hence greater than 21 years
        If dt_diff <= 21 Then
            dt = DateAdd("yyyy", -22, dt)
            dt_diff = DateDiff("yyyy", dt, Date)
        End If
        
        get_dob_as_per_plan = dt
    Else
        get_dob_as_per_plan = ""
    End If
End Function

Function GetCombinationValue_1(combination_text, valPlace)
    ' this function is to handle segments such as "COMB:SBSB_FIRST_NAME+ +SBSB_LAST_NAME"
    'combination_text = "SBSB_FIRST_NAME+ +SBSB_LAST_NAME"
    combined_text = ""
    combination_text = Split(combination_text, ":")(1)
    combination_text_arr = Split(combination_text, "+")
    For Each comb_txt In combination_text_arr
        'checking if comb_txt contains "REF:" which indicates that value needs to refer to another segment value
        If InStr(1, Left(comb_txt, 4), refIndicator) <> 0 Then
            ' checking if "|" is present to get the second part of the segment value
            ' example - "REF:2000-DTP03|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
            ' tmplSegmentValue_suffix will be set to -22991231
            tmplSegmentValue_suffix = ""
            If InStr(1, comb_txt, "|") <> 0 Then
                tmplSegmentValue_suffix = Split(comb_txt, "|")(1)
            End If
            comb_txt = GetReferenceSegmentValue(comb_txt, valPlace)
            comb_txt = Trim(comb_txt)
            If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
                comb_txt = comb_txt & tmplSegmentValue_suffix
            End If
            combined_text = comb_txt
        Else
            db_text = GetRandomValueFromDatabaseDict(comb_txt, valPlace)
            If Len(db_text) > 0 Then
                combined_text = combined_text & db_text
            Else
                combined_text = combined_text & comb_txt
            End If
        End If
    Next
    GetCombinationValue = combined_text
End Function

Function GetCombinationValue(combination_text, valPlace)
    ' this function is to handle segments such as "COMB:SBSB_FIRST_NAME+ +SBSB_LAST_NAME"
    'combination_text = "SBSB_FIRST_NAME+ +SBSB_LAST_NAME"
    combined_text = ""
    temp_comb_text = combination_text
    combination_text = Split(combination_text, ":", 2)(1)
    combination_text_arr = Split(combination_text, "+")
    For Each comb_txt In combination_text_arr
        'checking if comb_txt contains "REF:" which indicates that value needs to refer to another segment value
        If InStr(1, Left(comb_txt, 4), refIndicator) <> 0 Then
            ' checking if "|" is present to get the second part of the segment value
            ' example - "REF:2000-DTP03|-22991231", here value after pipe symbol will be appended to the tmplSegmentValue obtained using reference indicator
            ' tmplSegmentValue_suffix will be set to -22991231
            tmplSegmentValue_suffix = ""
            If InStr(1, comb_txt, "|") <> 0 Then
                tmplSegmentValue_suffix = Split(comb_txt, "|")(1)
            End If
            If InStr(1, comb_txt, "(") <> 0 Then
                valPlace = Replace(Split(comb_txt, "(")(1), ")", "")
                valPlace = CInt(Trim(valPlace))
                comb_txt = Split(comb_txt, "(")(0)
            End If
                        
            comb_txt = GetReferenceSegmentValue(comb_txt, valPlace)
            comb_txt = Trim(comb_txt)
            If Not IsNull(tmplSegmentValue_suffix) And Not IsEmpty(tmplSegmentValue_suffix) Then
                comb_txt = comb_txt & tmplSegmentValue_suffix
            End If
            If combined_text = "" Then
                combined_text = comb_txt
            Else
                combined_text = combined_text & comb_txt
            End If
        ElseIf InStr(1, Left(comb_txt, 4), refIndicator) <> 0 Then
        
        Else
            On Error Resume Next
            db_text = GetRandomValueFromDatabaseDict(comb_txt, valPlace)
            On Error GoTo 0
            If Len(db_text) > 0 Then
                combined_text = combined_text & db_text
            Else
                combined_text = combined_text & comb_txt
            End If
        End If
    Next
    GetCombinationValue = combined_text
End Function



Function GetReferenceSegmentValue_1(referenceSegment)
    'referenceSegment = "REF:2100A-NM103" or "REF:2000-DTP03|-22991231"
    LogStatement "starting function GetReferenceSegmentValue with argument " & referenceSegment
    reference = Split(referenceSegment, ":")(1)
    ' checking if pipe symbol is present
    If InStr(1, reference, "|") <> 0 Then
        reference = Split(reference, "|")(0)
    End If
    If segmentLoopDict.exists(reference) Then
        GetReferenceSegmentValue = segmentLoopDict(reference)
    Else
        GetReferenceSegmentValue = "Not found"
    End If
    LogStatement "Function completed"
End Function
Function GetReferenceSegmentValue(referenceSegment, valPlace)
    'referenceSegment = "REF:2100A-NM103" or "REF:0-DTP03-YYYYMMDD" or "REF:2000-DTP03|-22991231"
    LogStatement "starting function GetReferenceSegmentValue with argument " & referenceSegment
    reference = Split(referenceSegment, ":")(1)
    formatt_type = ""

    
    If InStr(1, reference, "(") <> 0 Then
        valPlace = Replace(Split(reference, "(")(1), ")", "")
        valPlace = CInt(Trim(valPlace))
        reference = Split(reference, "(")(0)
    End If
    
    ' checking if pipe symbol is present
    If InStr(1, reference, "|") <> 0 Then
        reference = Split(reference, "|")(0)
    End If
    ' checking if format type is also provided as part of referenceSegment - example - "REF:0-DTP03-YYYYMMDD"
    no_of_slashes = Len(reference) - Len(Replace(reference, "-", ""))
    If no_of_slashes > 1 Then
        If InStr(1, reference, "-") <> 0 Then
            reference_1 = Split(reference, "-")(0)
            reference_2 = Split(reference, "-")(1)
            formatt_type = Split(reference, "-")(2)
            reference = reference_1 & "-" & reference_2
        End If
    End If
    
    If segmentLoopDict.exists(reference) Then
        GetReferenceSegmentVal = segmentLoopDict(reference)
        If InStr(1, GetReferenceSegmentVal, "@@@@@") <> 0 Then
            GetReferenceSegmentValue = Split(GetReferenceSegmentVal, "@@@@@")(valPlace - 1)
        Else
            GetReferenceSegmentValue = GetReferenceSegmentVal
        End If
    
    Else
        If InStr(1, scenarioName, "Scenario_7") = 0 And InStr(1, scenarioName, "Scenario_8") = 0 And InStr(1, scenarioName, "Scenario_16") = 0 Then
            If storedDataDict(currentRandKey).exists(reference) Then
                GetReferenceSegmentVal = storedDataDict(currentRandKey)(reference)
                If InStr(1, GetReferenceSegmentVal, "@@@@@") <> 0 Then
                    GetReferenceSegmentValue = Split(GetReferenceSegmentVal, "@@@@@")(valPlace - 1)
                Else
                    GetReferenceSegmentValue = GetReferenceSegmentVal
                End If
            Else
                GetReferenceSegmentValue = "Not found"
            End If
        Else
            If dataToBeSavedToDbArr(currentRandKey).exists(reference) Then
                GetReferenceSegmentVal = dataToBeSavedToDbArr(currentRandKey)(reference)
                If InStr(1, GetReferenceSegmentVal, "@@@@@") <> 0 Then
                    GetReferenceSegmentValue = Split(GetReferenceSegmentVal, "@@@@@")(valPlace - 1)
                Else
                    GetReferenceSegmentValue = GetReferenceSegmentVal
                End If
            Else
                GetReferenceSegmentValue = "Not found"
            End If
        End If
    
    End If
    If formatt_type <> "" Then
        If InStr(1, formatt_type, "YY") <> 0 Or InStr(1, formatt_type, "DD") <> 0 Then
            GetReferenceSegmentValue = GetFormattedDate(GetReferenceSegmentValue, formatt_type)
        ElseIf InStr(1, formatt_type, "HH") <> 0 Then
            GetReferenceSegmentValue = GetFormattedTime(GetReferenceSegmentValue, formatt_type)
        End If
    End If
    LogStatement "Function completed"
End Function
Function CreateRandomValue(tmplSegmentValue, valuePlace)
    ' valuePlace indicates the index of the value to be taken when there are multiple values for the same segment
    LogStatement "starting function CreateRandomValue with argument " & tmplSegmentValue
        'code for testing
    '    GetRandomDataValues 2
    '    currentRandKey = 2
    '    tmplSegmentValue = "RAND:SSN"
    '    randIndicator = "RAND:"
    '    randDateIndicator = "DATE"
    '    randTimeIndicator = "TIME"
    '    randFNIndicator = "FN"
    '    randLNIndicator = "LN"
    '    randMIIndicator = "MI"
    '    randSSNIndicator = "SSN"
    '    randDOBIndicator = "DOB"
    '    randADDR1Indicator = "ADDR1"
    '    randADDR2Indicator = "ADDR2"
    '
        randomValueGenerated = ""
        randYes = "N"
        randDirection = ""
        randHowmany = ""
        randInterval_val = ""
        randValueFormat = ""
        ''msgbox "tmplSegmentValueInd is " & tmplSegmentValueInd
        tmplSegmentValueInd = Split(tmplSegmentValue, randIndicator)(1)
        increaseCurrentRandKey = False
        If InStr(1, tmplSegmentValueInd, "-") <> 0 Then
            LogStatement "Random indicator exists"
            randValueIndicator = Split(tmplSegmentValueInd, "-")(0)
            randValueFormat = Split(tmplSegmentValueInd, "-")(1)
            ''msgbox  "randValueIndicator " &  randValueIndicator & " randValueFormat " & randValueFormat
            If randValueIndicator = randDateIndicator Then
                'assigning default values for the formatted date function
                randYes = "N"
                randDirection = ""
                randHowmany = ""
                randInterval_val = ""
                ' checking if additional format conditions are present
                If InStr(1, tmplSegmentValueInd, "-P") <> 0 Then ' -P indicates past
                    randDirection = Split(tmplSegmentValueInd, "-")(2)
                    randYes = "Y"
                    randHowmany = Split(tmplSegmentValueInd, "-")(3)
                    randInterval_val = Split(tmplSegmentValueInd, "-")(4)
                End If
                If InStr(1, tmplSegmentValueInd, "-F") <> 0 Then ' -F indicates future
                    randDirection = Split(tmplSegmentValueInd, "-")(2)
                    randYes = "Y"
                    randHowmany = Split(tmplSegmentValueInd, "-")(3)
                    randInterval_val = Split(tmplSegmentValueInd, "-")(4)
                End If
                ''msgbox "Passing arguments " & randValueFormat + " " +  randYes + " " +  randDirection + " " + randHowmany + " " + randInterval_val
                randomValueGenerated = FormattedDate(randValueFormat, randYes, randDirection, randHowmany, randInterval_val)
            ElseIf randValueIndicator = randTimeIndicator Then
                'assigning default values for the formatted date function
                randYes = "N"
                randDirection = ""
                randHowmany = ""
                randInterval_val = ""
                ' checking if additional format conditions are present
                If InStr(1, tmplSegmentValueInd, "-P") <> 0 Then ' -P indicates past
                    randDirection = Split(tmplSegmentValueInd, "-")(2)
                    randYes = "Y"
                    randHowmany = Split(tmplSegmentValueInd, "-")(3)
                    randInterval_val = Split(tmplSegmentValueInd, "-")(4)
                End If
                If InStr(1, tmplSegmentValueInd, "-F") <> 0 Then ' -F indicates future
                    randDirection = Split(tmplSegmentValueInd, "-")(2)
                    randYes = "Y"
                    randHowmany = Split(tmplSegmentValueInd, "-")(3)
                    randInterval_val = Split(tmplSegmentValueInd, "-")(4)
                End If
                randomValueGenerated = FormattedTime(randValueFormat, randYes, randDirection, randHowmany, randInterval_val)
            ElseIf randValueIndicator = randNumIndicator Then
                randomValueGenerated = GenerateUniqueRandomNumber(randValueFormat)
            End If
        Else
            LogStatement "Random value to be generated for patient demographics"
            randValueIndicator = tmplSegmentValueInd
            '''''deug.print randValueIndicator
            '''''''deug.print "currentRandKey used to get random values " & currentRandKey
            If randValueIndicator = "START_DATE_OF_YEAR" Then
                ' this function call will return todays date
                randomValueGenerated = FormattedDate(randValueFormat, randYes, randDirection, randHowmany, randInterval_val)
                randomValueGenerated = GetStartDayOfYear(randomValueGenerated)
            Else
                randomValueGenerated = GetRandomValueFromDatabaseDict(randValueIndicator, valuePlace)
            End If
        End If
        LogStatement "Function completed"
        ''msgbox "Random value generated " & randomValueGenerated
        randomValueGenerated = Trim(CStr(randomValueGenerated))
        ' replacing empty characters that are getting added to the end of the random value
        randomValueGenerated = Replace(randomValueGenerated, Chr(160), "")
        ' assigning the random value generated to global variable MEME_MEDCD_NO, as this value will be used in function to store values to DB
        '''''debug.print "randValueIndicator " & randValueIndicator
        If randValueIndicator = "MEME_MEDCD_NO" Then
            If MEME_MEDCD_NO = "" Then
                MEME_MEDCD_NO = randomValueGenerated
            End If
        End If
        CreateRandomValue = randomValueGenerated
End Function
Function GetRandomValueFromDatabaseDict_(randValueIndicator, valuePlace)
    randomValueGenerated = ""

    
    If randomDataDict.exists(currentRandKey) Then
        ' getting value from dictionary which is an dictionary of random values
        Set rand_Values = randomDataDict.Item(currentRandKey)
        ' checking if the random value indicator is present as a key in the dictionary
        If rand_Values.exists(randValueIndicator) Then
            randomValueGenerated = rand_Values(randValueIndicator)
            'checking if comma is present in randomValueGenerated which indicates that multiple values are present for the segment
            ' in that case we get the value corresponding to the index from variable valuePlace passed as argument
            If InStr(1, randomValueGenerated, val_separator) <> 0 Then
                randomValueGeneratedArr = Split(randomValueGenerated, val_separator)
                randomValueGenerated = randomValueGeneratedArr(valuePlace)
            End If
            ' performing additional formatting
            If InStr(1, randValueIndicator, "SSN") <> 0 Then
                randomValueGenerated = CStr(randomValueGenerated)
                ' below loop is to add prefix of Zero to SSN if length is less than 9
                Do While Len(randomValueGenerated) < 9
                    randomValueGenerated = "0" & randomValueGenerated
                    If Len(randomValueGenerated) >= 9 Then
                        Exit Do
                    End If
                Loop
            'ElseIf InStr(1, randValueIndicator, "BIRTH_DT") <> 0 Then
            '    randomValueGenerated = GetFormattedDate(randomValueGenerated, "YYYYMMDD")
            'ElseIf InStr(1, randValueIndicator, "ZIP") <> 0 Then
            '    Do While Len(randomValueGenerated) > 5 And Len(randomValueGenerated) < 11
            '        randomValueGenerated = randomValueGenerated + "0"
            '    Loop
            '    If Len(randomValueGenerated) > 0 Then
            '        Do While Len(randomValueGenerated) < 5
            '            randomValueGenerated = randomValueGenerated + "0"
            '        Loop
            '    End If
            '    randomValueGenerated = CStr(randomValueGenerated)
            Else
                If IsNull(randomValueGenerated) Or IsEmpty(randomValueGenerated) Then
                    randomValueGenerated = ""
                Else
                    randomValueGenerated = CStr(randomValueGenerated)
                End If
            End If
        End If
    End If
    randomValueGenerated = Trim(CStr(randomValueGenerated))
    ' replacing empty characters that are getting added to the end of the random value
    randomValueGenerated = Replace(randomValueGenerated, Chr(160), "")
    GetRandomValueFromDatabaseDict = randomValueGenerated
End Function

Function GetRandomValueFromDatabaseDict(randValueIndicator, valuePlace)
    randomValueGenerated = ""

    
    If randomDataDict.exists(currentRandKey) Then
        ' getting value from dictionary which is an dictionary of random values
        If dep_ct = 0 Then
            Set rand_Values = randomDataDict.Item(currentRandKey)
        Else
            Set rand_Values = dependantRandomDataDict.Item(currentRandKey).Item(dep_ct)
        End If
        
        ' checking if the random value indicator is present as a key in the dictionary
        If rand_Values.exists(randValueIndicator) Then
            randomValueGenerated = rand_Values(randValueIndicator)
            'checking if comma is present in randomValueGenerated which indicates that multiple values are present for the segment
            ' in that case we get the value corresponding to the index from variable valuePlace passed as argument
            If InStr(1, randomValueGenerated, val_separator) <> 0 Then
                randomValueGeneratedArr = Split(randomValueGenerated, val_separator)
                randomValueGenerated = randomValueGeneratedArr(valuePlace)
            End If
            ' performing additional formatting
            If InStr(1, randValueIndicator, "SSN") <> 0 Then
                randomValueGenerated = CStr(randomValueGenerated)
                ' below loop is to add prefix of Zero to SSN if length is less than 9
                Do While Len(randomValueGenerated) < 9
                    randomValueGenerated = "0" & randomValueGenerated
                    If Len(randomValueGenerated) >= 9 Then
                        Exit Do
                    End If
                Loop
            'ElseIf InStr(1, randValueIndicator, "BIRTH_DT") <> 0 Then
            '    randomValueGenerated = GetFormattedDate(randomValueGenerated, "YYYYMMDD")
            'ElseIf InStr(1, randValueIndicator, "ZIP") <> 0 Then
            '    Do While Len(randomValueGenerated) > 5 And Len(randomValueGenerated) < 11
            '        randomValueGenerated = randomValueGenerated + "0"
            '    Loop
            '    If Len(randomValueGenerated) > 0 Then
            '        Do While Len(randomValueGenerated) < 5
            '            randomValueGenerated = randomValueGenerated + "0"
            '        Loop
            '    End If
            '    randomValueGenerated = CStr(randomValueGenerated)
            Else
                If IsNull(randomValueGenerated) Or IsEmpty(randomValueGenerated) Then
                    randomValueGenerated = ""
                Else
                    randomValueGenerated = CStr(randomValueGenerated)
                End If
            End If
        End If
    End If
    randomValueGenerated = Trim(CStr(randomValueGenerated))
    ' replacing empty characters that are getting added to the end of the random value
    randomValueGenerated = Replace(randomValueGenerated, Chr(160), "")
    GetRandomValueFromDatabaseDict = randomValueGenerated
End Function


Function ChooseRandomValue(tmplSegmentValue)
'values for testing
'tmplSegmentValue = "ANY:F,M,U,L,D"
'randChooseIndicator = "ANY:"
LogStatement "starting function ChooseRandomValue with argument " & tmplSegmentValue
tmplSegmentValueInd = Split(tmplSegmentValue, randChooseIndicator)(1)
tmplSegmentValuearr = Split(tmplSegmentValueInd, ",")
tmplSegmentValueArrLen = UBound(tmplSegmentValuearr) + 1
Randomize
r_n = Rnd(3) * 10
r_m = r_n Mod tmplSegmentValueArrLen
LogStatement "Function completed"
ChooseRandomValue = tmplSegmentValuearr(r_m)
End Function
Function create_beginning_segment(objTemplateWorksheet, objTemplateWorksheetRangeRows) '  working function
    Dim tmplLoop, tmplSubLoop, tmplSegmentName, tmplRequirement, tmplLoopCount, tmplSegmentValue, tmplValueLength
    Dim SegmentName
    Dim beginning_segment, beginning_segment_text
    ' variable to hold the entire beginning segment text that will be returned as output by the function
    beginning_segment_text = ""
    ' variable to hold one segment text
    segment_text = ""
    ' variable that will be crosschecked to identify the change in segments as we loop through the excel rows
    startSegmentName = ""
    For Row = 2 To objTemplateWorksheetRangeRows
        ' checking if the value in first column indicates a beginning segment
        tmplSegmentLoop = objTemplateWorksheet.Cells(Row, 1)
        If tmplSegmentLoop <> "B" Then
            LogStatement "Function completed"
            create_beginning_segment = beginning_segment_text
            Exit Function
        End If
        tmplSegmentLoopName = objTemplateWorksheet.Cells(Row, 3)
        tmplSegmentName = objTemplateWorksheet.Cells(Row, 4)
        tmplRequirement = objTemplateWorksheet.Cells(Row, 5)
        ' moving to the next row if the current row segment is not to be included in the edi text,
        ' indicated by absence of "Y" in column
        ' proceeding to include the segment value only if "Y"
        If tmplRequirement = "Y" Then
            tmplValueLength = objTemplateWorksheet.Cells(Row, 8)
            tmplLoopCount = objTemplateWorksheet.Cells(Row, 6)
            tmplSegmentValue = objTemplateWorksheet.Cells(Row, 7)
            If tmplValueLength <> "" Then
                tmplValueLength = CInt(tmplValueLength)
                tmplSegmentValue = RequiredLengthString(tmplSegmentValue, tmplValueLength)
            End If
            SegmentName = RootSegment(tmplSegmentName)
            If startSegmentName <> SegmentName Then
                startSegmentName = SegmentName
                If Len(segment_text) > 0 Then
                    beginning_segment_text = beginning_segment_text + segment_text + "~"
                    segment_text = ""
                Else
                    segment_text = SegmentName & "*" & tmplSegmentValue
                End If
            Else
                segment_text = segment_text & "*" & tmplSegmentValue
            End If
        End If
    Next
LogStatement "Function completed"
End Function
Function FindSegments(objTemplateWorksheet, startLoopName, endLoopName)
    LogStatement "starting function FindSegments with argument startLoopName -" & startLoopName & " endLoopName- " & endLoopName
    endLoopName = CStr(endLoopName)
    endLoopName = Trim(endLoopName)
    row_numbers_range = ""
    start_segment = ""
    objTemplateWorksheetRangeRows = objTemplateWorksheet.UsedRange.Rows.Count
    'default value for row from which segments will be parsed for edi text
    start_row = 2
    For x = 2 To objTemplateWorksheetRangeRows
        tmplSegLoopName = objTemplateWorksheet.Cells(x, 3)
        If CStr(tmplSegLoopName) = CStr(startLoopName) Then
            start_row = x
            Exit For
        End If
    Next
    end_row = 2
    loop_segment = 0
    reassign_start_row = True
    For x = start_row To objTemplateWorksheetRangeRows
        end_row = x
        tmplSegmentLoop = objTemplateWorksheet.Cells(x, 1)
        tmplSegmentLoopName = objTemplateWorksheet.Cells(x, 3)
        ' getting the value of the loop name in the next row, this is used in logic for handling segment with only one row
        nxttmplSegmentLoopName = objTemplateWorksheet.Cells(x + 1, 3)
        prevtmplSegmentLoopName = objTemplateWorksheet.Cells(x - 1, 3)
        tmplSegmentName = objTemplateWorksheet.Cells(x, 4)
        tmplRequirement = objTemplateWorksheet.Cells(x, 5)
        SegmentName = RootSegment(tmplSegmentName)
        If startSegmentName <> SegmentName And start_row <> end_row Then
            startSegmentName = SegmentName
            end_row = end_row - 1
            tmplSegmentValue = objTemplateWorksheet.Cells(start_row, 7)
            If InStr(1, tmplSegmentValue, ",") <> 0 Then
                loop_segment = Split(tmplSegmentValue, ",")
                loop_segment = UBound(loop_segment) + 1
            Else
                loop_segment = 1
            End If
            For lp = 1 To loop_segment
                If row_numbers_range = "" Then
                    row_numbers_range = CStr(start_row) & "to" & CStr(end_row) & "$" & lp
                Else
                    row_numbers_range = row_numbers_range & ":" & CStr(start_row) & "to" & CStr(end_row) & "$" & lp
                End If
            Next
            start_row = x
            If reassign_start_row = False Then
                reassign_start_row = True
            End If
        Else
            startSegmentName = SegmentName
        End If
        If CStr(tmplSegmentLoopName) = endLoopName Then
            FindSegments = row_numbers_range
            Exit Function
        End If
    Next
    LogStatement "Function completed"
    FindSegments = row_numbers_range
End Function
Function removeNumbersFromString(inString)
    ' Function takes a alpha numeric string as input and returns a string with only alphabets
    ' without any numbers and special characters
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    outString = ""
    For x = 1 To Len(inString)
        If InStr(1, chars, Mid(inString, x, 1)) <> 0 Then
            outString = outString + Mid(inString, x, 1)
        End If
    Next
    LogStatement "Function removeNumbersFromString completed"
    removeNumbersFromString = outString
End Function
Function RootSegment(tmplSegmentName)
    ' function removes the numeric suffix in the segment name and returns only the root segment
    RootSegment = Left(tmplSegmentName, Len(tmplSegmentName) - 2)
    LogStatement "Function RootSegment completed"
End Function
Function RequiredLengthString(inString, required_len)
    'inString - Input string
    'required_len - The required length of the string
    'Function checks if the length of the input string is less than the required length and
    ' adds spaces at the end till the string is of the required length
    ct = 0
    Do While Len(inString) < required_len
        inString = inString & " "
    Loop
    RequiredLengthString = inString
    LogStatement "Function RequiredLengthString completed"
End Function
Function FormattedDate(dtFormat, random, direction, howmany, interval_val)
    'random - False, True
    'howmany - any integer, negative integers to generate dates in past
    'interval_val - string representing the code for the part of date to be modified, possibles values are give below
    'Optional random = False, Optional howmany = 10, Optional interval_val = "yyyy"
    'possible values for interval_val
    'yyyy    Year
    'q Quarter
    'm Month
    'y   Day of year
    'd Day
    'w Weekday
    'ww Week
    'h Hour
    'n Minute
    's Second
    'if random is False, todays date will be generated and returned in the specified format
    If random = "N" Then
        dt = Date
    Else
        If direction = "P" Then
            howmany = -howmany
        End If
        dt = DateAdd(interval_val, howmany, Date)
    End If
        dd = CStr(Day(dt)) ' day of date without zero prefix
        mm = CStr(Month(dt)) ' month of day without zero prefix
        yyyy = CStr(Year(dt)) ' four digit year
        If Len(dd) = 1 Then
            dd = "0" & dd
        End If
        If Len(mm) = 1 Then
            mm = "0" & mm
        End If
        If Len(yyyy) = 4 Then
            yy = Right(yyyy, 2)
        End If
        If Len(yyyy) = 2 Then
            yyyy = "20" & yyyy
        End If
        'FormattedDate = Format(dt, cStr(dtFormat))
        If dtFormat = "DDMMYY" Then
            FormattedDate = dd & mm & yy
        ElseIf dtFormat = "YYMMDD" Then
            FormattedDate = yy & mm & dd
        ElseIf dtFormat = "YYYYMMDD" Then
            FormattedDate = yyyy & mm & dd
        Else
            FormattedDate = dt
        End If
End Function
Function GetStartDayOfYear(dt)
    yr = Year(dt)
    GetStartDayOfYear = CStr(yr) & "0101"
End Function
Function FormattedTime(dtFormat, random, direction, howmany, interval_val)
    'Optional random = False, Optional howmany = 10, Optional interval_val = "n"
     'possible values for interval_val
    'h Hour
    'n Minute
    's Second
    If random = "N" Then
        dt = Time
    Else
        If direction = "P" Then
            howmany = -howmany
        End If
        dt = DateAdd(interval_val, howmany, Time)
    End If
    hh = CStr(Hour(dt))
    Min = CStr(Minute(dt))
    ss = CStr(Second(dt))
    If Len(hh) = 1 Then
        hh = "0" & hh
    End If
    If Len(Min) = 1 Then
        Min = "0" & Min
    End If
    If Len(ss) = 1 Then
        ss = "0" & ss
    End If
    If dtFormat = "HHMM" Then
        FormattedTime = hh & Min
    ElseIf dtFormat = "HHMMSS" Then
        FormattedTime = hh & Min & ss
    End If
End Function
Function GenerateUniqueRandomNumber_1()
    firstNum = FormattedDate("DDMMYY", "N", "", "", "")
    secondNum = FormattedTime("HHMMSS", "N", "", "", "")
    'Application.Wait Now() + TimeValue("00:00:01")
    'WScript.Sleep 500
    GenerateUniqueRandomNumber = CStr(firstNum) & CStr(secondNum)
End Function
Function GenerateUniqueRandomNumber(fmt)
    If IsNull(fmt) Or IsEmpty(fmt) Then
        fmt = ""
    End If
    dt = Date
    dd = CStr(Day(dt))
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    mm = CStr(Month(dt))
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    yyyy = CStr(Year(dt))
    If Len(yyyy) = 4 Then
        yy = Right(yyyy, 2)
    End If
    tm = Time
    hh = CStr(Hour(tm))
    If Len(hh) = 1 Then
        hh = "0" & hh
    End If
    mins = CStr(Minute(tm))
    If Len(mins) = 1 Then
        mins = "0" & mins
    End If
    ss = CStr(Second(tm))
    If Len(ss) = 1 Then
        ss = "0" & ss
    End If
    If fmt = "" Then
        GenerateUniqueRandomNumber = dd & mm & yy & hh & mins & ss
    ElseIf fmt = "DDMMYY" Then
        GenerateUniqueRandomNumber = dd & mm & yy
    ElseIf fmt = "DDMMYYYY" Then
        GenerateUniqueRandomNumber = dd & mm & yyyy
    ElseIf fmt = "DDMMYYHH" Then
        GenerateUniqueRandomNumber = dd & mm & yy & hh
    ElseIf fmt = "DDMMYYHHMM" Then
        GenerateUniqueRandomNumber = dd & mm & yy & hh & mins
    ElseIf fmt = "DDMMYYHHSS" Then
        GenerateUniqueRandomNumber = dd & mm & yy & hh & ss
    ElseIf fmt = "DDMMYYHHMMSS" Then
        GenerateUniqueRandomNumber = dd & mm & yy & hh & mins & ss
    ElseIf fmt = "HHMM" Then
        GenerateUniqueRandomNumber = hh & mins
    ElseIf fmt = "HHMMSS" Then
        GenerateUniqueRandomNumber = hh & mins & ss
    End If
End Function
Function GetRandomDataValues(noOfData, randomDataFilePath)
'Sub GetRandomDataValues()
'    noOfData = 2
'    randomDataFilePath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Supporting Files\Random_Data_DB.xlsx"
'
    Dim objConnection, objRecordset
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordset = CreateObject("ADODB.Recordset")
    Set randomDataDict = CreateObject("Scripting.Dictionary")
    dbConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & randomDataFilePath & ";Extended Properties = excel 12.0;" ' excel 8.0 for excel 2003
    If objConnection.State = 1 Then
        objConnection.Close
    End If
    With objConnection
        .Open dbConnectStr
    End With
    ' fetching records that have not already been used for generating edi records
    'objRecordset.Open "select * from [Master$] where [Used] = 'N'", objConnection
    objRecordset.Open "select Top " & noOfData & " * from [Master$] where [Used] = 'N' and LEN([MEME_MEDCD_NO]) = " & memberIdLength & ";", objConnection
    '''On Error Resume Next
    fields_count = objRecordset.Fields.Count
    ' looping numberOfRecordsNeeded times to create dictionaries with column names as keys
    For rec_ct = 1 To noOfData
        ' creating dictionay to hold values from each row of the recordset with column names as keys
        Set randomSubDict = CreateObject("Scripting.Dictionary")
        ' looping to add individual column names as keys into dictionary
        For Each field_name In objRecordset.Fields
            ''''''deug.print field_name.Name
            randomSubDict.Add field_name.Name, ""
        Next
        '''''''deug.print rec_ct
        randomDataDict.Add rec_ct, randomSubDict
    Next
    ct = 1
    ssn_field_name = ""
    Do While objRecordset.EOF = False
        For Each subDict In randomDataDict.items
            For Each k In subDict.keys ' k => keys within the dict, in this case keys will be the names of columns in database
                ''''''''deug.print k
                If k = Empty Then
                    Exit For
                End If
                randomDataDict(ct)(k) = objRecordset.Fields(k)
                If Len(ssn_field_name) = 0 Then
                    If InStr(1, k, "SSN") <> 0 Then
                        ssn_field_name = k
                    End If
                End If
            Next
            'randomDataDict(ct) = subDict
            If ct >= noOfData Then
                '''''''''deug.print ct
                Exit Do
            End If
            ct = ct + 1
            objRecordset.MoveNext
        Next
    Loop
    objRecordset.Close
    '' marking the records that have already been used for generating edi records as "Y" so they will not be used during next run to generate edi records
    ssns = "("
    ' getting the list of SSNs that are to be marked as Y
    For Each d In randomDataDict.items
        ssns = ssns & "'" & d(ssn_field_name) & "'" & ","
    Next
    ssns = Left(ssns, Len(ssns) - 1) & ")"
    ''''''''deug.print "update [Master$] set [Used] = 'Y' where [" & ssn_field_name & "] in " & ssns & ";"
    objRecordset.Open "update [Master$] set [Used] = 'Y' where [" & ssn_field_name & "] in " & ssns & ";", objConnection
    'objRecordset.Open "update [Master$] set [Used] = 'Y' where [SSN] = '" & ssn & "'", objConnection
    ''On Error Resume Next
    objRecordset.Close
    objConnection.Close
End Function

Sub GetRandomDataValues_2(noOfData, randomDataFilePath)

'Sub GetRandomDataValues_2()
'    noOfData = 2
'    numberOfRecordsNeeded = 2
'    memberIdLength = 12
'    randomDataFilePath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Supporting Files\Random_Data_DB.xlsx"
    
       
    Dim randomDataArray(25)
    Dim MEME_MEDCD_NO_ARR, SBSB_FIRST_NAME_ARR, SBSB_LAST_NAME_ARR, SBSB_MID_INIT_ARR, RES_PARTY_FIRST_NAME_ARR, RES_PARTY_LAST_NAME_ARR
    Dim RES_PARTY_MID_INIT_ARR, SBSB_TELEPHONE_ARR, SBAD_ADDR1_ARR, SBAD_ADDR2_ARR, SBAD_CITY_ARR, SBAD_ZIP_ARR, SSN_ARR, BIRTH_DT_ARR
    Dim MEME_SEX_ARR, MEME_SSNS_ARR, RES_PARTY_SSNS_ARR
    
    Set randomDataDict = CreateObject("Scripting.Dictionary")
    
    If dependantDataRequired = True Then
        Set dependantRandomDataDict = CreateObject("Scripting.Dictionary")
    End If
    
    
    
    randomColumnNames = "Used,MEME_MEDCD_NO,SBSB_FIRST_NAME,SBSB_LAST_NAME,SBSB_MID_INIT,RES_PARTY_FIRST_NAME,RES_PARTY_LAST_NAME,RES_PARTY_MID_INIT,SBSB_TELEPHONE,PLAN_SPONSOR_FIRST_NAME,PLAN_SPONSOR_LAST_NAME,PLAN_SPONSOR_TIN,SBAD_ADDR1,SBAD_ADDR2,SBAD_ADDR3,SBAD_CITY,SBAD_STATE,SBAD_ZIP,SBAD_COUNTY,MEME_FAM_LINK_ID,MEME_SSN,MEME_BIRTH_DT,MEME_SEX,MEME_ORIG_EFF_DT,MEPE_TERM_DT,RES_PARTY_SSN"
    randomColumnNamesArr = Split(randomColumnNames, ",")
    
    If lineOfBusiness <> "DCH" Then
        MEME_MEDCD_NO_ARR = Split(getRandomMemberIds(), ",")
    Else
        MEME_MEDCD_NO_ARR = Split(getRandomMemberIdsForDCH(), ",")
    End If
    
    SBSB_FIRST_NAME_ARR = Split(getRandomName(True, True), ",")
    SBSB_LAST_NAME_ARR = Split(getRandomName(False, True), ",")
    SBSB_MID_INIT_ARR = Split(getMiddleInitial(), ",")
    RES_PARTY_FIRST_NAME_ARR = Split(getRandomName(True, False), ",")
    RES_PARTY_LAST_NAME_ARR = Split(getRandomName(False, False), ",")
    RES_PARTY_MID_INIT_ARR = Split(getMiddleInitial(), ",")
    SBSB_TELEPHONE_ARR = Split(getTelephone(), ",")
    SBAD_ADDR1_ARR = Split(getRandomAddresses1(), ",")
    SBAD_ADDR2_ARR = Split(getRandomAddresses2(), ",")
    SBAD_CITY_ARR = Split(getCities(), ",")
    SBAD_ZIP_ARR = Split(getZipCodes(), ",")
    MEME_SSNS_ARR = Split(getRandomSsn(), ",")
    RES_PARTY_SSNS_ARR = Split(getRandomSsn(), ",")
    BIRTH_DT_ARR = Split(getRandomDOB(), ",")
    MEME_SEX_ARR = Split(getGenders(), ",")
    
    states = ""
    For x = 1 To numberOfRecordsNeeded
        states = states & "MA" & ","
    Next
    
    If Right(states, 1) = "," Then
        states = Left(states, Len(states) - 1)
    End If
    
    statesArr = Split(states, ",")
    

    randomDataArray(1) = MEME_MEDCD_NO_ARR
    randomDataArray(2) = SBSB_FIRST_NAME_ARR
    randomDataArray(3) = SBSB_LAST_NAME_ARR
    randomDataArray(4) = SBSB_MID_INIT_ARR
    randomDataArray(5) = RES_PARTY_FIRST_NAME_ARR
    randomDataArray(6) = RES_PARTY_LAST_NAME_ARR
    randomDataArray(7) = RES_PARTY_MID_INIT_ARR
    randomDataArray(8) = SBSB_TELEPHONE_ARR
    randomDataArray(12) = SBAD_ADDR1_ARR
    randomDataArray(13) = SBAD_ADDR2_ARR
    randomDataArray(15) = SBAD_CITY_ARR
    randomDataArray(16) = statesArr
    randomDataArray(17) = SBAD_ZIP_ARR
    randomDataArray(20) = MEME_SSNS_ARR
    randomDataArray(21) = BIRTH_DT_ARR
    randomDataArray(22) = MEME_SEX_ARR
    randomDataArray(25) = RES_PARTY_SSNS_ARR
    
    ' looping numberOfRecordsNeeded times to create dictionaries with column names as keys
    For rec_ct = 1 To noOfData
        ' creating dictionay to hold values from each row of the recordset with column names as keys
        Set randomSubDict = CreateObject("Scripting.Dictionary")
        If dependantDataRequired = True Then
            Set depRandomSubDict = CreateObject("Scripting.Dictionary")
            depRandomSubDict.Add 1, CreateObject("Scripting.Dictionary")
            depRandomSubDict.Add 2, CreateObject("Scripting.Dictionary")
        End If
        ' looping to add individual column names as keys into dictionary
        For x = 0 To UBound(randomColumnNamesArr)
            
            randomSubDict.Add randomColumnNamesArr(x), ""
            If dependantDataRequired = True Then
                depRandomSubDict(1).Add randomColumnNamesArr(x), ""
                depRandomSubDict(2).Add randomColumnNamesArr(x), ""
            End If
        Next
        '''''''deug.print rec_ct
        randomDataDict.Add rec_ct, randomSubDict
        If dependantDataRequired = True Then
            
            dependantRandomDataDict.Add rec_ct, depRandomSubDict
            
        End If
    Next
    
    
    
    ct = 1
    edi_record_ct = 1
    ssn_field_name = ""
   
    On Error Resume Next
    For Each subDict In randomDataDict.items
        arrCt = 0
        For Each k In subDict.keys ' k => keys within the dict, in this case keys will be the names of columns in database
            ''''''''deug.print k
            If k = Empty Then
                Exit For
            End If
            
            randomDataDict(edi_record_ct)(k) = randomDataArray(arrCt)(edi_record_ct - 1)
            If dependantDataRequired = True Then
                'example of array data generated when noOf Data is 3 = 1,2,3,4,5,6,7,8,9 - below code is to distribute an array of data generated between its member and 2 dependants
                dependantRandomDataDict(edi_record_ct)(1)(k) = randomDataArray(arrCt)((ct + 1) - 1)
                dependantRandomDataDict(edi_record_ct)(2)(k) = randomDataArray(arrCt)((ct + 2) - 1)
            End If
            arrCt = arrCt + 1
            If Len(ssn_field_name) = 0 Then
                If InStr(1, k, "SSN") <> 0 Then
                    ssn_field_name = k
                End If
            End If
        Next
        'randomDataDict(ct) = subDict
        If ct >= noOfData * how_many_random_data Then
            '''''''''deug.print ct
            Exit For
        End If
        ct = ct + how_many_random_data
        edi_record_ct = edi_record_ct + 1
        
    Next
    On Error GoTo 0
    

End Sub

Function getStartOfMonthDate()

    dt = Date
    new_dt = CDate(Month(dt) & "/01/" & Year(dt))
    new_dt = GetFormattedDate(new_dt, "YYYYMMDD")
    'Debug.Print new_dt
    getStartOfMonthDate = new_dt
    
End Function



Function GetFormattedDate(dt, dtFormat)
    ' for now the function logic is used only for formatting the dob
    ' logic needs to be extended to format date for all types of date format
    If dt = "Not found" Then
        GetFormattedDate = dt
        Exit Function
    End If
    dd = CStr(Day(dt)) ' day of date without zero prefix
    mm = CStr(Month(dt)) ' month of day without zero prefix
    yy = CStr(Year(dt)) ' four digit year
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    If Len(yy) = 4 And dtFormat = "YYMMDD" Then
        yy = Right(yy, 2)
    Else
       yyyy = yy
    End If
    If Len(yy) = 2 Then
        yyyy = "20" & yy
    End If
    'FormattedDate = Format(dt, cStr(dtFormat))
    If dtFormat = "DDMMYY" Then
        GetFormattedDate = dd & mm & yy
    ElseIf dtFormat = "YYMMDD" Then
        GetFormattedDate = yy & mm & dd
    ElseIf dtFormat = "YYYYMMDD" Then
        GetFormattedDate = yyyy & mm & dd
    ElseIf dtFormat = "DDMMYYYY" Then
        GetFormattedDate = dd & mm & yyyy
    End If
End Function
Function GetFormattedTime(dt, dtFormat)
    'Optional random = False, Optional howmany = 10, Optional interval_val = "n"
     'possible values for interval_val
    'h Hour
    'n Minute
    's Second
    If dt = "Not found" Then
        GetFormattedTime = dt
        Exit Function
    End If
    hh = CStr(Hour(dt))
    Min = CStr(Minute(dt))
    ss = CStr(Second(dt))
    If Len(hh) = 1 Then
        hh = "0" & hh
    End If
    If Len(Min) = 1 Then
        Min = "0" & Min
    End If
    If Len(ss) = 1 Then
        ss = "0" & ss
    End If
    If dtFormat = "HHMM" Then
        GetFormattedTime = hh & Min
    ElseIf dtFormat = "HHMMSS" Then
        GetFormattedTime = hh & Min & ss
    End If
End Function
Function ReadEdiTextFile(filePath)
    'wscript.echo "inside readEdiTextFile " & filePath
    'filePath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Output\MH\MH_Scenario 1_260521154539_2.txt"
    ediText = ""
    If FSO.FileExists(filePath) Then
        Set objFile = FSO.OpenTextFile(filePath)
        ediText = objFile.ReadAll
        objFile.Close
    End If
    ReadEdiTextFile = ediText
End Function
Function CreateSaveDataDictionary(strDataPath, scenario)
'Sub CreateSaveDataDictionary()
'numberOfRecordsNeeded = 1
'strDataPath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Scenarios\CCA.xlsx"
'scenario = "Scenario 1"
'Set dataToBeSavedToDbDict = CreateObject("Scripting.Dictionary")
    'templateDict - dictionary variable to hold the keys (loopName + segmentName) with empty string as value for each  key,
    ' this will serve as a template for individual dictionary to be added to dataToBeSavedToDbArr
    Dim templateDict
    Set templateDict = CreateObject("Scripting.Dictionary")
    Dim objConnection, objRecordset
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordset = CreateObject("ADODB.Recordset")
    dbConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strDataPath & ";Extended Properties = excel 12.0;" ' excel 8.0 for excel 2003
    If objConnection.State = 1 Then
        objConnection.Close
    End If
    With objConnection
        .Open dbConnectStr
    End With
    ' fetching records that have not already been used for generating edi records
    'objRecordset.Open "select * from [Scenario 1$] where [Requirement] = 'Y'", objConnection
    ''''debug.print "scenarioName " & scenarioName
    If dependantDataRequired = False Then
        objRecordset.Open "select * from [" & scenarioName & "$] where [Requirement] = 'Y' and [LoopName] not in ('0','2800','10000') ", objConnection
    Else
        objRecordset.Open "select * from [" & scenarioName & "$] where [Requirement] in ('Y','Y,Y,Y') and [LoopName] not in ('0','2800','10000') ", objConnection
    End If
    Do While objRecordset.EOF = False
        ' dictionary object that will store values to be inserted to db for each record
        key1 = objRecordset.Fields("LoopName")
        key2 = objRecordset.Fields("Segment")
        dbKey = key1 & "-" & key2
        On Error Resume Next
        templateDict.Add dbKey, ""
        objRecordset.MoveNext
    Loop
    '''On Error Resume Next
    objRecordset.Close
    objConnection.Close
    For rec_ct = 1 To numberOfRecordsNeeded
        Set dataToBeSavedToDbDict = CreateObject("Scripting.Dictionary")
        If dependantDataRequired = True Then
            Set dataToBeSavedToDbDict1 = CreateObject("Scripting.Dictionary")
            Set dataToBeSavedToDbDict2 = CreateObject("Scripting.Dictionary")
        End If
        ' looping through the templateDict and copying the keys to new dictionaries created to hold values for each records to be created
        For Each tmpl_key In templateDict
            '''''debug.print "Key is " & tmpl_key
            dataToBeSavedToDbDict.Add tmpl_key, ""
            If dependantDataRequired = True Then
                dataToBeSavedToDbDict1.Add tmpl_key, ""
                dataToBeSavedToDbDict2.Add tmpl_key, ""
            End If
        Next
        dataToBeSavedToDbArr.Add rec_ct, dataToBeSavedToDbDict
        If dependantDataRequired = True Then
            Set tempdepDataToBeSavedToDbArr = CreateObject("Scripting.Dictionary")
            tempdepDataToBeSavedToDbArr.Add 1, dataToBeSavedToDbDict1 ' dictionary to save segment information of first dependant
            tempdepDataToBeSavedToDbArr.Add 2, dataToBeSavedToDbDict2 ' dictionary to save segment information of second dependant
            depDataToBeSavedToDbArr.Add rec_ct, tempdepDataToBeSavedToDbArr
        End If
        
    Next
'LogStatement "Function completed"
End Function
Function InsertDictValuesIntoDb(updateDbPath, lob)
' example query - Insert into [MH$]([0-ISA01],[0-ISA06]) values('abc','bcd');
'Dim dataToBeSavedToDbArr
'Dim dataToBeSavedToDbDict
'Set dataToBeSavedToDbDict = CreateObject("Scripting.Dictionary")
'Set dataToBeSavedToDbArr = CreateObject("Scripting.Dictionary")
'
'strDataPath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Supporting Files\Random_Updated_DB.xlsx"
'lob = "MH"
'
'For x = 0 To 2
'    dataToBeSavedToDbDict.RemoveAll
'    dataToBeSavedToDbDict.Add "0-ISA01", "abc"
'    dataToBeSavedToDbDict.Add "0-ISA02", "abc"
'    dataToBeSavedToDbDict.Add "0-ISA03", "abc"
'    dataToBeSavedToDbArr.Add x, dataToBeSavedToDbDict
'Next
'
    Dim objConnection, objRecordset
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordset = CreateObject("ADODB.Recordset")
    dbConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & updateDbPath & ";Extended Properties = excel 12.0;" ' excel 8.0 for excel 2003
    If objConnection.State = 1 Then
        objConnection.Close
    End If
    With objConnection
        .Open dbConnectStr
    End With
    ' looping through the dataToBeSavedToDbArr to create query to be executed to update record in database
    For Each d In dataToBeSavedToDbArr
        If d = Empty Then
            Exit For
        End If
        Query = "Insert into [" & lob & "$]"
        colNamesStr = ""
        colValuesStr = ""
        For Each k In dataToBeSavedToDbArr(d)
            colNamesStr = colNamesStr & "[" & k & "],"
            tempval = dataToBeSavedToDbArr(d)(k)
            If Left(tempval, 1) = "," Then
                tempval = Mid(tempval, 2)
            End If
            colValuesStr = colValuesStr & "'" & tempval & "',"
        Next
        ' removing extra commas at the end of colNameStr and colValuesStr
        If Right(colNamesStr, 1) = "," Then
            colNamesStr = Left(colNamesStr, Len(colNamesStr) - 1)
        End If
        If Right(colValuesStr, 1) = "," Then
            colValuesStr = Left(colValuesStr, Len(colValuesStr) - 1)
        End If
        Query = "Insert into [" & "Main" & "$](" & colNamesStr & ") values(" & colValuesStr & "); "
        ''''deug.print Query
        objConnection.Execute Query
    Next
    objConnection.Close
    Set objRecordset = Nothing
    Set objConnection = Nothing
    ''''''''deug.print "x"
    LogStatement "Function completed"
End Function
Function InsertDictValuesIntoDb_1_(updateDbPath, lob)
' example query - Insert into [MH$]([0-ISA01],[0-ISA06]) values('abc','bcd');
'Dim dataToBeSavedToDbArr
'Dim dataToBeSavedToDbDict
'Set dataToBeSavedToDbDict = CreateObject("Scripting.Dictionary")
'Set dataToBeSavedToDbArr = CreateObject("Scripting.Dictionary")
'
'    updateDbPath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Supporting Files\Random_Updated_DB.xlsx"
'    lob = "MH"
'    MEME_MEDCD_NO = "1234567"
    Dim objExcel
    Dim objDBWorkbook
    Dim objDBSht
    Dim objDBShtRng1Text
    Dim objDBShtRng2Text
    objDBShtRng1Text = ""
    objDBShtRng2Text = ""
    Set objExcel = CreateObject("Excel.Application")
    If FSO.FileExists(updateDbPath) Then
        Set objDBWorkbook = objExcel.Workbooks.Open(updateDbPath)
        Set objDBSht = objDBWorkbook.Worksheets("ColNames")
        Set objDBShtRng1 = objDBSht.Range("A:A")
        Set objDBShtRng2 = objDBSht.Range("B:B")
        For Each c In objDBShtRng1
            If IsNull(c) Or IsEmpty(c) Then
                Exit For
            End If
            objDBShtRng1Text = objDBShtRng1Text & c.Value
        Next
    End If
    objDBWorkbook.Close
    objExcel.Quit
    Set objExcel = Nothing
    Set objDBWorkbook = Nothing
    Set objDBSht = Nothing
    Set objDBShtRng1 = Nothing
    Set objDBShtRng2 = Nothing
    Dim objConnection, objRecordset
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordset = CreateObject("ADODB.Recordset")
    dbConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & updateDbPath & ";Extended Properties = excel 12.0;" ' excel 8.0 for excel 2003
    If objConnection.State = 1 Then
        objConnection.Close
    End If
    With objConnection
        .Open dbConnectStr
    End With
    Query = "Insert into [Main$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N');"
    objConnection.Execute Query
    Query = "Insert into [Another$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N');"
    objConnection.Execute Query
    query_string = ""
    For Each d In dataToBeSavedToDbArr
        column_count = 0
        If d = Empty Then
            Exit For
        End If
        tbl_name = "Main"
        second_tbl_name = "Another"
        query_string = "update [" & tbl_name & "$] set "
        For Each k In dataToBeSavedToDbArr(d)
            If InStr(1, objDBShtRng1Text, k) = 0 And tbl_name = "Main" Then
                ' removing extra commas at the end of colNameStr and colValuesStr
                If Right(query_string, 1) = "," Then
                    query_string = Left(query_string, Len(query_string) - 1)
                End If
                query_string = query_string & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
                ''''deug.print query_string
                objConnection.Execute query_string
                tbl_name = "Another"
                query_string = "update [" & tbl_name & "$] set "
                tempval = dataToBeSavedToDbArr(d)(k)
                If Left(tempval, 1) = "," Then
                    tempval = Mid(tempval, 2)
                End If
                If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                    query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                End If
            Else
                tempval = dataToBeSavedToDbArr(d)(k)
                If Left(tempval, 1) = "," Then
                    tempval = Mid(tempval, 2)
                End If
                If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                    query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                End If
            End If
        Next
        If Right(query_string, 1) = "," Then
            query_string = Left(query_string, Len(query_string) - 1)
        End If
        query_string = query_string & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        ''''deug.print query_string
        objConnection.Execute query_string
    Next
    ''''deug.print "Done"
    objConnection.Close
    Set objRecordset = Nothing
    Set objConnection = Nothing
End Function
Function InsertDictValuesIntoDb_1(updateDbPath, lob)
' example query - Insert into [MH$]([0-ISA01],[0-ISA06]) values('abc','bcd');
'Dim dataToBeSavedToDbArr
'Dim dataToBeSavedToDbDict
'Set dataToBeSavedToDbDict = CreateObject("Scripting.Dictionary")
'Set dataToBeSavedToDbArr = CreateObject("Scripting.Dictionary")
'
'    updateDbPath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Supporting Files\Random_Updated_DB.xlsx"
'    lob = "MH"
'    MEME_MEDCD_NO = "1234567"
    Dim objExcel
    Dim objDBWorkbook
    Dim objDBSht
    Dim objDBShtRng1Text
    Dim objDBShtRng2Text
    objDBShtRng1Text = ""
    objDBShtRng2Text = ""
    Set objExcel = CreateObject("Excel.Application")
    If FSO.FileExists(updateDbPath) Then
        Set objDBWorkbook = objExcel.Workbooks.Open(updateDbPath)
        Set objDBSht = objDBWorkbook.Worksheets("ColNames")
        Set objDBShtRng1 = objDBSht.Range("A:A")
        Set objDBShtRng2 = objDBSht.Range("B:B")
        For Each c In objDBShtRng1
            If IsNull(c) Or IsEmpty(c) Then
                Exit For
            End If
            objDBShtRng1Text = objDBShtRng1Text & c.Value
        Next
    End If
    objDBWorkbook.Close
    objExcel.Quit
    Set objExcel = Nothing
    Set objDBWorkbook = Nothing
    Set objDBSht = Nothing
    Set objDBShtRng1 = Nothing
    Set objDBShtRng2 = Nothing
    Dim objConnection, objRecordset
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordset = CreateObject("ADODB.Recordset")
    dbConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & updateDbPath & ";Extended Properties = excel 12.0;" ' excel 8.0 for excel 2003
    If objConnection.State = 1 Then
        objConnection.Close
    End If
    With objConnection
        .Open dbConnectStr
    End With
    
    If Not dataToBeSavedToDbArr(currentRandKey).exists("2310-LX01") And dependantDataRequired = False Then
        'indicates that enrollment is done without PCP
        Query = "Insert into [Main$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[Dependants],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','Y','N','N');"
        objConnection.Execute Query
        Query = "Insert into [Another$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[Dependants],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','Y','N','N');"
        objConnection.Execute Query
    ElseIf Not dataToBeSavedToDbArr(currentRandKey).exists("2310-LX01") And dependantDataRequired = True Then
        'indicates that enrollment is done without PCP
        Query = "Insert into [Main$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[Dependants],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','Y','Y','N');"
        objConnection.Execute Query
        Query = "Insert into [Another$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[Dependants],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','Y','Y','N');"
        objConnection.Execute Query
    Else
        
        'indicates that enrollment is done with PCP and no dependants
        If dependantDataRequired = False Then
            Query = "Insert into [Main$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[Dependants],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','Y','N','Y');"
            objConnection.Execute Query
            Query = "Insert into [Another$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[Dependants],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','Y','N','Y');"
            objConnection.Execute Query
        Else
            'indicates that enrollment is done with PCP and dependants
            Query = "Insert into [Main$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[Dependants],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','Y','Y','Y');"
            objConnection.Execute Query
            Query = "Insert into [Another$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[Dependants],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','Y','Y','Y');"
            objConnection.Execute Query
        End If
        
    End If
   
    query_string = ""
    For Each d In dataToBeSavedToDbArr
        column_count = 0
        If d = Empty Then
            Exit For
        End If
        tbl_name = "Main"
        second_tbl_name = "Another"
        query_string = "update [" & tbl_name & "$] set "
        For Each k In dataToBeSavedToDbArr(d)
            If InStr(1, objDBShtRng1Text, k) = 0 And tbl_name = "Main" Then
                ' removing extra commas at the end of colNameStr and colValuesStr
                If Right(query_string, 1) = "," Then
                    query_string = Left(query_string, Len(query_string) - 1)
                End If
                query_string = query_string & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
                ''''deug.print query_string
                objConnection.Execute query_string
                tbl_name = "Another"
                query_string = "update [" & tbl_name & "$] set "
                tempval = dataToBeSavedToDbArr(d)(k)
                If Left(tempval, 1) = "," Then
                    tempval = Mid(tempval, 2)
                End If
                If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                    query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                End If
            Else
                tempval = dataToBeSavedToDbArr(d)(k)
                If Left(tempval, 1) = "," Then
                    tempval = Mid(tempval, 2)
                End If
                If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                    query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                End If
            End If
        Next
        If Right(query_string, 1) = "," Then
            query_string = Left(query_string, Len(query_string) - 1)
        End If
        query_string = query_string & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        ''''deug.print query_string
        objConnection.Execute query_string
    Next
    
    MEME_MEDCD_NOS = Str(MEME_MEDCD_NO)
    
    If dependantDataRequired = True Then
        For dependant_ct = 1 To 2
            DEP_MEME_MEDCD_NO = depDataToBeSavedToDbArr(currentRandKey)(dependant_ct)("2000-REF02")
            If InStr(1, DEP_MEME_MEDCD_NO, "@@@@@") <> 0 Then
                MEME_MEDCD_NO_LIST = Split(DEP_MEME_MEDCD_NO, "@@@@@")
                For Each mem_no In MEME_MEDCD_NO_LIST
                    If mem_no <> MEME_MEDCD_NO And Len(mem_no) = Len(MEME_MEDCD_NO) Then
                        If InStr(1, CStr(MEME_MEDCD_NOS), CStr(mem_no)) = 0 Then
                            MEME_MEDCD_NO = mem_no
                            MEME_MEDCD_NOS = MEME_MEDCD_NOS & "@" & MEME_MEDCD_NO
                            Exit For
                        End If
                    End If
                    
                Next
                
                
            End If
            If Not dataToBeSavedToDbArr(currentRandKey).exists("2310-LX01") Then
                'indicates that enrollment is done without PCP
                Query = "Insert into [Main$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','N','N');"
                objConnection.Execute Query
                Query = "Insert into [Another$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','N','N');"
                objConnection.Execute Query
            Else
                'indicates that enrollment is done with PCP
                Query = "Insert into [Main$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','N','Y');"
                objConnection.Execute Query
                Query = "Insert into [Another$]([UniqueID],[LOB],[Used],[Termed],[Mem_Termed],[Subscriber],[PCP]) values ('" & MEME_MEDCD_NO & "','" & lineOfBusiness & "','N','N','N','N','Y');"
                objConnection.Execute Query
                
            End If
            query_string = ""
            For Each d In depDataToBeSavedToDbArr(currentRandKey)(dependant_ct).keys
                column_count = 0
                If d = Empty Then
                    Exit For
                End If
                tbl_name = "Main"
                second_tbl_name = "Another"
                query_string = "update [" & tbl_name & "$] set "
                For Each k In depDataToBeSavedToDbArr(currentRandKey)(dependant_ct).keys
                    If InStr(1, objDBShtRng1Text, k) = 0 And tbl_name = "Main" Then
                        ' removing extra commas at the end of colNameStr and colValuesStr
                        If Right(query_string, 1) = "," Then
                            query_string = Left(query_string, Len(query_string) - 1)
                        End If
                        query_string = query_string & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
                        ''''deug.print query_string
                        objConnection.Execute query_string
                        tbl_name = "Another"
                        query_string = "update [" & tbl_name & "$] set "
                        tempval = depDataToBeSavedToDbArr(currentRandKey)(dependant_ct)(k)
                        If Left(tempval, 1) = "," Then
                            tempval = Mid(tempval, 2)
                        End If
                        If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                            query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                        End If
                    Else
                        tempval = depDataToBeSavedToDbArr(currentRandKey)(dependant_ct)(k)
                        If Left(tempval, 1) = "," Then
                            tempval = Mid(tempval, 2)
                        End If
                        If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                            query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                        End If
                    End If
                Next
                If Right(query_string, 1) = "," Then
                    query_string = Left(query_string, Len(query_string) - 1)
                End If
                query_string = query_string & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
                ''''deug.print query_string
                objConnection.Execute query_string
                Exit For
            Next
        
            
        Next
    End If
    
    
    
    ''''deug.print "Done"
    objConnection.Close
    Set objRecordset = Nothing
    Set objConnection = Nothing
End Function

Function InsertDictValuesIntoDb_2_(updateDbPath, lob)
' example query - Insert into [MH$]([0-ISA01],[0-ISA06]) values('abc','bcd');
'Dim dataToBeSavedToDbArr
'Dim dataToBeSavedToDbDict
'Set dataToBeSavedToDbDict = CreateObject("Scripting.Dictionary")
'Set dataToBeSavedToDbArr = CreateObject("Scripting.Dictionary")
'
'    updateDbPath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Supporting Files\Random_Updated_DB.xlsx"
'    lob = "MH"
'    MEME_MEDCD_NO = "1234567"
    Dim objExcel
    Dim objDBWorkbook
    Dim objDBSht
    Dim objDBShtRng1Text
    Dim objDBShtRng2Text
    objDBShtRng1Text = ""
    objDBShtRng2Text = ""
    Set objExcel = CreateObject("Excel.Application")
    If FSO.FileExists(updateDbPath) Then
        Set objDBWorkbook = objExcel.Workbooks.Open(updateDbPath)
        Set objDBSht = objDBWorkbook.Worksheets("ColNames")
        Set objDBShtRng1 = objDBSht.Range("A:A")
        Set objDBShtRng2 = objDBSht.Range("B:B")
        For Each c In objDBShtRng1
            If IsNull(c) Or IsEmpty(c) Then
                Exit For
            End If
            objDBShtRng1Text = objDBShtRng1Text & c.Value
        Next
    End If
    objDBWorkbook.Close
    objExcel.Quit
    Set objExcel = Nothing
    Set objDBWorkbook = Nothing
    Set objDBSht = Nothing
    Set objDBShtRng1 = Nothing
    Set objDBShtRng2 = Nothing
    Dim objConnection, objRecordset
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordset = CreateObject("ADODB.Recordset")
    dbConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & updateDbPath & ";Extended Properties = excel 12.0;" ' excel 8.0 for excel 2003
    If objConnection.State = 1 Then
        objConnection.Close
    End If
    With objConnection
        .Open dbConnectStr
    End With
    If InStr(1, scenarioName, "Scenario_2") <> 0 Then
        Query = "update [Main$] set [Used] = 'Y' " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        objConnection.Execute Query
        Query = "update [Another$] set [Used] = 'Y' " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
    ElseIf InStr(1, scenarioName, "Scenario_3") <> 0 Then
        Query = "update [Main$] set [Termed] = 'Y' " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        objConnection.Execute Query
        Query = "update [Another$] set [Termed] = 'Y'  " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
    ElseIf InStr(1, scenarioName, "Scenario_4") <> 0 Then
        Query = "update [Main$] set [Mem_Termed] = 'Y' " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        objConnection.Execute Query
        Query = "update [Another$] set [Mem_Termed] = 'Y'  " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
    ElseIf InStr(1, scenarioName, "Scenario_5") <> 0 Or InStr(1, scenarioName, "Scenario_6") <> 0 Or InStr(1, scenarioName, "Scenario_8") <> 0 Then
        Query = "update [Main$] set [Mem_Termed] = 'N' " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        objConnection.Execute Query
        Query = "update [Another$] set [Mem_Termed] = 'N'  " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
    End If
    objConnection.Execute Query
    query_string = ""
    For Each d In dataToBeSavedToDbArr
        column_count = 0
        If d = Empty Then
            Exit For
        End If
        tbl_name = "Main"
        second_tbl_name = "Another"
        query_string = "update [" & tbl_name & "$] set "
        For Each k In dataToBeSavedToDbArr(d)
            If InStr(1, objDBShtRng1Text, k) = 0 And tbl_name = "Main" Then
                ' removing extra commas at the end of colNameStr and colValuesStr
                If Right(query_string, 1) = "," Then
                    query_string = Left(query_string, Len(query_string) - 1)
                End If
                query_string = query_string & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
                ''''deug.print query_string
                objConnection.Execute query_string
                tbl_name = "Another"
                query_string = "update [" & tbl_name & "$] set "
                tempval = dataToBeSavedToDbArr(d)(k)
                If Left(tempval, 1) = "," Then
                    tempval = Mid(tempval, 2)
                End If
                If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                    query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                End If
            Else
                tempval = dataToBeSavedToDbArr(d)(k)
                If Left(tempval, 1) = "," Then
                    tempval = Mid(tempval, 2)
                End If
                If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                    query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                End If
            End If
        Next
        If Right(query_string, 1) = "," Then
            query_string = Left(query_string, Len(query_string) - 1)
        End If
        query_string = query_string & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        ''''deug.print query_string
        objConnection.Execute query_string
    Next
    ''''deug.print "Done"
    objConnection.Close
    Set objRecordset = Nothing
    Set objConnection = Nothing
End Function


Function InsertDictValuesIntoDb_2(updateDbPath, lob)
' example query - Insert into [MH$]([0-ISA01],[0-ISA06]) values('abc','bcd');
'Dim dataToBeSavedToDbArr
'Dim dataToBeSavedToDbDict
'Set dataToBeSavedToDbDict = CreateObject("Scripting.Dictionary")
'Set dataToBeSavedToDbArr = CreateObject("Scripting.Dictionary")
'
'    updateDbPath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Supporting Files\Random_Updated_DB.xlsx"
'    lob = "MH"
'    MEME_MEDCD_NO = "1234567"
    Dim objExcel
    Dim objDBWorkbook
    Dim objDBSht
    Dim objDBShtRng1Text
    Dim objDBShtRng2Text
    objDBShtRng1Text = ""
    objDBShtRng2Text = ""
    Set objExcel = CreateObject("Excel.Application")
    If FSO.FileExists(updateDbPath) Then
        Set objDBWorkbook = objExcel.Workbooks.Open(updateDbPath)
        Set objDBSht = objDBWorkbook.Worksheets("ColNames")
        Set objDBShtRng1 = objDBSht.Range("A:A")
        Set objDBShtRng2 = objDBSht.Range("B:B")
        For Each c In objDBShtRng1
            If IsNull(c) Or IsEmpty(c) Then
                Exit For
            End If
            objDBShtRng1Text = objDBShtRng1Text & c.Value
        Next
    End If
    objDBWorkbook.Close
    objExcel.Quit
    Set objExcel = Nothing
    Set objDBWorkbook = Nothing
    Set objDBSht = Nothing
    Set objDBShtRng1 = Nothing
    Set objDBShtRng2 = Nothing
    Dim objConnection, objRecordset
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordset = CreateObject("ADODB.Recordset")
    dbConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & updateDbPath & ";Extended Properties = excel 12.0;" ' excel 8.0 for excel 2003
    If objConnection.State = 1 Then
        objConnection.Close
    End If
    With objConnection
        .Open dbConnectStr
    End With
    
    
    
    If InStr(1, scenarioName, "Scenario_2") <> 0 Or InStr(1, scenarioName, "Scenario_11") <> 0 Then
        Query = "update [Main$] set [Used] = 'Y' " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        objConnection.Execute Query
        Query = "update [Another$] set [Used] = 'Y' " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
    ElseIf InStr(1, scenarioName, "Scenario_3") <> 0 Or InStr(1, scenarioName, "Scenario_12") <> 0 Then
        Query = "update [Main$] set [Termed] = 'Y' " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        objConnection.Execute Query
        Query = "update [Another$] set [Termed] = 'Y'  " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
    ElseIf InStr(1, scenarioName, "Scenario_4") <> 0 Or InStr(1, scenarioName, "Scenario_10") <> 0 Then
        Query = "update [Main$] set [Mem_Termed] = 'Y' " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        objConnection.Execute Query
        Query = "update [Another$] set [Mem_Termed] = 'Y'  " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
    ElseIf InStr(1, scenarioName, "Scenario_5") <> 0 Or InStr(1, scenarioName, "Scenario_6") <> 0 Or InStr(1, scenarioName, "Scenario_9") <> 0 Or InStr(1, scenarioName, "Scenario_17") <> 0 Or InStr(1, scenarioName, "Scenario_13") <> 0 Or InStr(1, scenarioName, "Scenario_14") <> 0 Then
        Query = "update [Main$] set [Mem_Termed] = 'N' " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        objConnection.Execute Query
        Query = "update [Another$] set [Mem_Termed] = 'N'  " & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
    End If
    objConnection.Execute Query
    query_string = ""
    For Each d In dataToBeSavedToDbArr
        column_count = 0
        If d = Empty Then
            Exit For
        End If
        tbl_name = "Main"
        second_tbl_name = "Another"
        query_string = "update [" & tbl_name & "$] set "
        For Each k In dataToBeSavedToDbArr(d)
            If InStr(1, objDBShtRng1Text, k) = 0 And tbl_name = "Main" Then
                ' removing extra commas at the end of colNameStr and colValuesStr
                If Right(query_string, 1) = "," Then
                    query_string = Left(query_string, Len(query_string) - 1)
                End If
                query_string = query_string & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
                ''''deug.print query_string
                objConnection.Execute query_string
                tbl_name = "Another"
                query_string = "update [" & tbl_name & "$] set "
                tempval = dataToBeSavedToDbArr(d)(k)
                If Left(tempval, 1) = "," Then
                    tempval = Mid(tempval, 2)
                End If
                If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                    query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                End If
            Else
                tempval = dataToBeSavedToDbArr(d)(k)
                If Left(tempval, 1) = "," Then
                    tempval = Mid(tempval, 2)
                End If
                If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                    query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                End If
            End If
        Next
        If Right(query_string, 1) = "," Then
            query_string = Left(query_string, Len(query_string) - 1)
        End If
        query_string = query_string & " where [UniqueID] = '" & MEME_MEDCD_NO & "';"
        ''''deug.print query_string
        objConnection.Execute query_string
    Next
    
    ' the policy holder member id
    MEME_MEDCD_NOS = Str(MEME_MEDCD_NO) ' variable to store the list of the policy holder's and dependant policy numbers
    
    ' starting to insert the dependants details
    If dependantDataRequired = True Then
        For dependant_ct = 1 To 2
            DEP_MEME_MEDCD_NO = depDataToBeSavedToDbArr(currentRandKey)(dependant_ct)("2000-REF02")
            If InStr(1, DEP_MEME_MEDCD_NO, "@@@@@") <> 0 Then
                MEME_MEDCD_NO_LIST = Split(DEP_MEME_MEDCD_NO, "@@@@@")
                For Each mem_no In MEME_MEDCD_NO_LIST
                    If mem_no <> MEME_MEDCD_NO And Len(mem_no) = Len(MEME_MEDCD_NO) Then
                        If InStr(1, CStr(MEME_MEDCD_NOS), CStr(mem_no)) = 0 Then
                            DEP_MEME_MEDCD_NO = mem_no
                            MEME_MEDCD_NOS = MEME_MEDCD_NOS & "@" & DEP_MEME_MEDCD_NO
                            Exit For
                        End If
                    End If
                    
                Next
            End If
            
            If InStr(1, scenarioName, "Scenario_2") <> 0 Or InStr(1, scenarioName, "Scenario_11") <> 0 Then
                Query = "update [Main$] set [Used] = 'Y' " & " where [UniqueID] = '" & DEP_MEME_MEDCD_NO & "';"
                objConnection.Execute Query
                Query = "update [Another$] set [Used] = 'Y' " & " where [UniqueID] = '" & DEP_MEME_MEDCD_NO & "';"
            ElseIf InStr(1, scenarioName, "Scenario_3") <> 0 Or InStr(1, scenarioName, "Scenario_12") <> 0 Then
                Query = "update [Main$] set [Termed] = 'Y' " & " where [UniqueID] = '" & DEP_MEME_MEDCD_NO & "';"
                objConnection.Execute Query
                Query = "update [Another$] set [Termed] = 'Y'  " & " where [UniqueID] = '" & DEP_MEME_MEDCD_NO & "';"
            ElseIf InStr(1, scenarioName, "Scenario_4") <> 0 Or InStr(1, scenarioName, "Scenario_10") <> 0 Then
                Query = "update [Main$] set [Mem_Termed] = 'Y' " & " where [UniqueID] = '" & DEP_MEME_MEDCD_NO & "';"
                objConnection.Execute Query
                Query = "update [Another$] set [Mem_Termed] = 'Y'  " & " where [UniqueID] = '" & DEP_MEME_MEDCD_NO & "';"
            ElseIf InStr(1, scenarioName, "Scenario_5") <> 0 Or InStr(1, scenarioName, "Scenario_6") <> 0 Or InStr(1, scenarioName, "Scenario_9") <> 0 Or InStr(1, scenarioName, "Scenario_17") <> 0 Or InStr(1, scenarioName, "Scenario_13") <> 0 Or InStr(1, scenarioName, "Scenario_14") <> 0 Then
                Query = "update [Main$] set [Mem_Termed] = 'N' " & " where [UniqueID] = '" & DEP_MEME_MEDCD_NO & "';"
                objConnection.Execute Query
                Query = "update [Another$] set [Mem_Termed] = 'N'  " & " where [UniqueID] = '" & DEP_MEME_MEDCD_NO & "';"
            End If
            
            objConnection.Execute Query
            query_string = ""
            
            For Each d In depDataToBeSavedToDbArr(currentRandKey)(dependant_ct).keys
                column_count = 0
                If d = Empty Then
                    Exit For
                End If
                tbl_name = "Main"
                second_tbl_name = "Another"
                query_string = "update [" & tbl_name & "$] set "
                For Each k In depDataToBeSavedToDbArr(currentRandKey)(dependant_ct).keys
                    If InStr(1, objDBShtRng1Text, k) = 0 And tbl_name = "Main" Then
                        ' removing extra commas at the end of colNameStr and colValuesStr
                        If Right(query_string, 1) = "," Then
                            query_string = Left(query_string, Len(query_string) - 1)
                        End If
                        query_string = query_string & " where [UniqueID] = '" & DEP_MEME_MEDCD_NO & "';"
                        ''''deug.print query_string
                        objConnection.Execute query_string
                        tbl_name = "Another"
                        query_string = "update [" & tbl_name & "$] set "
                        tempval = depDataToBeSavedToDbArr(currentRandKey)(dependant_ct)(k)
                        If Left(tempval, 1) = "," Then
                            tempval = Mid(tempval, 2)
                        End If
                        If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                            query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                        End If
                    Else
                        tempval = depDataToBeSavedToDbArr(currentRandKey)(dependant_ct)(k)
                        If Left(tempval, 1) = "," Then
                            tempval = Mid(tempval, 2)
                        End If
                        If Not IsNull(tempval) And Not IsEmpty(tempval) And Not tempval = "" Then
                            query_string = query_string & "[" & k & "] = " & "'" & tempval & "',"
                        End If
                    End If
                Next
                If Right(query_string, 1) = "," Then
                    query_string = Left(query_string, Len(query_string) - 1)
                End If
                query_string = query_string & " where [UniqueID] = '" & DEP_MEME_MEDCD_NO & "';"
                ''''deug.print query_string
                objConnection.Execute query_string
                Exit For
            Next
        
            
        Next
    End If
    
    
    
    ''''deug.print "Done"
    objConnection.Close
    Set objRecordset = Nothing
    Set objConnection = Nothing
End Function



Function PreprocessTemplateFile(rootPath, lob, scenarioName)
    'Set FSO = CreateObject("Scripting.Filesystemobject")
        LogStatement "Starting to pre-process template file"
        templateFilePath = rootPath & Chr(92) & templateFileName & ".xlsx"
        newTemplateFilePath = rootPath & Chr(92) & templateFileName & "_temp.xlsx"
        On Error Resume Next
        If FSO.FileExists(newTemplateFilePath) Then
            FSO.DeleteFile newTemplateFilePath
        End If
        Dim objExcel
        ' variable for Excel workbook
        Dim objTemplateWorkbook
        ' variable for Excel worksheet
        Dim objTemplateWorksheet
        ' variable for Excel worksheet Range
        Dim objTemplateWorksheetRange, objTemplateWorksheetRangeRows
        Set objExcel = CreateObject("Excel.Application")
        If FSO.FileExists(templateFilePath) Then
            Set objTemplateWorkbook = objExcel.Workbooks.Open(templateFilePath)
            Set objTemplateWorksheet = objTemplateWorkbook.Worksheets(scenarioName)
        Else
            LogStatement "Template excel file not found "
        End If
        Set objTemplateWorksheetRange = objTemplateWorksheet.UsedRange
        objTemplateWorksheetRangeRows = objTemplateWorksheetRange.Rows.Count
        st_row = 0
        ed_row = 0
        st_val = ""
        same_rows_count = 0
        number_of_values = 0
        LogStatement "getting the range of rows that need to be treated as same"
        ' getting the range of rows that need to be treated as same even if the segment names are different indicated by "S" in first column
        For row_ct = 2 To objTemplateWorksheetRangeRows
            If objTemplateWorksheet.Cells(row_ct, 1) = "S" And st_row = 0 Then
                st_row = row_ct
                st_val = objTemplateWorksheet.Cells(row_ct, 7)
            ElseIf objTemplateWorksheet.Cells(row_ct, 1) <> "S" And st_row <> 0 And ed_row = 0 Then
                ed_row = row_ct - 1
            End If
        Next
        If st_row <> 0 Then
            ' getting the difference
            same_rows_count = ed_row - st_row + 1
            LogStatement CStr(same_rows_count)
        End If
        ' checking if multiple values are present in st_val, this indicates that the row range needs to be repeated multiple times
        If InStr(1, st_val, ",") <> 0 Then
            number_of_values = UBound(Split(st_val, ","))
        End If
        saveAsNewFile = False
        If number_of_values > 0 Then
            saveAsNewFile = True
            For r_ct = 1 To number_of_values
                insert_row = st_row + (same_rows_count * r_ct)
                LogStatement "insert_row " & insert_row
                objTemplateWorksheet.Range("A" & st_row, "A" & ed_row).EntireRow.Copy
                objTemplateWorksheet.Range("A" & insert_row).Insert
            Next
        End If
        If number_of_values > 0 Then
            LogStatement "number_of_values " & CStr(number_of_values)
            For r_ct = 0 To number_of_values
                For ch_row = st_row To ed_row
                    new_val = objTemplateWorksheet.Cells(ch_row, 7)
                    If InStr(1, new_val, ",") <> 0 Then
                        On Error Resume Next
                        new_val = Split(new_val, ",")(r_ct)
                        ' updating the value only if the is value to be included in that row
                        ' in some scenarios, some of the segments within 2700A are not repeated the number of times we have LX
                        ' example is edi template for CCA where LX segment has 4 values but DTP segment has only 3
                        ' in such cases we will exclude that segment by marking that segment as "N"
                        If Err.Number = 0 Then
                            objTemplateWorksheet.Cells(ch_row, 7) = new_val
                        Else
                            objTemplateWorksheet.Cells(ch_row, 5) = "N"
                            objTemplateWorksheet.Cells(ch_row, 7) = ""
                        End If
                        On Error GoTo 0
                    End If
                Next
                st_row = ed_row + 1
                ed_row = ed_row + same_rows_count
            Next
        End If
        LogStatement "st_row " & st_row
        LogStatement "ed_row " & ed_row
        LogStatement "same_rows_count " & same_rows_count
        If saveAsNewFile = True Then
            objTemplateWorkbook.SaveAs newTemplateFilePath
            objTemplateWorkbook.Close
        Else
            objTemplateWorkbook.Close
        End If
        objExcel.Quit
        Set objTemplateWorkbook = Nothing
        Set objExcel = Nothing
        If Err.Number <> 0 Then
            LogStatement Err.Description
        End If
        LogStatement "PreprocessTemplateFile Function completed"
        On Error GoTo 0
End Function

Sub PreprocessTemplateFile_()

rootPath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project"
lob = "DCH"
scenarioName = "DCH_Scenario_1"
Set FSO = CreateObject("Scripting.Filesystemobject")
templateFileName = "EDITemplate"
    
        'logstatement "Starting to pre-process template file"
        templateFilePath = rootPath & Chr(92) & templateFileName & ".xlsx"
        newTemplateFilePath = rootPath & Chr(92) & templateFileName & "_temp.xlsx"
        On Error Resume Next
        If FSO.FileExists(newTemplateFilePath) Then
            FSO.DeleteFile newTemplateFilePath
        End If
        Dim objExcel
        ' variable for Excel workbook
        Dim objTemplateWorkbook
        ' variable for Excel worksheet
        Dim objTemplateWorksheet
        ' variable for Excel worksheet Range
        Dim objTemplateWorksheetRange, objTemplateWorksheetRangeRows
        Set objExcel = CreateObject("Excel.Application")
        If FSO.FileExists(templateFilePath) Then
            Set objTemplateWorkbook = objExcel.Workbooks.Open(templateFilePath)
            Set objTemplateWorksheet = objTemplateWorkbook.Worksheets(scenarioName)
        Else
            'logstatement "Template excel file not found "
        End If
        Set objTemplateWorksheetRange = objTemplateWorksheet.UsedRange
        objTemplateWorksheetRangeRows = objTemplateWorksheetRange.Rows.Count
        st_row = 0
        ed_row = 0
        st_val = ""
        same_rows_count = 0
        number_of_values = 0
        'logstatement "getting the range of rows that need to be treated as same"
        ' getting the range of rows that need to be treated as same even if the segment names are different indicated by "S" in first column
        For row_ct = 2 To objTemplateWorksheetRangeRows
            If objTemplateWorksheet.Cells(row_ct, 1) = "S" And st_row = 0 Then
                st_row = row_ct
                st_val = objTemplateWorksheet.Cells(row_ct, 7)
            ElseIf objTemplateWorksheet.Cells(row_ct, 1) <> "S" And st_row <> 0 And ed_row = 0 Then
                ed_row = row_ct - 1
            End If
        Next
        If st_row <> 0 Then
            ' getting the difference
            same_rows_count = ed_row - st_row + 1
            'logstatement CStr(same_rows_count)
        End If
        ' checking if multiple values are present in st_val, this indicates that the row range needs to be repeated multiple times
        If InStr(1, st_val, ",") <> 0 Then
            number_of_values = UBound(Split(st_val, ","))
        End If
        saveAsNewFile = False
        If number_of_values > 0 Then
            saveAsNewFile = True
            For r_ct = 1 To number_of_values
                insert_row = st_row + (same_rows_count * r_ct)
                'logstatement "insert_row " & insert_row
                objTemplateWorksheet.Range("A" & st_row, "A" & ed_row).EntireRow.Copy
                objTemplateWorksheet.Range("A" & insert_row).Insert
            Next
        End If
        
        
        
        
        If number_of_values > 0 Then
            'logstatement "number_of_values " & CStr(number_of_values)
            For r_ct = 0 To number_of_values
                For ch_row = st_row To ed_row
                    new_val = objTemplateWorksheet.Cells(ch_row, 7)
                    If InStr(1, new_val, ",") <> 0 Then
                        On Error Resume Next
                        new_val = Split(new_val, ",")(r_ct)
                        ' updating the value only if the is value to be included in that row
                        ' in some scenarios, some of the segments within 2700A are not repeated the number of times we have LX
                        ' example is edi template for CCA where LX segment has 4 values but DTP segment has only 3
                        ' in such cases we will exclude that segment by marking that segment as "N"
                        If Err.Number = 0 Then
                            objTemplateWorksheet.Cells(ch_row, 7) = new_val
                        Else
                            objTemplateWorksheet.Cells(ch_row, 5) = "N"
                            objTemplateWorksheet.Cells(ch_row, 7) = ""
                        End If
                        On Error GoTo 0
                    End If
                Next
                st_row = ed_row + 1
                ed_row = ed_row + same_rows_count
            Next
        End If
        'logstatement "st_row " & st_row
        'logstatement "ed_row " & ed_row
        'logstatement "same_rows_count " & same_rows_count
        If saveAsNewFile = True Then
            objTemplateWorkbook.SaveAs newTemplateFilePath
            objTemplateWorkbook.Close
        Else
            objTemplateWorkbook.Close
        End If
        objExcel.Quit
        Set objTemplateWorkbook = Nothing
        Set objExcel = Nothing
        If Err.Number <> 0 Then
            'logstatement Err.Description
        End If
        'logstatement "PreprocessTemplateFile Function completed"
        On Error GoTo 0
End Sub




Function changeName(nameStr)
    ' function will convert the string argument to uppercase and get the length of the string
    ' using the length choose a random number and choose corresponding character from index of the input argument
    ' get the ascii code for the character and get the next character by adding 1 to it and convert it to character and replace the original character in the input
    ' string with the next character and return as output
    nameStr = UCase(nameStr)
    lenOfNameStr = Len(nameStr)
    newNameStr = ""
    If lenOfNameStr > 1 Then
        Randomize
        r_1 = CInt(Rnd(lenOfNameStr) * 10) + 1
        If r_1 > lenOfNameStr Then
            r_1 = lenOfNameStr
        End If
        chr_1 = Mid(nameStr, r_1, 1)
        If Asc(chr_1) >= 65 And Asc(chr_1) < 90 Then
            chr_1 = Chr(Asc(chr_1) + 1)
        ElseIf Asc(chr_1) = 90 Then
            chr_1 = "A"
        End If
        For x = 1 To lenOfNameStr
            If x = r_1 Then
                newNameStr = newNameStr & chr_1
            Else
                newNameStr = newNameStr & Mid(nameStr, x, 1)
            End If
        Next
    ElseIf lenOfNameStr = 1 Then
        If Asc(nameStr) >= 65 And Asc(nameStr) < 90 Then
            newNameStr = Chr(Asc(nameStr) + 1)
        ElseIf Asc(nameStr) = 90 Then
            newNameStr = "A"
        End If
    End If
    ''''deug.print newNameStr
    changeName = newNameStr
End Function

Function changeDOB(dobStr)
    newdobStr = ""
    lastDigit = CInt(Right(dobStr, 1))
    secondlastDigit = CInt(Mid(dobStr, Len(dobStr) - 1, 1))
    If secondlastDigit <> 3 Then
        newlastDigit = Left(CStr(lastDigit + 1), 1)
        newdobStr = Left(dobStr, Len(dobStr) - 1) & newlastDigit
    Else
        newsecondlastDigit = "2"
        newdobStr = Left(dobStr, Len(dobStr) - 2) & newsecondlastDigit & lastDigit
    End If
    changeDOB = newdobStr
End Function

Function changeSSN(inSsn)
    lenSsn = Len(inSsn)
    newSsn = ""
    For x = lenSsn To 1 Step -1
        newSsn = newSsn & Mid(inSsn, x, 1)
    Next
    ''''deug.print newSsn
    changeSSN = newSsn
End Function
Function changeAddr1(addr1Str)
    newaddr1Str = ""
    Randomize
    r_1 = CStr(CInt(Rnd() * 10) + 1)
    If InStr(1, addr1Str, " ") = 0 Then
        newaddr1Str = r_1 & addr1Str
    Else
        ' getting the numeric part of the address and string part in two separate variables
        addr1Str1 = Split(addr1Str, Chr(32), 2)(0) ' chr(32) = space
        addr1Str2 = Split(addr1Str, " ", 2)(1)
        If IsNumeric(addr1Str1) Then
            ct = 0
            Do Until r_1 <> Left(addr1Str1, 1)
                r_1 = CStr(CInt(Rnd() * 10) + 1)
                ct = ct + 1
                If ct > 250 Then
                    Exit Do
                End If
            Loop
            newaddr1Str = r_1 & Mid(addr1Str1, 2) & " " & addr1Str2
        End If
    End If
    changeAddr1 = newaddr1Str
End Function
Function GetStoredRandomDataValues(noOfData, randomDataFilePath)

'Sub GetStoredRandomDataValues()
'   lineOfBusiness = "CCA"
'   memberIdLength = 12
'    noOfData = 1
'    scenarioName = "CCA_Scenario_17"
'    randomDataFilePath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Random_Updated_DB.xlsx"
    
    LogStatement "Started execution of function GetStoredRandomDataValues"
    Dim objConnection, objRecordset
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordset = CreateObject("ADODB.Recordset")
    Set storedDataDict = CreateObject("Scripting.Dictionary")
    Set storedDataDict1 = CreateObject("Scripting.Dictionary")
    Set storedDataDict2 = CreateObject("Scripting.Dictionary")
    
    If dependantDataRequired = True Then
        Set dependantRandomDataDict = CreateObject("Scripting.Dictionary")
    End If
    
    
    dbConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & randomDataFilePath & ";Extended Properties = excel 12.0;" ' excel 8.0 for excel 2003
    If objConnection.State = 1 Then
        objConnection.Close
    End If
    With objConnection
        .Open dbConnectStr
    End With
    Dim memberids
    'this variable is the primary key for the database
    uid_field_name = "UniqueID"
    
    memberids = "("
    ' fetching records that have not already been used for generating edi records from Main sheet
    If InStr(1, scenarioName, "Scenario_2") <> 0 Then
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Used] = 'N' and [Termed] = 'N' and [Mem_Termed] = 'N' and [Subscriber] = 'Y' and [Dependants] = 'N' and [PCP] = 'N' and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
    ElseIf InStr(1, scenarioName, "Scenario_3") <> 0 Then
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Termed] = 'N' and [Mem_Termed] = 'N' and [Subscriber] = 'Y' and [Dependants] = 'N' and [PCP] = 'N' and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
    ElseIf InStr(1, scenarioName, "Scenario_4") <> 0 Or InStr(1, scenarioName, "Scenario_6") <> 0 Then
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Termed] = 'N' and [Mem_Termed] = 'N' and [Subscriber] = 'Y' and [Dependants] = 'N' and [PCP] = 'N' and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
    ElseIf InStr(1, scenarioName, "Scenario_5") <> 0 Then
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Termed] = 'N' and [Mem_Termed] = 'Y' and [Subscriber] = 'Y' and [Dependants] = 'N' and [PCP] = 'N' and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
    ElseIf InStr(1, scenarioName, "Scenario_9") <> 0 Then
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Termed] = 'N' and [Mem_Termed] = 'Y' and [Subscriber] = 'Y' and [Dependants] = 'Y' and [PCP] = 'N' and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
    ElseIf InStr(1, scenarioName, "Scenario_17") <> 0 Then ' getting only members details who are subscribers and enrolled with pcp
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Termed] = 'N' and [Mem_Termed] = 'N' and [Subscriber] = 'Y' and [Dependants] = 'N' and [PCP] = 'Y'  and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
    ElseIf InStr(1, scenarioName, "Scenario_10") <> 0 Or InStr(1, scenarioName, "Scenario_6") <> 0 Then
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Termed] = 'N' and [Mem_Termed] = 'N' and [Subscriber] = 'Y' and [Dependants] = 'Y' and [PCP] = 'N' and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
    ElseIf InStr(1, scenarioName, "Scenario_11") <> 0 Then
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Used] = 'N' and [Termed] = 'N' and [Mem_Termed] = 'N' and [Subscriber] = 'Y' and [Dependants] = 'Y' and [PCP] = 'N' and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
    ElseIf InStr(1, scenarioName, "Scenario_12") <> 0 Then
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Termed] = 'N' and [Mem_Termed] = 'N' and [Subscriber] = 'Y' and [Dependants] = 'Y' and [PCP] = 'N' and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
    ElseIf InStr(1, scenarioName, "Scenario_13") <> 0 Then ' getting only members details who are subscribers and enrolled with pcp
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Termed] = 'N' and [Mem_Termed] = 'N' and [Subscriber] = 'Y' and [Dependants] = 'Y' and [PCP] = 'Y'  and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
    ElseIf InStr(1, scenarioName, "Scenario_14") <> 0 Then ' getting only members details who are subscribers and enrolled with pcp
        objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Termed] = 'N' and [Mem_Termed] = 'N' and [Subscriber] = 'Y' and [Dependants] = 'Y' and [PCP] = 'N'  and [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & ";", objConnection
   
    End If

    
    'objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Used] = 'N'", objConnection
    '''On Error Resume Next
    fields_count = objRecordset.Fields.Count
    ' looping numberOfRecordsNeeded times to create dictionaries with column names as keys
    For rec_ct = 1 To noOfData
        ' creating dictionay to hold values from each row of the recordset with column names as keys
        Set randomSubDict = CreateObject("Scripting.Dictionary")
        
        
        ' looping to add individual column names as keys into dictionary
        For Each field_name In objRecordset.Fields
            ''''''deug.print field_name.Name
            randomSubDict.Add field_name.Name, ""
            
        Next
        '''''''deug.print rec_ct
        storedDataDict1.Add rec_ct, randomSubDict
        
    Next
    ct = 1
    Do While objRecordset.EOF = False
        For Each subDict In storedDataDict1.items
            For Each k In subDict.keys ' k => keys within the dict, in this case keys will be the names of columns in database
                '''''deug.print k
                If k = Empty Then
                    Exit For
                End If
                storedDataDict1(ct)(k) = objRecordset.Fields(k)
                If k = uid_field_name Then
                    memberids = memberids & "'" & storedDataDict1(ct)(k) & "'" & ","
                End If
            Next
            'storedDataDict(ct) = subDict
            If ct >= noOfData Then
                '''''''''deug.print ct
                Exit Do
            End If
            ct = ct + 1
            objRecordset.MoveNext
        Next
    Loop
    objRecordset.Close
    memberids = Left(memberids, Len(memberids) - 1) & ")"
    If memberids = ")" Then
        LogStatement "Need to enroll few members for the LOB " & lineOfBusiness & " as all the previously enrolled members data have been used up"
        
        On Error Resume Next
        objRecordset.Close
    
        objConnection.Close
        Set objRecordset = Nothing
        Set objConnection = Nothing
        MsgBox "Script failed in function GetStoredRandomDataValues. Please review log file for more information"
'        oShell.Exec "taskkill /f /im excel.exe"
'        wscript.Quit
        On Error GoTo 0
        
    End If
    
    Debug.Print memberids
    ' fetching records that have not already been used for generating edi records from "Another" sheet
    objRecordset.Open "select Top " & noOfData & " * from [Another$] where [LOB] = '" & lineOfBusiness & "' and LEN([UniqueID]) = " & memberIdLength & " and  [" & uid_field_name & "] in " & memberids & ";", objConnection
    '''On Error Resume Next
    fields_count = objRecordset.Fields.Count
    ' looping numberOfRecordsNeeded times to create dictionaries with column names as keys
    For rec_ct = 1 To noOfData
        ' creating dictionay to hold values from each row of the recordset with column names as keys
        Set randomSubDict = CreateObject("Scripting.Dictionary")
        ' looping to add individual column names as keys into dictionary
        For Each field_name In objRecordset.Fields
            ''''''deug.print field_name.Name
            randomSubDict.Add field_name.Name, ""
        Next
        '''''''deug.print rec_ct
        storedDataDict2.Add rec_ct, randomSubDict
    Next
    ct = 1
    Do While objRecordset.EOF = False
        For Each subDict In storedDataDict2.items
            For Each k In subDict.keys ' k => keys within the dict, in this case keys will be the names of columns in database
                ''''''''deug.print k
                If k = Empty Then
                    Exit For
                End If
                storedDataDict2(ct)(k) = objRecordset.Fields(k)
                ''''deug.print k & " " & storedDataDict2(ct)(k)
            Next
            'storedDataDict(ct) = subDict
            If ct >= noOfData Then
                '''''''''deug.print ct
                Exit Do
            End If
            ct = ct + 1
            objRecordset.MoveNext
        Next
    Loop
    objRecordset.Close
    ' starting to consolidate the data from storedDataDict1 and storedDataDict2 into one dictionary object storedDataDict
    For rec_ct = 1 To noOfData
        Set randomSubDict = CreateObject("Scripting.Dictionary")
        For Each k In storedDataDict1.keys
            For Each k_1 In storedDataDict1(k).keys
                If Not randomSubDict.exists(k_1) Then
                    randomSubDict.Add k_1, storedDataDict1(k)(k_1)
                End If
                If k_1 = "UniqueID" Then
                    ''''deug.print rec_ct & " " & k_1 & " " & storedDataDict1(k)(k_1)
                End If
            Next
        Next
        For Each k In storedDataDict2.keys
            For Each k_1 In storedDataDict2(k).keys
                If Not randomSubDict.exists(k_1) Then
                    randomSubDict.Add k_1, storedDataDict2(k)(k_1)
                End If
                If k_1 = "UniqueID" Then
                    ''''deug.print rec_ct & " " & k_1 & " " & storedDataDict2(k)(k_1)
                End If
            Next
        Next
        storedDataDict.Add rec_ct, randomSubDict
    Next
    'msgbox "Need to edit the below query to update the excel database as Y"
    'objRecordset.Open "update [Master$] set [Used] = 'Y' where [" & uid_field_name & "] in " & ssns & ";", objConnection
    On Error Resume Next
    objRecordset.Close
    objConnection.Close
    Set objRecordset = Nothing
    Set objConnection = Nothing
    On Error GoTo 0
    
    If dependantDataRequired = True Then
        memberids = Replace(memberids, "(", "")
        memberids = Replace(memberids, ")", "")
        memberids_list = Split(memberids, ",")
        Set dependantRandomDataDict = CreateObject("Scripting.Dictionary")
        memid_ct = 1
        For Each memberid In memberids_list
            dependantRandomDataDict.Add memid_ct, CreateObject("Scripting.Dictionary")
            'getting data for dependants
           
            GetStoredRandomDataValuesForDependants 2, randomDataFilePath, memberid, memid_ct ' passing 2 as noOfData as we are getting 2 dependants data
            
            
        Next
        
    End If
    
    LogStatement "Completed execution of function GetStoredRandomDataValues"
End Function

Sub testingdbconn()

noOfData = 2
randomDataFilePath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Random_Updated_DB.xlsx"
memberid = "162716686901"
lineOfBusiness = "CCA"
currentRandKey = 1
Set dependantRandomDataDict = CreateObject("Scripting.Dictionary")
dependantRandomDataDict.Add 1, CreateObject("Scripting.Dictionary")
GetStoredRandomDataValuesForDependants noOfData, randomDataFilePath, memberid
End Sub

Function GetStoredRandomDataValuesForDependants(noOfData, randomDataFilePath, memberid, memid_ct)
'Sub GetStoredRandomDataValues()
'   lineOfBusiness = "MH"
'   memberIdLength = 12
'    noOfData = 2
'    randomDataFilePath = "C:\Users\karthik.thangaraj\Documents\My files\BMC Project\BMC Project\Supporting Files\Random_Updated_DB.xlsx"
    LogStatement "Started execution of function GetStoredRandomDataValuesForDependants"
    Dim objConnection, objRecordset
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordset = CreateObject("ADODB.Recordset")
    'Set newDependantRandomDataDict = CreateObject("Scripting.Dictionary")
    Set storedDataDict1 = CreateObject("Scripting.Dictionary")
    Set storedDataDict2 = CreateObject("Scripting.Dictionary")
    
       
    
    dbConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & randomDataFilePath & ";Extended Properties = excel 12.0;" ' excel 8.0 for excel 2003
    If objConnection.State = 1 Then
        objConnection.Close
    End If
    With objConnection
        .Open dbConnectStr
    End With
    Dim memberids
    'this variable is the primary key for the database
    uid_field_name = "UniqueID"
    
    memberids = "("
    memberid = Replace(memberid, "'", "")
    objRecordset.Open "select Top " & noOfData & " * from [Main$] where  [UniqueID] <> '" & memberid & "' and [2000-REF02] LIKE '%" & memberid & "%' and [LOB] = '" & lineOfBusiness & "';", objConnection
    
    
    
    'objRecordset.Open "select Top " & noOfData & " * from [Main$] where [Used] = 'N'", objConnection
    '''On Error Resume Next
    fields_count = objRecordset.Fields.Count
    ' looping numberOfRecordsNeeded times to create dictionaries with column names as keys
    For rec_ct = 1 To noOfData
    
        ' creating dictionay to hold values from each row of the recordset with column names as keys
        Set randomSubDict = CreateObject("Scripting.Dictionary")
        
        
        ' looping to add individual column names as keys into dictionary
        For Each field_name In objRecordset.Fields
            ''''''deug.print field_name.Name
            randomSubDict.Add field_name.Name, ""
            
        Next
        '''''''deug.print rec_ct
        storedDataDict1.Add rec_ct, randomSubDict
        
    Next
    ct = 1
    Do While objRecordset.EOF = False
        For Each subDict In storedDataDict1.items
            For Each k In subDict.keys ' k => keys within the dict, in this case keys will be the names of columns in database
                '''''deug.print k
                If k = Empty Then
                    Exit For
                End If
                storedDataDict1(ct)(k) = objRecordset.Fields(k)
                If k = uid_field_name Then
                    memberids = memberids & "'" & storedDataDict1(ct)(k) & "'" & ","
                End If
            Next
            'storedDataDict(ct) = subDict
            If ct >= noOfData Then
                '''''''''deug.print ct
                Exit Do
            End If
            ct = ct + 1
            objRecordset.MoveNext
        Next
    Loop
    objRecordset.Close
    memberids = Left(memberids, Len(memberids) - 1) & ")"
    If memberids = ")" Then
        LogStatement "Need to enroll few members for the LOB " & lineOfBusiness & " as all the previously enrolled members data have been used up"
        
        On Error Resume Next
        objRecordset.Close
    
        objConnection.Close
        Set objRecordset = Nothing
        Set objConnection = Nothing
        MsgBox "Script failed in function GetStoredRandomDataValuesForDependants. Please review log file for more information"
        'oShell.Exec "taskkill /f /im excel.exe"
        'wscript.Quit
        On Error GoTo 0
        
    End If
    
    ''''deug.print memberids
    ' fetching records that have not already been used for generating edi records from "Another" sheet
    objRecordset.Open "select Top " & noOfData & " * from [Another$] where  [UniqueID] in " & memberids & " and [LOB] = '" & lineOfBusiness & "';", objConnection
    
    '''On Error Resume Next
    fields_count = objRecordset.Fields.Count
    ' looping numberOfRecordsNeeded times to create dictionaries with column names as keys
    For rec_ct = 1 To noOfData
        ' creating dictionay to hold values from each row of the recordset with column names as keys
        Set randomSubDict = CreateObject("Scripting.Dictionary")
        ' looping to add individual column names as keys into dictionary
        For Each field_name In objRecordset.Fields
            ''''''deug.print field_name.Name
            randomSubDict.Add field_name.Name, ""
        Next
        '''''''deug.print rec_ct
        storedDataDict2.Add rec_ct, randomSubDict
    Next
    ct = 1
    Do While objRecordset.EOF = False
        For Each subDict In storedDataDict2.items
            For Each k In subDict.keys ' k => keys within the dict, in this case keys will be the names of columns in database
                ''''''''deug.print k
                If k = Empty Then
                    Exit For
                End If
                storedDataDict2(ct)(k) = objRecordset.Fields(k)
                ''''deug.print k & " " & storedDataDict2(ct)(k)
            Next
            'storedDataDict(ct) = subDict
            If ct >= noOfData Then
                '''''''''deug.print ct
                Exit Do
            End If
            ct = ct + 1
            objRecordset.MoveNext
        Next
    Loop
    objRecordset.Close
    ' starting to consolidate the data from storedDataDict1 and storedDataDict2 into one dictionary object storedDataDict
    For rec_ct = 1 To noOfData
        
        For Each k In storedDataDict1.keys
            ' storedDataDict1 will have 2 dictionary objects (ie) for 2 dependants
            'below if condition is to make sure that we loop through contents of only one dependant at a time and save it to dependantRandomDataDict
            If k = rec_ct Then
                Debug.Print "K is " & k
                Set randomSubDict = CreateObject("Scripting.Dictionary")
                For Each k_1 In storedDataDict1(k).keys
                    If Not randomSubDict.exists(k_1) Then
                        randomSubDict.Add k_1, storedDataDict1(k)(k_1)
                    End If
                    If k_1 = "UniqueID" Then
                        Debug.Print rec_ct & " " & k_1 & " " & storedDataDict1(k)(k_1)
                    End If
                Next
            End If
        Next
        For Each k In storedDataDict2.keys
            ' storedDataDict2 will have 2 dictionary objects (ie) for 2 dependants
            'below if condition is to make sure that we loop through contents of only one dependant at a time and save it to dependantRandomDataDict
            
            If k = rec_ct Then
                For Each k_1 In storedDataDict2(k).keys
                    If Not randomSubDict.exists(k_1) Then
                        randomSubDict.Add k_1, storedDataDict2(k)(k_1)
                    End If
                    If k_1 = "UniqueID" Then
                        Debug.Print rec_ct & " " & k_1 & " " & storedDataDict2(k)(k_1)
                    End If
                Next
            End If
        Next
        dependantRandomDataDict(memid_ct).Add rec_ct, randomSubDict
    Next
    
    'msgbox "Need to edit the below query to update the excel database as Y"
    'objRecordset.Open "update [Master$] set [Used] = 'Y' where [" & uid_field_name & "] in " & ssns & ";", objConnection
    On Error Resume Next
    objRecordset.Close
    objConnection.Close
    Set objRecordset = Nothing
    Set objConnection = Nothing
    On Error GoTo 0
    LogStatement "Completed execution of function GetStoredRandomDataValues"
End Function


Function getRandomName(f, m)
    fname = ""
    If f = True And m = True Then ' this is for members first name
        namesList = "ROY,MARVIN,JAMIE,ANDREW,TOMMY,ZACHARY,MARIO,TONYA,EDGAR,MALLORY,GARRETT,JESSE,BRYAN,RUBEN,DWAYNE,ESTHER,JULIA,NANCY,HENRY,GLENN,PHILLIP,JOHNSON,DESIREE,BRADY,GABRIELLE,RANDI,DALE,JAY,NINA,COREY,PRESTON,KARLA,TYLER,CHRISTY,MARTIN,MARISA,RICHARD,COLLEEN,HECTOR,RONNIE,CANDACE,TARA,ALAN,LILLIAN,CARRIE,TIMOTHY,JERRY,RICARDO,BRENDAN,CHRISTIAN,DEBORAH,MIKE,JACK,ERIN,ANA,JOSEPH,LATASHA,MARIA,LISA,TASHA,THOMAS,LACEY,NATALIE,KEITH,GLORIA,EDWIN,BLAKE,JUSTINE,SYLVIA,LATOYA,BRANDI,EMMA,RUTH,ADRIANA,KELLIE,BRITNEY,ALEJANDRO,TRISTAN,BRAD,CURTIS,MAILI,AIMEE,KAYLA,CANDICE,JACQUELINE,LAWRENCE,JOHNNY,CALEB,DARRELL,GERALD,MONIQUE,WILLIE"
    ElseIf f = True And m = False Then ' this is for sponsor first name
        namesList = "KIMBERLY,SOPHIA,CALEB,MICHAEL,SUMMER,RANDI,CASSANDRA,MELODY,MELANIE,LATASHA,ROGER,TAMARA,CHRISTINE,JEFFERY,DESIREE,SAMUEL,CURTIS,ALEXANDER,ALEXANDRA,KAREN,KENDRA,ISAAC,VICTORIA,GEOFFREY,MAILI,NICOLAS,KRISTEN,CARLOS,AMANDA,BRIANNA,LOGAN,DIANE,GARRETT,TYSON,BRUCE,CESAR,BRENT,THOMAS,ALICIA,ALBERT,ALANA,JOHN,TERESA,BROOKE,KAITLIN,TRACY,RAYMOND,THEODORE,BRIDGET,BARRY,LEE,TIFFANY,AIMEE,HEIDI,GARY,COLE,EDGAR,OSCAR,STACEY,SHANNA,CAROL,CARLA,LILLIAN,JACQUELYN,BRANDI,ERIK,CALVIN,JUAN,CANDICE,SHAWN,ANNE,DIANA,ERNEST,BARBARA,ROBERT,JENNY,JESUS,GRACE,RACHAEL,KATIE,WHITNEY,DAWN,DONNA,CODY,KRISTA,SEAN,VINCENT,NATHAN,SAVANNAH"
    ElseIf f = False And m = True Then ' this is for members last name
        namesList = "BLAINE,DANELLE,SHARITA,NICOLETTE,SANTOS,ALIA,BRAIN,MEGGAN,STEVIE,LOIS,DAPHNE,CANDY,LESTER,MADISON,ROSEMARIE,KAYLEIGH,MARI,TIERA,TITUS,CAMILLA,JED,RAYMUNDO,EUNICE,SHERITA,JOELLE,LORETTA,SHANNAN,SEBASTIAN,MARYANN,SELENA,TUCKER,DANICA,MARSHA,JANELL,TED,MALINDA,KIRBY,CARI,GIOVANNI,CONNOR,AMIR,IESHA,AVA,DAMION,SONJA,MARCELLA,LOUISE,MARIAN,RASHAWN,JULIETTE,DULCE,CARTER,CYRUS,BRITTANIE,BEATRIZ,LATRICIA,CASSI,SCOTTY,PHOEBE,SHARA,ROYCE,KRYSTIN,DONNY,GERARD,CHEYENNE,REYNA,RENA,DARCI,KAMINSKI,CHERISH,ROBBIE,LEILA,RILEY,STEPHANY,KESHA,MATT,CARSON,MALORIE,CORINA,DENISHA,TEENA,YAJAIRA,SOFIA,MISTI,DIANDRA,JUSTEN,JANETTE,MARGO,MARKITA,JOSH,JEANINE,ECHO,SHANAE,SHIRA,LEONEL,SOLOMON,PETE,MONA,BO"

    ElseIf f = False And m = False Then ' this is for sponsor last name
       namesList = "JANESSA,JEANNE,CRUZ,MAURA,FAWN,MARION,ZANE,SHANTELL,SHALONDA,ANTON,ARON,EVERETT,TENISHA,SONDRA,JEROMY,FRANCINE,DAMION,TAISHA,DEIDRE,ARNOLD,SHARINA,ARCHIE,YASMIN,MYLES,PAULINE,KENISHA,EDMOND,RUFUS,JAZMIN,LYNSEY,HERMAN,SHAYNE,MARTA,STEVIE,NICHELLE,VICKIE,MARYANN,NATHALIE,KODY,ANN,AMIR,AMIT,TEENA,ALI,DARRON,DEIDRA,LATORIA,RUSSELL,VALENCIA,ALONZO,MEGHANN,TOBY,SHAMIKA,LIZETTE,ADOLFO,RODRIGO,MALORIE,MCKENZIE,RONDA,COTY,SAMATHA,MARANDA,GREGG,EMILEE,ADAN,COLETTE,CORTEZ,MARTY,JEANIE,RENA,LAKIA,BRITANY,AMOS,BRYON,JEANNA,MOHAMMAD,ELMER,CHERI,DOMINGO,EUGENIA,ASHLYN,WANDA,JESS,KATHERYN,CARYN,DEANA,JUSTEN,DONNY,LISSETTE,SHARITA,JOSIE,FREDDY,KOREY,MARIAM,YADIRA"
    End If
    
    namesListArr = Split(namesList, ",")
    namesListArrLen = UBound(namesListArr) - 1
    fnames = ""
    Randomize
    
    
    fnames = ""
    For x = 1 To numberOfRecordsNeeded * 1 * how_many_random_data
        r_1 = CInt(Rnd() * 100) Mod namesListArrLen
        ''On Error Resume Next
        fname = namesListArr(r_1)
        If Err.Number <> 0 Then
            If f = True Then
                fname = "ABCDEFG"
            Else
                fname = "PQRSTU"
            End If
        End If
        fnames = fnames & fname & ","
    Next
    If Right(fnames, 1) = "," Then
        fnames = Left(fnames, Len(fnames) - 1)
    End If
    
    'on error goto 0
    getRandomName = fnames

End Function


Function getRandomDOB()


    dt = Date
    Randomize
    r_1 = CInt(Rnd() * 100) Mod 65
    r_2 = CInt(Rnd() * 100) Mod 12
    r_3 = CInt(Rnd() * 100) Mod 30
    '''''debug.print r_1
    dt = DateAdd("yyyy", -r_1, dt)
    dt = DateAdd("m", -r_2, dt)
    dt = DateAdd("d", -r_3, dt)
    
    dt_diff = DateDiff("yyyy", dt, Date)
    If dt_diff <= 21 Then
        dt = DateAdd("yyyy", -23, dt)
        
    End If
    
    ' checking if the new dob generated results in age being less than 21 and providing a new dob that is atleast 22 years more then generated date
    ' and hence greater than 21 years
    
    dobs = ""
    
    For x = 1 To numberOfRecordsNeeded * 2 * how_many_random_data
        new_dt = DateAdd("d", -x, dt)
        new_dt = DateAdd("m", x, new_dt)
        
        If x Mod 3 = 0 Then
            new_dt = DateAdd("yyyy", 3, new_dt)
        ElseIf x Mod 3 = 2 Then
            new_dt = DateAdd("yyyy", 1, new_dt)
        ElseIf x Mod 3 = 1 Then
            new_dt = DateAdd("yyyy", 2, new_dt)
        End If
        
        If x Mod 3 = 1 Then
            dt_diff = DateDiff("yyyy", new_dt, Date)
            If dt_diff <= 21 Then
                new_dt = DateAdd("yyyy", -35, new_dt)
                dt = new_dt
            Else
                dt = new_dt
            End If
        
        End If
        
        new_dt = GetFormattedDate(new_dt, "YYYYMMDD")
        
        'Debug.Print x & " : " & new_dt & " : " & dt & " : " & dt_diff
        
        dobs = dobs & CStr(new_dt) & ","
        
    Next
    If Right(dobs, 1) = "," Then
        dobs = Left(dobs, Len(dobs) - 1)
    End If
    
    
    getRandomDOB = dobs
End Function

Function getRandomNumberOfLength(req_len)


    dt = Date
    Randomize
    r_1 = CInt(Rnd() * 100) Mod 65
    r_2 = CInt(Rnd() * 100) Mod 12
    r_3 = CInt(Rnd() * 100) Mod 30
    '''''debug.print r_1
    dt = DateAdd("yyyy", -r_1, dt)
    dt = DateAdd("m", -r_2, dt)
    dt = DateAdd("d", -r_3, dt)
    
    dt_diff = DateDiff("s", dt, Date)
    
    If Len(Trim(Str(dt_diff))) > CInt(req_len) Then
        Do While Len(Trim(Str(dt_diff))) > CInt(req_len)
            dt_diff = CLng(dt_diff / 10)
        
        Loop
        
    End If
        
    If Len(Trim(Str(dt_diff))) < CInt(req_len) Then
        Do While Len(Trim(Str(dt_diff))) < CInt(req_len)
            dt_diff = CDbl(dt_diff * 10)
        
        Loop
        
    End If
        
    getRandomNumberOfLength = Trim(Str(dt_diff))
End Function





Function getRandomSsn()
    
    posix_dt = #1/1/1970#
    dt = Now
    
    dt_diff = CStr(DateDiff("s", posix_dt, dt))
    
    If Len(CStr(dt_diff)) = 10 Then
        dt_diff = Mid(dt_diff, 2)
    ElseIf Len(CStr(dt_diff)) = 11 Then
        dt_diff = Mid(dt_diff, 3)
    ElseIf Len(CStr(dt_diff)) = 12 Then
        dt_diff = Mid(dt_diff, 4)
            
    End If
    
    ssns = ""
    For x = 1 To numberOfRecordsNeeded * 2 * how_many_random_data
        ssns = ssns & CStr(dt_diff + x) & ","
        Application.Wait Now() + TimeValue("00:00:01")
    Next
    If Right(ssns, 1) = "," Then
        ssns = Left(ssns, Len(ssns) - 1)
    End If
    getRandomSsn = ssns

End Function

Function getRandomMemberIds_old()
    
    posix_dt = #1/1/1970#
    dt = Now
    
    dt_diff = CLng(DateDiff("s", posix_dt, dt))
    
    dt_diff = CLngLng(dt_diff * (10 ^ Abs(memberIdLength - Len(CStr(dt_diff)))))
    
    memberids = ""
    For x = 1 To numberOfRecordsNeeded * 1
        memberids = memberids & CStr(dt_diff + x) & ","
    Next
    If Right(memberids, 1) = "," Then
        memberids = Left(memberids, Len(memberids) - 1)
    End If
    getRandomMemberIds = memberids

End Function
Function getRandomMemberIds()
    
    posix_dt = #1/1/1970#
    dt = Now
    
    dt_diff = CLng(DateDiff("s", posix_dt, dt))
    
    addSuffix = 10 ^ Abs(memberIdLength - Len(CStr(dt_diff)))
    
    dt_diff = CStr(dt_diff) & Mid(addSuffix, 2)
    
    memberids = ""
    For x = 1 To numberOfRecordsNeeded * 1 * how_many_random_data
        If x < 10 Then
            newid = Left(dt_diff, Len(dt_diff) - 1) & x
        Else
            newid = Left(dt_diff, Len(dt_diff) - 2) & x
        End If
        
        memberids = memberids & newid & ","
        Application.Wait Now() + TimeValue("00:00:01")
    Next
    If Right(memberids, 1) = "," Then
        memberids = Left(memberids, Len(memberids) - 1)
    End If
    getRandomMemberIds = memberids
    
End Function

Function getRandomEmail(val_Place)
    f_name = dataToBeSavedToDbArr(currentRandKey)("2100A-NM103")
    If InStr(1, f_name, "@@@@@") <> 0 Then
        f_name = Split(f_name, "@@@@@")(val_Place - 1)
    End If
    l_name = dataToBeSavedToDbArr(currentRandKey)("2100A-NM104")
    If InStr(1, l_name, "@@@@@") <> 0 Then
        l_name = Split(l_name, "@@@@@")(val_Place - 1)
    End If
    mem_id = dataToBeSavedToDbArr(currentRandKey)("2000-REF02")
    If InStr(1, mem_id, "@@@@@") <> 0 Then
        mem_id = Split(mem_id, "@@@@@")(val_Place - 1)
    End If
    Email_id = f_name & l_name & mem_id
    Email_id = Email_id & "@email.com"
    getRandomEmail = Email_id
End Function


Function getRandomMemberIdsForDCH()
    memberids = ""
    For x = 1 To numberOfRecordsNeeded * 1 * how_many_random_data
        memberids = memberids & getRandomNumberOfLength(memberIdLength) & ","
        Application.Wait Now() + TimeValue("00:00:01")
    Next
    If Right(memberids, 1) = "," Then
        memberids = Left(memberids, Len(memberids) - 1)
    End If
    getRandomMemberIdsForDCH = memberids
    
End Function


Function getNpi()
    numberOfRecordsNeeded = 1
    dt = Now
    
    npis = ""
    posix_dt = #1/1/1970#
    dt = Now
    
    

    For x = 1 To numberOfRecordsNeeded * 1
        dt_diff = CLng(DateDiff("s", posix_dt, dt))
        
        npis = npis & dt_diff & ","
        Application.Wait Now() + TimeValue("00:00:01")
    Next
    
    If Right(npis, 1) = "," Then
        npis = Left(npis, Len(npis) - 1)
    End If
    getNpi = npis
    
End Function


Function getMiddleInitial()
    midinit = "A"
    inits = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
    initsArr = Split(inits, ",")
    Randomize
    
    midinits = ""
    For x = 1 To numberOfRecordsNeeded * 1 * how_many_random_data
        r_1 = CInt(Rnd() * 100) Mod 26
        midinit = initsArr(r_1)
        midinits = midinits & midinit & ","
    Next
    If Right(midinits, 1) = "," Then
        midinits = Left(midinits, Len(midinits) - 1)
    End If
    
    getMiddleInitial = midinits
End Function


Function getTelephone()
    tele = ""
    dt = Now
    dd = CStr(Day(dt))
    mon = CStr(Month(dt))
    yr = Right(CStr(Year(dt)), 2)
    ''''''debug.print yr
    hh = CStr(Hour(dt))
    MI = CStr(Minute(dt))
    ss = CStr(Second(dt))
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    If Len(mon) = 1 Then
        mon = "0" & mon
    End If
    If Len(hh) = 1 Then
        hh = "0" & hh
    End If
    If Len(MI) = 1 Then
        MI = "0" & MI
    End If
    If Len(ss) = 1 Then
        ss = "0" & ss
    End If
    
    tele = mon & dd & hh & MI & ss
    
    teles = ""
    For x = 1 To numberOfRecordsNeeded * 1 * how_many_random_data
        teles = teles & CStr(CLng(tele) + x) & ","
    Next
    If Right(teles, 1) = "," Then
        teles = Left(teles, Len(teles) - 1)
    End If
    
    getTelephone = teles

End Function

Function getRandomAddresses1()
    addr_mid = "CAR,BUS,VAN,LORRY,AUTO,CYCLE,MINIVAN,SUV,HATCH,LUXURY,COOK,DOCTOR,TEACHER"
    suffix = "ST,AVE,RD,STREET,ROAD,AVENUE"
    
    Randomize
    
    dt = Time
    hh = CStr(Hour(dt))
    MI = CStr(Minute(dt))
    ss = CStr(Second(dt))
    If Len(hh) = 1 Then
        hh = "0" & hh
    End If
    If Len(MI) = 1 Then
        MI = "0" & MI
    End If
    If Len(ss) = 1 Then
        ss = "0" & ss
    End If
    addresses = ""
    
    For x = 1 To numberOfRecordsNeeded * 1 * how_many_random_data
        r_1 = CInt(Rnd() * 100) Mod 13
        r_2 = CInt(Rnd() * 100) Mod 6
        
        addresses = addresses & (CLng(hh & MI & ss) + x) & " " & Split(addr_mid, ",")(r_1) & " " & Split(suffix, ",")(r_2) & ","
    Next
    If Right(addresses, 1) = "," Then
        addresses = Left(addresses, Len(addresses) - 1)
    End If
    
    
    getRandomAddresses1 = addresses
End Function
Function getRandomAddresses2()
    addr_mid = "CAR,BUS,VAN,LORRY,AUTO,CYCLE,MINIVAN,SUV,HATCH,LUXURY,COOK,DOCTOR,TEACHER"
    suffix = "ST,AVE,RD,STREET,ROAD,AVENUE"
    
    Randomize

    addresses = ""
    
    For x = 1 To numberOfRecordsNeeded * 1 * how_many_random_data
        r_1 = CInt(Rnd() * 100) Mod 13
        r_2 = CInt(Rnd() * 100) Mod 6
        
        addresses = addresses & Split(addr_mid, ",")(r_1) & " " & Split(suffix, ",")(r_2) & ","
    Next
    If Right(addresses, 1) = "," Then
        addresses = Left(addresses, Len(addresses) - 1)
    End If
    
    
    getRandomAddresses2 = addresses
End Function

Function getCities()
    cityList = "WALES,  TISBURY,    STURBRIDGE, BUCKLAND,   MALDEN, TYRINGHAM,  SAVOY,  ROCHESTER,  CLARKSBURG, WORTHINGTON,    SALEM,  UPTON,  LONGMEADOW, SANDISFIELD,    PLYMPTON,   WESTMINSTER,    BARNSTABLE, BRAINTREE,  PAXTON, COLRAIN,    WESTFORD,   HAMPDEN,    DUDLEY, TAUNTON,    WRENTHAM,   TEMPLETON,  CAMBRIDGE,  WINDSOR,    WENDELL,    NORTON, BARRE,  LEYDEN, GOSHEN, ASHFIELD,   IPSWICH,    MILFORD,    CONCORD,    SOUTHBRIDGE,    MONTEREY,   AMESBURY,   CHESTER,    BECKET, SHARON, STOUGHTON,  MASHPEE,    WESTBOROUGH,    NEWBURYPORT,    ABINGTON,   SOUTHBOROUGH,   BRIMFIELD,  MARION, MONTGOMERY, STERLING,   CHELMSFORD, GRANBY, CHILMARK,   GROTON, HARDWICK,   MILTON, EGREMONT,   CHELSEA,    DORCHESTER, LYNNFIELD,  WEYMOUTH,   HEATH,  WESTHAMPTON,    HADLEY, WOBURN, NANTUCKET,  COHASSET,   NORTHFIELD, LANESBOROUGH,   LANCASTER,  PHILLIPSTON,    SOMERSET,   BROCKTON,   SHERBORN,   SHELBURNE,  MANSFIELD,  WATERTOWN"
    cityListArr = Split(cityList, ",")
    
    Randomize
    
    cities = ""
    On Error Resume Next
    
    For x = 1 To numberOfRecordsNeeded * 1 * how_many_random_data
        r_1 = CInt(Rnd() * 100) Mod UBound(cityListArr) - 1
        citi = Trim(cityListArr(r_1))
        If Err.Number <> 0 Then
            citi = "WALES"
        End If
        cities = cities & citi & ","
    Next
    If Right(cities, 1) = "," Then
        cities = Left(cities, Len(cities) - 1)
    End If
    On Error GoTo 0
    
    getCities = cities


End Function

Function getZipCodes()
    zipsList = "023480000,027730000,027950000,027920000,026160000,027330000,027360000,024410000,026800000,025120000,025440000,023590000,026510000,025140000,027500000,024880000,025010000,027370000,024290000,023710000,027620000,026210000,026050000,024330000,026040000,026920000,023830000,023330000,023440000,025430000,026450000,024060000,027970000,023200000,026830000,026930000,027140000,023280000,027250000,026180000,027860000,027190000,025830000,027400000,023880000,026480000,023900000,024530000,023260000,024600000,023210000,025110000,025610000,025680000,025170000,027200000,024500000,027790000,025670000,024050000,025240000,024930000,027720000,025490000,026910000,026360000,026000000,025090000,025480000,026640000,027430000,024080000,027580000,023500000,023820000,025050000,024810000,027680000,026540000,024560000,026700000,027660000,026530000,024450000,023780000"
    zipsListArr = Split(zipsList, ",")
    
    Randomize
    
    zips = ""
    On Error Resume Next
    
    For x = 1 To numberOfRecordsNeeded * 1 * how_many_random_data
        r_1 = CInt(Rnd() * 100) Mod UBound(zipsListArr) - 1
        zip = Trim(zipsListArr(r_1))
        If Err.Number <> 0 Then
            zip = "023480000"
        End If
        zips = zips & zip & ","
    Next
    If Right(zips, 1) = "," Then
        zips = Left(zips, Len(zips) - 1)
    End If
    On Error GoTo 0
    
    getZipCodes = zips


End Function

Function getGenders()
    gendersList = "F,M,U"
    
    Randomize

    genders = ""
    
    For x = 1 To numberOfRecordsNeeded * 1 * how_many_random_data
        
        r_2 = CInt(Rnd() * 100) Mod 3
        
        genders = genders & Split(gendersList, ",")(r_2) & ","
    Next
    If Right(genders, 1) = "," Then
        genders = Left(genders, Len(genders) - 1)
    End If
    
    getGenders = genders

End Function

Function cleanUp(kill)
    Dim objExcel
    ' variable for Excel workbook
    Dim objWbk
    
    Set objExcel = CreateObject("Excel.Application")
    
    For Each objWbk In objExcel.Workbooks
        MsgBox objWbk.Name
        objWbk.Close False
    Next
    If Not oShell Then
        Set oShell = CreateObject("WScript.Shell")
    End If
    oShell.Exec "taskkill /f /im excel.exe"
    If kill = True Then
        wscript.Quit
    End If
End Function

Sub test333()
    Call cleanUp(False)

End Sub

Function DateFromString(dtStr)
    ' this function takes a date string in YYYYMMDD format and convert that into a date so date calculations can be done
    newDt = ""
    'On Error Resume Next
    If Len(dtStr) = 8 Then
        yr = Left(dtStr, 4)
        mon = Mid(dtStr, 5, 2)
        dd = Right(dtStr, 2)
        newDtStr = mon & "/" & dd & "/" & yr
        newDt = CDate(newDtStr)
        ''''debug.print newDtStr & " " & newDt
        
    
    End If
    On Error GoTo 0
    DateFromString = newDt
End Function

Function GetPreviousValue(loopSegmentKey, previousValueIndicator, valuePlace)
    'example of previousValueIndicator - PREVIOUS_VALUE(-1)
    ' default values to variables
    
    
    incrementVal = 0
    prevSegmentValue = ""
    ' checking if "(" is present in previousValueIndicator
    If InStr(1, previousValueIndicator, "(") <> 0 Then
        
        incrementVal = Replace(Split(previousValueIndicator, "(")(1), ")", "")
        incrementVal = CInt(Trim(incrementVal))
    
    End If
    newvaluePlace = valuePlace - 1
    refloopSegmentKey = "REF:" & loopSegmentKey
    prevSegmentValue = GetReferenceSegmentValue(refloopSegmentKey, newvaluePlace)
    ' incrementVal is passed only for date values in segments, positive or negative values can be passed from template
    If incrementVal <> 0 Then
        prevSegmentValue = DateAdd("d", incrementVal, DateFromString(prevSegmentValue))
        prevSegmentValue = GetFormattedDate(prevSegmentValue, "YYYYMMDD")
    End If
    
    GetPreviousValue = prevSegmentValue
        
End Function



