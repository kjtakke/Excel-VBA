Public UniqueList As Variant, TableRange As Variant, Working As Variant, DurationArray As Variant, UniqueListAreaArea As Variant, UniqueFinalListArea As Variant
Public StartYear As String, StartFinicialYear As String
Public TotalEvents As Long, CurrentEventCount
Public placeholder As Double

Sub RefreshForm()

MasterScheduleRefresh.ProgressBar.Value = 0
MasterScheduleRefresh.ProgressBar1.Value = 0


Worksheets("data").Range("AN2").ListObject.QueryTable.Refresh BackgroundQuery:=False
Worksheets("data").Range("B2").ListObject.QueryTable.Refresh BackgroundQuery:=False
Worksheets("data").Range("BD2").ListObject.QueryTable.Refresh BackgroundQuery:=False
Worksheets("data").Range("BF2").ListObject.QueryTable.Refresh BackgroundQuery:=False
Worksheets("Data").Rows("31:31").RowHeight = 25
MasterScheduleRefresh.Show
Worksheets("DateSheet").Range("C2").Select
End Sub

Sub StartMasterSchedule()
Application.ScreenUpdating = False
CurrentEventCount = 0
Call DateSheet
MasterScheduleRefresh.ProgressBar1.Value = 40

Range("C2").Select
ActiveWindow.FreezePanes = True
Dim cell As Range
Dim CellArea As Range

For Each CellArea In Range("B4:B200")

Select Case CellArea

Case "Enviromental Testing Facility"
CellArea.Interior.Color = RGB(Worksheets("Data").Range("AG2").Value, Worksheets("Data").Range("AH2").Value, Worksheets("Data").Range("AI2").Value)
Case "Graytown"
CellArea.Interior.Color = RGB(Worksheets("Data").Range("AG3").Value, Worksheets("Data").Range("AH3").Value, Worksheets("Data").Range("AI3").Value)
Case "Headquarters"
CellArea.Interior.Color = RGB(Worksheets("Data").Range("AG4").Value, Worksheets("Data").Range("AH4").Value, Worksheets("Data").Range("AI4").Value)
Case "Munitions Test Center"
CellArea.Interior.Color = RGB(Worksheets("Data").Range("AG5").Value, Worksheets("Data").Range("AH5").Value, Worksheets("Data").Range("AI5").Value)
Case "Port Wakefield"
CellArea.Interior.Color = RGB(Worksheets("Data").Range("AG6").Value, Worksheets("Data").Range("AH6").Value, Worksheets("Data").Range("AI6").Value)

End Select


Next CellArea
MasterScheduleRefresh.ProgressBar1.Value = 60
Columns("B").ColumnWidth = 37
Columns("C:CC").ColumnWidth = 6






Worksheets("DateSheet").Range("C1").Interior.Color = RGB(Worksheets("Data").Range("AG7").Value, Worksheets("Data").Range("AH7").Value, Worksheets("Data").Range("AI7").Value)
Worksheets("DateSheet").Range("F1").Interior.Color = RGB(Worksheets("Data").Range("AG8").Value, Worksheets("Data").Range("AH8").Value, Worksheets("Data").Range("AI8").Value)
Worksheets("DateSheet").Range("H1").Interior.Color = RGB(Worksheets("Data").Range("AG9").Value, Worksheets("Data").Range("AH9").Value, Worksheets("Data").Range("AI9").Value)
Worksheets("DateSheet").Range("L1").Interior.Color = RGB(Worksheets("Data").Range("AG10").Value, Worksheets("Data").Range("AH10").Value, Worksheets("Data").Range("AI10").Value)
Worksheets("DateSheet").Range("Q1").Interior.Color = RGB(Worksheets("Data").Range("AG11").Value, Worksheets("Data").Range("AH11").Value, Worksheets("Data").Range("AI11").Value)
Worksheets("DateSheet").Range("T1").Interior.Color = RGB(Worksheets("Data").Range("AG12").Value, Worksheets("Data").Range("AH12").Value, Worksheets("Data").Range("AI12").Value)
Worksheets("DateSheet").Range("W1").Interior.Color = RGB(Worksheets("Data").Range("AG13").Value, Worksheets("Data").Range("AH13").Value, Worksheets("Data").Range("AI13").Value)
Worksheets("DateSheet").Range("Z1").Interior.Color = RGB(Worksheets("Data").Range("AG14").Value, Worksheets("Data").Range("AH14").Value, Worksheets("Data").Range("AI14").Value)
Worksheets("DateSheet").Range("AD1").Interior.Color = RGB(Worksheets("Data").Range("AG15").Value, Worksheets("Data").Range("AH15").Value, Worksheets("Data").Range("AI14").Value)

Worksheets("DateSheet").Range("C1").font.Color = Worksheets("Data").Range("AL7").Value
Worksheets("DateSheet").Range("F1").font.Color = Worksheets("Data").Range("AL8").Value
Worksheets("DateSheet").Range("H1").font.Color = Worksheets("Data").Range("AL9").Value
Worksheets("DateSheet").Range("L1").font.Color = Worksheets("Data").Range("AL10").Value
Worksheets("DateSheet").Range("Q1").font.Color = Worksheets("Data").Range("AL11").Value
Worksheets("DateSheet").Range("T1").font.Color = Worksheets("Data").Range("AL12").Value
Worksheets("DateSheet").Range("W1").font.Color = Worksheets("Data").Range("AL13").Value
Worksheets("DateSheet").Range("Z1").font.Color = Worksheets("Data").Range("AL14").Value
Worksheets("DateSheet").Range("AD1").font.Color = Worksheets("Data").Range("AL15").Value



Worksheets("DateSheet").Range("BL1").Interior.Color = RGB(Worksheets("Data").Range("AG7").Value, Worksheets("Data").Range("AH7").Value, Worksheets("Data").Range("AI7").Value)
Worksheets("DateSheet").Range("BO1").Interior.Color = RGB(Worksheets("Data").Range("AG8").Value, Worksheets("Data").Range("AH8").Value, Worksheets("Data").Range("AI8").Value)
Worksheets("DateSheet").Range("BQ1").Interior.Color = RGB(Worksheets("Data").Range("AG9").Value, Worksheets("Data").Range("AH9").Value, Worksheets("Data").Range("AI9").Value)
Worksheets("DateSheet").Range("BU1").Interior.Color = RGB(Worksheets("Data").Range("AG10").Value, Worksheets("Data").Range("AH10").Value, Worksheets("Data").Range("AI10").Value)
Worksheets("DateSheet").Range("BZ1").Interior.Color = RGB(Worksheets("Data").Range("AG11").Value, Worksheets("Data").Range("AH11").Value, Worksheets("Data").Range("AI11").Value)
Worksheets("DateSheet").Range("CC1").Interior.Color = RGB(Worksheets("Data").Range("AG12").Value, Worksheets("Data").Range("AH12").Value, Worksheets("Data").Range("AI12").Value)
Worksheets("DateSheet").Range("CF1").Interior.Color = RGB(Worksheets("Data").Range("AG13").Value, Worksheets("Data").Range("AH13").Value, Worksheets("Data").Range("AI13").Value)
Worksheets("DateSheet").Range("CI1").Interior.Color = RGB(Worksheets("Data").Range("AG14").Value, Worksheets("Data").Range("AH14").Value, Worksheets("Data").Range("AI14").Value)
Worksheets("DateSheet").Range("CM1").Interior.Color = RGB(Worksheets("Data").Range("AG15").Value, Worksheets("Data").Range("AH15").Value, Worksheets("Data").Range("AI14").Value)

Worksheets("DateSheet").Range("BL1").font.Color = Worksheets("Data").Range("AL7").Value
Worksheets("DateSheet").Range("BO1").font.Color = Worksheets("Data").Range("AL8").Value
Worksheets("DateSheet").Range("BQ1").font.Color = Worksheets("Data").Range("AL9").Value
Worksheets("DateSheet").Range("BU1").font.Color = Worksheets("Data").Range("AL10").Value
Worksheets("DateSheet").Range("BZ1").font.Color = Worksheets("Data").Range("AL11").Value
Worksheets("DateSheet").Range("CC1").font.Color = Worksheets("Data").Range("AL12").Value
Worksheets("DateSheet").Range("CF1").font.Color = Worksheets("Data").Range("AL13").Value
Worksheets("DateSheet").Range("CI1").font.Color = Worksheets("Data").Range("AL14").Value
Worksheets("DateSheet").Range("CM1").font.Color = Worksheets("Data").Range("AL15").Value
MasterScheduleRefresh.ProgressBar1.Value = 70
Dim TodaysDate As String
Call DateColor
MasterScheduleRefresh.ProgressBar1.Value = 80
Call BlankAreaFill
MasterScheduleRefresh.ProgressBar1.Value = 95
Call MonthMerge
Worksheets("DateSheet").Range("AH1").Value = "Correct as at " & Worksheets("Data").Range("AD5")
Application.ScreenUpdating = True
MasterScheduleRefresh.ProgressBar1.Value = 100
End Sub

Sub PlaceEventsOnSheet()

'Call UniqueListAreaCaptureArea
'Call UniqueListStatusCaptureStatus
'Call CaptureDuration

'Call DateSheet

'Declorations
Dim StartingCell As Range, CalendarEvent As Range, EventStartingCell As Range, FullEventMerge As Range
Dim CounterQuarter As Long, CounterStatus As Long, CounterRow As Long, Counter As Long, NumberOfStatusItems As Long, CounterAreas As Long
Dim QuarterStartDate As Date, QuarterEndDate As Date, QuarterEndDateWorking As Range
Dim DaysBeforeQuarterStartDate, DaysAfterQuarterEndDate As Integer, Duration As Integer
Dim TestCounter As Long
Dim EventStartInt As Integer, EventEndInt As Integer
Dim EventStartDate As Date, EventEndDate As Date
Dim EventStartCell As Range, EventEndCell As Range
Dim Status As String, EventStartAddress As String
Dim RowOffset As Integer, CounterRowOffset As Integer, RowOffsetNextAreaStart As Integer
Dim cell As Range, EventRangeFree As Boolean, AreaHasEvents As Boolean
Dim EventStartRange As Range, StartingCellConstant As Range
Dim NextEventStartInt As Integer
Dim RangeCollection As Collection
Dim EventCollect As Range
Dim SubmittedRed As Integer, SubmittedGreen As Integer, SubmittedBlue As Integer
Dim PendingRed As Integer, PendingGreen As Integer, PendingBlue As Integer
Dim SchedulingConflictRed As Integer, SchedulingConflictGreen As Integer, SchedulingConflictBlue As Integer
Dim ResubmitwithChangesRed As Integer, ResubmitwithChangesGreen As Integer, ResubmitwithChangesBlue As Integer
Dim ApprovedRed As Integer, ApprovedGreen As Integer, ApprovedBlue As Integer
Dim RejectedRed As Integer, RejectedGreen As Integer, RejectedBlue As Integer
Dim CompletedRed As Integer, CompletedGreen As Integer, CompletedBlue As Integer
Dim CancelledRed As Integer, CancelledGreen As Integer, CancelledBlue As Integer
Dim DelayedRed As Integer, DelayedGreen As Integer, DelayedBlue As Integer
Dim OtherRed As Integer, OtherGreen As Integer, OtherBlue As Integer
Dim StatusColorRed As Integer, StatusColorGreen As Integer, StatusColorBlue As Integer
Dim ExcludeSubmitted As Boolean
Dim ExcludePending As Boolean
Dim ExcludeSchedulingConflict As Boolean
Dim ExcludeResubmitwithChanges As Boolean
Dim ExcludeApproved As Boolean
Dim ExcludeRejected As Boolean
Dim ExcludeCompleted As Boolean
Dim ExcludeCancelled As Boolean
Dim ExcludeDelayed As Boolean
Dim ExcludeOther As Boolean
Dim ExcludeEnviromentalTestingFacility As Boolean
Dim ExcludeGraytown As Boolean
Dim ExcludeHeadquarters As Boolean
Dim ExcludeMunitionsTestCenter As Boolean
Dim ExcludePortWakefield As Boolean
Dim FontColorSubmitted As String
Dim FontColorPending As String
Dim FontColorSchedulingConflict As String
Dim FontColorResubmitwithChanges As String
Dim FontColorApproved As String
Dim FontColorRejected As String
Dim FontColorCompleted As String
Dim FontColorCancelled As String
Dim FontColorDelayed As String
Dim FontColorOther As String
Dim FontColorEnviromentalTestingFacility As String
Dim FontColorGraytown As String
Dim FontColorHeadquarters As String
Dim FontColorMunitionsTestCenter As String
Dim FontColorPortWakefield As String
Dim FontColor As Integer

ExcludeSubmitted = Worksheets("Data").Range("AJ7").Value
ExcludePending = Worksheets("Data").Range("AJ8").Value
ExcludeSchedulingConflict = Worksheets("Data").Range("AJ9").Value
ExcludeResubmitwithChanges = Worksheets("Data").Range("AJ10").Value
ExcludeApproved = Worksheets("Data").Range("AJ11").Value
ExcludeRejected = Worksheets("Data").Range("AJ12").Value
ExcludeCompleted = Worksheets("Data").Range("AJ13").Value
ExcludeCancelled = Worksheets("Data").Range("AJ14").Value
ExcludeDelayed = Worksheets("Data").Range("AJ15").Value
ExcludeOther = Worksheets("Data").Range("AJ16").Value
ExcludeEnviromentalTestingFacility = Worksheets("Data").Range("AJ2").Value
ExcludeGraytown = Worksheets("Data").Range("AJ3").Value
ExcludeHeadquarters = Worksheets("Data").Range("AJ4").Value
ExcludeMunitionsTestCenter = Worksheets("Data").Range("AJ5").Value
ExcludePortWakefield = Worksheets("Data").Range("AJ6").Value

FontColorSubmitted = Worksheets("Data").Range("AK7").Value
FontColorPending = Worksheets("Data").Range("AK8").Value
FontColorSchedulingConflict = Worksheets("Data").Range("AK9").Value
FontColorResubmitwithChanges = Worksheets("Data").Range("AK10").Value
FontColorApproved = Worksheets("Data").Range("AK11").Value
FontColorRejected = Worksheets("Data").Range("AK12").Value
FontColorCompleted = Worksheets("Data").Range("AK13").Value
FontColorCancelled = Worksheets("Data").Range("AK14").Value
FontColorDelayed = Worksheets("Data").Range("AK15").Value
FontColorOther = Worksheets("Data").Range("AK16").Value
FontColorEnviromentalTestingFacility = Worksheets("Data").Range("AK2").Value
FontColorGraytown = Worksheets("Data").Range("AK3").Value
FontColorHeadquarters = Worksheets("Data").Range("AK4").Value
FontColorMunitionsTestCenter = Worksheets("Data").Range("AK5").Value
FontColorPortWakefield = Worksheets("Data").Range("AK6").Value

SubmittedRed = Worksheets("Data").Range("AG7").Value
SubmittedGreen = Worksheets("Data").Range("AH7").Value
SubmittedBlue = Worksheets("Data").Range("AI7").Value


PendingRed = Worksheets("Data").Range("AG8").Value
PendingGreen = Worksheets("Data").Range("AH8").Value
PendingBlue = Worksheets("Data").Range("AI8").Value


SchedulingConflictRed = Worksheets("Data").Range("AG9").Value
SchedulingConflictGreen = Worksheets("Data").Range("AH9").Value
SchedulingConflictBlue = Worksheets("Data").Range("AI9").Value


ResubmitwithChangesRed = Worksheets("Data").Range("AG10").Value
ResubmitwithChangesGreen = Worksheets("Data").Range("AH10").Value
ResubmitwithChangesBlue = Worksheets("Data").Range("AI10").Value


ApprovedRed = Worksheets("Data").Range("AG11").Value
ApprovedGreen = Worksheets("Data").Range("AH11").Value
ApprovedBlue = Worksheets("Data").Range("AI11").Value


RejectedRed = Worksheets("Data").Range("AG12").Value
RejectedGreen = Worksheets("Data").Range("AH12").Value
RejectedBlue = Worksheets("Data").Range("AI12").Value


CompletedRed = Worksheets("Data").Range("AG13").Value
CompletedGreen = Worksheets("Data").Range("AH13").Value
CompletedBlue = Worksheets("Data").Range("AI13").Value


CancelledRed = Worksheets("Data").Range("AG14").Value
CancelledGreen = Worksheets("Data").Range("AH14").Value
CancelledBlue = Worksheets("Data").Range("AI14").Value


DelayedRed = Worksheets("Data").Range("AG15").Value
DelayedGreen = Worksheets("Data").Range("AH15").Value
DelayedBlue = Worksheets("Data").Range("AI15").Value

OtherRed = Worksheets("Data").Range("AG16").Value
OtherGreen = Worksheets("Data").Range("AH16").Value
OtherBlue = Worksheets("Data").Range("AI16").Value





Set RangeCollection = New Collection
NextEventStartInt = 0
For Each cell In Range("A2:A1000")
    If cell.Interior.ThemeColor = xlThemeColorAccent1 Then
        NextEventStartInt = NextEventStartInt + 1
    Else
        Exit For
    End If
Next cell

Set StartingCell = Range("A2").Offset(NextEventStartInt, 2) 'May need to offset by -1,x
Set StartingCellConstant = StartingCell.Offset(-1, 0)
QuarterStartDate = StartingCell.Offset(-2, 0).Value
Set QuarterEndDateWorking = StartingCell.Offset(-2, 0)
    'QuarterEndDateWorking.Select
QuarterEndDate = QuarterEndDateWorking.End(xlToRight).Value

'Road Map
    'Statue Working(x, 5)
    'Start Date Working(x, 3)
    'End date Working(x, 4)
    'Duration Working(x, 7)
    'Title Working(x, 8)
    ' CounterUniqueItems
    'UniqueFinalList(x,1)

'To be placed between each Range/Area
RowOffsetNextAreaStart = 1


'***Calculate days in quarter !!!Issues with it
'StartingCell.Select
 Dim CounterQuarterDaysStartingCellAddress As String, CountOfDaysInQuarter As Integer
 CounterQuarterDaysStartingCell = StartingCell.Address(ReferenceStyle:=xlA1)
 For Each cell In Range(StartingCellConstant, StartingCellConstant.End(xlToRight))
 
     If cell.Interior.ThemeColor = xlThemeColorAccent1 Then
            CountOfDaysInQuarter = CountOfDaysInQuarter + 1
        Else
            Exit For
        End If
 
 Next cell

  
'***Area Loop using Array UniqueFinalListArea to UniqueFinalListArea Lenght
For CounterAreas = 1 To UBound(UniqueFinalListArea)
RowOffsetNextAreaStart = 0
AreaHasEvents = False
'***
        For Counter = 1 To UBound(Working)


  

'***Identify Qualifying Items by Start and End Date***
    '*** Start Date is Before Quarter End Date and End Date after Quarter Start Date
                If Working(Counter, 3) < QuarterEndDate And _
                Working(Counter, 4) > QuarterStartDate _
                Then
                          'Debug.Print (Counter & " - " & Working(Counter, 6) & " - " & Working(Counter, 3) & " - " & Working(Counter, 4) & " - " & Working(Counter, 5))



If Working(Counter, 5) = "1. Submitted" And ExcludeSubmitted = True Then GoTo NotInclude:
If Working(Counter, 5) = "2. Pending" And ExcludePending = True Then GoTo NotInclude:
If Working(Counter, 5) = "3. Scheduling Conflict" And ExcludeSchedulingConflict = True Then GoTo NotInclude:
If Working(Counter, 5) = "4. Resubmit with Changes" And ExcludeResubmitwithChanges = True Then GoTo NotInclude:
If Working(Counter, 5) = "5. Approved" And ExcludeApproved = True Then GoTo NotInclude:
If Working(Counter, 5) = "6. Rejected" And ExcludeRejected = True Then GoTo NotInclude:
If Working(Counter, 5) = "7. Completed" And ExcludeCompleted = True Then GoTo NotInclude:
If Working(Counter, 5) = "8. Cancelled" And ExcludeCancelled = True Then GoTo NotInclude:
If Working(Counter, 5) = "9. Delayed" And ExcludeDelayed = True Then GoTo NotInclude:
If Working(Counter, 5) = "Other" And ExcludeOther = True Then GoTo NotInclude:
If Working(Counter, 1) = "Enviromental Testing Facility" And ExcludeEnviromentalTestingFacility = True Then GoTo NotInclude:
If Working(Counter, 1) = "Graytown" And ExcludeGraytown = True Then GoTo NotInclude:
If Working(Counter, 1) = "Headquarters" And ExcludeHeadquarters = True Then GoTo NotInclude:
If Working(Counter, 1) = "Munitions Test Center" And ExcludeMunitionsTestCenter = True Then GoTo NotInclude:
If Working(Counter, 1) = "Port Wakefield" And ExcludePortWakefield = True Then GoTo NotInclude:


Select Case True
Case Working(Counter, 5) = "1. Submitted"
        Select Case True
            Case FontColorSubmitted = "White"
            FontColor = -2
            Case FontColorSubmitted = "Black"
            FontColor = -1
        End Select
Case Working(Counter, 5) = "2. Pending"

        Select Case True
            Case FontColorPending = "White"
            FontColor = -2
            Case FontColorPending = "Black"
            FontColor = -1
        End Select

Case Working(Counter, 5) = "3. Scheduling Conflict"

        Select Case True
            Case FontColorSchedulingConflict = "White"
            FontColor = -2
            Case FontColorSchedulingConflict = "Black"
            FontColor = -1
        End Select

Case Working(Counter, 5) = "4. Resubmit with Changes"

        Select Case True
            Case FontColorResubmitwithChanges = "White"
            FontColor = -2
            Case FontColorResubmitwithChanges = "Black"
            FontColor = -1
        End Select

Case Working(Counter, 5) = "5. Approved"

        Select Case True
            Case FontColorApproved = "White"
            FontColor = -2
            Case FontColorApproved = "Black"
            FontColor = -1
        End Select

Case Working(Counter, 5) = "6. Rejected"

        Select Case True
            Case FontColorRejected = "White"
            FontColor = -2
            Case FontColorRejected = "Black"
            FontColor = -1
        End Select

Case Working(Counter, 5) = "7. Completed"

        Select Case True
            Case FontColorCompleted = "White"
            FontColor = -2
            Case FontColorCompleted = "Black"
            FontColor = -1
        End Select

Case Working(Counter, 5) = "8. Cancelled"

        Select Case True
            Case FontColorCancelled = "White"
            FontColor = -2
            Case FontColorCancelled = "Black"
            FontColor = -1
        End Select

Case Working(Counter, 5) = "9. Delayed"

        Select Case True
            Case FontColorDelayed = "White"
            FontColor = -2
            Case FontColorDelayed = "Black"
            FontColor = -1
        End Select
        
Case Working(Counter, 5) = "Other"

        Select Case True
            Case FontColorOther = "White"
            FontColor = -2
            Case FontColorOther = "Black"
            FontColor = -1
        End Select

Case Else
    FontColor = -1
End Select



If Working(Counter, 5) = "1. Submitted" Then
StatusColorRed = SubmittedRed
    StatusColorGreen = SubmittedGreen
    StatusColorBlue = SubmittedBlue

ElseIf Working(Counter, 5) = "2. Pending" Then
    StatusColorRed = PendingRed
    StatusColorGreen = PendingGreen
    StatusColorBlue = PendingBlue

ElseIf Working(Counter, 5) = "3. Scheduling Conflict" Then
    StatusColorRed = SchedulingConflictRed
    StatusColorGreen = SchedulingConflictGreen
    StatusColorBlue = SchedulingConflictBlue

ElseIf Working(Counter, 5) = "4. Resubmit with Changes" Then
    StatusColorRed = ResubmitwithChangesRed
    StatusColorGreen = ResubmitwithChangesGreen
    StatusColorBlue = ResubmitwithChangesBlue

ElseIf Working(Counter, 5) = "5. Approved" Then
    StatusColorRed = ApprovedRed
    StatusColorGreen = ApprovedGreen
    StatusColorBlue = ApprovedBlue

ElseIf Working(Counter, 5) = "6. Rejected" Then
    StatusColorRed = RejectedRed
    StatusColorGreen = RejectedGreen
    StatusColorBlue = RejectedBlue

ElseIf Working(Counter, 5) = "7. Completed" Then
    StatusColorRed = CompletedRed
    StatusColorGreen = CompletedGreen
    StatusColorBlue = CompletedBlue

ElseIf Working(Counter, 5) = "8. Cancelled" Then
    StatusColorRed = CancelledRed
    StatusColorGreen = CancelledGreen
    StatusColorBlue = CancelledBlue

ElseIf Working(Counter, 5) = "9. Delayed" Then
    StatusColorRed = DelayedRed
    StatusColorGreen = DelayedGreen
    StatusColorBlue = DelayedBlue

ElseIf Working(Counter, 5) = "Other" Then
    StatusColorRed = OtherRed
    StatusColorGreen = OtherGreen
    StatusColorBlue = OtherBlue

End If



'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
'***Split relevant items into four categories
    '***1. Start Date is Before Quarter Start Date and End Date after Quarter End Date
    '***2. Start Date is Before Quarter Start
    '***3. End Date after Quarter End Date
    '***4. Start and end date are between Quarter Start Date and Quarter End Date
                    If Working(Counter, 3) < QuarterStartDate And Working(Counter, 4) > QuarterEndDate Then
                         'Debug.Print ("Both " & Counter & " - " & Working(Counter, 6) & " - " & Working(Counter, 3) & " - " & Working(Counter, 4))
                         
                         
                         
                         
                         
                         
If Working(Counter, 1) = UniqueFinalListArea(CounterAreas, 1) Then
AreaHasEvents = True
Else
GoTo CounterEnd:
End If

                        EventStartInt = 0
                        RowOffset = 0
                        NextEventStartInt = 0
                        For Each cell In Range("A2:A1000")
                            If cell.Interior.ThemeColor = xlThemeColorAccent1 Then
                                NextEventStartInt = NextEventStartInt + 1
                            Else
                                Exit For
                            End If
                        Next cell
                        Set StartingCell = Range("A2").Offset(NextEventStartInt - 1, 2)
                        Set EventStartCell = StartingCell.Offset(RowOffset, EventStartInt)
                        EventStartAddress = EventStartCell.Address(ReferenceStyle:=xlA1)


    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Calculating Event Starting Addres
                        For CounterRowOffset = 1 To 50
                            EventRangeFree = True
                            EventStartAddress = EventStartCell.Offset(CounterRowOffset, 0).Address(ReferenceStyle:=xlA1)

                            'Range(EventStartAddress).Select
                   'Range(EventStartAddress).Select

         
         
             


         
         
         
         
         
         
         'Worksheets("DateSheet").Range(EventStartAddress, Range(EventStartAddress).Offset(RowOffset, CountOfDaysInQuarter - 1)).Select
         '***Identify if Event Cell Range is available
                            For Each cell In Worksheets("DateSheet").Range(EventStartAddress, Range(EventStartAddress).Offset(RowOffset, CountOfDaysInQuarter - 1))
                                    
                                    If cell.MergeCells = True Or cell.Interior.Pattern <> xlNone Then
                                        EventRangeFree = False
                                    End If
                                Next cell
                                
    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Identifying if Selected Range is available
                                If EventRangeFree = True Then
                                    If RowOffsetNextAreaStart < CounterRowOffset Then RowOffsetNextAreaStart = CounterRowOffset
                                          On Error Resume Next
                                          
    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Range/Event Formatting
    
    'Dim RangeCollection As Collection
    'Dim EventCollect As Range
    
    RangeCollection.Add Worksheets("DateSheet").Range(EventStartAddress, Range(EventStartAddress).Offset(RowOffset, CountOfDaysInQuarter - 1))
    
                                    With Worksheets("DateSheet").Range(EventStartAddress, Range(EventStartAddress).Offset(RowOffset, CountOfDaysInQuarter - 1))
                                        '.Select
                                        '.Merge
                                        '.Name = Working(Counter, 12)
                                        .Interior.Color = RGB(StatusColorRed, StatusColorGreen, StatusColorBlue)
'                                        .Interior.Pattern = xlSolid
'                                        .Interior.PatternColorIndex = xlAutomatic
'                                        .Interior.ThemeColor = xlThemeColorAccent1
'                                        .Interior.TintAndShade = 0.799981688894314
'                                        .Interior.PatternTintAndShade = 0
                                        '.Value = Working(Counter, 3) & "-" & Working(Counter, 4) & "                                  "
                                        .HorizontalAlignment = xlFill
                                        .VerticalAlignment = xlBottom
                                        .WrapText = False
                                        .Orientation = 0
                                        .AddIndent = False
                                        .IndentLevel = 0
                                        .ShrinkToFit = False
                                        '.ReadingOrder = xlContext
                                        .Borders(xlDiagonalDown).LineStyle = xlNone
                                        .Borders(xlDiagonalUp).LineStyle = xlNone
                                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                                        .Borders(xlEdgeLeft).ColorIndex = 0
                                        .Borders(xlEdgeLeft).TintAndShade = 0
                                        .Borders(xlEdgeLeft).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                                        .Borders(xlEdgeRight).ColorIndex = 0
                                        .Borders(xlEdgeRight).TintAndShade = 0
                                        .Borders(xlEdgeRight).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                                        .Borders(xlEdgeTop).ColorIndex = 0
                                        .Borders(xlEdgeTop).TintAndShade = 0
                                        .Borders(xlEdgeTop).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                                        .Borders(xlEdgeBottom).ColorIndex = 0
                                        .Borders(xlEdgeBottom).TintAndShade = 0
                                        .Borders(xlEdgeBottom).Weight = xlThin ' xlMedium
                                    End With
                                    Set EventStartCell = Range(EventStartAddress)
                                    With EventStartCell
                                        .Value = Working(Counter, 2)
                                        .font.Color = FontColor
                                        .font.Bold = True
                                        .font.Size = 12
                                        .AddComment
                                        .Comment.Visible = False
                                        .Comment.Shape.Height = 300
                                        .Comment.Shape.Width = 600
                                        .Comment.Shape.Fill.Visible = msoTrue
                                        .Comment.Shape.Fill.ForeColor.SchemeColor = 0 '0 = black 1 = white 2 = red 3 = green 4 = blue 5 = yellow
                                        .Comment.Shape.TextFrame.Characters.font.Color = -2
                                        .Comment.Shape.TextFrame.Characters.font.Bold = True
                                        .Comment.Shape.TextFrame.Characters.font.Size = 16
                                        '.Comment.Shape.TextFrame.AutoSize = True
                                        .Comment.Text "Task Name: " & Working(Counter, 2) & Chr(10) & Chr(10) & _
                                                      "Range: " & Working(Counter, 1) & Chr(10) & _
                                                      "Date Raised: " & Working(Counter, 8) & Chr(10) & _
                                                      "Date Range: " & Working(Counter, 3) & "-" & Working(Counter, 4) & Chr(10) & _
                                                      "Duration: " & Working(Counter, 20) & " days" & Chr(10) & _
                                                      "Status: " & Working(Counter, 5) & Chr(10) & _
                                                      "Category: " & Working(Counter, 9) & Chr(10) & _
                                                      "Task Type: " & Working(Counter, 11) & Chr(10) & _
                                                      "Old Task Number: " & Working(Counter, 10) & Chr(10) & _
                                                      "Task Number: " & Working(Counter, 15) & Chr(10) & _
                                                      "Request Type: " & Working(Counter, 16) & Chr(10) & _
                                                      "Objective Ref: " & Working(Counter, 12) & Chr(10) & _
                                                      "Customer: " & Working(Counter, 17)


                                    End With
                                    Exit For
                                End If
                            Next CounterRowOffset


                         
'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
                      
                         
                    ElseIf Working(Counter, 3) < QuarterStartDate Then
                        'Debug.Print ("Befo " & Counter & " - " & Working(Counter, 6) & " - " & Working(Counter, 3) & " - " & Working(Counter, 4))
If Working(Counter, 1) = UniqueFinalListArea(CounterAreas, 1) Then
AreaHasEvents = True
Else
GoTo CounterEnd:
End If

                        EventStartInt = 0
                        RowOffset = 0
                        NextEventStartInt = 0
                        For Each cell In Range("A2:A1000")
                            If cell.Interior.ThemeColor = xlThemeColorAccent1 Then
                                NextEventStartInt = NextEventStartInt + 1
                            Else
                                Exit For
                            End If
                        Next cell
                        
                        
                        Set StartingCell = Range("A2").Offset(NextEventStartInt - 1, 2)
                        'StartingCell.Select
                        Set EventStartCell = StartingCell.Offset(RowOffset, EventStartInt)
                        EventStartAddress = EventStartCell.Address(ReferenceStyle:=xlA1)


    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Calculating Event Starting Addres
                        For CounterRowOffset = 1 To 50
                        'StartingCell.Select
                            EventRangeFree = True
                            EventStartAddress = EventStartCell.Offset(CounterRowOffset, 0).Address(ReferenceStyle:=xlA1)
'Range(EventStartAddress).Select
                            'Range(EventStartAddress).Select
                   'Range(EventStartAddress).Select
         '***Identify if Event Cell Range is available
         
'Fix This
Dim DurationForEventStartBeforeQuarter As Integer
DurationForEventStartBeforeQuarter = Working(Counter, 20) - (QuarterStartDate - Working(Counter, 3))
         'Worksheets("DateSheet").Range(EventStartAddress, Worksheets("DateSheet").Range(EventStartAddress).Offset(RowOffset, DurationForEventStartBeforeQuarter - 1)).Select
         
                            For Each cell In Worksheets("DateSheet").Range(EventStartAddress, _
                                             Worksheets("DateSheet").Range(EventStartAddress).Offset(RowOffset, DurationForEventStartBeforeQuarter - 1))
                                             
                                    If cell.MergeCells = True Or cell.Interior.Pattern <> xlNone Then
                                        EventRangeFree = False
                                    End If
                                Next cell
                                
    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Identifying if Selected Range is available
                                If EventRangeFree = True Then
                                    If RowOffsetNextAreaStart < CounterRowOffset Then RowOffsetNextAreaStart = CounterRowOffset
                                          On Error Resume Next
                                          
    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Range/Event Formatting
    
    'Dim RangeCollection As Collection
    'Dim EventCollect As Range
    RangeCollection.Add Worksheets("DateSheet").Range(EventStartAddress, Worksheets("DateSheet").Range(EventStartAddress).Offset(0, DurationForEventStartBeforeQuarter - 1))
    
                                    With Worksheets("DateSheet").Range(EventStartAddress, Worksheets("DateSheet").Range(EventStartAddress).Offset(0, DurationForEventStartBeforeQuarter - 1))
                                        '.Select
                                        '.Merge
                                        '.Name = Working(Counter, 12)
                                        .Interior.Color = RGB(StatusColorRed, StatusColorGreen, StatusColorBlue)
'                                        .Interior.Pattern = xlSolid
'                                        .Interior.PatternColorIndex = xlAutomatic
'                                        .Interior.ThemeColor = xlThemeColorAccent1
'                                        .Interior.TintAndShade = 0.799981688894314
'                                        .Interior.PatternTintAndShade = 0
                                        '.Value = Working(Counter, 3) & "-" & Working(Counter, 4) & "  -  "
                                        .HorizontalAlignment = xlFill
                                        .VerticalAlignment = xlBottom
                                        .WrapText = False
                                        .Orientation = 0
                                        .AddIndent = False
                                        .IndentLevel = 0
                                        .ShrinkToFit = False
                                        '.ReadingOrder = xlContext
                                        .Borders(xlDiagonalDown).LineStyle = xlNone
                                        .Borders(xlDiagonalUp).LineStyle = xlNone
                                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                                        .Borders(xlEdgeLeft).ColorIndex = 0
                                        .Borders(xlEdgeLeft).TintAndShade = 0
                                        .Borders(xlEdgeLeft).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                                        .Borders(xlEdgeRight).ColorIndex = 0
                                        .Borders(xlEdgeRight).TintAndShade = 0
                                        .Borders(xlEdgeRight).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                                        .Borders(xlEdgeTop).ColorIndex = 0
                                        .Borders(xlEdgeTop).TintAndShade = 0
                                        .Borders(xlEdgeTop).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                                        .Borders(xlEdgeBottom).ColorIndex = 0
                                        .Borders(xlEdgeBottom).TintAndShade = 0
                                        .Borders(xlEdgeBottom).Weight = xlThin ' xlMedium
                                    End With
                                    Set EventStartCell = Range(EventStartAddress)
                                    With EventStartCell
                                        .Value = Working(Counter, 2)
                                        .font.Color = FontColor
                                        .font.Bold = True
                                        .font.Size = 12
                                        .AddComment
                                        .Comment.Visible = False
                                        .Comment.Shape.Height = 300
                                        .Comment.Shape.Width = 600
                                        .Comment.Shape.Fill.Visible = msoTrue
                                        .Comment.Shape.Fill.ForeColor.SchemeColor = 0 '0 = black 1 = white 2 = red 3 = green 4 = blue 5 = yellow
                                        .Comment.Shape.TextFrame.Characters.font.Color = -2
                                        .Comment.Shape.TextFrame.Characters.font.Bold = True
                                        .Comment.Shape.TextFrame.Characters.font.Size = 16
                                        '.Comment.Shape.TextFrame.AutoSize = True
                                        .Comment.Text "Task Name: " & Working(Counter, 2) & Chr(10) & Chr(10) & _
                                                      "Range: " & Working(Counter, 1) & Chr(10) & _
                                                      "Date Raised: " & Working(Counter, 8) & Chr(10) & _
                                                      "Date Range: " & Working(Counter, 3) & "-" & Working(Counter, 4) & Chr(10) & _
                                                      "Duration: " & Working(Counter, 20) & " days" & Chr(10) & _
                                                      "Status: " & Working(Counter, 5) & Chr(10) & _
                                                      "Category: " & Working(Counter, 9) & Chr(10) & _
                                                      "Task Type: " & Working(Counter, 11) & Chr(10) & _
                                                      "Old Task Number: " & Working(Counter, 10) & Chr(10) & _
                                                      "Task Number: " & Working(Counter, 15) & Chr(10) & _
                                                      "Request Type: " & Working(Counter, 16) & Chr(10) & _
                                                      "Objective Ref: " & Working(Counter, 12) & Chr(10) & _
                                                      "Customer: " & Working(Counter, 17)


                                    End With
                                    Exit For
                                End If
                            Next CounterRowOffset

                        
                        
                        
'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
                          
                    ElseIf Working(Counter, 4) > QuarterEndDate Then
                        'Debug.Print ("Afte " & Counter & " - " & Working(Counter, 6) & " - " & Working(Counter, 3) & " - " & Working(Counter, 4))
                        
                        
                        
                        
If Working(Counter, 1) = UniqueFinalListArea(CounterAreas, 1) Then
AreaHasEvents = True
Else
GoTo CounterEnd:
End If

                        EventStartInt = Working(Counter, 3) - QuarterStartDate
                        RowOffset = 0
                        NextEventStartInt = 0
                        For Each cell In Range("A2:A1000")
                            If cell.Interior.ThemeColor = xlThemeColorAccent1 Then
                                NextEventStartInt = NextEventStartInt + 1
                            Else
                                Exit For
                            End If
                        Next cell
                        
                        
                        Set StartingCell = Range("A2").Offset(NextEventStartInt - 1, 2)
                        'StartingCell.Select
                        Set EventStartCell = StartingCell.Offset(RowOffset, EventStartInt)
                        EventStartAddress = EventStartCell.Address(ReferenceStyle:=xlA1)


    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Calculating Event Starting Addres
                        For CounterRowOffset = 1 To 50
                        'StartingCell.Select
                            EventRangeFree = True
                            EventStartAddress = EventStartCell.Offset(CounterRowOffset, 0).Address(ReferenceStyle:=xlA1)
'Range(EventStartAddress).Select
                            'Range(EventStartAddress).Select
                   'Range(EventStartAddress).Select
         '***Identify if Event Cell Range is available
         'From the first If   ************For Each Cell In Worksheets("DateSheet").Range(EventStartAddress, Range(EventStartAddress).Offset(RowOffset, CountOfDaysInQuarter - 1))
                            Dim EventStartDateLessQuarterStartDate As Integer
                            EventStartDateLessQuarterStartDate = CountOfDaysInQuarter - (Working(Counter, 3) - QuarterStartDate + 1)
    
    
                            For Each cell In Worksheets("DateSheet").Range(EventStartAddress, _
                                             Worksheets("DateSheet").Range(EventStartAddress).Offset(RowOffset, EventStartDateLessQuarterStartDate))
                                    If cell.MergeCells = True Or cell.Interior.Pattern <> xlNone Then
                                        EventRangeFree = False
                                    End If
                                Next cell
                                
    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Identifying if Selected Range is available
                                If EventRangeFree = True Then
                                    If RowOffsetNextAreaStart < CounterRowOffset Then RowOffsetNextAreaStart = CounterRowOffset
                                          On Error Resume Next
                                          
    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Range/Event Formatting
    
    'Dim RangeCollection As Collection
    'Dim EventCollect As Range

'Fix up here
    
    RangeCollection.Add Worksheets("DateSheet").Range(EventStartAddress, Worksheets("DateSheet").Range(EventStartAddress).Offset(RowOffset, EventStartDateLessQuarterStartDate))
    
                                    With Worksheets("DateSheet").Range(EventStartAddress, Worksheets("DateSheet").Range(EventStartAddress).Offset(RowOffset, EventStartDateLessQuarterStartDate))
                                        '.Select
                                        '.Merge
                                        '.Name = Working(Counter, 12)
                                        .Interior.Color = RGB(StatusColorRed, StatusColorGreen, StatusColorBlue)
'                                        .Interior.Pattern = xlSolid
'                                        .Interior.PatternColorIndex = xlAutomatic
'                                        .Interior.ThemeColor = xlThemeColorAccent1
'                                        .Interior.TintAndShade = 0.799981688894314
'                                        .Interior.PatternTintAndShade = 0
                                        '.Value = Working(Counter, 3) & "-" & Working(Counter, 4) & "  -  "
                                        .HorizontalAlignment = xlFill
                                        .VerticalAlignment = xlBottom
                                        .WrapText = False
                                        .Orientation = 0
                                        .AddIndent = False
                                        .IndentLevel = 0
                                        .ShrinkToFit = False
                                        '.ReadingOrder = xlContext
                                        .Borders(xlDiagonalDown).LineStyle = xlNone
                                        .Borders(xlDiagonalUp).LineStyle = xlNone
                                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                                        .Borders(xlEdgeLeft).ColorIndex = 0
                                        .Borders(xlEdgeLeft).TintAndShade = 0
                                        .Borders(xlEdgeLeft).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                                        .Borders(xlEdgeRight).ColorIndex = 0
                                        .Borders(xlEdgeRight).TintAndShade = 0
                                        .Borders(xlEdgeRight).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                                        .Borders(xlEdgeTop).ColorIndex = 0
                                        .Borders(xlEdgeTop).TintAndShade = 0
                                        .Borders(xlEdgeTop).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                                        .Borders(xlEdgeBottom).ColorIndex = 0
                                        .Borders(xlEdgeBottom).TintAndShade = 0
                                        .Borders(xlEdgeBottom).Weight = xlThin ' xlMedium
                                    End With
                                    Set EventStartCell = Range(EventStartAddress)
                                    With EventStartCell
                                        .Value = Working(Counter, 2)
                                        .font.Color = FontColor
                                        .font.Bold = True
                                        .font.Size = 12
                                        .AddComment
                                        .Comment.Visible = False
                                        .Comment.Shape.Height = 300
                                        .Comment.Shape.Width = 600
                                        .Comment.Shape.Fill.Visible = msoTrue
                                        .Comment.Shape.Fill.ForeColor.SchemeColor = 0 '0 = black 1 = white 2 = red 3 = green 4 = blue 5 = yellow
                                        .Comment.Shape.TextFrame.Characters.font.Color = -2
                                        .Comment.Shape.TextFrame.Characters.font.Bold = True
                                        .Comment.Shape.TextFrame.Characters.font.Size = 16
                                        '.Comment.Shape.TextFrame.AutoSize = True
                                        .Comment.Text "Task Name: " & Working(Counter, 2) & Chr(10) & Chr(10) & _
                                                      "Range: " & Working(Counter, 1) & Chr(10) & _
                                                      "Date Raised: " & Working(Counter, 8) & Chr(10) & _
                                                      "Date Range: " & Working(Counter, 3) & "-" & Working(Counter, 4) & Chr(10) & _
                                                      "Duration: " & Working(Counter, 20) & " days" & Chr(10) & _
                                                      "Status: " & Working(Counter, 5) & Chr(10) & _
                                                      "Category: " & Working(Counter, 9) & Chr(10) & _
                                                      "Task Type: " & Working(Counter, 11) & Chr(10) & _
                                                      "Old Task Number: " & Working(Counter, 10) & Chr(10) & _
                                                      "Task Number: " & Working(Counter, 15) & Chr(10) & _
                                                      "Request Type: " & Working(Counter, 16) & Chr(10) & _
                                                      "Objective Ref: " & Working(Counter, 12) & Chr(10) & _
                                                      "Customer: " & Working(Counter, 17)


                                    End With
                                    Exit For
                                                                   
                                    
                                End If
                                        placeholder = (CurrentEventCount / TotalEvents) * 100
                                        CurrentEventCount = CurrentEventCount + 1
                                        MasterScheduleRefresh.ProgressBar.Value = (CurrentEventCount / TotalEvents) * 100
                                
                            Next CounterRowOffset


'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************



                    ElseIf Working(Counter, 3) >= QuarterStartDate _
                       And Working(Counter, 3) <= QuarterEndDate _
                       And Working(Counter, 4) >= QuarterStartDate _
                       And Working(Counter, 4) <= QuarterEndDate Then
                        'Debug.Print ("Neit " & Counter & " - " & Working(Counter, 6) & " - " & Working(Counter, 3) & " - " & Working(Counter, 4) & " - " & Working(Counter, 5))



If Working(Counter, 1) = UniqueFinalListArea(CounterAreas, 1) Then
AreaHasEvents = True
Else
GoTo CounterEnd:
End If

                        EventStartInt = Working(Counter, 3) - QuarterStartDate
                        RowOffset = 0
                        NextEventStartInt = 0
                        For Each cell In Range("A2:A1000")
                            If cell.Interior.ThemeColor = xlThemeColorAccent1 Then
                                NextEventStartInt = NextEventStartInt + 1
                            Else
                                Exit For
                            End If
                        Next cell
                        
                        
                        Set StartingCell = Range("A2").Offset(NextEventStartInt - 1, 2)
                        'StartingCell.Select
                        Set EventStartCell = StartingCell.Offset(RowOffset, EventStartInt)
                        EventStartAddress = EventStartCell.Address(ReferenceStyle:=xlA1)


    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Calculating Event Starting Addres
                        For CounterRowOffset = 1 To 50
                        'StartingCell.Select
                            EventRangeFree = True
                            EventStartAddress = EventStartCell.Offset(CounterRowOffset, 0).Address(ReferenceStyle:=xlA1)
'Range(EventStartAddress).Select
                            'Range(EventStartAddress).Select
                   'Range(EventStartAddress).Select
         '***Identify if Event Cell Range is available
         
         
                            For Each cell In Worksheets("DateSheet").Range(EventStartAddress, _
                                             Worksheets("DateSheet").Range(EventStartAddress).Offset(RowOffset, Working(Counter, 20) - 1))
                                    If cell.MergeCells = True Or cell.Interior.Pattern <> xlNone Then
                                        EventRangeFree = False
                                    End If
                                Next cell
                                
    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Identifying if Selected Range is available
                                If EventRangeFree = True Then
                                    If RowOffsetNextAreaStart < CounterRowOffset Then RowOffsetNextAreaStart = CounterRowOffset
                                          On Error Resume Next
                                          
    '***4. Start and end date are between Quarter Start Date and Quarter End Date
    '***Range/Event Formatting
    
    'Dim RangeCollection As Collection
    'Dim EventCollect As Range
    RangeCollection.Add Worksheets("DateSheet").Range(EventStartAddress, Worksheets("DateSheet").Range(EventStartAddress).Offset(0, Working(Counter, 20) - 1))
    
                                    With Worksheets("DateSheet").Range(EventStartAddress, Worksheets("DateSheet").Range(EventStartAddress).Offset(0, Working(Counter, 20) - 1))
                                        '.Select
                                        '.Merge
                                        '.Name = Working(Counter, 12)
                                        .Interior.Color = RGB(StatusColorRed, StatusColorGreen, StatusColorBlue)
'                                        .Interior.Pattern = xlSolid
'                                        .Interior.PatternColorIndex = xlAutomatic
'                                        .Interior.ThemeColor = xlThemeColorAccent1
'                                        .Interior.TintAndShade = 0.799981688894314
'                                        .Interior.PatternTintAndShade = 0
                                        '.Value = Working(Counter, 3) & "-" & Working(Counter, 4) & "  -  "
                                        .HorizontalAlignment = xlFill
                                        .VerticalAlignment = xlBottom
                                        .WrapText = False
                                        .Orientation = 0
                                        .AddIndent = False
                                        .IndentLevel = 0
                                        .ShrinkToFit = False
                                        '.ReadingOrder = xlContext
                                        .Borders(xlDiagonalDown).LineStyle = xlNone
                                        .Borders(xlDiagonalUp).LineStyle = xlNone
                                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                                        .Borders(xlEdgeLeft).ColorIndex = 0
                                        .Borders(xlEdgeLeft).TintAndShade = 0
                                        .Borders(xlEdgeLeft).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                                        .Borders(xlEdgeRight).ColorIndex = 0
                                        .Borders(xlEdgeRight).TintAndShade = 0
                                        .Borders(xlEdgeRight).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                                        .Borders(xlEdgeTop).ColorIndex = 0
                                        .Borders(xlEdgeTop).TintAndShade = 0
                                        .Borders(xlEdgeTop).Weight = xlThin ' xlMedium
                                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                                        .Borders(xlEdgeBottom).ColorIndex = 0
                                        .Borders(xlEdgeBottom).TintAndShade = 0
                                        .Borders(xlEdgeBottom).Weight = xlThin ' xlMedium
                                    End With
                                    Set EventStartCell = Range(EventStartAddress)
                                    With EventStartCell
                                        .Value = Working(Counter, 2)
                                        .font.Color = FontColor
                                        .font.Bold = True
                                        .font.Size = 12
                                        .AddComment
                                        .Comment.Visible = False
                                        .Comment.Shape.Height = 300
                                        .Comment.Shape.Width = 600
                                        .Comment.Shape.Fill.Visible = msoTrue
                                        .Comment.Shape.Fill.ForeColor.SchemeColor = 0 '0 = black 1 = white 2 = red 3 = green 4 = blue 5 = yellow
                                        .Comment.Shape.TextFrame.Characters.font.Color = -2
                                        .Comment.Shape.TextFrame.Characters.font.Bold = True
                                        .Comment.Shape.TextFrame.Characters.font.Size = 16
                                        '.Comment.Shape.TextFrame.AutoSize = True
                                        .Comment.Text "Task Name: " & Working(Counter, 2) & Chr(10) & Chr(10) & _
                                                      "Range: " & Working(Counter, 1) & Chr(10) & _
                                                      "Date Raised: " & Working(Counter, 8) & Chr(10) & _
                                                      "Date Range: " & Working(Counter, 3) & "-" & Working(Counter, 4) & Chr(10) & _
                                                      "Duration: " & Working(Counter, 20) & " days" & Chr(10) & _
                                                      "Status: " & Working(Counter, 5) & Chr(10) & _
                                                      "Category: " & Working(Counter, 9) & Chr(10) & _
                                                      "Task Type: " & Working(Counter, 11) & Chr(10) & _
                                                      "Old Task Number: " & Working(Counter, 10) & Chr(10) & _
                                                      "Task Number: " & Working(Counter, 15) & Chr(10) & _
                                                      "Request Type: " & Working(Counter, 16) & Chr(10) & _
                                                      "Objective Ref: " & Working(Counter, 12) & Chr(10) & _
                                                      "Customer: " & Working(Counter, 17)


                                    End With
                                    Exit For
                                End If
                            Next CounterRowOffset
                    End If
                End If


CounterEnd:
NotInclude:
        Next Counter
        

    
    
                                    If AreaHasEvents = True Then
                                    
                                    Dim AreaStartAddress As String, AreaStartRange As Range
                                    Dim NextAreaStartInt As Integer
                                    NextAreaStartInt = 0
                                    For Each cell In Range("A2:A1000")
                                        If cell.Interior.ThemeColor = xlThemeColorAccent1 Then
                                            NextAreaStartInt = NextAreaStartInt + 1
                                        Else
                                            Exit For
                                        End If
                                    Next cell
                                    Set AreaStartRange = Worksheets("DateSheet").Range("A2").Offset(NextAreaStartInt, 1)
                                    AreaStartAddress = AreaStartRange.Address(ReferenceStyle:=xlA1)
                                     'AreaStartRange.Select
                                            
                                            
                                            
                                            
'*******************************
                                    








                                    
                                            With Worksheets("DateSheet").Range(AreaStartAddress, _
                                                 Worksheets("DateSheet").Range(AreaStartAddress).Offset(RowOffsetNextAreaStart - 1, 0))
                                                '.Select
                                                .Merge
                                                .font.ThemeColor = FontColor
                                                .Value = UniqueFinalListArea(CounterAreas, 1)
                                                .Interior.Pattern = xlSolid
                                                .Interior.PatternColorIndex = xlAutomatic
                                                .Interior.ThemeColor = xlThemeColorAccent1
                                                .Interior.TintAndShade = 0
                                                .Interior.PatternTintAndShade = 0
                                                .WrapText = True
                                                .HorizontalAlignment = xlCenter
                                                .VerticalAlignment = xlCenter
                                                .font.Name = "Calibri"
                                                .font.FontStyle = "Bold"
                                                .font.Size = 11
                                                .font.Strikethrough = False
                                                .font.Superscript = False
                                                .font.Subscript = False
                                                .font.OutlineFont = False
                                                .font.Shadow = False
                                                .font.Underline = xlUnderlineStyleNone
                                                .font.TintAndShade = 0
                                                .font.ThemeFont = xlThemeFontMinor
                                            End With

                                            Dim NextColorInt As Integer, CellLeftAreaColor As Range, CounterColorNextToAtea As Integer
                                            NextColorInt = 0
                                            For Each cell In Range("A2:A1000")
                                                If cell.Interior.ThemeColor = xlThemeColorAccent1 Then
                                                    NextColorInt = NextColorInt + 1
                                                Else
                                                    Exit For
                                                End If
                                            
                                            Next cell
                                            
                                            Set CellLeftAreaColor = Worksheets("DateSheet").Range("A1").Offset(NextColorInt, 0)
                                            'CellLeftAreaColor.Select
                                            For CounterColorNextToAtea = 1 To RowOffsetNextAreaStart + 1 ' + 1 is to add a space between each area
                                            With CellLeftAreaColor.Offset(CounterColorNextToAtea, 0)
                                                'CellLeftAreaColor.Offset(CounterColorNextToAtea, 0).Select
                                                .Interior.Pattern = xlSolid
                                                .Interior.PatternColorIndex = xlAutomatic
                                                .Interior.ThemeColor = xlThemeColorAccent1
                                                .Interior.TintAndShade = 0
                                                .Interior.PatternTintAndShade = 0
                                            End With
                                            Next CounterColorNextToAtea
                                    

                                    
                                    End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Area Loop Exit





Next CounterAreas


                        'Road Map
                            'Statue Working(x, 5)
                            'Start Date Working(x, 3)
                            'End date Working(x, 4)
                            'Duration Working(x, 7)
                            'Title Working(x, 8)
                            ' CounterUniqueItems
                            'UniqueFinalList(x,1)
                        'StartingPlace (QuarterStartDate | EventStartingCell | FullEventMerge | EventStartDate | EventEndDate)

    For Each cell In Worksheets("DateSheet").Range("B5:B1000")
        If cell.Interior.ThemeColor = xlThemeColorAccent1 Then
            With cell
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).ColorIndex = 0
                .Borders(xlEdgeLeft).TintAndShade = 0
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).ColorIndex = 0
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeRight).Weight = xlMedium
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).ColorIndex = 0
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).ColorIndex = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).Weight = xlMedium
            End With
        Else
            Exit For
        End If
    Next cell

    'Change Cell colour in Area Column
For Each cell In Range("B2:B150")
        'Cell.Select
        Select Case True
        Case cell.Value = "Enviromental Testing Facility"
                    Select Case True
                    Case FontColorEnviromentalTestingFacility = "White"
                        cell.font.Color = -2
                        cell.font.TintAndShade = 0
                    Case Else
                        cell.font.Color = -16777216
                    End Select

        Case cell.Value = "Graytown"
                    Select Case True
                    Case FontColorGraytown = "White"
                        cell.font.Color = -2
                    Case Else
                        cell.font.Color = -16777216
                    End Select

        Case cell.Value = "Headquarters"
                    Select Case True
                    Case FontColorHeadquarters = "White"
                        cell.font.Color = -2
                    Case Else
                        cell.font.Color = -16777216
                    End Select

        Case cell.Value = "Munitions Test Center"
                    Select Case True
                    Case FontColorMunitionsTestCenter = "White"
                        cell.font.Color = -2
                    Case Else
                        cell.font.Color = -16777216
                    End Select

        Case cell.Value = "Port Wakefield"
                    Select Case True
                    Case FontColorPortWakefield = "White"
                        cell.font.Color = -2
                    Case Else
                        cell.font.Color = -16777216
                    End Select

        Case Else
            cell.font.Color = -2
        End Select
    Next cell

    'Dim RangeCollection As Collection
    'Dim EventCollect As Range
    
    For Each EventCollect In RangeCollection
        EventCollect.Merge
        EventCollect.HorizontalAlignment = xlCenter
        'Debug.Print (EventCollect)
    Next EventCollect
          
    




'Public TotalEvents As Long, CurrentEventCount





End Sub



Sub CaptureDuration()

    Dim CounterUniqueItems As Long, CounterLoops As Long, CounterInternalUniqueLoop As Long
    Dim Unique As Boolean
    Dim TableItem As String, FirstItem As String
    Dim ColumnCount As Integer
    Dim RowCount As Long, Counter As Long, CounterInner As Long, CounterStorage1 As Long, CounterStorage2 As Long
    Dim CellRange As Range, cell As Range, DurationPasteRange As Range

    'Row Count less headder
    Set CellRange = Worksheets("Data").Range("A2", Worksheets("Data").Range("A2").Offset(1000000))
    
    For Each cell In CellRange
        If cell.Value = "" Then Exit For
        Counter = Counter + 1
    Next cell
    CounterStorage1 = Counter
    
    'Allicate Data to TableRange
    TableRange = Worksheets("Data").Range("A2", Worksheets("Data").Range("Q3").End(xlDown))
    
    'Adding extra rows to the TableRange
    Dim ExtraTable As Variant, IntermediatTable As Variant
    Dim i As Integer, ii As Integer
    
    
    ExtraTable = Worksheets("Data").Range("BF2", Worksheets("Data").Range("BI3").End(xlDown))
    
    i = UBound(TableRange) + UBound(ExtraTable)
    ReDim IntermediatTable(1 To i, 1 To 17)

    For i = 1 + 1 To UBound(TableRange)
    IntermediatTable(i, 1) = TableRange(i, 1)
    IntermediatTable(i, 2) = TableRange(i, 2)
    IntermediatTable(i, 3) = TableRange(i, 3)
    IntermediatTable(i, 4) = TableRange(i, 4)
    IntermediatTable(i, 5) = TableRange(i, 5)
    IntermediatTable(i, 6) = TableRange(i, 6)
    IntermediatTable(i, 7) = TableRange(i, 7)
    IntermediatTable(i, 8) = TableRange(i, 8)
    IntermediatTable(i, 9) = TableRange(i, 9)
    IntermediatTable(i, 10) = TableRange(i, 10)
    IntermediatTable(i, 11) = TableRange(i, 11)
    IntermediatTable(i, 12) = TableRange(i, 12)
    IntermediatTable(i, 13) = TableRange(i, 13)
    IntermediatTable(i, 14) = TableRange(i, 14)
    IntermediatTable(i, 15) = TableRange(i, 15)
    IntermediatTable(i, 16) = TableRange(i, 16)
    IntermediatTable(i, 17) = TableRange(i, 17)
    Next i


    i = UBound(TableRange) + UBound(ExtraTable) + 1
    ii = 1
    For i = UBound(TableRange) + 1 To UBound(TableRange) + UBound(ExtraTable)
    IntermediatTable(i, 1) = ExtraTable(ii, 1)
    IntermediatTable(i, 2) = ExtraTable(ii, 2)
    IntermediatTable(i, 3) = ExtraTable(ii, 3)
    IntermediatTable(i, 4) = ExtraTable(ii, 4)
    IntermediatTable(i, 5) = "Other"
    
    ii = ii + 1
    Next i
    
    Erase TableRange
    TableRange = IntermediatTable
    
    i = UBound(IntermediatTable)
    ReDim Working(1 To i, 1 To 22)
    For Counter = 1 To 17
        For CounterInner = 1 To i
            Working(CounterInner, Counter) = TableRange(CounterInner, Counter)
        Next CounterInner
    Next Counter
    
    'Calculating Duration
    For Counter = 1 To i
        Working(Counter, 20) = Int((Working(Counter, 4) - Working(Counter, 3)) + 1)
    Next Counter
    
    'Calculating Title Displayed on Sheet
        For Counter = 1 To i
        Working(Counter, 21) = "ID=" & Str(Working(Counter, 6)) & "    " & Working(Counter, 2)
    Next Counter
    
    
    'Assigning Duration to DurationArray
    ReDim DurationArray(1 To Counter - 1)
    For Counter = 1 To i
        DurationArray(Counter) = Working(Counter, 6)
    Next Counter
    


    
    
    
    
    
End Sub



'List Item Address
'http://spapps.defence.gov.au/jcg/jpeu/bms/Lists/JPEU%20Task%20Request/DispForm.aspx?ID=93


Sub DateSheet()


Call UniqueListAreaCaptureArea
Call UniqueListStatusCaptureStatus
Call CaptureDuration


Worksheets("DateSheet").Select
Rows("2:1000").Delete
Columns("C:DH").ColumnWidth = 5
    'Declorations
    Dim DateText As String, StartDateAsText As String
    Dim QuarterOneStart As Date, QuarterTwoStart As Date, QuarterThreeStart As Date, QuarterFourStart As Date
    Dim QuarterOneEnd As Date, QuarterTwoEnd As Date, QuarterThreeEnd As Date, QuarterFourEnd As Date
    Dim StartDateByQuarter As Date, EndDateByQuarter As Date
    Dim DatePlusInt As Date
    Dim FinicalYear As Boolean
    Dim StartDate As Date
    Dim Counter As Long, CounterInner As Long
    Dim StartingCell As Range, cell As Range
    Dim QuarterCounter As Integer
    Dim DateRangeToFormat As Range, DateRangeformatOne As Range, DateRangeformatTwo As Range, DateRangeformatThree As Range
    Dim DateRangeStart As Range
    
    'Adds Test Areas
    'Call Macro11
    
    'Assignments
    Set StartingCell = Worksheets("DateSheet").Range("C2")
    StartYear = Worksheets("Data").Range("AE2").Value
    DateText = "Jan 01, " & StartYear
    StartDate = DateText
    StartDateAsText = StartDate
    StartFinicialYear = StartYear + 1
    
    'Transfered to UserForm
    FinicalYear = Worksheets("Data").Range("AE3").Value
    
    If FinicalYear = True Then
    
        QuarterOneStart = "Jul 01, " & StartYear
        QuarterOneEnd = "Sep 30, " & StartYear
        
        QuarterTwoStart = "Oct 01, " & StartYear
        QuarterTwoEnd = "Dec 31, " & StartYear
        
        QuarterThreeStart = "Jan 01, " & StartFinicialYear
        QuarterThreeEnd = "Mar 31, " & StartFinicialYear
        
        QuarterFourStart = "Apr 01, " & StartFinicialYear
        QuarterFourEnd = "Jun 30, " & StartFinicialYear


    Else
    
        QuarterOneStart = "Jan 01, " & StartYear
        QuarterOneEnd = "Mar 31, " & StartYear
        
        QuarterTwoStart = "Apr 01, " & StartYear
        QuarterTwoEnd = "Jun 30, " & StartYear
        
        QuarterThreeStart = "Jul 01, " & StartYear
        QuarterThreeEnd = "Sep 30, " & StartYear
        
        QuarterFourStart = "Oct 01, " & StartYear
        QuarterFourEnd = "Dec 31, " & StartYear

    End If
    

    
    
'***Count Total Items For Progress Bar
TotalEvents = 0
For Counter = 1 To UBound(Working)
If Working(Counter, 3) < QuarterFourEnd And Working(Counter, 4) > QuarterOneStart Then TotalEvents = TotalEvents + 1
Next Counter
    
    
    
    
    
    
    
    

    
For QuarterCounter = 1 To 4
'StartingCell

        If Range("B2").Value = "" Then
            Set StartingCell = Worksheets("DateSheet").Range("C2")
        ElseIf Range("B2") <> "" Then

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim NextDateStartInt As Integer
NextDateStartInt = 0
For Each cell In Range("A2:A1000")
    If cell.Interior.ThemeColor = xlThemeColorAccent1 Then
        NextDateStartInt = NextDateStartInt + 1
    Else
        Exit For
    End If
Next cell
Set StartingCell = Worksheets("DateSheet").Range("C2").Offset(NextDateStartInt, 0)
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        End If



            If QuarterCounter = 1 Then
                StartDateByQuarter = QuarterOneStart
                EndDateByQuarter = QuarterOneEnd

            ElseIf QuarterCounter = 2 Then
                StartDateByQuarter = QuarterTwoStart
                EndDateByQuarter = QuarterTwoEnd

            ElseIf QuarterCounter = 3 Then
                StartDateByQuarter = QuarterThreeStart
                EndDateByQuarter = QuarterThreeEnd

            ElseIf QuarterCounter = 4 Then
                StartDateByQuarter = QuarterFourStart
                EndDateByQuarter = QuarterFourEnd
            Else
                Exit Sub
            End If
            

'Print Dates to Sheet



'If QuarterCounter = 1 Then
    DatePlusInt = StartDateByQuarter + 1
For Counter = 1 To 3
    StartingCell.Offset(0, -1).Value = "Month"
    StartingCell.Offset(1, -1).Value = "Day"
    StartingCell.Offset(2, -1).Value = "Range"
    StartingCell.Offset(0, -1).Interior.ThemeColor = xlThemeColorAccent1
    StartingCell.Offset(1, -1).Interior.ThemeColor = xlThemeColorAccent1
    StartingCell.Offset(2, -1).Interior.ThemeColor = xlThemeColorAccent1
    StartingCell.Offset(0, -2).Interior.ThemeColor = xlThemeColorAccent1
    StartingCell.Offset(1, -2).Interior.ThemeColor = xlThemeColorAccent1
    StartingCell.Offset(2, -2).Interior.ThemeColor = xlThemeColorAccent1
    For CounterInner = 1 To 94
        If CounterInner = 1 And Counter < 3 Then
            StartingCell.Offset(Counter - 1, 0).Value = StartDateByQuarter
        Else
            If DatePlusInt - 1 = EndDateByQuarter Then
                'CounterInner = 94
                Exit For
            Else
                StartingCell.Offset(Counter - 1, CounterInner - 1).Value = StartDateByQuarter - 1 + CounterInner
                DatePlusInt = StartDateByQuarter + CounterInner
            End If
        End If
    Next CounterInner
    DatePlusInt = StartDateByQuarter + 1
Next Counter


'Formatting Date Rows

For Counter = 1 To 3



    If Range("B2").Value = "" Then
            Set DateRangeformatOne = Worksheets("DateSheet").Range(StartingCell, StartingCell.End(xlToRight))
            Set DateRangeformatTwo = Worksheets("DateSheet").Range(StartingCell.Offset(1, 0), StartingCell.Offset(1, 0).End(xlToRight))
            Set DateRangeformatThree = Worksheets("DateSheet").Range(StartingCell.Offset(2, 0), StartingCell.Offset(2, 0).End(xlToRight))
            Set DateRangeToFormat = Worksheets("DateSheet").Range(StartingCell.Offset(0, -1), StartingCell.Offset(2, 0).End(xlToRight))
    
    ElseIf Range("B2") <> "" Then

            Set DateRangeformatOne = Worksheets("DateSheet").Range(StartingCell, StartingCell.End(xlToRight))
            Set DateRangeformatTwo = Worksheets("DateSheet").Range(StartingCell.Offset(1, 0), StartingCell.Offset(1, 0).End(xlToRight))
            Set DateRangeformatThree = Worksheets("DateSheet").Range(StartingCell.Offset(2, 0), StartingCell.Offset(2, 0).End(xlToRight))
            Set DateRangeToFormat = Worksheets("DateSheet").Range(StartingCell.Offset(0, -1), StartingCell.Offset(2, 0).End(xlToRight))
    End If

    If Counter = 1 Then
        DateRangeformatOne.NumberFormat = "mmm"
        With DateRangeToFormat
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With DateRangeToFormat.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
            .PatternTintAndShade = 0
            '.Color = 15382741
        End With
        DateRangeToFormat.font.Bold = True
        With DateRangeToFormat.font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
    End If
    
    If Counter = 2 Then
        DateRangeformatTwo.NumberFormat = "ddd"
        With DateRangeToFormat
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With DateRangeToFormat.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        DateRangeToFormat.font.Bold = True
        With DateRangeToFormat.font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
    End If
    
    If Counter = 3 Then
        DateRangeformatThree.NumberFormat = "dd"
        With DateRangeToFormat
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With DateRangeToFormat.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        DateRangeToFormat.font.Bold = True
        With DateRangeToFormat.font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
    End If
Next Counter


'Date borders
    With DateRangeToFormat.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium 'xlThin
    End With
    With DateRangeToFormat.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium 'xlThin
    End With
    With DateRangeToFormat.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium 'xlThin
    End With
    With DateRangeToFormat.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium 'xlThin
    End With
    With DateRangeToFormat.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium 'xlThin
    End With
    With DateRangeToFormat.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium 'xlThin
    End With


    



    'Set StartingCell = StartingCell.Offset(3, 0)
'Call TempFillArea

Call PlaceEventsOnSheet
If QuarterCounter = 1 Then MasterScheduleRefresh.ProgressBar1.Value = 10
If QuarterCounter = 2 Then MasterScheduleRefresh.ProgressBar1.Value = 20
If QuarterCounter = 3 Then MasterScheduleRefresh.ProgressBar1.Value = 30
If QuarterCounter = 4 Then MasterScheduleRefresh.ProgressBar1.Value = 40
Next QuarterCounter


End Sub


Sub TempFillArea()
Dim Rng As Range
Set Rng = Range("B5").End(xlDown)
    If Range("B5").Value = "" Then
        Range("B2").Offset(0, 0).Value = "Test"
        Range("B2").Offset(1, 0).Value = "Test"
        Range("B2").Offset(2, 0).Value = "Test"
        Range("B2").Offset(3, 0).Value = "Test"
        Range("B2").Offset(4, 0).Value = "Test"
        Range("B2").Offset(5, 0).Value = "Test"
    ElseIf Range("B5") <> "" Then
        Rng.Offset(1, 0).Value = "Test"
        Rng.Offset(2, 0).Value = "Test"
        Rng.Offset(3, 0).Value = "Test"
        Rng.Offset(4, 0).Value = "Test"
        Rng.Offset(5, 0).Value = "Test"
        Rng.Offset(6, 0).Value = "Test"
        Rng.Offset(7, 0).Value = "Test"
    End If
End Sub





Sub Macro11()
'
' Macro11 Macro
'

'
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("B8").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("B12").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("B13").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("B14").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("B18").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("B19").Select
    ActiveCell.FormulaR1C1 = "Area"
    
    Range("B20").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("B21").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("B22").Select
    ActiveCell.FormulaR1C1 = "Area"


End Sub






Sub UniqueListAreaCaptureArea()
    Dim CounterUniqueItems As Long, CounterLoops As Long, CounterInternalUniqueLoop As Long
    Dim Unique As Boolean
    Dim TableItem As String, FirstItem As String
    CounterInternalUniqueLoop = 0
    CounterUniqueItems = 1
    TableRange = Worksheets("Data").Range("A2", Worksheets("Data").Range("F3").End(xlDown))
    
    
    Worksheets("Data").Range("K2:K132").ClearContents

For CounterLoops = 1 To UBound(TableRange, 1) '- 1
TableItem = TableRange(CounterLoops, 1)
'Debug.Print (TableItem)
    If CounterLoops = 1 Then
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ReDim UniqueListArea(10000, 0)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        UniqueListArea(0, 0) = TableItem
        CounterUniqueItems = 2
    Else
                
            For CounterInternalUniqueLoop = 1 To CounterLoops - 1
                    'CounterInternalUniqueLoop = CounterInternalUniqueLoop + 1
                    If TableRange(CounterLoops, 1) = UniqueListArea(CounterInternalUniqueLoop - 1, 0) Then
                        CounterInternalUniqueLoop = CounterLoops
                        Unique = False
                        'Debug.Print (TableRange(CounterLoops, 2) & "-" & UniqueListArea(CounterInternalUniqueLoop - 1, 0))
                    Else
                        Unique = True
                        'Debug.Print (TableRange(CounterLoops, 2) & "-" & UniqueListArea(CounterInternalUniqueLoop - 1, 0))
                    End If
    
            Next CounterInternalUniqueLoop

        If Unique = True Then
        UniqueListArea(CounterUniqueItems - 1, 0) = TableItem
        CounterUniqueItems = CounterUniqueItems + 1
        End If
        
    End If
  
'CounterLoops = CounterLoops + 1
Next CounterLoops
TableItem = ""

Dim CountUniqueRows As Long, NextCountUniqueRow As Long

For NextCountUniqueRow = 0 To 10000

If UniqueListArea(NextCountUniqueRow, 0) <> "" Then CountUniqueRows = CountUniqueRows + 1

Next NextCountUniqueRow
CountUniqueRows = CountUniqueRows + 1

'(Optional)


Worksheets("Data").Range("AA2", Worksheets("Data").Range("AA1").Offset(CountUniqueRows, 0)).Value = UniqueListArea
    Worksheets("Data").Sort.SortFields.Clear
    Worksheets("Data").Sort.SortFields.Add Key:=Worksheets("Data").Range("AA2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Worksheets("Data").Sort
        .SetRange Worksheets("Data").Range("AA2", Worksheets("Data").Range("AA1").Offset(CountUniqueRows, 0))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'Worksheets("Data").Range("H2").Select
UniqueFinalListArea = Worksheets("Data").Range("AA2", Worksheets("Data").Range("AA1").Offset(CountUniqueRows, 0)).Value


End Sub


Sub DateColor()
'Application.ScreenUpdating = False
Dim cell As Range
On Error Resume Next
Dim ws As Worksheet
Dim WeekdayRed As Integer, WeekdayGreen As Integer, WeekdayBlue As Integer
Dim WeekdendRed As Integer, WeekdendGreen As Integer, WeekendBlue As Integer
Dim WeekDayfontColor As String, WeekendfontColor As String
Dim PrgressCounter As Double
Dim WeekdayFontInt As Integer, WeekendFontInt As Integer



WeekdayRed = Worksheets("Data").Range("AG17").Value
WeekdayGreen = Worksheets("Data").Range("AH17").Value
WeekdayBlue = Worksheets("Data").Range("AI17").Value
WeekDayfontColor = Worksheets("Data").Range("AK17").Value

If WeekDayfontColor = White Then
WeekdayFontInt = -2
Else
WeekdayFontInt = -1
End If

WeekendRed = Worksheets("Data").Range("AG18").Value
WeekendGreen = Worksheets("Data").Range("AH18").Value
WeekendBlue = Worksheets("Data").Range("AI18").Value
WeekendfontColor = Worksheets("Data").Range("AK18").Value

If WeekendfontColor = White Then
WeekendFontInt = -2
Else
WeekendFontInt = -1
End If

PrgressCounter = 0
For Each cell In Worksheets("DateSheet").Range("C2:CR200")

                Dim TestCounter As Double
                PrgressCounter = PrgressCounter + 1
                TestCounter = Round(PrgressCounter / Range("C2:CR200").Count, 0) * 100 / 10
                MasterScheduleRefresh.ProgressBar1.Value = 70 + TestCounter
                
If IsDate(cell.Value) = True Then
    If IsWeekend(cell.Value) = True Then
    
    cell.Interior.Color = RGB(WeekendRed, WeekendGreen, WeekendBlue)
    cell.font.Color = WeekendFontInt
    cell.font.Bold = True
    cell.font.Size = 12
    
    
    
    
    
            Dim ii As Integer
            Dim iiDate As Integer
            Dim DateRowCell As Range
            iiDate = 0
            ii = 0
                Set DateRowCell = cell.Offset(3 + iiDate, 0)

                
                Do While IsDate(DateRowCell.Value) = False And ii <= 100
                Set DateRowCell = cell.Offset(3 + iiDate, 0)
                    If DateRowCell.Interior.Color = xlNone Or DateRowCell.Value = "" Then
                        With DateRowCell
                            .Interior.Color = RGB(WeekendRed, WeekendGreen, WeekendBlue)
                            .Value = " "
                        End With
                    End If
                    iiDate = iiDate + 1
                    ii = ii + 1
                Loop
                
                
                
    
    Else
    
    cell.Interior.Color = RGB(WeekdayRed, WeekdayGreen, WeekdayBlue)
    cell.font.Color = WeekdayFontInt
    cell.font.Bold = True
    cell.font.Size = 12
    End If
    

End If


Next cell
'Application.ScreenUpdating = True
End Sub

Public Function IsWeekend(InputDate As Date) As Boolean
    Select Case Weekday(InputDate)
        Case vbSaturday, vbSunday
            IsWeekend = True
        Case Else
            IsWeekend = False
        End Select
End Function


Sub BlankAreaFill()
'Application.ScreenUpdating = False
Dim cell As Range
Dim Area As Range
Dim i As Integer
Dim CellAddress As String
Dim WeekdayRed As Integer, WeekdayGreen As Integer, WeekdayBlue As Integer
Dim WeekDayfontColor As String
Dim WeekdayFontInt As Integer

WeekdayRed = Worksheets("Data").Range("AG17").Value
WeekdayGreen = Worksheets("Data").Range("AH17").Value
WeekdayBlue = Worksheets("Data").Range("AI17").Value
WeekDayfontColor = Worksheets("Data").Range("AK17").Value

If WeekDayfontColor = White Then
WeekdayFontInt = -2
Else
WeekdayFontInt = -1
End If





For Each cell In Worksheets("DateSheet").Range("B1:B1200")
    If cell.Value = "" Then
        With cell.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -4.99893185216834E-02
            .PatternTintAndShade = 0
        End With
    End If
Next cell

For Each cell In Worksheets("DateSheet").Range("B1:B200")
    If cell.Value <> "" And cell.Value <> "Month" And cell.Value <> "Day" And cell.Value <> "Area" Then
        Set Area = cell
        i = 0
        Do While Area.Value <> ""
        i = i + 1
        Set Area = Area.Offset(i, 0)
        Loop
        
        CellAddress = cell.Address(ReferenceStyle:=xlA1)
        Set Area = Range(CellAddress, Range(CellAddress).Offset(i - 1, 0))
        Area.Select
        
        With Selection
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).ColorIndex = 0
                .Borders(xlEdgeLeft).TintAndShade = 0
                .Borders(xlEdgeLeft).Weight = xlMedium
                
            .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).ColorIndex = 0
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeRight).Weight = xlMedium
                
            .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).ColorIndex = 0
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlMedium
                
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).ColorIndex = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).Weight = xlMedium
                
            .font.Bold = True
            .font.Size = 12
        End With

    End If
    
    
    If cell.Value = "Month" Or cell.Value = "Day" Or cell.Value = "Range" Then
    
    cell.font.Color = WeekdayFontInt
    cell.Interior.Color = RGB(WeekdayRed, WeekdayGreen, WeekdayBlue)
    cell.font.Size = 12
    
    End If
Next cell

'Application.ScreenUpdating = True
End Sub


Sub MonthMerge()
'Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim cell As Range, ActRange As Range, activve As Range, ActWorking As Range
Dim i As Integer, DateRange As Integer, ii As Integer
Dim MonthAddress As String
Dim Months As Collection
Dim Monthcollect As Range
Dim WeekdayRed As Integer, WeekdayGreen As Integer, WeekdayBlue As Integer
Dim WeekDayfontColor As String
Dim WeekdayFontInt As Integer

WeekdayRed = Worksheets("Data").Range("AG17").Value
WeekdayGreen = Worksheets("Data").Range("AH17").Value
WeekdayBlue = Worksheets("Data").Range("AI17").Value
WeekDayfontColor = Worksheets("Data").Range("AK17").Value

If WeekDayfontColor = White Then
WeekdayFontInt = -2
Else
WeekdayFontInt = -1
End If
Set Months = New Collection

For Each cell In Range("B2:B200")

    If cell.Value = "Month" Then
    Set act = cell.Offset(0, 1)
    Set ActWorking = act
    DateRange = 0
    act.Select
    Do While IsDate(ActiveCell.Value) = True
    act.Offset(0, DateRange).Select
    DateRange = DateRange + 1
    Loop
    DateRange = DateRange - 1
    
    
    For ii = 1 To 3
    act.Select
                Set ActWorking = act
                i = 0
                Do While Month(act.Value) = Month(ActWorking.Value) And ActWorking.Value <> ""
                ActWorking.Select
                Set ActWorking = act.Offset(0, i)
                i = i + 1
                Loop
 
                MonthAddress = act.Address(ReferenceStyle:=xlA1)
                Range(MonthAddress, Range(MonthAddress).Offset(0, i - 2)).Select
                Months.Add Range(MonthAddress, Range(MonthAddress).Offset(0, i - 2))
                Set act = ActWorking
    
    Next ii

End If
Next cell

For Each Monthcollect In Months
Monthcollect.Merge
Monthcollect.font.Color = WeekdayFontInt
Monthcollect.Interior.Color = RGB(WeekdayRed, WeekdayGreen, WeekdayBlue)
Monthcollect.NumberFormat = "mmmm"
Next Monthcollect
Application.DisplayAlerts = True
'Application.ScreenUpdating = True
End Sub
