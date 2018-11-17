{\rtf1\ansi\ansicpg1252\cocoartf1561\cocoasubrtf600
{\fonttbl\f0\fnil\fcharset0 Menlo-Regular;}
{\colortbl;\red255\green255\blue255;\red0\green0\blue128;\red255\green255\blue255;\red0\green0\blue0;
\red0\green128\blue0;}
{\*\expandedcolortbl;;\csgenericrgb\c0\c0\c50196;\csgenericrgb\c100000\c100000\c100000;\csgenericrgb\c0\c0\c0;
\csgenericrgb\c0\c50196\c0;}
\margl1440\margr1440\vieww21420\viewh8400\viewkind0
\pard\tx560\tx1120\tx1680\tx2240\tx2800\tx3360\tx3920\tx4480\tx5040\tx5600\tx6160\tx6720\pardirnatural\partightenfactor0

\f0\fs22 \cf2 \cb3 Sub\cf4  stockcode_easy():\
\cf0 \
\cf5 'variables\
\cf0 \
\cf2 Dim\cf4  Data(50, 20000, 2, 3, 2) \cf5 'data storage array: sheets, tickers, data volume, open and close date, open and close price\
'year, ticker name(x), and total stock volume trade(1), openCloseDate (1-open,2-close)\
'startEndPrice (0-start,1end)\
\cf2 Dim\cf4  data_year \cf2 As\cf4  \cf2 Integer\
Dim\cf4  data_ticker \cf2 As\cf4  \cf2 Integer\
Dim\cf4  data_totalVolume \cf2 As\cf4  \cf2 Integer\
Dim\cf4  openCloseDate \cf2 As\cf4  \cf2 Integer\
Dim\cf4  startEndPrice \cf2 As\cf4  \cf2 Integer\
Dim\cf4  SheetName \cf2 As\cf4  \cf2 String\
\cf0 \
\
\cf2 Dim\cf4  row \cf2 As\cf4  \cf2 Integer\
Dim\cf4  col \cf2 As\cf4  \cf2 Integer\
\cf0 \
\cf5 'initialize variables\
\cf4 data_year = 0\
data_ticker = 0\
data_totalVolume = 0\
openCloseDate = 0\
startEndPrice = 0\
\cf2 Erase\cf4  Data\
\cf0 \
\cf4 ws = ActiveWorkbook.Worksheets.Count\
row = 0\
col = 0\
\cf0 \
\
\cf5 '''''''''''''''''''''START EVALUATION FUNCTION'''''''''''''''''''\
\cf0 \
\cf5 'for loop for start for sheet sheet\
\cf2 For\cf4  data_year = 1 \cf2 To\cf4  ws\
    \cf5 'input sheet name (date year), into Data(x,0,0)\
\cf4     Data(data_year, data_ticker, data_totalVolume, openCloseDate, startEndPrice) = ActiveWorkbook.Worksheets(data_year).Name\
    \cf5 'use inputted name to activate worksheet\
\cf4     Worksheets(Data(data_year, data_ticker, data_totalVolume, openCloseDate, startEndPrice)).Activate\
    \cf5 'call out message box denoting which sheet and name of sheet\
\cf4     MsgBox ("Sheet number = " & data_year & vbNewLine & "Evaluating Sheet: " & Data(data_year, 0, 0, 0, 0))\
\cf0 \
\
\cf5 ''''''START EVALAUTION''''''\
\cf2 For\cf4  I = 2 \cf2 To\cf4  1500000\
\cf0 \
\cf5 'check if there is another ticker to add, end loop if not\
\cf2 If\cf4  Cells(I, 1).Value = "" \cf2 Then\
\cf4     MsgBox ("Evaluation ended due to blank ticker row")\
    \cf2 Exit\cf4  \cf2 For\
End\cf4  \cf2 If\
\cf0 \
\cf5 'EVALUATE TICKER NAME FOR CHANGE\
\cf2 If\cf4  Cells(I, 1).Value <> Cells(I - 1, 1).Value \cf2 Then\cf4  \cf5 ' compare current stored ticker name vs current cell\
\cf4     \cf5 'increment data_ticker to next value for next ticker\
\cf4     data_ticker = data_ticker + 1\
    \cf5 'set data_totalVolume to 0 for data ticker name\
\cf4     data_totalVolume = 0\
    \cf5 'input data_ticker name\
\cf4     Data(data_year, data_ticker, data_totalVolume, openCloseDate, startEndPrice) = Cells(I, 1).Value\
    \cf5 'check message for new data ticker name\
\cf4     \cf5 'MsgBox ("ticker changed to " + Data(data_year, data_ticker, data_totalVolume))\
\cf4     \cf5 'set data_totalVolume to 1 for data input\
\cf4     data_totalVolume = 1\
    \
    \cf5 'initialize moderate variables\
\cf4     \cf5 'get open date\
\cf4     Data(data_year, data_ticker, data_totalVolume, 1, 0) = Cells(I, 2).Value\
    \cf5 'get open price\
\cf4     Data(data_year, data_ticker, data_totalVolume, 1, 1) = Cells(I, 3).Value\
    \cf5 'get close date\
\cf4     Data(data_year, data_ticker, data_totalVolume, 2, 0) = Cells(I, 2).Value\
    \cf5 'get close price\
\cf4     Data(data_year, data_ticker, data_totalVolume, 2, 1) = Cells(I, 6).Value\
    \
\cf0 \
\cf2 End\cf4  \cf2 If\
\cf0 \
\
\cf5 'input data volume\
\cf4 Data(data_year, data_ticker, data_totalVolume, 0, 0) = Data(data_year, data_ticker, data_totalVolume, 0, 0) + Cells(I, 7).Value\
\cf0 \
\cf5 'EVALUATE MODERATE VARIABLES!!!!!\
\cf0 \
\cf4 openCloseDate = 1 \cf5 ' denotes open date\
\cf4 startEndPrice = 0 \cf5 ' denotes open date\
'if the date is smaller/earlier than the last noted open date, replace open date and open price\
\cf2 If\cf4  Cells(I, 2).Value < Data(data_year, data_ticker, data_totalVolume, 1, 0) \cf2 Then\
\cf4     \cf5 'update open date\
\cf4     Data(data_year, data_ticker, data_totalVolume, 1, 0) = Cells(I, 2).Value\
    startEndPrice = 1 \cf5 ' denote open price\
\cf4     \cf5 'update open price\
\cf4     Data(data_year, data_ticker, data_totalVolume, 1, 1) = Cells(I, 3).Value\
\cf2 End\cf4  \cf2 If\
\cf0 \
\cf4 openCloseDate = 2 \cf5 ' denotes close date\
\cf4 startEndPrice = 0 \cf5 ' denotes close date\
'if the date is larger/later than the last noted close date, replace close date and close price\
\cf2 If\cf4  Cells(I, 2).Value > Data(data_year, data_ticker, data_totalVolume, openCloseDate, startEndPrice) \cf2 Then\
\cf4     \cf5 'update close date\
\cf4     Data(data_year, data_ticker, data_totalVolume, 2, 0) = Cells(I, 2).Value\
    startEndPrice = 1 \cf5 ' denote close price\
\cf4     \cf5 'update close price\
\cf4     Data(data_year, data_ticker, data_totalVolume, 2, 1) = Cells(I, 6).Value\
\cf2 End\cf4  \cf2 If\
\cf0 \
\cf5 'reset openCloseDate and startEndPrice\
\cf4 openCloseDate = 0 \cf5 ' denotes close date\
\cf4 startEndPrice = 0 \cf5 ' denotes close date\
\cf0 \
\cf5 'checking write volume data to row\
'Cells(i, 11).Value = Data(data_year, data_ticker, data_totalVolume)\
\cf0 \
\cf2 Next\cf4  I\
\cf0 \
\
\cf5 '''''''''''''''''WRITE RESULTS'''''''''''''''''''\
'formatting - write total in column name\
'ticker name\
\cf4 Cells(1, 10) = "_ticker_name"\
\cf5 'yearly delta\
\cf4 Cells(1, 11) = "_yearly_change"\
\cf5 'yearly delta currency format for range\
\cf4 Range("K2:K10000").NumberFormat = "$#,##0.00"\
\cf5 'percent change\
\cf4 Cells(1, 12) = "_percent_change"\
\cf5 'percent change percent format for range\
\cf4 Range("L2:L10000").NumberFormat = "0.00%"\
\cf5 'total stock volume\
\cf4 Cells(1, 13) = "_total_stock_volume"\
\cf5 'check stock open price and currency format\
'Cells(1, 14) = "_stock_open_price"\
'Range("N2:N10000").NumberFormat = "$#,##0.00"\
'check stock close price and currency format\
'Cells(1, 15) = "_stock_close_price"\
'Range("O2:O10000").NumberFormat = "$#,##0.00"\
\cf0 \
\cf5 'format tittles for hard\
\cf4 Cells(2, 19) = "_greatest_%_increase"\
Cells(3, 19) = "_greatest_%_decrease"\
Cells(4, 19) = "_greatest_total_volume"\
Cells(1, 20) = "_ticker"\
Cells(1, 21) = "_value"\
Cells(2, 21).NumberFormat = "0.00%"\
Cells(3, 21).NumberFormat = "0.00%"\
\cf0 \
\
\
\
\cf4 MsgBox ("Writing Results")\
\cf0 \
\cf2 For\cf4  row = 1 \cf2 To\cf4  20000\
    \cf5 'if ticker name is empty, exit for loop\
\cf4     \cf2 If\cf4  IsEmpty(Data(data_year, row, 0, 0, 0)) = \cf2 True\cf4  \cf2 Then\
\cf4     \cf2 Exit\cf4  \cf2 For\
End\cf4  \cf2 If\
\cf0 \
\cf5 'write/print data TABLE\
'write ticker name - row+1 due to formatting in first row\
\cf4 Cells(row + 1, 10).Value = Data(data_year, row, 0, 0, 0)\
\cf5 'write yearly change open to close\
\cf4 Cells(row + 1, 11).Value = Data(data_year, row, 1, 2, 1) - Data(data_year, row, 1, 1, 1)\
\cf5 'conditional formatting for positive or negative change\
\cf2 If\cf4  Cells(row + 1, 11).Value < 0 \cf2 Then\
\cf4    Cells(row + 1, 11).Interior.ColorIndex = 3\
\cf2 ElseIf\cf4  Cells(row + 1, 11).Value >= 0 \cf2 Then\
\cf4     Cells(row + 1, 11).Interior.ColorIndex = 4\
\cf2 Else\
\cf4     Cells(row + 1, 11).Interior.ColorIndex = 2\
\cf2 End\cf4  \cf2 If\
\cf5 'calculate and write price change and percentage\
\cf2 If\cf4  Data(data_year, row, 1, 1, 1) <> 0 \cf2 Then\
\cf5 '=((close-open)/open)\
\cf4     Cells(row + 1, 12).Value = ((Data(data_year, row, 1, 2, 1) - Data(data_year, row, 1, 1, 1)) / Data(data_year, row, 1, 1, 1))\
\cf2 Else\
\cf4     Cells(row + 1, 12).Value = "n/a"\
\cf2 End\cf4  \cf2 If\
\cf5 'write total volume\
\cf4 Cells(row + 1, 13).Value = Data(data_year, row, 1, 0, 0) \cf5 ' good\
'write open price\
'Cells(row + 1, 14).Value = Data(data_year, row, 1, 1, 1) ' good\
'write close price\
'Cells(row + 1, 15).Value = Data(data_year, row, 1, 2, 1) ' good\
'write open date\
'Cells(row + 1, 16).Value = Data(data_year, row, 1, 1, 0) ' test\
'write close date\
'Cells(row + 1, 17).Value = Data(data_year, row, 1, 2, 0) ' test\
\cf0 \
\cf5 'if statement for greatest positive percent increase\
'compare last written row to current greatest positive percent increase with check if number\
\cf2 If\cf4  IsNumeric(Cells(row + 1, 12).Value) = \cf2 True\cf4  \cf2 And\cf4  Cells(row + 1, 12).Value > Cells(2, 21).Value \cf2 Then\
\cf4         Cells(2, 21).Value = Cells(row + 1, 12).Value\
        Cells(2, 20).Value = Cells(row + 1, 10).Value\
    \cf2 End\cf4  \cf2 If\
\cf5 'if statement for greatest negative percent increase\
'compare last written row to current greatest negative percent increase w check if number\
\cf2 If\cf4  IsNumeric(Cells(row + 1, 12).Value) = \cf2 True\cf4  \cf2 And\cf4  Cells(row + 1, 12).Value < Cells(3, 21).Value \cf2 Then\
\cf4         Cells(3, 21).Value = Cells(row + 1, 12).Value\
        Cells(3, 20).Value = Cells(row + 1, 10).Value\
    \cf2 End\cf4  \cf2 If\
\cf5 'if statement for greatest total volume w check if number\
\cf2 If\cf4  IsNumeric(Cells(row + 1, 13).Value) = \cf2 True\cf4  \cf2 And\cf4  Cells(row + 1, 13).Value > Cells(4, 21).Value \cf2 Then\
\cf4         Cells(4, 21).Value = Cells(row + 1, 13).Value\
        Cells(4, 20).Value = Cells(row + 1, 10).Value\
    \cf2 End\cf4  \cf2 If\
\cf0 \
\
\cf5 'advance to next row\
\cf2 Next\cf4  row\
\cf0 \
\cf5 'message showing finished write\
\cf4 MsgBox ("Results write ended" & vbNewLine & "Blank ticker at row " & row & vbNewLine & "Results table written for sheet " + Data(data_year, 0, 0, 0, 0))\
\cf0 \
\
\cf5 'reset counter variables\
\cf4 data_ticker = 0\
data_totalVolume = 0\
openCloseDate = 0\
startEndPrice = 0\
\cf0 \
\cf5 'format all cells to autofit text\
\cf4 Range("J1:Q1").Columns.AutoFit\
Columns("S").AutoFit\
Columns("U").AutoFit\
\cf0 \
\
\cf5 'increment year loop\
\cf2 Next\cf4  data_year\
\cf0 \
\cf4 MsgBox ("DONE!")\
\cf0 \
\
\cf2 End\cf4  \cf2 Sub\
}