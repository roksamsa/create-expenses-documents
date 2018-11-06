@ECHO OFF
setlocal enableDelayedExpansion

title Costs documents creator

ECHO ------------------------------------------------------------------------------------
ECHO Costs documents creator
ECHO Tool for making easier life with costs files.
ECHO Author: Rok Samsa, rok.samsa@gmail.com
ECHO ------------------------------------------------------------------------------------
ECHO[

REM Name of this files
SET me=%~n0

REM Current location
SET currentLocation=%0
SET log=%currentLocation%%me%.txt

SET adria="adria-airways"
SET goOpti="goopti"
SET hotelMotelOne="hotel-motel-one"
SET hotelBold="hotel-bold"
SET train="train"

SET /p year="For which year costs are? "
ECHO[
ECHO ----------
ECHO[
SET /p month="For which month costs are? "
ECHO[
ECHO ----------
ECHO[

SET nameOfMonth=None
IF %month%==1 (
	SET nameOfMonth=January
)
IF %month%==2 (
	SET nameOfMonth=February
)
IF %month%==3 (
	SET nameOfMonth=March
)
IF %month%==4 (
	SET nameOfMonth=April
)
IF %month%==5 (
	SET nameOfMonth=May
)
IF %month%==6 (
	SET nameOfMonth=June
)
IF %month%==7 (
	SET nameOfMonth=July
)
IF %month%==8 (
	SET nameOfMonth=August
)
IF %month%==9 (
	SET nameOfMonth=September
)
IF %month%==10 (
	SET nameOfMonth=October
)
IF %month%==11 (
	SET nameOfMonth=November
)
IF %month%==12 (
	SET nameOfMonth=December
)

SET /p tripsThisMonth="How many trips did you had in %nameOfMonth%? "
ECHO[
ECHO ----------
ECHO[

REM Only 1 trip this month
IF "%tripsThisMonth%"=="1" (
	SET /p firstTravel="(1) Date (just day) when you went to work? "
	ECHO ----------
	SET /p secondTravel="(2) Date (just day) when you went back home? "
	ECHO ----------
	SET /p travelStartTime="At which time did you went on trip? "
	ECHO ----------
	SET /p travelEndTime="At which time did you come home from trip? "
	ECHO ----------
	SET /p firstHotel="Name of 1st hotel? "

	REM Rename for Adria Airways
	REN %year%\%month%\1\PDF*.pdf Beleg-%year:~2,2%%month%01-%firstTravel%-%month%-%year%-%adria%.pdf

	REM Rename for Hotels
	REN %year%\%month%\1\hotel.pdf Beleg-%year:~2,2%%month%02-%firstTravel%-%month%-%year%-%firstHotel%.pdf

	REM Rename for GoOpti
	REN %year%\%month%\1\GoOpti-1.pdf Beleg-%year:~2,2%%month%03-%firstTravel%-%month%-%year%-%goOpti%.pdf
	REN %year%\%month%\1\GoOpti-2.pdf Beleg-%year:~2,2%%month%04-%secondTravel%-%month%-%year%-%goOpti%.pdf

	REM Rename for Trains
	REN %year%\%month%\1\train.pdf Beleg-%year:~2,2%%month%05-%firstTravel%-%month%-%year%-Beleg-%year:~2,2%%month%06-%secondTravel%-%month%-%year%-%train%.pdf

	REM Daily allowance
	if not exist "%year%\%month%\1\Daily-allowance_Spesenblatt_SamsaR*" (
		xcopy /s /y Template\Daily-allowance_Spesenblatt_SamsaR.xlsx %year%\%month%\1
		REN %year%\%month%\1\Daily-allowance_Spesenblatt_SamsaR.xlsx Daily-allowance_Spesenblatt_SamsaR_%month%%year%.xlsx
	)
)

REM 2 trips this month
IF "%tripsThisMonth%"=="2" (
	SET /p firstTravel="(1) First trip date (just day) when you went to work? "
	ECHO[
	ECHO ----------
	ECHO[
	SET /p secondTravel="(2) Second trip date (just day) when you went back home? "
	ECHO[
	ECHO ----------
	ECHO[
	SET /p thirdTravel="(3) Third trip date (just day) when you went to work? "
	ECHO[
	ECHO ----------
	ECHO[
	SET /p forthTravel="(4) Forth trip date (just day) when you went back home? "
	ECHO[
	ECHO ----------
	ECHO[
	SET /p travelStartTime1="(1) First trip time when did you went on trip? "
	ECHO[
	ECHO ----------
	ECHO[
	SET /p travelEndTime1="(2) Second trip time when did you come home from trip? "
	ECHO[
	ECHO ----------
	ECHO[
	SET /p travelStartTime2="(3) Third trip time when did you went on trip? "
	ECHO[
	ECHO ----------
	ECHO[
	SET /p travelEndTime2="(4) Forth trip time when did you come home from trip? "
	ECHO[
	ECHO ----------
	ECHO[
	SET /p firstHotel="Name of 1st hotel? "
	ECHO[
	ECHO ----------
	ECHO[
	SET /p secondHotel="Name of 2nd hotel? "

	REM Rename for Adria Airways
	REN %year%\%month%\1\PDF*.pdf Beleg-%year:~2,2%%month%01-%firstTravel%-%month%-%year%-%adria%.pdf
	REN %year%\%month%\2\PDF*.pdf Beleg-%year:~2,2%%month%07-%thirdTravel%-%month%-%year%-%adria%.pdf

	REM Rename for Hotels
	REN %year%\%month%\1\hotel.pdf Beleg-%year:~2,2%%month%02-%firstTravel%-%month%-%year%-%firstHotel%.pdf
	REN %year%\%month%\2\hotel.pdf Beleg-%year:~2,2%%month%08-%thirdTravel%-%month%-%year%-%secondHotel%.pdf

	REM Rename for GoOpti
	REN %year%\%month%\1\GoOpti-1.pdf Beleg-%year:~2,2%%month%03-%firstTravel%-%month%-%year%-%goOpti%.pdf
	REN %year%\%month%\1\GoOpti-2.pdf Beleg-%year:~2,2%%month%04-%secondTravel%-%month%-%year%-%goOpti%.pdf
	REN %year%\%month%\2\GoOpti-3.pdf Beleg-%year:~2,2%%month%09-%thirdTravel%-%month%-%year%-%goOpti%.pdf
	REN %year%\%month%\2\GoOpti-4.pdf Beleg-%year:~2,2%%month%10-%forthTravel%-%month%-%year%-%goOpti%.pdf

	REM Rename for Trains
	REN %year%\%month%\1\train.pdf Beleg-%year:~2,2%%month%05-%firstTravel%-%month%-%year%-Beleg-%year:~2,2%%month%06-%secondTravel%-%month%-%year%-%train%.pdf
	REN %year%\%month%\2\train.pdf Beleg-%year:~2,2%%month%11-%firstTravel%-%month%-%year%-Beleg-%year:~2,2%%month%12-%secondTravel%-%month%-%year%-%train%.pdf

	REM Daily allowance
	if not exist "%year%\%month%\1\Daily-allowance_Spesenblatt_SamsaR*" (
		xcopy /s /y Template\Daily-allowance_Spesenblatt_SamsaR.xlsx %year%\%month%\1
		REN %year%\%month%\1\Daily-allowance_Spesenblatt_SamsaR.xlsx Daily-allowance_Spesenblatt_SamsaR_%month%%year%_01.xlsx
	)
	if not exist "%year%\%month%\2\Daily-allowance_Spesenblatt_SamsaR*" (
		xcopy /s /y Template\Daily-allowance_Spesenblatt_SamsaR.xlsx %year%\%month%\2
		REN %year%\%month%\2\Daily-allowance_Spesenblatt_SamsaR.xlsx Daily-allowance_Spesenblatt_SamsaR_%month%%year%_02.xlsx
	)
)

cscript.exe //NoLogo create-costs.vbs /costsYear:"%year%" /costsMonth:"%month%" /tripsThisMonth:"%tripsThisMonth%" /firstTravel:"%firstTravel%" /secondTravel:"%secondTravel%" /thirdTravel:"%thirdTravel%" /forthTravel:"%forthTravel%" /travelStartTime1:"%travelStartTime1%" /travelEndTime1:"%travelEndTime1%" /travelStartTime2:"%travelStartTime2%" /travelEndTime2:"%travelEndTime2%" /travelStartTime:"%travelStartTime%" /travelEndTime:"%travelEndTime%"

REM Main Excel file
if not exist "%year%\%month%\Spesenblatt_SamsaR*" (
	xcopy /s /y Template\Spesenblatt_SamsaR.xlsx %year%\%month%\
	REN %year%\%month%\Spesenblatt_SamsaR.xlsx Spesenblatt_SamsaR_%month%%year%.xlsx
)

ECHO ------------------------------------------------------------------------------------
ECHO Creating documents for costs in "%nameOfMonth%" finished! Happy to the moon!

pause
