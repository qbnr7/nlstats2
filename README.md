# nlstats2
NL Stats - American Football Officiating Analysis
This project collects and analyses officiating grades from American football games in the Danish National League (NL). The goal is to provide officials, crew chiefs, and supervisors with structured performance data based on graded penalty calls.
How It Works
Each game is recorded in an Excel file with one row per play. When a penalty occurs, the row includes a foul code, a grade for the call, and a code that links the grade to the specific official(s) involved. A separate schedule file tracks which officials worked each game and in what positions.
Grading System
Each penalty call is graded on the following scale:
CodeNameScoreCC / CCorrect Call100MC / MMarginal Call75NC / NNo Call50IC / IIncorrect Call0NG / GNon-gradeableExcludedWWaivedExcluded
Official Position Codes
Grades are linked to officials using single-letter position codes — for example BC means the Back Judge received a Correct Call grade, and BCHC means both the Back Judge and Head Linesman received Correct Call grades.
CodePositionRRefereeUUmpireHHead LinesmanLLine JudgeBBack JudgeFField JudgeSSide JudgeCCenter Judge
Scripts
The analysis is built as a series of single-purpose scripts that can be run individually or in sequence. Each script is described in its own section below.
