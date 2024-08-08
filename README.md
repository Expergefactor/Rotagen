v.1.0

# Rotagen 
![alt text](https://repository-images.githubusercontent.com/839893011/f2262a0f-482d-47b1-9bd3-2d48a9a5d91a)
### Summary...

Rotagen will create a rota to determine who is going to be working on any given day between specified dates conforming to preset criteria.

It produces a simple excel spreadsheet comprising a single sheet with 4 headers:
  * 'Date',
  * 'Day',
  * 'On Call' &
  * 'Public Holiday'.

The date value is the date and will list every date in the UK calendar for the period requested. 
The day value is the correct day for the date value. 
The 'On Call' value is a person parameter which specifies the person to work on that day. 
The person parameters are user defined, comma seperated values. 
The 'Public Holiday' value will be populated with 'PH' for every UK public holiday on any given date.

The script asks you to:

  1. Specify the person parameters i.e. 'John Doe, Jane Doe,' etc.
  2. The person parameter that was working on the last day of the current rota.
  3. The date that the current rota ends (this will determine the date that the new rota begins).
  4. The date that the new rota should end.
 
 Rules that will be applied:
   * UK based calendar.
   * Each person gets allocated days to work as equally and fairly as is possible.
   * The person parameter that is allocated to any given Friday also gets allocated the Saturday and Sunday that immediately follow.
   * The Friday-Sunday blocks of working are spaced as far apart as is possible for each person so that each person does not work weekends in quick succession.

### Dependancies...
  * pandas
  * holidays
  * openpyxl

         pip install pandas holidays openpyxl

### Instructions
  * Install Python
  * Install dependancies
  * Download rotagen.py to a local directory
  
        python rotagen.py
