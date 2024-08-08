v.1.0 initial commit. 

# Rotagen

Rota to determine who is going to be working on any given day in the coming months conforming to preset criteria.
Produces an excel spreadsheet comprising a single sheet with 4 headers: 'Date', 'Day', 'On Call' & 'Public Holiday'.

The date value is the date and will list every date in the UK calendar for the period requested. 
The day value is the correct day for the date value. 
The 'On Call' value is a person parameter which specifies the person to work on that day. The person parameters are user defined, comma seperated values. 
The 'Public Holiday' value will be populated with 'PH' for every UK public holiday on any given date. 

Script asks to confirm 
1. The person parameters.
2. The date that the current rota ends (this will determine the date that the new rota begins).
3. The person parameter that was last working on the last day of the current rota.
4. The date to run the new rota to.

Rules that will be applied:
1. UK based calendar.
2. The person parameter that is allocated to any given Friday also needs to be allocated in the Saturday and Sunday that immediately follow.
3. Each person needs to be allocated days to work as equally and fairly as is possible.
4. The Friday-Sunday blocks of working need to be spaced as far apart as is possible for each person to that each person does not work weekends in quick succession. 
