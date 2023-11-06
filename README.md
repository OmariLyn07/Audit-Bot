# Audit-Bot
This bot accepts inputs for the following:
  Excel file location via harddrive
  Excel sheet to be read
  Download folder location
  CUNYfirst Username i.e. *Username*@login.cuny.edu
  CUNYfirst password (hidden as asterisk)

Update changes 11/6/2023:
  Fixed a bug where the target audit file would not download;
    
    *possible cause: after opening the pdf version of degree audit the browser closes too quickly for the "save" function to be fully utilized.
                     this may only be the case for when there is only one student in the excel sheet, but could also be the cause of others missing.
    
    *Fix(Temporary): Added a 10 second timer after line opening the pdf audit page
  
  Fixed a bug where the audit page downloaded was not fully up to date;

    *possible cause: due to the printer icon being visible during process new initialization the bot would grab an old audit file rather than current.

    *Fix(Temporary): Added nested while loop after process new.click() to wait for the button to reappear before proceeding.
