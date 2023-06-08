CREATE OR REPLACE PACKAGE cv_types AS
  
	TYPE WellData IS RECORD(  
		WellName		Varchar2(50),
		ResultsCount	Number
	);
	TYPE CV_WEllData IS REF CURSOR RETURN WellData;
    
End;
/


Create Or Replace Procedure WellCounting (    
	pWellName   OUT	VARCHAR2,
    pCount		OUT	NUMBER,
    rsWellData	IN OUT cv_types.CV_WEllData)

AS
	
BEGIN
	Open rsWellData For
		Select 
			Wells.WELLNAME,Count(RESULTS.WELLID) 
		Into 
			pWellName,
			pCount
		From 
			Wells, Results 
		Where 
			Wells.WellID = Results.WellID
		 group by 
		 	WEllName;
            
EXCEPTION 
  WHEN OTHERS THEN         
      ROLLBACK WORK;
      RAISE;

End WellCounting;
/


Create Or Replace Procedure OneWellCount (    
	pWellID		IN  Number,
	pWellName   OUT	VARCHAR2,
    pCount		OUT	NUMBER,
    rsWellData	IN OUT cv_types.CV_WEllData
    )

AS
BEGIN
	Open rsWellData For
		Select 
			Wells.WELLNAME,Count(RESULTS.WELLID) 
		Into 
			pWellName,
			pCount
		From 
			Wells, Results 
		Where 
			Wells.WellID = pWellID And 
			Wells.WellID = Results.WellID
		 group by 
		 	WEllName;
EXCEPTION 
  WHEN OTHERS THEN         
      ROLLBACK WORK;
      RAISE;

End OneWellCount;
/
