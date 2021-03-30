CREATE PROC [dbo].[AGING_REPORT]
AS
Begin 
SELECT OCRD."CardCode", OCRD."CardName",JDT1.CreatedBy [DocEntry],
JDT1."BaseRef" "Invoice No.",  JDT1."Ref2" "Customer Ref. No.", JDT1."TaxDate" "Posting Date",  JDT1."DueDate" "Due Date",  

SUM(CASE 
WHEN JDT1."DueDate" > DATEADD(dd, 0,GETDATE()) then JDT1."BalDueDeb" - JDT1."BalDueCred"  
ELSE 0 END) "Future", 

SUM(CASE 
WHEN JDT1."DueDate" <= DATEADD(dd,0,GETDATE()) AND JDT1."DueDate" >= DATEADD(dd, -30,GETDATE()) then JDT1."BalDueDeb" - JDT1."BalDueCred" 
ELSE 0 END) "0-30 Days", 

SUM(CASE 
WHEN JDT1."DueDate" <= DATEADD(dd, -31,GETDATE()) AND JDT1."DueDate" >= DATEADD(dd, -60,GETDATE()) then JDT1."BalDueDeb" - JDT1."BalDueCred" 
ELSE 0 END) "31-60 Days", 

SUM(CASE 
WHEN JDT1."DueDate" <= DATEADD(dd,-61,GETDATE()) AND JDT1."DueDate" >= DATEADD(dd, -90,GETDATE()) then JDT1."BalDueDeb" - JDT1."BalDueCred" 
ELSE 0 END) "61-90 Days",

SUM(CASE 
WHEN JDT1."DueDate" <= DATEADD(dd, -91,GETDATE()) AND JDT1."DueDate" >= DATEADD(dd, -120,GETDATE()) then JDT1."BalDueDeb" - JDT1."BalDueCred" 
ELSE 0 END) "91-120 Days",

SUM(CASE WHEN JDT1."DueDate" <= DATEADD(dd, -121,GETDATE()) then JDT1."BalDueDeb" - JDT1."BalDueCred" 
ELSE 0 END) "121+ Days", 

SUM(JDT1."BalDueDeb" - JDT1."BalDueCred") "Balance Due",
(SELECT TOP 1 U_AsColNotes FROM [@AW_B1CZHDR] Where U_DocEntry=JDT1.CreatedBy ORDER BY CAST(Code AS int) DESC) [Previous Collection Notes],
CAST('' AS NVARCHAR(254)) [As Collection Notes]

FROM OCRD
LEFT OUTER JOIN JDT1 ON OCRD."CardCode" = JDT1."ShortName"
 
WHERE  OCRD."CardType" = 'C' AND JDT1."BalDueDeb" - JDT1."BalDueCred" <> 0
GROUP BY GROUPING SETS((OCRD."CardCode", OCRD."CardName", JDT1."BaseRef",JDT1.CreatedBy,  JDT1."Ref2" , JDT1."TaxDate",  JDT1."DueDate"), (OCRD."CardCode", OCRD."CardName"),())
ORDER BY CASE WHEN OCRD."CardName" IS NULL THEN '' ELSE OCRD."CardName" END, CASE WHEN JDT1."BaseRef" IS NULL THEN '' ELSE JDT1."BaseRef" END
END