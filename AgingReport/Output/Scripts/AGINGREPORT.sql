CREATE Procedure "AGING_REPORT"
(in CardCode varchar(100))
AS
Begin 
SELECT OCRD."CardCode" AS "Customer Code", OCRD."CardName" AS "Customer Name",
CASE 
            WHEN JDT1."TransType" = 20 THEN 'PD' 
            WHEN JDT1."TransType" = 21 THEN 'PR' 
            WHEN JDT1."TransType" = 18 THEN 'PU' 
            WHEN JDT1."TransType" = 13 THEN 'IN' 
            WHEN JDT1."TransType" = 14 THEN 'CN' 
            WHEN JDT1."TransType" = 60 THEN 'SO' 
            WHEN JDT1."TransType" = 59 THEN 'SI' 
            WHEN JDT1."TransType" = 58 THEN 'ST' 
            WHEN JDT1."TransType" = 67 THEN 'IM' 
            WHEN JDT1."TransType" = 162 THEN 'MR' 
            WHEN JDT1."TransType" = 15 THEN 'DN' 
            WHEN JDT1."TransType" = 16 THEN 'RE' 
            WHEN JDT1."TransType" = 202 THEN 'PW' 
            WHEN JDT1."TransType" = 203 THEN 'DT' 
            WHEN JDT1."TransType" = 204 THEN 'DT' 
            WHEN JDT1."TransType" = 30 THEN 'JE' 
            WHEN JDT1."TransType" = 58 THEN 'ST' 
            WHEN JDT1."TransType" = 46 THEN 'PY' 
            WHEN JDT1."TransType" = 24 THEN 'RC' 
            WHEN JDT1."TransType" = 25 THEN 'DP' 
            WHEN JDT1."TransType" = 140000010 THEN 'IE' 
            WHEN JDT1."TransType" = 140000009 THEN 'OE' 
            WHEN JDT1."TransType" = 19 THEN 'PC' 
            WHEN JDT1."TransType" = 69 THEN 'IF' 
            WHEN JDT1."TransType" = 321 THEN 'JR' 
            WHEN JDT1."TransType" = 10000071 THEN 'ST' 
        END AS "Type",JDT1."TransType",
JDT1."CreatedBy" "DocEntry",JDT1."BaseRef" "Document No.",  JDT1."Ref2" "Customer Ref. No.", JDT1."TaxDate" "Posting Date",  JDT1."DueDate" "Due Date",  


SUM(CASE 
WHEN JDT1."DueDate" > ADD_DAYS(CURRENT_DATE, 0) then JDT1."BalDueDeb" - JDT1."BalDueCred"  
ELSE 0 END) "Future", 

SUM(CASE 
WHEN JDT1."DueDate" <= ADD_DAYS(CURRENT_DATE, 0) AND JDT1."DueDate" >= ADD_DAYS(CURRENT_DATE, -30) then JDT1."BalDueDeb" - JDT1."BalDueCred" 
ELSE 0 END) "0-30 Days", 

SUM(CASE 
WHEN JDT1."DueDate" <= ADD_DAYS(CURRENT_DATE, -31) AND JDT1."DueDate" >= ADD_DAYS(CURRENT_DATE, -60) then JDT1."BalDueDeb" - JDT1."BalDueCred" 
ELSE 0 END) "31-60 Days", 

SUM(CASE 
WHEN JDT1."DueDate" <= ADD_DAYS(CURRENT_DATE, -61) AND JDT1."DueDate" >= ADD_DAYS(CURRENT_DATE, -90) then JDT1."BalDueDeb" - JDT1."BalDueCred" 
ELSE 0 END) "61-90 Days", 

SUM(CASE 
WHEN JDT1."DueDate" <= ADD_DAYS(CURRENT_DATE, -91) AND JDT1."DueDate" >= ADD_DAYS(CURRENT_DATE, -120) then JDT1."BalDueDeb" - JDT1."BalDueCred" 
ELSE 0 END) "91-120 Days", 

SUM(CASE WHEN JDT1."DueDate" <= ADD_DAYS(CURRENT_DATE, -121) then JDT1."BalDueDeb" - JDT1."BalDueCred" 
ELSE 0 END) "121+ Days", 

SUM(JDT1."BalDueDeb" - JDT1."BalDueCred") "Balance Due"
,(SELECT MAX("DocEntry") FROM "@AW_B1CZHDR" Where "U_DocEntry"=JDT1."CreatedBy") "AWDocEntry"
,(SELECT "U_AsColNotes" FROM "@AW_B1CZDTL" Where "U_DocEntry"=JDT1."CreatedBy" 
AND "LineId"= (SELECT MAX("LineId") FROM "@AW_B1CZDTL" Where "U_DocEntry"=JDT1."CreatedBy")) "Previous Collection Notes"
,CAST('' AS TEXT) "Collection Notes"

FROM OCRD
LEFT OUTER JOIN JDT1 ON OCRD."CardCode" = JDT1."ShortName"
 
WHERE  OCRD."CardType" = 'C' AND JDT1."BalDueDeb" - JDT1."BalDueCred" <> 0 AND OCRD."CardCode" = IFNULL(:CardCode,OCRD."CardCode")
GROUP BY GROUPING SETS((OCRD."CardCode", OCRD."CardName",JDT1."TransType",JDT1."CreatedBy", JDT1."BaseRef",  JDT1."Ref2" , JDT1."TaxDate",  JDT1."DueDate"), (OCRD."CardCode", OCRD."CardName"),())
ORDER BY CASE WHEN OCRD."CardName" IS NULL THEN '' ELSE OCRD."CardName" END, CASE WHEN JDT1."BaseRef" IS NULL THEN '' ELSE JDT1."BaseRef" END;
END