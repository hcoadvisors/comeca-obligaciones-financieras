SELECT * FROM (
	SELECT ROW_NUMBER() OVER (PARTITION BY T0."DocEntry" ORDER BY T0."LineId") As "#", 'Y' "Check", T1."DocNum", T1."DocEntry", T1."U_HCO_CardCode", T0."U_HCO_Date",
	T0."U_HCO_PayAmt", T0."LineId", T1."U_HCO_AcctCode", T1."U_HCO_Currency",
	(SELECT "MainCurncy" FROM OADM) "MainCurncy",
	T1."U_HCO_OcrCode", T1."U_HCO_AccBank"
	FROM "@HCO_POL1" T0 
	INNER JOIN "@HCO_OPOL" T1 ON T0."DocEntry" = T1."DocEntry"
	WHERE T0."U_HCO_Date" <= CURRENT_DATE
	AND IFNULL(T0."U_HCO_TransId", 0) = 0
	AND T1."U_HCO_CardCode" BETWEEN '{0}' AND '{1}'
	ORDER BY T0."DocEntry"
) S0 WHERE "#" = 1