SELECT * FROM (
	SELECT ROW_NUMBER() OVER (PARTITION BY T0."DocEntry" ORDER BY T0."LineId") As "#", 'Y' "Check", T1."DocNum", T1."DocEntry", T1."U_HCO_CardCode", T0."U_HCO_Date", CURRENT_DATE "FechaContabilizacion", T0."U_HCO_InitalAmt", T0."U_HCO_PayAmt", T0."U_HCO_Capita",
	T0."U_HCO_Interes", T0."LineId", T1."U_HCO_AcctCode",
	CASE T0."LineId" WHEN 1 THEN (T1."U_HCO_Amount" * "U_HCO_Commission" / 100) ELSE 0 END "Comision", 
	CASE T0."LineId" WHEN 1 THEN T1."U_HCO_Insuran" ELSE 0 END "U_HCO_Insuran", 
	CASE T0."LineId" WHEN 1 THEN T1."U_HCO_Other" ELSE 0 END "U_HCO_Other",
	0.0 "InteresMora", 0.0 "PorcIntMora", T1."U_HCO_OcrCode", T1."U_HCO_Currency",
	(SELECT "MainCurncy" FROM OADM) "MainCurncy", T1."U_HCO_AccBank",
	CAST('' AS NVARCHAR(250)) "Referencia",
	T1."U_HCO_MnthPay", T1."U_HCO_AccComm",
	T1."U_HCO_AccIns", T1."U_HCO_AccOthe"
	FROM "@HCO_OFE1" T0 
	INNER JOIN "@HCO_OOFE" T1 ON T0."DocEntry" = T1."DocEntry"
	WHERE T0."U_HCO_Date" <= CURRENT_DATE
	AND IFNULL(T0."U_HCO_TransId", 0) = 0 AND IFNULL(T0."U_HCO_RDREntry", 0) = 0
	AND T1."U_HCO_CardCode" BETWEEN '{0}' AND '{1}'
	ORDER BY T0."DocEntry"
) S0 WHERE "#" = 1