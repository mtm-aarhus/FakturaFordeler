SELECT
	a.[AZIdent],
    a.[Kaldenavn],
    a.[EANnedarvet],
    b.[Medarbejdernummer],
    b.[LosID],
    b.[Maxgr�nse]
FROM [Opus].[intdebitor].[InterneDebitorer_BrugerInfo] a
JOIN [Opus].[brugerstyring].[BRS_GodkenderBel�b] b
    ON a.[AZIdent] = b.[Ident]
Where a.[EANnedarvet] LIKE 'EANNummer' 

