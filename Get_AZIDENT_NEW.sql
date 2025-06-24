SELECT
	a.[AZIdent],
    a.[Kaldenavn],
    a.[EANnedarvet],
    b.[Medarbejdernummer],
    b.[LosID],
    b.[Maxgrænse]
FROM [Opus].[intdebitor].[InterneDebitorer_BrugerInfo] a
JOIN [Opus].[brugerstyring].[BRS_GodkenderBeløb] b
    ON a.[AZIdent] = b.[Ident]
WHERE a.[EANnedarvet] LIKE 'EANVEJ' OR a.[EANnedarvet] LIKE 'EANNATUR'

