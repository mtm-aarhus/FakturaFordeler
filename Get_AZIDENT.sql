SELECT
    a.[Ident],
    a.[Fornavn],
    a.[Efternavn],
    a.[Displayname],
    a.[Mail],
    a.[Division],
    a.[Department],
    a.[PhysicaldeliveryofficeName],
    a.[ExtensionAttribute7],
    b.[Medarbejdernummer],
    b.[LosID],
    b.[Maxgr�nse]
FROM [FDW].[ad].[Brugeroplysninger_SenestInfo] a
JOIN [Opus].[brugerstyring].[BRS_GodkenderBel�b] b
    ON a.[Ident] = b.[Ident]
WHERE a.[Division] = 'Teknik og Milj�'
  AND a.[ExtensionAttribute7] LIKE '%11138%' or a.[ExtensionAttribute7] like '%11144%'

