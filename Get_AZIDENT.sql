SELECT
    a.[BrugerNavn] AS [AZIdent],

    a.[KaldeNavn]  AS [Kaldenavn],

    a.[Tjsted_EAN],

    b.[Medarbejdernummer],

    b.[LosID],

    b.[Maxgrænse]

FROM [ORG].[adm].[Bruger_AD_PrimærKonto_Aktuel] a

JOIN [Opus].[brugerstyring].[BRS_GodkenderBeløb] b

    ON a.[BrugerNavn] = b.[Ident]

WHERE a.[Tjsted_EAN] IN (
    {{EAN_LIST}}
)
ORDER BY a.[KaldeNavn] DESC;