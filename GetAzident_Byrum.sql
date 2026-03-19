SELECT

    a.[BrugerNavn],

    a.[KaldeNavn],

    a.[Tjsted_EAN],

    b.[Medarbejdernummer],

    b.[LosID],

    b.[Maxgrænse]

FROM [ORG].[adm].[Bruger_AD_Aktuel] a

JOIN [Opus].[brugerstyring].[BRS_GodkenderBeløb] b

    ON a.[BrugerNavn] = b.[Ident]

WHERE a.[Tjsted_EAN] = 5798005770329

;
