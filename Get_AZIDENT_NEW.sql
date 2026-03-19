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

WHERE a.[Tjsted_EAN] IN (

    '5798005770183',

    '5798005770190',

    '5798005770213',

    '5798005770220',

    '5798005770336',

    '5798005773597',

    '5798005774075',

    '5798005775706'

);

