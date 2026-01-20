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

WHERE a.[EANnedarvet] IN (

    '5798005770183',

    '5798005770190',

    '5798005770213',

    '5798005770220',

    '5798005770336',

    '5798005773597',

    '5798005774075',

    '5798005775706'

);