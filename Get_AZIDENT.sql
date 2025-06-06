SELECT TOP (1000) [BrugerNøgle]
      ,[Ident]
      ,[Fornavn]
      ,[Efternavn]
      ,[SammenstilletNavn]
  FROM [Opus].[brugerstyring].[BRS_Brugere]
  where SammenstilletNavn like 'Medarbejder'
