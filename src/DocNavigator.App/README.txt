Build & Run:
  cd src/DocNavigator.App
  dotnet restore
  dotnet build
  dotnet run

Edit DB connection in: src/DocNavigator.App/Config/profiles/pg.json


--- Remote .desc integration ---
Profiles now support DescBaseUrl, DescUrlTemplate, DescVersion. MainViewModel fetches .desc over HTTP with session cache.
